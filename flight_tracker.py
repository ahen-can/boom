"""
Flight price tracker: Goa -> Mumbai, with price-drop alerts.

Replicates the search behind this Skyscanner page:
  https://www.skyscanner.co.in/transport/flights/igoi/ibom/260830/
      ?adultsv2=1&cabinclass=economy&rtn=0&preferdirects=true

...but instead of scraping Skyscanner (which is JS-rendered, bot-blocked, and
against their ToS), it queries a real flight-search API for the same route and
filters, records the price on every run, and alerts you when the price drops.

Search parameters (edit in the CONFIG section below):
  * Origin airports:      GOI (Dabolim) or GOX (Mopa/Manohar Intl), Goa
  * Destination airports: BOM (Chhatrapati Shivaji) or NMI (Navi Mumbai Intl)
  * Date:                 2026-08-30  (the 260830 in the URL)
  * Passengers:           1 adult, economy, one-way
  * Non-stop only:        yes (preferdirects=true)
  * Departure window:     14:00 <= departure < 22:00 (local time)

Data source:
  * "mock"    -> no keys needed, returns synthetic prices. Great for testing
                 the pipeline and the alerting end to end.
  * "amadeus" -> real prices via the Amadeus Self-Service API (free test tier).
                 Set AMADEUS_CLIENT_ID / AMADEUS_CLIENT_SECRET and
                 FLIGHT_PROVIDER=amadeus.

Usage:
    python flight_tracker.py            # one run: fetch, filter, record, alert
    FLIGHT_PROVIDER=amadeus python flight_tracker.py
"""

from __future__ import annotations

import json
import os
import random
import smtplib
import sys
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from email.mime.text import MIMEText
from itertools import product
from pathlib import Path
from typing import Optional

import requests

# --------------------------------------------------------------------------- #
# CONFIG  -- everything you'd normally tweak lives here.
# --------------------------------------------------------------------------- #

ORIGINS = ["GOI", "GOX"]            # Goa: Dabolim, Mopa (Manohar Intl)
DESTINATIONS = ["BOM", "NMI"]       # Mumbai: Chhatrapati Shivaji, Navi Mumbai
DEPARTURE_DATE = "2026-08-30"       # from the "260830" in the Skyscanner URL
ADULTS = 1
CABIN_CLASS = "ECONOMY"
NON_STOP_ONLY = True                # preferdirects=true
CURRENCY = "INR"

# Departure-time window (local, 24h). Evening flights only: 2pm–10pm.
DEPART_AFTER_HOUR = 14              # inclusive, 14:00
DEPART_BEFORE_HOUR = 22            # exclusive, 22:00

# Alert when the best fare is at/under this many INR, OR drops vs. last run.
PRICE_THRESHOLD_INR: Optional[float] = None   # e.g. 3000 to only alert under ₹3000
ALERT_ON_ANY_DROP = True                      # alert whenever price falls run-to-run

HISTORY_FILE = Path(__file__).with_name("price_history.json")

# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #


@dataclass
class Offer:
    origin: str
    destination: str
    depart_at: str          # ISO local datetime, e.g. "2026-08-30T18:45:00"
    arrive_at: str
    carrier: str
    price: float
    currency: str

    @property
    def depart_hour(self) -> int:
        return datetime.fromisoformat(self.depart_at).hour


# --------------------------------------------------------------------------- #
# Providers  -- each returns a list[Offer] for one origin/destination pair.
# --------------------------------------------------------------------------- #


class MockProvider:
    """Deterministic-ish synthetic prices so the whole pipeline is testable
    with no API keys. Prices wobble run-to-run so drops actually happen."""

    CARRIERS = ["6E", "AI", "UK", "SG"]  # IndiGo, Air India, Vistara, SpiceJet

    def search(self, origin: str, destination: str) -> list[Offer]:
        rng = random.Random(f"{origin}{destination}{datetime.now().minute}")
        offers: list[Offer] = []
        for hour in (7, 15, 18, 21, 23):          # a spread across the day
            price = round(rng.uniform(2200, 6500), 2)
            offers.append(
                Offer(
                    origin=origin,
                    destination=destination,
                    depart_at=f"{DEPARTURE_DATE}T{hour:02d}:{rng.choice([5, 20, 45]):02d}:00",
                    arrive_at=f"{DEPARTURE_DATE}T{hour:02d}:55:00",
                    carrier=rng.choice(self.CARRIERS),
                    price=price,
                    currency=CURRENCY,
                )
            )
        return offers


class AmadeusProvider:
    """Real fares via the Amadeus Self-Service Flight Offers Search API.

    Free test credentials: https://developers.amadeus.com/  ->  create an app.
    Set AMADEUS_CLIENT_ID and AMADEUS_CLIENT_SECRET. The test environment uses
    the test.api.amadeus.com host; flip AMADEUS_ENV=production for live data
    (requires a production app + billing on their side).
    """

    def __init__(self) -> None:
        self.client_id = os.environ["AMADEUS_CLIENT_ID"]
        self.client_secret = os.environ["AMADEUS_CLIENT_SECRET"]
        env = os.getenv("AMADEUS_ENV", "test")
        self.host = (
            "https://api.amadeus.com"
            if env == "production"
            else "https://test.api.amadeus.com"
        )
        self._token: Optional[str] = None

    def _access_token(self) -> str:
        if self._token:
            return self._token
        resp = requests.post(
            f"{self.host}/v1/security/oauth2/token",
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
            },
            timeout=30,
        )
        resp.raise_for_status()
        self._token = resp.json()["access_token"]
        return self._token

    def search(self, origin: str, destination: str) -> list[Offer]:
        resp = requests.get(
            f"{self.host}/v2/shopping/flight-offers",
            headers={"Authorization": f"Bearer {self._access_token()}"},
            params={
                "originLocationCode": origin,
                "destinationLocationCode": destination,
                "departureDate": DEPARTURE_DATE,
                "adults": ADULTS,
                "travelClass": CABIN_CLASS,
                "nonStop": "true" if NON_STOP_ONLY else "false",
                "currencyCode": CURRENCY,
                "max": 50,
            },
            timeout=45,
        )
        # A 400 for an unserved route/date is expected for some pairs (e.g. a
        # brand-new airport not yet in inventory) -- skip it rather than crash.
        if resp.status_code == 400:
            return []
        resp.raise_for_status()

        offers: list[Offer] = []
        for item in resp.json().get("data", []):
            seg = item["itineraries"][0]["segments"]
            offers.append(
                Offer(
                    origin=origin,
                    destination=destination,
                    depart_at=seg[0]["departure"]["at"],
                    arrive_at=seg[-1]["arrival"]["at"],
                    carrier=seg[0]["carrierCode"],
                    price=float(item["price"]["grandTotal"]),
                    currency=item["price"].get("currency", CURRENCY),
                )
            )
        return offers


def get_provider():
    name = os.getenv("FLIGHT_PROVIDER", "mock").lower()
    if name == "amadeus":
        return AmadeusProvider()
    if name == "mock":
        return MockProvider()
    raise SystemExit(f"Unknown FLIGHT_PROVIDER: {name!r} (use 'mock' or 'amadeus')")


# --------------------------------------------------------------------------- #
# Core logic
# --------------------------------------------------------------------------- #


def in_evening_window(offer: Offer) -> bool:
    return DEPART_AFTER_HOUR <= offer.depart_hour < DEPART_BEFORE_HOUR


def collect_offers(provider) -> list[Offer]:
    """Query every origin x destination pair, keep only offers that match the
    evening departure window."""
    matching: list[Offer] = []
    for origin, dest in product(ORIGINS, DESTINATIONS):
        try:
            found = provider.search(origin, dest)
        except requests.RequestException as exc:
            print(f"  ! {origin}->{dest}: request failed: {exc}", file=sys.stderr)
            continue
        matching.extend(o for o in found if in_evening_window(o))
    return sorted(matching, key=lambda o: o.price)


def load_history() -> list[dict]:
    if HISTORY_FILE.exists():
        return json.loads(HISTORY_FILE.read_text())
    return []


def previous_best(history: list[dict]) -> Optional[float]:
    for record in reversed(history):
        if record.get("best_price") is not None:
            return record["best_price"]
    return None


def save_run(history: list[dict], best: Optional[Offer], all_offers: list[Offer]) -> None:
    history.append(
        {
            "checked_at": datetime.now(timezone.utc).isoformat(),
            "best_price": best.price if best else None,
            "currency": best.currency if best else CURRENCY,
            "best_offer": asdict(best) if best else None,
            "num_matching_offers": len(all_offers),
        }
    )
    HISTORY_FILE.write_text(json.dumps(history, indent=2))


# --------------------------------------------------------------------------- #
# Notifications
# --------------------------------------------------------------------------- #


def send_email(subject: str, body: str) -> None:
    """Optional email alert. Set SMTP_HOST, SMTP_USER, SMTP_PASSWORD, ALERT_TO
    (and optionally SMTP_PORT). For Gmail, use an App Password."""
    host = os.getenv("SMTP_HOST")
    to_addr = os.getenv("ALERT_TO")
    if not (host and to_addr):
        return
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = os.getenv("SMTP_USER", to_addr)
    msg["To"] = to_addr
    with smtplib.SMTP(host, int(os.getenv("SMTP_PORT", "587"))) as server:
        server.starttls()
        if os.getenv("SMTP_USER"):
            server.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        server.send_message(msg)
    print("  -> email alert sent")


def send_webhook(text: str) -> None:
    """Optional Slack/Discord-style webhook alert. Set ALERT_WEBHOOK_URL."""
    url = os.getenv("ALERT_WEBHOOK_URL")
    if not url:
        return
    try:
        requests.post(url, json={"text": text, "content": text}, timeout=15)
        print("  -> webhook alert sent")
    except requests.RequestException as exc:
        print(f"  ! webhook failed: {exc}", file=sys.stderr)


def notify(best: Offer, prev: Optional[float], reason: str) -> None:
    subject = f"✈️ Goa→Mumbai fare {reason}: {best.currency} {best.price:,.0f}"
    body = (
        f"{subject}\n\n"
        f"Route:      {best.origin} -> {best.destination}\n"
        f"Departs:    {best.depart_at}  (carrier {best.carrier})\n"
        f"Date:       {DEPARTURE_DATE}, evening window "
        f"{DEPART_AFTER_HOUR}:00-{DEPART_BEFORE_HOUR}:00\n"
        f"Price:      {best.currency} {best.price:,.2f}\n"
    )
    if prev is not None:
        body += f"Previous:   {best.currency} {prev:,.2f}\n"
    print("\n" + body)
    send_email(subject, body)
    send_webhook(body)


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #


def main() -> int:
    provider = get_provider()
    print(
        f"[{datetime.now():%Y-%m-%d %H:%M}] "
        f"Checking {'/'.join(ORIGINS)} -> {'/'.join(DESTINATIONS)} "
        f"on {DEPARTURE_DATE} "
        f"(evening {DEPART_AFTER_HOUR}:00-{DEPART_BEFORE_HOUR}:00) "
        f"via {provider.__class__.__name__}"
    )

    offers = collect_offers(provider)
    if not offers:
        print("No non-stop offers matched the evening window this run.")
        save_run(load_history(), None, offers)
        return 0

    best = offers[0]
    print(f"Best matching fare: {best.currency} {best.price:,.2f} "
          f"({best.origin}->{best.destination}, departs {best.depart_at})")

    history = load_history()
    prev = previous_best(history)

    dropped = ALERT_ON_ANY_DROP and prev is not None and best.price < prev
    under_threshold = (
        PRICE_THRESHOLD_INR is not None and best.price <= PRICE_THRESHOLD_INR
    )
    if dropped:
        notify(best, prev, "dropped")
    elif under_threshold:
        notify(best, prev, "under threshold")
    else:
        msg = "no drop" if prev is not None else "first recorded price"
        print(f"  ({msg} — no alert)")

    save_run(history, best, offers)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
