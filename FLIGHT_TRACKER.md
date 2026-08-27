# Flight price tracker — Goa → Mumbai

Tracks the fare behind this Skyscanner search and alerts you on a **price drop**:

> https://www.skyscanner.co.in/transport/flights/igoi/ibom/260830/?adultsv2=1&cabinclass=economy&rtn=0&preferdirects=true

**Filters applied:**

| Setting | Value |
|---|---|
| Origin | **GOI** (Dabolim) or **GOX** (Mopa / Manohar Intl), Goa |
| Destination | **BOM** (Chhatrapati Shivaji) or **NMI** (Navi Mumbai Intl) |
| Date | 2026-08-30 (the `260830` in the URL) |
| Passengers | 1 adult, economy, one-way |
| Stops | Non-stop only (`preferdirects=true`) |
| Departure window | **14:00 ≤ departure < 22:00** (evening, local time) |

## Why not just scrape the Skyscanner page?

Skyscanner renders prices with JavaScript and actively blocks bots (Cloudflare,
fingerprinting, CAPTCHAs); scraping also violates their ToS. A scraper would be
fragile and get IP-banned. Instead this queries a **flight-search API** for the
same route + filters — stable and permitted.

## Quick start (no API keys)

```bash
pip install requests
python flight_tracker.py
```

Runs in **mock mode** with synthetic prices so you can see the full pipeline —
search → evening-window filter → record → drop detection → alert — end to end.
Each run appends to `price_history.json`.

## Real prices (Amadeus, free tier)

1. Create a free app at https://developers.amadeus.com/ → copy the API key/secret.
2. Set env vars and run:

   ```bash
   export FLIGHT_PROVIDER=amadeus
   export AMADEUS_CLIENT_ID=your_key
   export AMADEUS_CLIENT_SECRET=your_secret
   python flight_tracker.py
   ```

The free **test** environment has limited/simulated inventory (especially for a
brand-new airport like NMI). For full live fares, use a production Amadeus app
(`AMADEUS_ENV=production`) or swap in another provider — the `Provider` classes
in `flight_tracker.py` are a small, pluggable interface (`search(origin, dest)`
→ `list[Offer]`). Other options you can drop in the same way: Kiwi/Tequila,
Google Flights via SerpApi, or a RapidAPI Skyscanner endpoint.

## Getting alerts

Alerts fire when the best evening fare **drops vs. the last run**
(`ALERT_ON_ANY_DROP`), or falls at/under `PRICE_THRESHOLD_INR` (set it near the
top of `flight_tracker.py`). Delivery channels — all optional, set via env:

- **Email:** `SMTP_HOST`, `SMTP_USER`, `SMTP_PASSWORD`, `ALERT_TO`
  (Gmail: use an [App Password](https://myaccount.google.com/apppasswords)).
- **Webhook** (Slack/Discord): `ALERT_WEBHOOK_URL`.
- The console always prints the result.

## Running it on a schedule (no server needed)

`.github/workflows/flight-price-tracker.yml` runs the tracker **every 6 hours**
via GitHub Actions and commits the updated `price_history.json` back to the repo,
giving you a price log over time. Add your keys as repository **Secrets**
(Settings → Secrets and variables → Actions): `FLIGHT_PROVIDER`,
`AMADEUS_CLIENT_ID`, `AMADEUS_CLIENT_SECRET`, and any alert secrets above.

Prefer your own machine? Use cron instead:

```cron
0 */6 * * *  cd /path/to/boom && FLIGHT_PROVIDER=amadeus python flight_tracker.py >> tracker.log 2>&1
```

## Tuning

Everything you'd change lives in the `CONFIG` block at the top of
`flight_tracker.py`: airports, date, the evening window hours, currency,
non-stop toggle, alert threshold. When your travel date changes, update
`DEPARTURE_DATE`.
