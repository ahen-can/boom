"""Add two tabs to life_insurance_category_analysis_Q4FY26.xlsx without touching the existing ones:
  'Bharti portfolio' : Bharti's products by market category with FY26 new business premium (NBP),
                       share of Bharti's category NBP, flagship, notable features, observations, suggestions.
  'My view'          : input tab for the user (most talked-about companies and products).
Inputs: data.json (phase 1), li-product-docs (product_features_v1.csv, features/ulip/*.json,
        sources/bharti-kpi/product_performance_FY24_to_H2FY27.pdf numbers typed below).
Usage: python3 -I add_bharti_tabs.py <xlsx> <data.json> <li-product-docs dir>"""
import sys, csv, json, os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

XLSX, DATA, LPD = sys.argv[1:4]
wb = load_workbook(XLSX)
for t in ('Bharti portfolio', 'My view'):
    if t in wb.sheetnames:
        del wb[t]

HEAD = PatternFill('solid', fgColor='1F3864'); HF = Font(bold=True, color='FFFFFF')
SUB = PatternFill('solid', fgColor='D9E1F2'); EST = PatternFill('solid', fgColor='FFF2CC'); REAL = PatternFill('solid', fgColor='E2EFDA')
INP = PatternFill('solid', fgColor='FFFBEA')
WR = Alignment(wrap_text=True, vertical='top'); B = Font(bold=True)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)

# ---- Product Performance disclosure (Bharti KPI page), FY26 columns; NBP and premium in Rs crore ----
PP_URL = 'https://www.bhartiaxa.com/sites/default/files/kpi-disclosures/10-2026/Product%20Performance%20FY24%20-%20H2%20FY%2027.pdf'
PP = {  # uin: (FY26 new policies, FY26 NBP, FY26 renewal premium, renewal proportion, exit proportion, FY25 NBP, FY24 NBP)
    '130N107V13': (16545, 113.23, 416.60, 0.79, 0.04, 157.86, 164.89),   # Guaranteed Wealth Pro
    '130L100V03': (3459, 81.55, 160.72, 0.66, 0.04, 79.13, 101.89),      # Wealth Pro
    '130L123V01': (4034, 79.58, 37.99, 0.32, 0.01, 53.84, None),         # Growth Shield Plus
    '130L125V01': (3074, 50.91, 0.0, 0.0, 0.01, None, None),             # Dream Shield Plus
    '130N145V03': (1375, 50.82, 0.0, 0.0, 0.01, None, None),             # Insta Gain
    '130N144V01': (9689, 46.56, 0.0, 0.0, 0.02, None, None),             # Guaranteed Bachat
    '130N106V02': (5448, 43.85, 85.51, 0.66, 0.07, 57.26, 24.25),        # Unnati
    '130N109V04': (673, 38.21, 24.69, 0.39, 0.04, 20.14, 6.38),          # Swabhimaan Retirement
}
DISCLOSED = sum(v[1] for v in PP.values())          # products that make up ~90% of NBP
TOTAL_NBP_EST = DISCLOSED / 0.90
REMAINDER = TOTAL_NBP_EST - DISCLOSED

# ---- Bharti products and features ----
prods = [r for r in csv.DictReader(open(os.path.join(LPD, 'analysis/product_features_v1.csv'), encoding='utf-8')) if r['insurer'] == 'Bharti Life' and not r['market_category'].startswith('OUT')]
NOTES = {  # notable features, from deep data (ULIP) or the brochure (others); short and factual
    '130L100V03': 'Growth variant returns mortality and allocation charges at maturity; Legacy variant covers to age 99; unlimited free switches',
    '130L123V01': 'Returns 2x year 1-4 allocation charges and rising multiples (1x, 2x, 3x) of mortality charges from year 11',
    '130L125V01': 'Rising multiples of mortality charges returned from year 11; no admin charge in years 1-4',
    '130L128V01': 'Refund of allocation, admin and mortality charges; boosters for women, ECS payers and existing customers; index funds',
    '130L121V01': 'Refunds 118% of all charges (FMC included) after the premium term; best net yield in our calculator (6.59% at 8%)',
    '130L088V04': 'Joint-life option (single pay); no allocation charge',
    '130L049V05': 'No allocation charge; separate male and female mortality tables',
}
PRI = ['return_of_premium', 'whole_life_cover', 'life_stage_increase', 'critical_illness_inbuilt_or_option', 'terminal_illness', 'waiver_of_premium',
       'spouse_joint_cover', 'women_specific_rates', 'death_benefit_as_income', 'loyalty_additions', 'premium_offset_or_holiday', 'child_payor_benefit',
       'zero_cost_exit', 'pay_till_60', 'online_discount', 'increasing_cover', 'ppt_single_pay']
cat = {f['feature']: f['definition'] for f in json.load(open(os.path.join(LPD, 'analysis/feature_catalogue_v1.json')))}


def notable(r):
    if r['uin'] in NOTES:
        return NOTES[r['uin']], 'policy document (deep discovery)'
    hits = [cat.get(f, f) for f in PRI if r.get('f_' + f) == 'Y'][:4]
    return ('; '.join(hits) if hits else 'see feature list v1'), 'brochure keyword scan (v1)'

# ---- Market context from phase 1 (data.json values are Rs lakh) ----
d = json.load(open(DATA))
def seg(v, keys): return sum(v['fy26'][k] for k in keys) / 100
PDSEG = {'Par Life': ['PL', 'PP', 'PA'], 'Non-Par Life': ['NL', 'NA', 'NP'], 'Linked': ['LL', 'LP']}
ind = {k: sum(seg(v, ks) for v in d.values()) for k, ks in PDSEG.items()}
pvt = {k: sum(seg(v, ks) for n, v in d.items() if n != 'LIC') for k, ks in PDSEG.items()}
bh = {k: seg(d['Bharti Life'], ks) for k, ks in PDSEG.items()}
pa_keys = ['PP', 'PA', 'NP', 'LP']
bh_pa = seg(d['Bharti Life'], pa_keys); ind_pa = sum(seg(v, pa_keys) for v in d.values())

CATS = [
    ('ULIP', 'Linked', [
        'Bharti\'s three new-generation ULIPs (Wealth Pro, Growth Shield Plus, Dream Shield Plus) bring in Rs 212 cr of NBP, about the same as non-par savings (Rs 210 cr from its three disclosed products).',
        'Growth Shield Plus NBP rose from Rs 53.8 cr (FY25, launch year) to Rs 79.6 cr and nearly matches Wealth Pro. Its renewal proportion is low (32%) because the book is young.',
        'Feature list v2: Bharti leads on refunds of charges (4 of 7 products return mortality charges; Wealth Maximizer refunds 118% of all charges) and has unlimited free switches.',
        'Gaps against peers: no ULIP pension, no waiver of premium or child ULIP, no top-ups, no capital guarantee, no built-in accident cover.'],
     ['Launch a ULIP pension with an assured vesting option (HDFC Click 2 Retire Plus II and SBI Retire Smart Plus are the reference designs).',
      'Add a premium-waiver option or a child variant to a TULIP product (SBI Smart Scholar Plus has built-in payor waiver and accident cover).',
      'Allow top-ups in Wealth Pro, the largest ULIP by in-force policies.']),
    ('Non-Par Savings', 'Non-Par Life', [
        'Guaranteed Wealth Pro is the largest product in the company (Rs 113 cr NBP, Rs 530 cr total premium, 79% renewals), but its NBP fell from Rs 165 cr (FY24) and Rs 158 cr (FY25).',
        'Insta Gain and Guaranteed Bachat, both first disclosed in FY26, add Rs 97 cr of NBP between them.',
        'Exact IRR is possible here because Bharti publishes factor tables (Income Laabh: about 4.2-4.3%). SBI, HDFC and Tata flagships do not publish them.',
        'White space (v1): no whole-life guaranteed income product.'],
     ['Check Guaranteed Wealth Pro\'s IRR against peers (SBI Smart Platina Supreme, HDFC Sanchay Plus) once BIs are available. The NBP decline may be a pricing signal.',
      'Consider a whole-life guaranteed income variant, a common design in leaders\' portfolios.']),
    ('Par', 'Par Life', [
        'Unnati is the only par product in the 90% list (Rs 43.9 cr NBP, down from Rs 57.3 cr in FY25).',
        'Par is 19% of Bharti\'s total premium, down from 38% in FY21 (phase 1 trend).',
        'White space (v1): no par whole life and no par pension.'],
     ['Decide whether par is a growth line or a run-off line. If growth: a par whole-life plan with cash bonuses, the most common par design among leaders.']),
    ('Annuity', None, [
        'Swabhimaan Retirement (deferred pension) NBP rose from Rs 6.4 cr (FY24) to Rs 38.2 cr (FY26).',
        'The only immediate annuity is Saral Pension, the IRDAI standard product.',
        'Pension and annuity premium is about 2% of Bharti\'s book against about 18% for the industry.'],
     ['Add a non-Saral immediate annuity (joint life, return of purchase price, increasing annuity options) to sell alongside Swabhimaan at vesting.']),
    ('Protection', 'Non-Par Life', [
        'No protection product is in the 90% NBP list, so protection NBP is below about Rs 56 cr in total.',
        'Flexi Term Pro has a return-of-premium option as well as pure term; Suraksha is a savings-oriented protection plan.'],
     ['Premium quotes for a standard profile are needed before judging competitiveness (see the ULIP-style comparison planned for protection).']),
]

ws = wb.create_sheet('Bharti portfolio')
for col, w in zip('ABCDEFGHIJKL', [32, 13, 34, 13, 12, 11, 11, 11, 10, 10, 62, 34]):
    ws.column_dimensions[col].width = w
r = 1
ws.cell(r, 1, 'Bharti Life: product portfolio by category, FY26').font = Font(bold=True, size=14); r += 1
ws.cell(r, 1, ('New business premium (NBP) by product is from Bharti\'s official "Product Performance" disclosure (KPI disclosures page), which lists the products '
               f'making up 90% of NBP. Those 8 products total Rs {DISCLOSED:.1f} cr, so total NBP is about Rs {TOTAL_NBP_EST:.0f} cr and the other products share about '
               f'Rs {REMAINDER:.0f} cr. Green = disclosed figure. Yellow = estimate (the remainder split evenly); overwrite yellow cells if you get real numbers, and the shares update.')).alignment = WR
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 48; r += 1
ws.cell(r, 1, f'Source: {PP_URL}').font = Font(italic=True, size=9); r += 2

n_rest = sum(1 for p in prods if p['uin'] not in PP)
est_each = REMAINDER / n_rest
hdr = ['Product', 'UIN', 'Sub-segment', 'FY26 NBP (Rs cr)', 'Share of Bharti category NBP', 'FY25 NBP', 'FY24 NBP', 'FY26 new policies', 'Renewal proportion', 'Exit proportion', 'Notable features', 'Feature source']
summary_rows = []
for cname, pdcat, obs, sugg in CATS:
    ps = [p for p in prods if p['market_category'] == cname]
    ps.sort(key=lambda p: -(PP.get(p['uin'], (0, 0))[1]))
    c = ws.cell(r, 1, cname); c.font = Font(bold=True, size=12, color='1F3864'); r += 1
    if pdcat:
        ctx = f'PD category {pdcat}: Bharti total premium FY26 Rs {bh[pdcat]:,.0f} cr = {bh[pdcat]/ind[pdcat]:.2%} of industry, {bh[pdcat]/pvt[pdcat]:.2%} of private insurers.'
    else:
        ctx = f'Pension and annuity lines: Bharti total premium FY26 Rs {bh_pa:,.0f} cr = {bh_pa/ind_pa:.2%} of industry.'
    if cname == 'Protection':
        ctx += ' Protection sits inside the Non-Par Life PD line with non-par savings.'
    ws.cell(r, 1, ctx).font = Font(italic=True); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); r += 1
    for j, h in enumerate(hdr, 1):
        x = ws.cell(r, j, h); x.fill = HEAD; x.font = HF; x.alignment = WR
    ws.row_dimensions[r].height = 30
    r += 1; first = r
    for p in ps:
        v = PP.get(p['uin'])
        note, src = notable(p)
        name = p['product_name'].replace('Bharti AXA Life ', '')
        if p['standard_or_pos_tag']: name += f" ({p['standard_or_pos_tag']})"
        vals = [name, p['uin'], p['sub_segment_v1'], v[1] if v else round(est_each, 1), None,
                v[5] if v else None, v[6] if v else None, v[0] if v else None, v[3] if v else None, v[4] if v else None, note, src]
        for j, val in enumerate(vals, 1):
            x = ws.cell(r, j, val); x.alignment = WR; x.border = BOX
        ws.cell(r, 4).fill = REAL if v else EST
        for j in (6, 7, 8):
            if v: ws.cell(r, j).fill = REAL
        for j in (9, 10, 5):
            ws.cell(r, j).number_format = '0%'
        r += 1
    last = r - 1
    for rr in range(first, last + 1):
        ws.cell(rr, 5, f'=D{rr}/SUM(D${first}:D${last})')
    ws.cell(r, 1, 'Category total').font = B
    ws.cell(r, 4, f'=SUM(D{first}:D{last})').font = B
    flag = max(ps, key=lambda p: PP.get(p['uin'], (0, 0))[1])
    ws.cell(r, 11, f"Flagship by NBP: {flag['product_name'].replace('Bharti AXA Life ', '')}" if flag['uin'] in PP else 'Flagship: no product in the disclosed 90% list').font = B
    summary_rows.append((cname, len(ps), f'D{r}', flag['product_name'].replace('Bharti AXA Life ', '') if flag['uin'] in PP else 'none disclosed'))
    r += 1
    ws.cell(r, 1, 'Observations').font = B; r += 1
    for o in obs:
        ws.cell(r, 1, '• ' + o).alignment = WR; ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 30; r += 1
    ws.cell(r, 1, 'Suggestions for discussion').font = B; r += 1
    for s in sugg:
        x = ws.cell(r, 1, '• ' + s); x.alignment = WR; x.fill = SUB
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 30; r += 1
    r += 1

ws.cell(r, 1, 'Summary').font = Font(bold=True, size=12, color='1F3864'); r += 1
for j, h in enumerate(['Category', 'Products', 'FY26 NBP (Rs cr)', 'Share of Bharti NBP', 'Flagship by NBP'], 1):
    x = ws.cell(r, j, h); x.fill = HEAD; x.font = HF
r += 1; s0 = r
for cname, n, ref, flag in summary_rows:
    ws.cell(r, 1, cname); ws.cell(r, 2, n); ws.cell(r, 3, f'={ref}'); ws.cell(r, 4, f'=C{r}/SUM(C${s0}:C${s0 + len(summary_rows) - 1})').number_format = '0%'; ws.cell(r, 5, flag); r += 1
ws.cell(r, 1, 'Total').font = B; ws.cell(r, 3, f'=SUM(C{s0}:C{r - 1})').font = B
r += 2
ws.cell(r, 1, ('Notes: Micro product Grameen Jeevan Bima is out of scope. Categories come from each product\'s IRDAI description line (feature list v1). '
               'APE is not disclosed by product; NBP is the closest official measure. Renewal proportion = renewal premium / total premium for the year; '
               'exit proportion = maturities, surrenders, withdrawals and cancellations / policies at the start plus new.')).alignment = WR
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12); ws.row_dimensions[r].height = 45
ws.freeze_panes = 'B5'

# ---- My view (user input) ----
mv = wb.create_sheet('My view')
for col, w in zip('ABCDEFGHI', [24, 18, 16, 40, 22, 12, 12, 12, 30]):
    mv.column_dimensions[col].width = w
r = 1
mv.cell(r, 1, 'My view: what I hear is popular or important').font = Font(bold=True, size=14); r += 1
mv.cell(r, 1, ('Fill the yellow cells. These views decide which 5 products per category go in the deck and which 10-15 go in the supporting Excel. '
               'Market share columns are filled from the public disclosures for reference.')).alignment = WR
mv.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); mv.row_dimensions[r].height = 32; r += 2
mv.cell(r, 1, 'Part 1. Companies').font = Font(bold=True, size=12, color='1F3864'); r += 1
h1 = ['Insurer', 'Category', 'Category share FY26', 'Why people talk about it', 'Where I heard it', 'My rating (1-5)']
for j, h in enumerate(h1, 1):
    x = mv.cell(r, j, h); x.fill = HEAD; x.font = HF; x.alignment = WR
r += 1
lk = ['LL', 'LP']
tot_l = sum(seg(v, lk) for v in d.values())
order = sorted(d, key=lambda n: -seg(d[n], lk))
for n in order[:16]:
    mv.cell(r, 1, n); mv.cell(r, 2, 'ULIP'); mv.cell(r, 3, seg(d[n], lk) / tot_l).number_format = '0.0%'
    for j in (4, 5, 6): mv.cell(r, j).fill = INP
    r += 1
for _ in range(8):
    for j in range(1, 7): mv.cell(r, j).fill = INP
    r += 1
r += 1
mv.cell(r, 1, 'Part 2. Products').font = Font(bold=True, size=12, color='1F3864'); r += 1
h2 = ['Insurer', 'Product', 'UIN', 'What makes it notable', 'Popularity signal (sales, ads, agents, press)', 'Category', 'In deck (5)? Y/N', 'In Excel (10-15)? Y/N', 'Comment']
for j, h in enumerate(h2, 1):
    x = mv.cell(r, j, h); x.fill = HEAD; x.font = HF; x.alignment = WR
mv.row_dimensions[r].height = 30
r += 1
for _ in range(30):
    for j in range(1, 10): mv.cell(r, j).fill = INP
    r += 1
r += 1
mv.cell(r, 1, 'Part 3. Notes and hypotheses').font = Font(bold=True, size=12, color='1F3864'); r += 1
for _ in range(10):
    mv.cell(r, 1).fill = INP; mv.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9); r += 1

wb.save(XLSX)
print('disclosed NBP', round(DISCLOSED, 2), 'est total', round(TOTAL_NBP_EST, 1), 'remainder', round(REMAINDER, 1), 'per other product', round(est_each, 2), 'n other', n_rest)
print([(c, n, f) for c, n, _, f in summary_rows])
