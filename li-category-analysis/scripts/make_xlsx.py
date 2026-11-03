"""Build the product-category workbook from data.json (values in Rs lakh)."""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

D = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
NOTES = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else {}

SEGS = ['PL', 'PP', 'PA', 'NL', 'NA', 'NP', 'LL', 'LP', 'H', 'V', 'GT']
SEGNAME = {'PL': 'Par Life', 'PP': 'Par Pension', 'PA': 'Par Annuity', 'NL': 'Non-Par Life', 'NA': 'Non-Par Annuity',
           'NP': 'Non-Par Pension', 'LL': 'Linked Life', 'LP': 'Linked Pension', 'H': 'Health (excl.)',
           'V': 'Variable (excl.)', 'GT': 'Grand Total (L-1)'}
CATS = [  # (tab, segments, definition)
    ('Par Life', ['PL'], 'Participating life (non-linked). Excludes par pension/annuity (in Pension & Annuity).'),
    ('Non-Par Life', ['NL'], 'Non-participating life (non-linked): term/protection + guaranteed savings + group term/credit life. Excludes non-par pension/annuity, health, variable.'),
    ('Linked', ['LL'], 'Unit-linked life (ULIPs, incl. group linked life). Excludes linked pension (in Pension & Annuity), linked health & variable.'),
    ('Pension & Annuity', ['PP', 'PA', 'NA', 'NP', 'LP'], 'Par pension + par annuity + non-par pension + non-par annuity + linked pension (incl. group pension/annuity/fund business as reported in L-1).'),
]
PRIVATE_SECTOR = lambda n: 'Public' if n.startswith('LIC') else 'Private'

# ---------- styles
NAVY = '1F3864'; BH = 'FFF2CC'; GREY = 'F2F2F2'
hfont = Font(bold=True, color='FFFFFF'); hfill = PatternFill('solid', fgColor=NAVY)
bfill = PatternFill('solid', fgColor=BH); gfill = PatternFill('solid', fgColor=GREY)
thin = Side(style='thin', color='BFBFBF'); box = Border(top=thin, bottom=thin, left=thin, right=thin)
CR = '#,##0;(#,##0);"-"'; PCT = '0.0%;(0.0%);"-"'

def hdr(ws, r, vals, widths=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.font = hfont; c.fill = hfill; c.border = box
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.row_dimensions[r].height = 32
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[CL(i)].width = w

def title(ws, t, sub):
    ws['A1'] = t; ws['A1'].font = Font(bold=True, size=14, color=NAVY)
    ws['A2'] = sub; ws['A2'].font = Font(italic=True, size=9, color='595959')

cr = lambda lakh: round(lakh / 100.0, 2)
names = sorted(D, key=lambda n: -D[n]['fy26']['GT'])
wb = Workbook()

# ---------- Data sheet (single source; every other tab references it)
wsD = wb.active; wsD.title = 'Data'
title(wsD, 'Segment premium data — FY26 & FY25 (₹ crore)', 'Source: Form L-1-A-RA line (a) Premium (gross), each insurer\'s Q4 FY26 public disclosure; group from L-25(ii) YTD grand total. Rs lakh ÷ 100.')
cols = ['Insurer', 'Sector'] + [f'FY26 {SEGNAME[s]}' for s in SEGS] + [f'FY25 {SEGNAME[s]}' for s in SEGS] + ['FY26 Group premium (L-25(ii))']
hdr(wsD, 4, cols, [26, 9] + [12] * (len(cols) - 2))
ROW = {}
for i, n in enumerate(names):
    r = 5 + i; ROW[n] = r; o = D[n]
    wsD.cell(r, 1, n); wsD.cell(r, 2, PRIVATE_SECTOR(n))
    for j, s in enumerate(SEGS):
        wsD.cell(r, 3 + j, cr(o['fy26'][s])).number_format = CR
        if o.get('fy25'): wsD.cell(r, 3 + len(SEGS) + j, cr(o['fy25'][s])).number_format = CR
    g = o.get('group_fy26')
    if g is not None: wsD.cell(r, 3 + 2 * len(SEGS), cr(g)).number_format = CR
    if n.startswith('Bharti'):
        for c in range(1, len(cols) + 1): wsD.cell(r, c).fill = bfill
LAST = 5 + len(names) - 1
wsD.freeze_panes = 'C5'
COL26 = {s: CL(3 + j) for j, s in enumerate(SEGS)}
COL25 = {s: CL(3 + len(SEGS) + j) for j, s in enumerate(SEGS)}
COLG = CL(3 + 2 * len(SEGS))

def f_sum(segs, n, yr=26):
    col = COL26 if yr == 26 else COL25
    return '+'.join(f"Data!{col[s]}{ROW[n]}" for s in segs)

def val(segs, n, yr=26):
    o = D[n].get('fy26' if yr == 26 else 'fy25')
    return None if not o else sum(o[s] for s in segs) / 100.0

ALLC = [s for _, segs, _ in CATS for s in segs]
ind_tot = sum(val(ALLC, n) for n in names)

# ---------- Summary
wsS = wb.create_sheet('Summary', 0)
title(wsS, 'Life insurance premium by product category — FY26 (₹ crore)',
      'Gross premium (first year + renewal + single) from L-1-A-RA. Category total excludes health & variable. Group business sits inside each product segment as reported (see Group tab for the overlay).')
H = ['Insurer', 'Sector'] + [c[0] for c in CATS] + ['Total (4 categories)', 'Health + Variable (excl.)', 'Grand total (L-1)'] + [f'{c[0]} % of book' for c in CATS] + ['Group % of grand total', 'Share of industry (4 cat.)']
hdr(wsS, 4, H, [26, 9, 12, 12, 12, 12, 13, 12, 12, 10, 10, 10, 10, 10, 11])
r0 = 5
for i, n in enumerate(names):
    r = r0 + i
    wsS.cell(r, 1, n); wsS.cell(r, 2, PRIVATE_SECTOR(n))
    for j, (_, segs, _) in enumerate(CATS):
        wsS.cell(r, 3 + j, '=' + f_sum(segs, n)).number_format = CR
    wsS.cell(r, 7, f'=SUM(C{r}:F{r})').number_format = CR
    wsS.cell(r, 8, '=' + f_sum(['H', 'V'], n)).number_format = CR
    wsS.cell(r, 9, f'=Data!{COL26["GT"]}{ROW[n]}').number_format = CR
    for j in range(4):
        wsS.cell(r, 10 + j, f'=IF($G{r}=0,0,{CL(3 + j)}{r}/$G{r})').number_format = PCT
    wsS.cell(r, 14, f'=IF(I{r}=0,0,Data!{COLG}{ROW[n]}/I{r})').number_format = PCT
    wsS.cell(r, 15, f'=G{r}/G${r0 + len(names)}').number_format = PCT
    if n.startswith('Bharti'):
        for c in range(1, 16): wsS.cell(r, c).fill = bfill
rt = r0 + len(names)
wsS.cell(rt, 1, 'Industry total (covered insurers)').font = Font(bold=True)
for c in range(3, 10):
    wsS.cell(rt, c, f'=SUM({CL(c)}{r0}:{CL(c)}{rt - 1})').number_format = CR
for j in range(4):
    wsS.cell(rt, 10 + j, f'={CL(3 + j)}{rt}/$G{rt}').number_format = PCT
wsS.cell(rt, 15, f'=G{rt}/G{rt}').number_format = PCT
for c in range(1, 16): wsS.cell(rt, c).font = Font(bold=True); wsS.cell(rt, c).fill = gfill
wsS.freeze_panes = 'C5'

# summary observations
obs = []
for tab, segs, _ in CATS:
    t = sum(val(segs, n) for n in names)
    obs.append(f'{tab}: ₹{t:,.0f} cr = {t / ind_tot:.1%} of industry premium (excl. health & variable).')
b = next(n for n in names if n.startswith('Bharti'))
bt = val(ALLC, b)
mix = ', '.join(f'{tab} {val(segs, b) / bt:.0%}' for tab, segs, _ in CATS)
imix = ', '.join(f'{tab} {sum(val(segs, n) for n in names) / ind_tot:.0%}' for tab, segs, _ in CATS)
rk = sorted(names, key=lambda n: -val(ALLC, n)).index(b) + 1
obs.append(f'Bharti Life: ₹{bt:,.0f} cr, rank {rk} of {len(names)} by total premium, {bt / ind_tot:.2%} industry share. Book mix — {mix} vs industry {imix}.')
obs += NOTES.get('Summary', [])
r = rt + 2
wsS.cell(r, 1, 'Key observations').font = Font(bold=True, size=12, color=NAVY)
for k, o in enumerate(obs, 1):
    c = wsS.cell(r + k, 1, f'{k}. {o}'); c.alignment = Alignment(wrap_text=False)

# ---------- category tabs
def cat_tab(tab, segs, definition):
    ws = wb.create_sheet(tab)
    title(ws, f'{tab} — FY26 premium ranking (₹ crore)', definition)
    hdr(ws, 4, ['Rank', 'Insurer', 'Sector', 'FY26 premium', 'Share of category', 'Cumulative share', 'FY25 premium', 'YoY growth',
                f'{tab} as % of insurer\'s book (4 cat.)', 'Share of private sector', 'Private rank'], [6, 26, 9, 13, 11, 11, 13, 10, 14, 11, 8])
    order = sorted(names, key=lambda n: -val(segs, n))
    order = [n for n in order if val(segs, n) > 0]
    first = 5; last = first + len(order) - 1; tr = last + 1
    for i, n in enumerate(order):
        r = first + i
        ws.cell(r, 1, i + 1); ws.cell(r, 2, n); ws.cell(r, 3, PRIVATE_SECTOR(n))
        ws.cell(r, 4, '=' + f_sum(segs, n)).number_format = CR
        ws.cell(r, 5, f'=D{r}/D${tr}').number_format = PCT
        ws.cell(r, 6, f'=SUM(E${first}:E{r})').number_format = PCT
        if D[n].get('fy25'):
            ws.cell(r, 7, '=' + f_sum(segs, n, 25)).number_format = CR
            ws.cell(r, 8, f'=IF(G{r}=0,"n.m.",D{r}/G{r}-1)').number_format = PCT
        else:
            ws.cell(r, 8, 'n.a.')
        ws.cell(r, 9, f'=D{r}/({f_sum(ALLC, n)})').number_format = PCT
        if PRIVATE_SECTOR(n) == 'Private':
            ws.cell(r, 10, f'=D{r}/SUMIF(C${first}:C${last},"Private",D${first}:D${last})').number_format = PCT
            ws.cell(r, 11, f'=COUNTIFS(C${first}:C{r},"Private")')
        if n.startswith('Bharti'):
            for c in range(1, 12): ws.cell(r, c).fill = bfill; ws.cell(r, c).font = Font(bold=True)
    ws.cell(tr, 2, 'Total').font = Font(bold=True)
    ws.cell(tr, 4, f'=SUM(D{first}:D{last})').number_format = CR
    ws.cell(tr, 5, f'=SUM(E{first}:E{last})').number_format = PCT
    ws.cell(tr, 7, f'=SUM(G{first}:G{last})').number_format = CR
    ws.cell(tr, 8, f'=D{tr}/G{tr}-1').number_format = PCT
    ws.cell(tr, 9, f'=D{tr}/Summary!G{rt}').number_format = PCT
    ws.cell(tr, 10, f'=SUM(J{first}:J{last})').number_format = PCT
    for c in range(1, 12): ws.cell(tr, c).font = Font(bold=True); ws.cell(tr, c).fill = gfill
    ws.freeze_panes = 'C5'
    # observations (computed here, stated with numbers)
    v = {n: val(segs, n) for n in order}; T = sum(v.values())
    o = [f'{tab} is ₹{T:,.0f} cr, {T / ind_tot:.1%} of industry premium (4 categories, excl. health & variable).']
    def top(k): return sum(v[n] for n in order[:k]) / T
    o.append(f'Concentration: top 3 ({", ".join(order[:3])}) = {top(3):.1%}; top 5 = {top(5):.1%}; top 10 = {top(10):.1%}. {len(order)} insurers write this business.')
    lic = [n for n in order if n.startswith('LIC')]
    if lic:
        po = [n for n in order if not n.startswith('LIC')]; P = sum(v[n] for n in po)
        o.append(f'LIC holds {v[lic[0]] / T:.1%}; private insurers together {P / T:.1%} (₹{P:,.0f} cr). Within the private sector, top 3 ({", ".join(po[:3])}) = {sum(v[n] for n in po[:3]) / P:.1%}; top 5 = {sum(v[n] for n in po[:5]) / P:.1%}.')
        if b in v: o.append(f'Bharti is #{po.index(b) + 1} of {len(po)} private insurers here with {v[b] / P:.2%} of private-sector premium.')
    have25 = [n for n in order if D[n].get('fy25') and val(segs, n, 25) and val(segs, n, 25) > 0]
    T25 = sum(val(segs, n, 25) for n in have25); T26c = sum(v[n] for n in have25)
    if T25: o.append(f'Category grew {T26c / T25 - 1:.1%} YoY across the {len(have25)} insurers with an FY25 split (LIC excluded: its FY25 segment split was not supplied).')
    big = [n for n in have25 if v[n] >= 0.01 * T]
    if big:
        g = sorted(big, key=lambda n: -(v[n] / val(segs, n, 25)))
        o.append('Fastest growers (≥1% share): ' + '; '.join(f'{n} {v[n] / val(segs, n, 25) - 1:+.0%}' for n in g[:3]) +
                 '. Slowest: ' + '; '.join(f'{n} {v[n] / val(segs, n, 25) - 1:+.0%}' for n in g[-2:]) + '.')
    if b in v:
        bmix = v[b] / val(ALLC, b); imix_ = T / ind_tot
        o.append(f'Bharti Life: rank {order.index(b) + 1} of {len(order)}, ₹{v[b]:,.0f} cr, {v[b] / T:.2%} share'
                 + (f', YoY {v[b] / val(segs, b, 25) - 1:+.1%}' if val(segs, b, 25) else '')
                 + f'. {tab} is {bmix:.0%} of Bharti\'s book vs {imix_:.0%} for the industry — ' + ('over-indexed.' if bmix > imix_ else 'under-indexed.'))
        peers = [n for n in order if abs(order.index(n) - order.index(b)) <= 2 and n != b]
        o.append('Nearest peers by size in this category: ' + ', '.join(f'{n} (₹{v[n]:,.0f} cr)' for n in peers) + '.')
    o += NOTES.get(tab, [])
    r = tr + 2
    ws.cell(r, 1, 'Key observations').font = Font(bold=True, size=12, color=NAVY)
    for k, t in enumerate(o, 1): ws.cell(r + k, 1, f'{k}. {t}')

for c in CATS: cat_tab(*c)

# ---------- Group overlay
ws = wb.create_sheet('Group')
title(ws, 'Group business — FY26 total group premium (₹ crore)',
      'OVERLAY VIEW: source L-25(ii) YTD grand total (new business + renewal). Group premium is ALSO included in the product tabs (L-1 does not split individual/group for most insurers) — do not add to them.')
hdr(ws, 4, ['Rank', 'Insurer', 'Sector', 'FY26 group premium', 'Share of group market', 'Cumulative share', 'Group as % of insurer grand total', 'Individual premium (grand total − group)'],
    [6, 26, 9, 13, 11, 11, 13, 14])
gnames = [n for n in names if D[n].get('group_fy26') is not None]
order = sorted(gnames, key=lambda n: -D[n]['group_fy26'])
first = 5; tr = first + len(order)
for i, n in enumerate(order):
    r = first + i
    ws.cell(r, 1, i + 1); ws.cell(r, 2, n); ws.cell(r, 3, PRIVATE_SECTOR(n))
    ws.cell(r, 4, f'=Data!{COLG}{ROW[n]}').number_format = CR
    ws.cell(r, 5, f'=D{r}/D${tr}').number_format = PCT
    ws.cell(r, 6, f'=SUM(E${first}:E{r})').number_format = PCT
    ws.cell(r, 7, f'=D{r}/Data!{COL26["GT"]}{ROW[n]}').number_format = PCT
    ws.cell(r, 8, f'=Data!{COL26["GT"]}{ROW[n]}-D{r}').number_format = CR
    if n.startswith('Bharti'):
        for c in range(1, 9): ws.cell(r, c).fill = bfill; ws.cell(r, c).font = Font(bold=True)
ws.cell(tr, 2, 'Total').font = Font(bold=True)
ws.cell(tr, 4, f'=SUM(D{first}:D{tr - 1})').number_format = CR
ws.cell(tr, 8, f'=SUM(H{first}:H{tr - 1})').number_format = CR
for c in range(1, 9): ws.cell(tr, c).font = Font(bold=True); ws.cell(tr, c).fill = gfill
ws.freeze_panes = 'C5'
gv = {n: D[n]['group_fy26'] / 100 for n in order}; GT_ = sum(gv.values())
gt_all = sum(D[n]['fy26']['GT'] for n in gnames) / 100
o = [f'Group premium is ₹{GT_:,.0f} cr = {GT_ / gt_all:.1%} of the covered insurers\' grand-total premium.',
     f'Top 3 ({", ".join(order[:3])}) = {sum(gv[n] for n in order[:3]) / GT_:.1%}; top 5 = {sum(gv[n] for n in order[:5]) / GT_:.1%}.',
     'Group-led models (group ≥ 50% of book): ' + ', '.join(f'{n} {gv[n] / (D[n]["fy26"]["GT"] / 100):.0%}' for n in order if gv[n] / (D[n]['fy26']['GT'] / 100) >= 0.5) + ' — mostly credit-life / bancassurance-led group term.']
if b in gv:
    o.append(f'Bharti Life: ₹{gv[b]:,.0f} cr group premium, rank {order.index(b) + 1} of {len(order)}, {gv[b] / GT_:.2%} share; group is only {gv[b] / (D[b]["fy26"]["GT"] / 100):.0%} of Bharti\'s book vs {GT_ / gt_all:.0%} for the industry.')
o += NOTES.get('Group', [])
r = tr + 2
ws.cell(r, 1, 'Key observations').font = Font(bold=True, size=12, color=NAVY)
for k, t in enumerate(o, 1): ws.cell(r + k, 1, f'{k}. {t}')

# ---------- Mapping & notes
ws = wb.create_sheet('Mapping & Notes')
title(ws, 'Method, column mapping and reconciliation', 'All figures from IRDAI-mandated public disclosures, Q4 FY26 (year ended 31 March 2026).')
lines = NOTES.get('_method', [])
r = 4
for t in lines:
    ws.cell(r, 1, t); r += 1
r += 1
hdr(ws, r, ['Insurer', 'Source file', 'L-1-A-RA columns, in order (code; *G = group column)', 'Sum of segments (₹ cr)', 'L-1 grand total (₹ cr)', 'Difference (₹ lakh)'], [26, 40, 90, 14, 14, 12])
for n in names:
    r += 1; o = D[n]
    s = sum(v for k, v in o['fy26'].items() if k != 'GT')
    ws.cell(r, 1, n); ws.cell(r, 2, o.get('file', '')); ws.cell(r, 3, ' | '.join(x or '(subtotal)' for x in o.get('layout', [])))
    ws.cell(r, 4, s / 100).number_format = CR; ws.cell(r, 5, o['fy26']['GT'] / 100).number_format = CR
    ws.cell(r, 6, round(s - o['fy26']['GT'], 1))
r += 2
ws.cell(r, 1, 'Codes: PL par life, PP par pension, PA par annuity, NL non-par life, NA non-par annuity, NP non-par pension, LL linked life, LP linked pension, H health, V variable insurance, GT grand total.')
for t in NOTES.get('_sources', []):
    r += 1; ws.cell(r, 1, t)

wb.move_sheet('Data', offset=len(wb.sheetnames))
wb.save(OUT)
print('saved', OUT)
