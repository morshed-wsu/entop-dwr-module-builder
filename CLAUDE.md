# SDL DWR Engine — CLAUDE.md

## Mission
Generate the SDL Daily Work Report (.docx) by populating the uploaded template binary with data from the uploaded daily Excel file. No Dry Run. No confirmation prompts. Execute and deliver the file immediately.

## Input Files (both uploaded directly into chat each session)
- **Daily Excel** — primary data source for all values
- **Template .docx** — `Daily_Work_Report_Template_-_enTop_v1_5_0.docx` — binary only; the Knowledge Base text extract is ignored

## Execution Protocol — run in a single Python script after unpacking

### Step 1 — Unpack
```
python3 /mnt/skills/public/docx/scripts/office/unpack.py template.docx unpacked/
```

### Step 2 — Extract ALL data from Excel in one pass
Use `openpyxl` with `data_only=True`. Read everything needed in a single script:

**Column map (Production_Center, rows 8–108):**
| Col | Letter | Field |
|-----|--------|-------|
| 4 | D | Type (Q/O) |
| 5 | E | Quote/Order No |
| 11 | K | Value (NZD) |
| 12 | L | Customer |
| 15 | O | Grp Work |
| 21 | U | RM Time (hrs) |
| 25 | Y | Status |

**⚠️ CRITICAL — Job No normalization (col E):** Mixed int/str types in the same column cause silent equality failures. Always normalize immediately on read:
```python
job_no_raw = ws_pc.cell(r, 5).value
job_no = str(int(job_no_raw)) if isinstance(job_no_raw, (int, float)) else str(job_no_raw) if job_no_raw is not None else ''
```

**⚠️ CRITICAL — Customer name normalization (col L):** Customer names must be stripped and normalised before any grouping, leaderboard, or chart computation. The same customer can appear with a trailing space, leading space, or inconsistent spacing in different rows — causing it to split into two separate leaderboard entries. Always normalize:
```python
customer_raw = ws_pc.cell(r, 12).value
customer = ' '.join(str(customer_raw).split()) if customer_raw else ''
```
`' '.join(str.split())` strips leading/trailing whitespace AND collapses internal multiple spaces in one step. Apply this before all deduplication, grouping, leaderboard ranking, and chart data construction.

**Filters:**
- Active analytics: Status in {C, IP} AND Grp Work ≠ "Continued"
- Completed: Status = C
- Incomplete: Status in {IP, NA, GU}

**⚠️ CRITICAL — "Continued" row exclusion:** Grp Work = "Continued" is the first entry of a split job; its hours are zero and must be excluded from all analytics, charts, and the Completed Tasks table. The paired "Finished" row carries the actual hours — never exclude it.

**Summary sheet cells:**
- `B1` = DWR number | `C1` = date
- `C3:I6` = Resource Utilization (4 rows × 7 cols)
- `A13:I13` = Quotations row (9 cells: target, done, WIP, pending, actual_hrs, overdue, parts_done, parts_left, value)
- `J13` = Orders target qty (col 10, NOT col 11) | `K13:R13` = Orders rest (8 cells)
- `A19:J19` = Total Jobs row (10 cells)
- `C115` = completed count | `D115` = IP count | `E115` = NA/GU count
- `G115` = completed hrs | `H115` = IP hrs | `I115` = NA hrs

**Compute in same script:**
- Order/quote leaderboards by value, count, RM time (top 3 + OTHER — see Leaderboard Standard below)
- Chart 1–9 + chartEx1 data arrays
- Descriptive statistics (count, sum, mean, median, min, max, range, Q1, Q3, stddev) for all/orders/quotes RM times
- Value frequency distribution: $0-500, $501-2k, $2k-5k, $5k-10k, $10k-15k, $15k-25k, $25k-50k, $50k-75k, $75k+
- Log-normal fit (μ, σ of ln(times)) for Chart 9
- Key Observations (anchor check ≥85%, top efficiency job $/hr, time-heavy orders, capacity vs 40hr baseline)

## Leaderboard Standard — TOP 3 + OTHER CUSTOMERS

**All leaderboards, pie charts, and leaderboard tables use TOP 3 customers + OTHER CUSTOMERS.** This applies to:
- Orders by count (chart3 / pie chart "Top Customers by Number of Orders")
- Orders by value (chart4)
- Orders by RM time (chart5)
- Quotes by value (chart6)
- Quotes by RM time (chart7)
- Ordering customers leaderboard table in document.xml
- Quoted customers leaderboard table in document.xml

```python
TOP_N = 3  # always 3; remaining go into OTHER CUSTOMERS

def make_leaderboard(jobs, sort_key, top_n=3):
    """Returns [(customer, value, count, time), ...] top_n + OTHER row."""
    from collections import defaultdict
    cust = defaultdict(lambda: {'value':0,'count':0,'time':0})
    for j in jobs:
        cust[j['customer']]['value'] += j['value']
        cust[j['customer']]['count'] += 1
        cust[j['customer']]['time']  += j['rm_time']
    ranked = sorted(cust.items(), key=lambda x: -x[1][sort_key])
    top = ranked[:top_n]
    other_v = sum(d['value'] for _,d in ranked[top_n:])
    other_c = sum(d['count'] for _,d in ranked[top_n:])
    other_t = sum(d['time']  for _,d in ranked[top_n:])
    rows = [(c, d['value'], d['count'], d['time']) for c,d in top]
    rows.append(('OTHER CUSTOMERS', other_v, other_c, other_t))
    return rows  # 4 rows total: 3 named + OTHER
```

### Step 3 — Update document.xml in one sequential pass
Find each table/section **once**, top-to-bottom. Work on the string in memory — never re-read between steps.

1. Global string replace: old DWR# → new; old dates → new (all formats: `DD May YYYY`, `Day, DD May YYYY`, `DD/MM/YY`, `YYYY-MM-DDT...Z`)
2. Quotations table — 9 data cells (position-based replacement)
3. Orders table — 9 data cells (position-based replacement)
4. Total Jobs table — 10 data cells (position-based replacement)
5. Resource Utilization table — rows 2,3,4
6. Completed Tasks table — replace all 38 template rows; blank extras
7. Incomplete Tasks table — replace all 9 template rows; add/remove rows; update footer
8. Ordering leaderboard table — top 3 customers + OTHER (4 data rows)
9. Quoted leaderboard table — top 3 customers + OTHER (4 data rows)
10. Orders mini-table (by count / value / time + totals)
11. Quotes mini-table (by count / value / time + totals)
12. Descriptive Statistics table — all 10 stat rows × 3 columns
13. Key Observations — replace all 4 bullet text blocks

**⚠️ CRITICAL — Position-based replacement for summary tables (Steps 2–4):**
Sequential `str.replace` on cells with identical values (0, 1) corrupts columns. Use index-based replacement:
```python
def replace_nth_wt(row_xml, position, new_value):
    matches = list(re.compile(r'(<w:t[^>]*>)([^<]*)(</w:t>)').finditer(row_xml))
    m = matches[position]
    return row_xml[:m.start()] + m.group(1) + new_value + m.group(3) + row_xml[m.end():]
```

**⚠️ CRITICAL — Steps 10 & 11: Orders and Quotes mini-tables (6 data cells per row).**
Each mini-table has 3 data rows + 1 totals row. Every row has 6 cells: name_count, count, name_value, value, name_time, time. All 6 must be populated from the correct leaderboard — do NOT reuse the same leaderboard for all three columns. Use `replace_nth_wt()` at positions 0–5 for each row:

```python
def fmt_k(v): return f'${v/1000:.1f}k'

# Compute leaderboards
ord_lb_count = make_leaderboard(active_O, 'count')
ord_lb_value = make_leaderboard(active_O, 'value')
ord_lb_time  = make_leaderboard(active_O, 'time')
q_lb_count   = make_leaderboard(active_Q, 'count')
q_lb_value   = make_leaderboard(active_Q, 'value')
q_lb_time    = make_leaderboard(active_Q, 'time')

# Totals for footer row
cnt_total_o   = sum(r[2] for r in ord_lb_count[:3])
val_total_o   = sum(r[1] for r in ord_lb_value[:3])
time_total_o  = sum(r[3] for r in ord_lb_time[:3])
cnt_pct_o     = cnt_total_o / len(active_O) * 100 if active_O else 0
val_pct_o     = val_total_o / ord_val * 100 if ord_val else 0
time_pct_o    = time_total_o / ord_hrs * 100 if ord_hrs else 0

cnt_total_q   = sum(r[2] for r in q_lb_count[:3])
val_total_q   = sum(r[1] for r in q_lb_value[:3])
time_total_q  = sum(r[3] for r in q_lb_time[:3])
q_cnt_pct     = cnt_total_q / len(active_Q) * 100 if active_Q else 0
q_val_pct     = val_total_q / q_val * 100 if q_val else 0
q_time_pct    = time_total_q / q_hrs * 100 if q_hrs else 0

# Build row tuples: (name_count, count, name_value, value, name_time, time)
new_o_rows = [
    (ord_lb_count[i][0], str(int(ord_lb_count[i][2])),
     ord_lb_value[i][0], fmt_k(ord_lb_value[i][1]),
     ord_lb_time[i][0],  f'{ord_lb_time[i][3]:.2f}')
    for i in range(3)
] + [(f'Totals ({int(cnt_pct_o)}% orders)', str(int(cnt_total_o)),
      f' Totals ({int(val_pct_o)}% Value)', fmt_k(val_total_o),
      f'Totals ({int(time_pct_o)}% Hrs) ', f'{time_total_o:.2f}')]

new_q_rows = [
    (q_lb_count[i][0], str(int(q_lb_count[i][2])),
     q_lb_value[i][0], fmt_k(q_lb_value[i][1]),
     q_lb_time[i][0],  f'{q_lb_time[i][3]:.2f}')
    for i in range(3)
] + [(f'Totals ({int(q_cnt_pct)}% quotes)',
      str(int(cnt_total_q)) if cnt_total_q > 0 else '-',
      f'Totals ({int(q_val_pct)}% Value)', fmt_k(val_total_q),
      f'Totals ({int(q_time_pct)}% Hrs) ', f'{time_total_q:.2f}')]

# Locate each mini-table by its unique footer text, then replace rows 1–4
# (skip header row 0) using replace_nth_wt at positions 0–5
def replace_mini_table(doc, footer_anchor, new_rows):
    idx = doc.find(footer_anchor)
    tbl_start = doc.rfind('<w:tbl', 0, idx)
    tbl_end = doc.find('</w:tbl>', idx) + len('</w:tbl>')
    tbl_xml = doc[tbl_start:tbl_end]
    tr_matches = list(re.finditer(r'<w:tr\b.*?</w:tr>', tbl_xml, re.DOTALL))
    new_tbl = tbl_xml
    for i, new_vals in enumerate(new_rows):
        old_row = tr_matches[i + 1].group()  # +1 skips header
        new_row = old_row
        for pos, val in enumerate(new_vals):
            new_row = replace_nth_wt(new_row, pos, val)
        new_tbl = new_tbl.replace(old_row, new_row, 1)
    return doc[:tbl_start] + new_tbl + doc[tbl_end:]

doc = replace_mini_table(doc, '% orders)', new_o_rows)
doc = replace_mini_table(doc, '% quotes)', new_q_rows)
```

The footer anchor `'% orders)'` uniquely identifies the Orders mini-table; `'% quotes)'` identifies the Quotes mini-table. Using `rfind('<w:tbl')` from that anchor position reliably finds the containing table without ambiguity.

**⚠️ CRITICAL — Step 12: Descriptive Statistics — use position-based replacement for every cell.**
The stats table has 10 data rows (rows 1–10; row 0 = header). Each row has 4 cells: label (pos 0), All (pos 1), Orders (pos 2), Quotes (pos 3). Because many values are identical across rows (e.g. "1.50 hrs" may appear in both Median-Q and Q3-All), sequential `str.replace` will corrupt the wrong cell. Always use `replace_nth_wt()` targeting the specific row and position:

```python
import statistics

def desc_stats(times):
    """Compute all 10 descriptive statistics for a list of RM times."""
    if not times: return {k: '0' for k in ['count','sum','mean','median','min','max','range','q1','q3','stddev']}
    n = len(times)
    s = sorted(times)
    q1 = statistics.quantiles(s, n=4)[0] if n >= 2 else s[0]
    q3 = statistics.quantiles(s, n=4)[2] if n >= 2 else s[0]
    return {
        'count':  str(n),
        'sum':    f'{sum(s):.2f} hrs',
        'mean':   f'{statistics.mean(s):.2f} hrs',
        'median': f'{statistics.median(s):.2f} hrs',
        'min':    f'{min(s):.2f} hrs',
        'max':    f'{max(s):.2f} hrs',
        'range':  f'{max(s)-min(s):.2f} hrs',
        'q1':     f'{q1:.2f} hrs',
        'q3':     f'{q3:.2f} hrs',
        'stddev': f'{statistics.stdev(s):.2f} hrs' if n > 1 else '0.00 hrs',
    }

all_times = [j['rm_time'] for j in active if j['rm_time'] > 0]
o_times   = [j['rm_time'] for j in active_O if j['rm_time'] > 0]
q_times   = [j['rm_time'] for j in active_Q if j['rm_time'] > 0]
stats_all = desc_stats(all_times)
stats_o   = desc_stats(o_times)
stats_q   = desc_stats(q_times)

stat_keys = ['count','sum','mean','median','min','max','range','q1','q3','stddev']

# Locate the stats table (it follows the heading "Descriptive Statistics")
idx_ds = doc.find('Descriptive Statistics')
tbl_start_ds = doc.find('<w:tbl', idx_ds)
tbl_end_ds = doc.find('</w:tbl>', tbl_start_ds) + len('</w:tbl>')
tbl_xml_ds = doc[tbl_start_ds:tbl_end_ds]
tr_ds = list(re.finditer(r'<w:tr\b.*?</w:tr>', tbl_xml_ds, re.DOTALL))

new_tbl_ds = tbl_xml_ds
for i, key in enumerate(stat_keys):
    old_row = tr_ds[i + 1].group()  # +1 skips header
    new_row = old_row
    new_row = replace_nth_wt(new_row, 1, stats_all[key])  # All column
    new_row = replace_nth_wt(new_row, 2, stats_o[key])    # Orders column
    new_row = replace_nth_wt(new_row, 3, stats_q[key])    # Quotes column
    new_tbl_ds = new_tbl_ds.replace(old_row, new_row, 1)

doc = doc[:tbl_start_ds] + new_tbl_ds + doc[tbl_end_ds:]
```

Note: descriptive statistics are computed from **individual job-level RM times** (one entry per active job), not per-customer aggregates. The Max for Quotes is the single longest quote job, not the top customer's total time.

### Step 4 — Update ancillary XML files
- `docProps/core.xml` → `<cp:contentStatus>` = new DWR#
- `customXml/item1.xml` → `<PublishDate>` = `YYYY-MM-DDT00:00:00`
- `word/footer1.xml`, `word/footer2.xml`, `word/header1.xml` → date/DWR# replacements

### Step 5 — Update chart XML files

| File | Data | Leaderboard sort |
|------|------|-----------------|
| chart1.xml | Value, Value%, Time, Time%, Count — O and Q series | n/a |
| chart2.xml | Order Frequency + Quote Frequency — 9 brackets each | n/a |
| chart3.xml | Order Count: top 3 + OTHER (names + counts) | by count |
| chart4.xml | Order Value: top 3 + OTHER (names + values + counts) | by value |
| chart5.xml | Order RM Time: top 3 + OTHER (names + times) | by time |
| chart6.xml | Quote Value: top 3 + OTHER (names + values + counts) | by value |
| chart7.xml | Quote RM Time: top 3 + OTHER (names + times) | by time |
| chart8.xml | Histogram 1-hr bins: ≤1 through ≤8 (8 values) | n/a |
| chart9.xml | Histogram 0.5-hr bins: 14 freq + 14 log-normal expected | n/a |
| chartEx1.xml | strDim lvl (Q/O labels) + numDim lvl (RM times) | n/a |

**chart1.xml special fields:**
- Count of Q/O (series 4): use **active analytics count** (C+IP, excl Continued) — not just completed
- Y-axis max: `round_up_to_nearest(max(ord_val, q_val), 50000)` — remove hardcoded `<c:max val="100000"/>`
- Data label position: `<c:dLblPos val="inEnd"/>`

**⚠️ CRITICAL — Chart 1 visible data table percentage formatting**

For chart1.xml, the percentage rows in the visible chart data table MUST display as formatted percentages, not raw decimal values.

Affected Chart 1 series:
- `% of Total Value` from `Sheet1!$L$3:$L$5`
- `% of ∑ Time` from `Sheet1!$N$3:$N$5`

Required behaviour:
- Store percentage values as decimals in XML/workbook, e.g. `0.586706614`
- Display them as percentages with two decimals, e.g. `58.67%`
- Never display these rows as raw decimals such as `0.586706614`

Implementation requirements:
1. In the embedded workbook for Chart 1, force cells `L3:L5` and `N3:N5` to use percentage number format `0.00%`.
2. If adding or modifying workbook styles is required, add a custom `<numFmt formatCode="0.00%">`, add/assign an `<xf>` using that `numFmtId`, and apply that style to `L3:L5` and `N3:N5`.

**⚠️ CRITICAL — ALL chart XML files: ALWAYS update cached values directly.**
Word renders charts entirely from the XML cache. It only falls back to the embedded Excel when the user manually clicks "Edit Data". If cached values are not updated, the chart displays old template data visually even when the embedded Excel is correct. This applies to **every chart** — chart1 through chart9 and chartEx1.

Use `replace_numref_values()` for standard `<c:numRef>` charts (chart1, chart2, chart8, chart9):
```python
def replace_numref_values(xml, formula, new_vals):
    """Replace cached <c:pt> values in the <c:numCache> of the numRef identified by formula."""
    def replace_block(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        fc = re.search(r'<c:formatCode>[^<]*</c:formatCode>', block)
        fmt = fc.group() if fc else ''
        pts_xml = f'<c:ptCount val="{len(new_vals)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_vals)
        )
        new_cache = f'<c:numCache>{fmt}{pts_xml}</c:numCache>'
        return re.sub(r'<c:numCache>.*?</c:numCache>', new_cache, block, flags=re.DOTALL)
    return re.sub(r'<c:numRef>.*?</c:numRef>', replace_block, xml, flags=re.DOTALL)
```

**chart1.xml — 5 numRef blocks to update (identify each by its formula string):**
| Formula | Series | Values |
|---------|--------|--------|
| `Sheet1!$J$3:$J$5` | ∑ Value (NZD) | `[ord_val, q_val]` |
| `Sheet1!$L$3:$L$5` | % of Total Value | `[ord_val/total_val, q_val/total_val]` (16 decimal places; formatCode=`0.00%`) |
| `Sheet1!$M$3:$M$5` | ∑ Time | `[ord_hrs, q_hrs]` |
| `Sheet1!$N$3:$N$5` | % of ∑ Time | `[ord_hrs/all_hrs, q_hrs/all_hrs]` (16 decimal places; formatCode=`0.00%`) |
| `Sheet1!$K$3:$K$5` | Count of Q/O | `[active_O_count, active_Q_count]` |

**chart2.xml — 2 numRef blocks to update:**
| Formula | Series | Values |
|---------|--------|--------|
| `Sheet1!$Q$2:$Q$10` | Order Frequency | `ord_freq[0..8]` (9 bracket counts) |
| `Sheet1!$R$2:$R$10` | Quote Frequency | `q_freq[0..8]` (9 bracket counts) |

**chart8.xml — 1 numRef block to update:**
| Formula | Series | Values |
|---------|--------|--------|
| `Sheet1!$B$2:$B$9` | Frequency (1-hr bins) | `hist_1hr[0..7]` (8 bin counts: ≤1 through ≤8) |

**chart9.xml — 2 numRef blocks to update (no embedded workbook — XML cache only):**
| Formula | Series | Values |
|---------|--------|--------|
| `Descriptive_Stats!$Z$15:$Z$28` | Frequency (0.5-hr bins) | `hist_half[0..13]` (14 bin counts) |
| `Descriptive_Stats!$AA$15:$AA$28` | Log-normal Expected | `lognorm_expected[0..13]` (14 values) |

**⚠️ NOTE — chart9 has NO embedded workbook.** It renders purely from its XML cache. Do not search for or patch a WS9 embedding — none exists.

**charts 3–7 — BOTH strCache (customer names) AND numCache (values/counts/times) must be updated.**
Each pie/donut chart in `word/charts/chart3.xml` through `chart7.xml` contains its own independent XML cache. Word renders these charts from this cache WITHOUT consulting the embedded workbook. Both caches must be updated together, identified by their `<c:f>` formula strings:

```python
def replace_strcache(xml, formula, new_names):
    """Replace cached <c:pt> string values in a <c:strCache> identified by formula."""
    def repl(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        pts_xml = f'<c:ptCount val="{len(new_names)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_names)
        )
        new_cache = f'<c:strCache>{pts_xml}</c:strCache>'
        return re.sub(r'<c:strCache>.*?</c:strCache>', new_cache, block, flags=re.DOTALL)
    return re.sub(r'<c:strRef>.*?</c:strRef>', repl, xml, flags=re.DOTALL)

def replace_numcache(xml, formula, new_vals):
    """Replace cached <c:pt> values in a <c:numCache> identified by formula."""
    def repl(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        fc = re.search(r'<c:formatCode>[^<]*</c:formatCode>', block)
        fmt = fc.group() if fc else ''
        pts_xml = f'<c:ptCount val="{len(new_vals)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_vals)
        )
        new_cache = f'<c:numCache>{fmt}{pts_xml}</c:numCache>'
        return re.sub(r'<c:numCache>.*?</c:numCache>', new_cache, block, flags=re.DOTALL)
    return re.sub(r'<c:numRef>.*?</c:numRef>', repl, xml, flags=re.DOTALL)
```

**chart3–7 formula lookup table:**

| Chart | strRef formula (names) | numRef formula (values) | Leaderboard |
|-------|------------------------|-------------------------|-------------|
| chart3 | `Summary_Charts!$AB$47:$AB$50` | `Summary_Charts!$AC$47:$AC$50` | ord_lb_count: names, counts |
| chart4 | `Sheet1!$O$29:$O$32` (appears twice — patch both occurrences) | `Sheet1!$P$29:$P$32`, `Sheet1!$Q$29:$Q$32` | ord_lb_value: names, values, counts |
| chart5 | `Summary_Charts!$AB$67:$AB$70` | `Summary_Charts!$AC$67:$AC$70` | ord_lb_time: names, times |
| chart6 | `Sheet1!$O$37:$O$40` (appears twice — patch both occurrences) | `Sheet1!$P$37:$P$40`, `Sheet1!$Q$37:$Q$40` | q_lb_value: names, values, counts |
| chart7 | `Summary_Charts!$AB$77:$AB$80` (or similar) | `Summary_Charts!$AC$77:$AC$80` | q_lb_time: names, times |

Always call `replace_strcache()` once per strRef formula. When a formula appears twice in the same chart XML (chart4 and chart6), call `replace_strcache()` twice with the same formula and names — the function replaces the next unpatched occurrence each time.

**chartEx1.xml — uses `<cx:lvl>` not `<c:numCache>` (different namespace).** Replace the `<cx:lvl>` blocks inside `<cx:strDim>` and `<cx:numDim>` directly. Data is individual job-level — one entry per active job, Q jobs first (sorted ascending by RM time), then O jobs (sorted ascending). Also update the formula range references to match the new row count:
```python
q_sorted = sorted([j['rm_time'] for j in active_Q if j['rm_time'] > 0])
o_sorted = sorted([j['rm_time'] for j in active_O if j['rm_time'] > 0])
all_labels = ['Q'] * len(q_sorted) + ['O'] * len(o_sorted)
all_values = q_sorted + o_sorted
total_pts = len(all_labels)
last_row = 1 + total_pts  # row 1 = header; data starts at row 2

str_pts = ''.join(f'<cx:pt idx="{i}">{lbl}</cx:pt>' for i, lbl in enumerate(all_labels))
num_pts = ''.join(f'<cx:pt idx="{i}">{val}</cx:pt>' for i, val in enumerate(all_values))
new_str_lvl = f'<cx:lvl ptCount="{total_pts}">{str_pts}</cx:lvl>'
new_num_lvl = f'<cx:lvl ptCount="{total_pts}" formatCode="General">{num_pts}</cx:lvl>'

cex = re.sub(r'(<cx:strDim[^>]*>.*?<cx:f>[^<]*</cx:f>)<cx:lvl[^>]*>.*?</cx:lvl>',
             lambda m: m.group(1) + new_str_lvl, cex, flags=re.DOTALL)
cex = re.sub(r'(<cx:numDim[^>]*>.*?<cx:f>[^<]*</cx:f>)<cx:lvl[^>]*>.*?</cx:lvl>',
             lambda m: m.group(1) + new_num_lvl, cex, flags=re.DOTALL)
cex = re.sub(r'<cx:f>Sheet1!\$A\$2:\$A\$\d+</cx:f>', f'<cx:f>Sheet1!$A$2:$A${last_row}</cx:f>', cex)
cex = re.sub(r'<cx:f>Sheet1!\$B\$2:\$B\$\d+</cx:f>', f'<cx:f>Sheet1!$B$2:$B${last_row}</cx:f>', cex)
```

### Step 6 — Patch embedded Excel workbooks

**All nine embeddings (WS0–WS8) need patching.** WS0–WS1 drive chart1/chart2; WS2–WS6 drive charts 3–7; WS7 drives chart8; WS8 drives chartEx1. chart9 has no embedding.

| Embedding | Filename | Chart | Sheet to patch |
|-----------|----------|-------|----------------|
| WS0 | Microsoft_Excel_Worksheet.xlsx | chart1 | sheet5.xml |
| WS1 | Microsoft_Excel_Worksheet1.xlsx | chart2 | sheet5.xml |
| WS2 | Microsoft_Excel_Worksheet2.xlsx | chart3 | sheet5.xml + sheet6.xml |
| WS3 | Microsoft_Excel_Worksheet3.xlsx | chart4 | sheet5.xml |
| WS4 | Microsoft_Excel_Worksheet4.xlsx | chart5 | sheet6.xml |
| WS5 | Microsoft_Excel_Worksheet5.xlsx | chart6 | sheet5.xml |
| WS6 | Microsoft_Excel_Worksheet6.xlsx | chart7 | sheet6.xml |
| WS7 | Microsoft_Excel_Worksheet7.xlsx | chart8 | sheet1.xml |
| WS8 | Microsoft_Excel_Worksheet8.xlsx | chartEx1 | sheet1.xml |

#### WS0 + WS1 (Microsoft_Excel_Worksheet.xlsx and Worksheet1.xlsx) — chart1 and chart2

Patch `xl/worksheets/sheet5.xml` (Sheet1) in each:
- `J3`=ord_val, `K3`=active_O_count, `L3`=ord_val/total_val, `M3`=ord_hrs, `N3`=ord_hrs/all_hrs
- `J4`=q_val, `K4`=active_Q_count, `L4`=q_val/total_val, `M4`=q_hrs, `N4`=q_hrs/all_hrs
- `J5`=total_val, `K5`=total_active_count, `M5`=all_hrs
- `Q2:Q10` = ord_freq[0..8] (static, no `<f>` element)
- `R2:R10` = q_freq[0..8] (static, no `<f>` element)

**⚠️ CRITICAL — FREQUENCY array formula freezing:** `Q2` and `R2` contain array formulas with `ref="Q2:Q10"` — cells Q3:Q10 and R3:R10 don't exist as individual XML nodes. After removing the `<f>` from Q2/R2, explicitly INSERT `<c>` elements for Q3:Q10 and R3:R10 into their rows:
```python
def insert_static_cell(xml, row_num, cell_ref, val):
    if re.search(rf'<c r="{cell_ref}"', xml):
        return re.sub(rf'<c r="{re.escape(cell_ref)}"[^>]*>.*?</c>',
                      f'<c r="{cell_ref}"><v>{val}</v></c>', xml, 1, re.DOTALL)
    col = cell_ref[0]
    def add_cell(m):
        content = m.group(2)
        cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', content))
        pos = next((cm.start() for cm in cells if cm.group(1) >= col), len(content))
        return m.group(1) + content[:pos] + f'<c r="{cell_ref}"><v>{val}</v></c>' + content[pos:] + m.group(3)
    return re.sub(rf'(<row r="{row_num}"[^>]*>)(.*?)(</row>)', add_cell, xml, 1, re.DOTALL)
```

#### WS2–WS6 — leaderboard charts — CRITICAL RULES

**⚠️ CRITICAL RULE 1 — Always start from the ORIGINAL TEMPLATE bytes** for WS2–WS6. Never patch a previously-patched version. Read template bytes fresh each time:
```python
with zipfile.ZipFile('template.docx', 'r') as z:
    t_ws2 = z.read('word/embeddings/Microsoft_Excel_Worksheet2.xlsx')
    # ... etc for WS3–WS6
```

**⚠️ CRITICAL RULE 2 — String cells must use `t="inlineStr"`** with `<is><t>text</t></is>`. Never write a customer name as `<c r="O29"><v>Customer X</v></c>` — that is invalid OOXML (Word treats `<v>` without `t="s"` as numeric) and corrupts the workbook. Always use:
```python
def str_cell(cell_ref, text, style=''):
    sa = f' s="{style}"' if style else ''
    return f'<c r="{cell_ref}"{sa} t="inlineStr"><is><t>{text}</t></is></c>'

def num_cell(cell_ref, value, style=''):
    sa = f' s="{style}"' if style else ''
    return f'<c r="{cell_ref}"{sa}><v>{value}</v></c>'
```
Preserve the original style attributes (e.g. `s="132"`) by reading them from the template sheet before patching:
```python
def get_style(xml, cell_ref):
    m = re.search(rf'<c r="{cell_ref}"[^>]*>', xml)
    s = re.search(r' s="(\d+)"', m.group()) if m else None
    return s.group(1) if s else ''
```

**⚠️ CRITICAL RULE 3 — Fix ZIP flag_bits after writing.** When Python's `zipfile` writes entries via `writestr(info, data)`, it resets the general-purpose bit flag from 6 to 0. Word's COM/OLE layer checks this flag and reports "linked file isn't available" when it is 0. After building the new xlsx bytes, patch all local file header and central directory flag fields back to 6:
```python
import struct

def fix_zip_flags(xlsx_bytes):
    """Set flag_bits=6 on all entries in an xlsx zip file."""
    raw = bytearray(xlsx_bytes)
    for sig, offset in [(b'PK\x03\x04', 6), (b'PK\x01\x02', 8)]:
        pos = 0
        while True:
            pos = bytes(raw).find(sig, pos)
            if pos == -1: break
            if struct.unpack_from('<H', raw, pos + offset)[0] == 0:
                struct.pack_into('<H', raw, pos + offset, 6)
            pos += 4
    return bytes(raw)
```
Call `fix_zip_flags()` on the bytes of every rewritten WS2–WS6 before inserting back into the output docx.

**Rewrite helper combining all three rules:**
```python
def rewrite_ws(template_bytes, sheet_patches):
    """
    Patch specific cells in an embedded xlsx, starting from template bytes.
    sheet_patches: {sheet_path: [(cell_ref, cell_xml_string), ...]}
    Preserves original compression and fixes flag_bits.
    """
    t_zip = zipfile.ZipFile(io.BytesIO(template_bytes), 'r')
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w') as z_out:
        for info in t_zip.infolist():
            data = t_zip.read(info.filename)
            if info.filename in sheet_patches:
                xml = data.decode('utf-8')
                for cell_ref, cell_xml in sheet_patches[info.filename]:
                    xml = patch_cell_in_xml(xml, cell_ref, cell_xml)
                data = xml.encode('utf-8')
            z_out.writestr(info, data)  # preserves original compress_type
    t_zip.close()
    return fix_zip_flags(out_buf.getvalue())
```

**WS2 patches (chart3 — order count):**

sheet5.xml (Sheet1): O29:Q32
- O29–O31: `str_cell(f'O{29+i}', name, style_from_template)` for top 3 names by count
- P29–P31: `num_cell(f'P{29+i}', value, style)` — their order values
- Q29–Q31: `num_cell(f'Q{29+i}', count, style)` — their order counts
- O32: `str_cell('O32', 'OTHER CUSTOMERS', style)`, P32/Q32: remaining value/count

sheet6.xml (Summary_Charts): AB47:AC50
- AB47–AB49: `str_cell(f'AB{47+i}', name)` for top 3 names by count
- AC47–AC49: `num_cell(f'AC{47+i}', count)`
- AB50: `str_cell('AB50', 'OTHER CUSTOMERS', style_from_template)`, AC50: remaining count

**WS3 patches (chart4 — order value):** sheet5.xml O29:R32 (sort by value; R = value/ord_val)

**WS4 patches (chart5 — order RM time):** sheet6.xml AB67:AC70 (sort by time; AC = rm_time)

**WS5 patches (chart6 — quote value):** sheet5.xml O37:R40 (sort by value; R = value/q_val)

**WS6 patches (chart7 — quote RM time):** sheet6.xml AB77:AC80 (sort by time; AC = rm_time)

#### WS7 (Microsoft_Excel_Worksheet7.xlsx) — chart8 (1-hr histogram)

Patch `xl/worksheets/sheet1.xml`. Structure: column A = bin upper bound (1–8), column B = frequency count.
- `B2:B9` = `hist_1hr[0..7]` (8 values: count of active jobs in each 1-hr bin)

```python
for i, val in enumerate(hist_1hr):
    row = i + 2
    cell_ref = f'B{row}'
    pattern = re.compile(rf'<c r="{cell_ref}"[^>]*>.*?</c>', re.DOTALL)
    new_cell = f'<c r="{cell_ref}"><v>{val}</v></c>'
    if pattern.search(sheet_xml):
        sheet_xml = pattern.sub(new_cell, sheet_xml, count=1)
```

#### WS8 (Microsoft_Excel_Worksheet8.xlsx) — chartEx1 (boxplot)

Patch `xl/worksheets/sheet1.xml`. Structure: row 1 = header (`B1` = shared string index 0 = "RM Time (hrs)"); rows 2–N = one row per active job. Column A = group label as shared string index (`1`=Q, `2`=O), Column B = RM time value.

**Shared strings (do not modify `xl/sharedStrings.xml`):** index `0`="RM Time (hrs)", `1`="Q", `2`="O".

Completely rebuild `<sheetData>` — Q jobs first (sorted ascending by RM time), then O jobs (sorted ascending). Update `<dimension ref>` to match the new row count:

```python
q_sorted = sorted([j['rm_time'] for j in active_Q if j['rm_time'] > 0])
o_sorted = sorted([j['rm_time'] for j in active_O if j['rm_time'] > 0])
last_row = 1 + len(q_sorted) + len(o_sorted)

new_rows_xml = '<row r="1" spans="1:2" x14ac:dyDescent="0.35"><c r="B1" t="s"><v>0</v></c></row>'
row_num = 2
for t in q_sorted:
    new_rows_xml += f'<row r="{row_num}" spans="1:2" x14ac:dyDescent="0.35"><c r="A{row_num}" t="s"><v>1</v></c><c r="B{row_num}"><v>{t}</v></c></row>'
    row_num += 1
for t in o_sorted:
    new_rows_xml += f'<row r="{row_num}" spans="1:2" x14ac:dyDescent="0.35"><c r="A{row_num}" t="s"><v>2</v></c><c r="B{row_num}"><v>{t}</v></c></row>'
    row_num += 1

sheet_xml = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:B{last_row}"/>', sheet_xml)
sheet_xml = re.sub(r'<sheetData>.*?</sheetData>', f'<sheetData>{new_rows_xml}</sheetData>', sheet_xml, flags=re.DOTALL)
```

### Step 7 — Pack
```
python3 /mnt/skills/public/docx/scripts/office/pack.py unpacked/ output.docx --original template.docx --validate false
```
`--validate false` required: lxml incorrectly rejects c15/c16/c16r3 namespaces in chartEx1 and histogram charts.

### Step 8 — Fix embedding compression (MANDATORY final step)
All embedded `.xlsx` files in the outer docx must be `ZIP_STORED` (uncompressed) so Word can use direct byte-offset access. Run after pack.py every time:
```python
import zipfile, os
def fix_embedding_compression(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        file_data = {i.filename: (z.read(i.filename), i.compress_type) for i in z.infolist()}
    tmp = docx_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as z_out:
        for fname, (data, orig_ct) in file_data.items():
            ct = zipfile.ZIP_STORED if ('embeddings' in fname and fname.endswith('.xlsx')) else orig_ct
            z_out.writestr(fname, data, compress_type=ct)
    os.replace(tmp, docx_path)
fix_embedding_compression('output.docx')
```

## Standards
- Percentages: 2 decimal places
- Currency: `$#,##0.00`
- Hours: 2 decimal places
- Load%: 2 decimal places
- Tone: professional, executive-level
- Q = Quotation, O = Order — never conflate

## Key Observations Logic
Four bullets:
1. **Revenue leader** — top customer by total active value; if ≥85% → "Revenue Anchor" warning; otherwise state share and top job $/hr
2. **Quote of the Day** — highest $/hr quote; state value, time, rate; flag for conversion
3. **Time-Heavy Orders** — highest RM time relative to value; name IP risk if any job still active
4. **Capacity & Backlog** — actual hrs vs 40hr baseline, load%, NA count and deferred revenue risk

## Output filename
`DWR_[DWR#]_[DD-Mon-YYYY].docx` — e.g. `DWR_2026-91_22-May-2026.docx`
