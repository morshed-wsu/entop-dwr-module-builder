# SDL DWR Step 3 Instructions — Leaderboards, Descriptive Statistics, Histograms, Boxplot

## Mission
Generate the Step 3 Daily Work Report `.docx` using `Daily Work Report Step 3 Template - enTop v1.0.0.docx`. This step outputs the leaderboard table section, descriptive statistics, histogram charts, and the boxplot/time-distribution chart if present.


## Shared Input + Extraction Rules

Input files for this step:
- Daily Excel file: `DD-MM-YYYY_SDL_Daily Work Report_AI.xlsx`
- Matching step template `.docx`

Use the template binary directly. Do not rely on Knowledge Base text extracts.

Unpack / pack pattern:
```bash
python3 /mnt/skills/public/docx/scripts/office/unpack.py template.docx unpacked/
# patch XML files inside unpacked/
python3 /mnt/skills/public/docx/scripts/office/pack.py unpacked/ output.docx --original template.docx --validate false
```
Use `--validate false` because Word chart namespaces may fail strict validation even when Word opens the document correctly.

Read the Excel once with `openpyxl.load_workbook(..., data_only=True)`.

### Production_Center column map, rows 8-108
| Column | Field |
|---|---|
| D | Type: `Q` or `O` |
| E | Quote/Order No |
| K | Value (NZD) |
| L | Customer |
| O | Grp Work |
| U | RM Time (hrs) |
| Y | Status |

### Mandatory normalization
```python
job_no_raw = ws_pc.cell(r, 5).value
job_no = str(int(job_no_raw)) if isinstance(job_no_raw, (int, float)) else str(job_no_raw) if job_no_raw is not None else ''

customer_raw = ws_pc.cell(r, 12).value
customer = ' '.join(str(customer_raw).split()) if customer_raw else ''
```

### Mandatory filters
- `active`: `Status in {'C', 'IP'}` and `Grp Work != 'Continued'`
- `completed`: `Status == 'C'` and `Grp Work != 'Continued'`
- `incomplete`: `Status in {'IP', 'NA', 'GU'}` and `Grp Work != 'Continued'`
- Never exclude `Grp Work = 'Finished'`; it carries the actual hours for split jobs.

### Summary sheet cells
- `B1` = DWR number
- `C1` = report date
- `C3:I6` = Resource Utilization table rows
- `A13:I13` = Quotations summary row
- `J13` = Orders target qty; `K13:R13` = remaining Orders summary values
- `A19:J19` = Total Jobs summary row
- `C115` = completed count; `D115` = IP count; `E115` = NA/GU count
- `G115` = completed hrs; `H115` = IP hrs; `I115` = NA/GU hrs

### Formatting standards
- Currency: `$#,##0.00`
- Hours: 2 decimal places
- Percentages / Load%: 2 decimal places
- Q = Quotation and O = Order. Never conflate them.
- Preserve the template’s existing page size, margins, fonts, colors, tables, charts, image positions, headers, footers, and footer page-number style.


## Leaderboard Standard: Top 3 + OTHER CUSTOMERS

All leaderboard tables and related charts use exactly 4 rows: top 3 customers plus `OTHER CUSTOMERS`.

```python
from collections import defaultdict
TOP_N = 3

def make_leaderboard(jobs, sort_key, top_n=3):
    cust = defaultdict(lambda: {'value': 0, 'count': 0, 'time': 0})
    for j in jobs:
        cust[j['customer']]['value'] += j['value']
        cust[j['customer']]['count'] += 1
        cust[j['customer']]['time']  += j['rm_time']
    ranked = sorted(cust.items(), key=lambda x: -x[1][sort_key])
    top = ranked[:top_n]
    other_v = sum(d['value'] for _, d in ranked[top_n:])
    other_c = sum(d['count'] for _, d in ranked[top_n:])
    other_t = sum(d['time'] for _, d in ranked[top_n:])
    rows = [(c, d['value'], d['count'], d['time']) for c, d in top]
    rows.append(('OTHER CUSTOMERS', other_v, other_c, other_t))
    return rows
```


## XML Replacement Helpers

Use position-based replacement wherever a row/table contains repeated values such as `0`, `1`, or blank text. Do not use uncontrolled sequential `str.replace()` for repeated numeric cells.

```python
import re

def replace_nth_wt(row_xml, position, new_value):
    matches = list(re.compile(r'(<w:t[^>]*>)([^<]*)(</w:t>)').finditer(row_xml))
    m = matches[position]
    return row_xml[:m.start()] + m.group(1) + str(new_value) + m.group(3) + row_xml[m.end():]
```

When the number of table rows can vary, locate the containing table by a nearby unique heading/anchor, clone or remove `<w:tr>` rows as needed, and keep the original cell styles from the template row.


## Step 3 Template Scope
Populate only content that exists in this template:
1. `The top customers of every leaderboard`
2. `ORDERS — Top Customers` mini-table
3. `QUOTES — Top Customers` mini-table
4. `Descriptive Statistics` table
5. Chart 8: Histogram — RM Time (hrs), 1-hour bins
6. Chart 9: Histogram — RM Time (hrs), 0.5-hour bins with log-normal expected line
7. Boxplot / chartEx1: RM Time by Type, if present
8. Related embedded workbook data for charts in this step

## Mini-Table Calculations
Use separate leaderboards for each mini-table column. Do not reuse the same leaderboard for count, value, and time.

```python
def fmt_k(v):
    return f'${v/1000:.1f}k'

active_O = [j for j in active if j['type'] == 'O']
active_Q = [j for j in active if j['type'] == 'Q']
ord_val = sum(j['value'] for j in active_O)
q_val = sum(j['value'] for j in active_Q)
ord_hrs = sum(j['rm_time'] for j in active_O)
q_hrs = sum(j['rm_time'] for j in active_Q)

ord_lb_count = make_leaderboard(active_O, 'count')
ord_lb_value = make_leaderboard(active_O, 'value')
ord_lb_time  = make_leaderboard(active_O, 'time')
q_lb_count   = make_leaderboard(active_Q, 'count')
q_lb_value   = make_leaderboard(active_Q, 'value')
q_lb_time    = make_leaderboard(active_Q, 'time')
```

Each data row has 6 cells:
`name_count`, `count`, `name_value`, `value`, `name_time`, `time`.
Use `replace_nth_wt()` positions `0-5` in each data row.

Footer calculations:
- Orders count footer = sum of top 3 count leaders / active order count
- Orders value footer = sum of top 3 value leaders / total active order value
- Orders time footer = sum of top 3 time leaders / total active order time
- Quotes follow the same logic using quote-only denominators.

## Descriptive Statistics
Compute from individual job-level RM times, not customer aggregates.

```python
import statistics

def desc_stats(times):
    if not times:
        return {k: '0' for k in ['count','sum','mean','median','min','max','range','q1','q3','stddev']}
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
o_times = [j['rm_time'] for j in active_O if j['rm_time'] > 0]
q_times = [j['rm_time'] for j in active_Q if j['rm_time'] > 0]
```

Stats table rows:
`Count`, `Sum`, `Mean`, `Median`, `Min`, `Max`, `Range`, `Q1 (25%)`, `Q3 (75%)`, `Std Dev`.

Use position-based replacement for every stats cell:
- position 1 = All
- position 2 = Orders
- position 3 = Quotes

## Histogram Calculations

### Chart 8: 1-hour bins
Use active job-level RM times.
Bins: `≤1.0`, `≤2.0`, `≤3.0`, `≤4.0`, `≤5.0`, `≤6.0`, `≤7.0`, `≤8.0`.
Produce `hist_1hr[0..7]`.

Chart XML formula:
- `Sheet1!$B$2:$B$9` = `hist_1hr`

Embedded workbook:
- Patch workbook cell range `B2:B9`.

### Chart 9: 0.5-hour bins
Bins: `≤0.5`, `≤1.0`, `≤1.5`, `≤2.0`, `≤2.5`, `≤3.0`, `≤3.5`, `≤4.0`, `≤4.5`, `≤5.0`, `≤5.5`, `≤6.0`, `≤6.5`, `≤7.0`.
Produce:
- `hist_half[0..13]`
- `lognorm_expected[0..13]` from log-normal fit of positive active RM times.

Chart XML formulas:
- `Descriptive_Stats!$Z$15:$Z$28` = `hist_half`
- `Descriptive_Stats!$AA$15:$AA$28` = `lognorm_expected`

Important: Chart 9 has no embedded workbook. Patch XML cache only.

## Chart XML Cache Rules

```python
def replace_numref_values(xml, formula, new_vals):
    import re
    def replace_block(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        fc = re.search(r'<c:formatCode>[^<]*</c:formatCode>', block)
        fmt = fc.group() if fc else ''
        pts_xml = f'<c:ptCount val="{len(new_vals)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_vals)
        )
        return re.sub(r'<c:numCache>.*?</c:numCache>', f'<c:numCache>{fmt}{pts_xml}</c:numCache>', block, flags=re.DOTALL)
    return re.sub(r'<c:numRef>.*?</c:numRef>', replace_block, xml, flags=re.DOTALL)
```

## Boxplot / chartEx1 Rules
If the Step 3 template contains `chartEx1.xml`, patch `<cx:lvl>` blocks directly. Data is one entry per active job, Q jobs first sorted ascending by RM time, then O jobs sorted ascending by RM time.

```python
q_sorted = sorted([j['rm_time'] for j in active_Q if j['rm_time'] > 0])
o_sorted = sorted([j['rm_time'] for j in active_O if j['rm_time'] > 0])
all_labels = ['Q'] * len(q_sorted) + ['O'] * len(o_sorted)
all_values = q_sorted + o_sorted
last_row = 1 + len(all_labels)
```

Patch:
- `<cx:strDim>` level points = Q/O labels
- `<cx:numDim>` level points = RM times
- Formula ranges to `Sheet1!$A$2:$A${last_row}` and `Sheet1!$B$2:$B${last_row}`

Patch the embedded chartEx workbook if present:
- Rebuild `sheet1.xml` `<sheetData>` with row 1 header, then Q rows, then O rows.
- Use shared string index `1` for Q and `2` for O.
- Update `<dimension ref="A1:B[last_row]"/>`.

## Validation Checklist
- Mini-tables use separate count, value, and time leaderboards.
- Descriptive statistics are based on individual active jobs, not per-customer totals.
- Quote max time is the longest single quote job.
- Chart 9 is XML-cache-only; no WS9 workbook is expected.
- Chart 8 and chartEx workbook/cache values match where applicable.


## Final Packaging Rule

After packing, if the step template contains embedded Excel workbooks, force all embedded `.xlsx` files in the outer `.docx` to `ZIP_STORED`.

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
```

Output filename pattern:
`DWR_[DWR#]_[DD-Mon-YYYY]_STEP1.docx`, `..._STEP2.docx`, `..._STEP3.docx`, or `..._STEP4.docx`.
