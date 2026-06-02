# SDL DWR Step 2 Instructions — Daily Metrics and Customer Charts

## Mission
Generate the Step 2 Daily Work Report `.docx` using `Daily Work Report Step 2 Template - enTop v1.0.0.docx`. This step outputs the Daily Metrics section: resource-utilization text/table if present, Total Quote & Order Overview, Quote/Order frequency distribution, top ordering/quoted customer charts, and the two customer analysis tables.


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


## Step 2 Template Scope
Populate only content that exists in this template:
1. `Resource Utilization & Capacity Allocation` paragraph/table if present
2. `Daily Metrics at a Glance`
3. Chart 1: Total Quote & Order Overview
4. Chart 2: Quote & Order NZD Frequency Distribution
5. Chart 3: Top Customers by Number of Orders
6. Chart 4: Top Ordering Customers
7. Chart 5: Top Customers by Order RM Time, if present
8. Chart 6: Top Quoted Customers by $Value, if present
9. Chart 7: Top Customers by Quote RM Time, if present
10. Analysis of Top Ordering Customers table
11. Analysis of Top Quoted Customers table
12. Any visible chart data tables and embedded workbooks connected to those charts

## Required Data Sets
```python
active_O = [j for j in active if j['type'] == 'O']
active_Q = [j for j in active if j['type'] == 'Q']
ord_val = sum(j['value'] for j in active_O)
q_val = sum(j['value'] for j in active_Q)
total_val = ord_val + q_val
ord_hrs = sum(j['rm_time'] for j in active_O)
q_hrs = sum(j['rm_time'] for j in active_Q)
all_hrs = ord_hrs + q_hrs
active_O_count = len(active_O)
active_Q_count = len(active_Q)
```

Frequency brackets for orders and quotes:
`$0-500`, `$501-2000`, `$2001-5000`, `$5001-10000`, `$10001-15000`, `$15001-25000`, `$25001-50000`, `$50001-75000`, `$75001+`.

## Customer Analysis Tables

### Top Ordering Customers table
Use order leaderboard by value: `make_leaderboard(active_O, 'value')`.
Rows:
1. top order customer by value
2. second
3. third
4. `Top Ordering Customers Total`
5. `OTHER CUSTOMERS`

`% of ∑ Value` is each row value divided by the active grand total value or, where the existing master template expects it, by total active value across O+Q. Keep the same denominator behaviour as the current master `CLAUDE.md` output and existing DWR examples.

### Top Quoted Customers table
Use quote leaderboard by value: `make_leaderboard(active_Q, 'value')`.
Rows:
1. top quote customer by value
2. second
3. third
4. `Top Quoted Customers Total`
5. `OTHER CUSTOMERS`

## Chart XML Rules

### General chart cache rule
Always update chart XML cached values directly. Word renders charts from XML caches and may not refresh from embedded workbooks automatically.

```python
def replace_numref_values(xml, formula, new_vals, force_format_code=None):
    import re
    def replace_block(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        if force_format_code is not None:
            fmt = f'<c:formatCode>{force_format_code}</c:formatCode>'
        else:
            fc = re.search(r'<c:formatCode>[^<]*</c:formatCode>', block)
            fmt = fc.group() if fc else ''
        pts_xml = f'<c:ptCount val="{len(new_vals)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_vals)
        )
        new_cache = f'<c:numCache>{fmt}{pts_xml}</c:numCache>'
        return re.sub(r'<c:numCache>.*?</c:numCache>', new_cache, block, flags=re.DOTALL)
    return re.sub(r'<c:numRef>.*?</c:numRef>', replace_block, xml, flags=re.DOTALL)

def replace_strcache(xml, formula, new_names):
    import re
    def repl(m):
        block = m.group()
        if f'<c:f>{formula}</c:f>' not in block:
            return block
        pts_xml = f'<c:ptCount val="{len(new_names)}"/>' + ''.join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(new_names)
        )
        return re.sub(r'<c:strCache>.*?</c:strCache>', f'<c:strCache>{pts_xml}</c:strCache>', block, flags=re.DOTALL)
    return re.sub(r'<c:strRef>.*?</c:strRef>', repl, xml, flags=re.DOTALL)
```

### Chart 1: Total Quote & Order Overview
Update these formula caches:
| Formula | Values |
|---|---|
| `Sheet1!$J$3:$J$5` | `[ord_val, q_val]` |
| `Sheet1!$L$3:$L$5` | `[ord_val/total_val, q_val/total_val]` |
| `Sheet1!$M$3:$M$5` | `[ord_hrs, q_hrs]` |
| `Sheet1!$N$3:$N$5` | `[ord_hrs/all_hrs, q_hrs/all_hrs]` |
| `Sheet1!$K$3:$K$5` | `[active_O_count, active_Q_count]` |

Chart 1 special formatting:
- Count of Q/O uses active counts: C + IP, excluding Continued.
- Y-axis max = round up `max(ord_val, q_val)` to nearest 50,000.
- Data labels position = `inEnd`.

### Critical: Chart 1 visible data-table percentages
The visible data table must display percentages, not raw decimals.

Affected rows:
- `% of Total Value` from `Sheet1!$L$3:$L$5`
- `% of ∑ Time` from `Sheet1!$N$3:$N$5`

Required behaviour:
- Store values as decimals, e.g. `0.586706614`
- Display values as `58.67%`
- Never display raw decimals such as `0.586706614`

Implementation requirements:
1. In `word/charts/chart1.xml`, force percentage caches to `<c:formatCode>0.00%</c:formatCode>`.
2. In the embedded workbook used by Chart 1, force cells `L3:L5` and `N3:N5` to percentage style `0.00%`.
3. If needed, create a custom workbook style with `<numFmt formatCode="0.00%">`, assign an `<xf>` to it, and apply that style to `L3:L5` and `N3:N5`.
4. Final visual validation: chart data table must show `xx.xx%` for percentage rows.

### Chart 2: Frequency Distribution
Update caches:
- `Sheet1!$Q$2:$Q$10` = order frequency, 9 values
- `Sheet1!$R$2:$R$10` = quote frequency, 9 values

### Charts 3-7: Customer charts
Patch both name and value caches:
| Chart | Name formula | Value formula(s) | Leaderboard |
|---|---|---|---|
| chart3 | `Summary_Charts!$AB$47:$AB$50` | `Summary_Charts!$AC$47:$AC$50` | orders by count |
| chart4 | `Sheet1!$O$29:$O$32` | `Sheet1!$P$29:$P$32`, `Sheet1!$Q$29:$Q$32` | orders by value |
| chart5 | `Summary_Charts!$AB$67:$AB$70` | `Summary_Charts!$AC$67:$AC$70` | orders by RM time |
| chart6 | `Sheet1!$O$37:$O$40` | `Sheet1!$P$37:$P$40`, `Sheet1!$Q$37:$Q$40` | quotes by value |
| chart7 | `Summary_Charts!$AB$77:$AB$80` | `Summary_Charts!$AC$77:$AC$80` | quotes by RM time |

If the same `strRef` formula appears twice in a chart, patch both occurrences.

## Embedded Workbook Rules for Step 2
Patch only the embedded workbooks used by the charts present in this Step 2 template.

- Chart 1 workbook: patch `J3:N5`, `L3:L5`, `N3:N5` styles as percentage.
- Chart 2 workbook: patch `Q2:Q10` and `R2:R10`; remove frozen array formulas and insert missing static cells.
- Charts 3-7 workbooks: patch customer names as `inlineStr`; patch numeric cells as numeric cells.

String cell rule:
```python
def str_cell(cell_ref, text, style=''):
    sa = f' s="{style}"' if style else ''
    return f'<c r="{cell_ref}"{sa} t="inlineStr"><is><t>{text}</t></is></c>'

def num_cell(cell_ref, value, style=''):
    sa = f' s="{style}"' if style else ''
    return f'<c r="{cell_ref}"{sa}><v>{value}</v></c>'
```

## Validation Checklist
- Chart 1 percentages display as `xx.xx%`, not decimals.
- Chart 1 counts use active O/Q counts, not completed-only counts.
- Top 3 + OTHER appears consistently in charts and tables.
- Customer names are normalized before grouping.
- Chart XML cache and embedded workbook values match.


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
