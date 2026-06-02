# SDL DWR Step 4 Instructions — Key Observations

## Mission
Generate the Step 4 Daily Work Report `.docx` using `Daily Work Report Step 4 Template - enTop v1.0.0.docx`. This step outputs only the `Key Observations` section, using validated row-level analytics from the daily Excel file.


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


## Step 4 Template Scope
Populate only the `Key Observations` section found in the Step 4 template. Preserve the exact heading style, bullet style, spacing, footer, and page layout.

## Required Active Job Sets
```python
active_jobs = [
    j for j in jobs
    if j['status'] in {'C', 'IP'}
    and j['grp_work'] != 'Continued'
    and j['rm_time'] > 0
    and j['value'] > 0
]
active_O = [j for j in active_jobs if j['type'] == 'O']
active_Q = [j for j in active_jobs if j['type'] == 'Q']
```

## Four Required Bullets

### Bullet 1 — Revenue leader
Find the top customer by total active value across all active jobs.
Report:
- customer name
- share of total active value
- order/quote count as appropriate
- total value
- best job for that customer by efficiency `$ / hr`, selected from that customer’s individual active job rows

If the top customer represents at least 85% of total active value, use a `Revenue Anchor` style warning. Otherwise, state the share and top job efficiency normally.

### Bullet 2 — Quote of the Day
Select the highest-efficiency active quote job from `active_Q` only.
```python
best_quote_eff = max(active_Q, key=lambda j: j['value'] / j['rm_time']) if active_Q else None
```
Report:
- `Q [job_no]`
- customer
- value
- RM time
- `$ / hr`
- conversion/watch note

### Bullet 3 — Time-Heavy Orders
Use active order jobs only. Identify the order/customer that consumes high RM time relative to its value. Prioritize:
1. IP/WIP order risk if any active order is still `IP`
2. otherwise longest active order RM time
3. include low $/hr wording only when value/time supports it

Report:
- customer or job number
- order count if customer-level
- total RM hours
- total value
- relevant high-time job(s)

### Bullet 4 — Capacity & Backlog
Use Summary and active/incomplete sets:
- Actual hours vs 40-hour baseline
- Load % = actual hours / 40 × 100
- NA count / pending order count
- IP count
- deferred revenue or backlog risk statement

## Critical Safeguard: Best-Job Row-Level Consistency
Never mix job number, value, RM time, customer, or type from different rows.

Required row-level objects:
```python
best_order_eff = max(active_O, key=lambda j: j['value'] / j['rm_time']) if active_O else None
best_quote_eff = max(active_Q, key=lambda j: j['value'] / j['rm_time']) if active_Q else None
```

Validation rules:
- If the sentence says `Orders`, `order`, or starts with `O`, the job number, customer, value, RM time, and `$ / hr` must all come from the same `best_order_eff` row or from the same selected active order row.
- If the sentence says `Quote`, `quotation`, or starts with `Q`, the job number, customer, value, RM time, and `$ / hr` must all come from the same `best_quote_eff` row or from the same selected active quote row.
- Do not infer best jobs from customer leaderboard aggregates.
- Do not combine the job number from one row with value/time from another row.
- Do not reuse order variables in quote sentences or quote variables in order sentences.

Mandatory assertions before writing final text:
```python
if best_order_eff:
    assert best_order_eff['type'] == 'O'
    assert round(best_order_eff['value'] / best_order_eff['rm_time'], 2) == round(best_order_eff['rate'], 2) if 'rate' in best_order_eff else True

if best_quote_eff:
    assert best_quote_eff['type'] == 'Q'
    assert round(best_quote_eff['value'] / best_quote_eff['rm_time'], 2) == round(best_quote_eff['rate'], 2) if 'rate' in best_quote_eff else True
```

Recommended safer reporting pattern:
```python
def efficiency_text(j):
    rate = j['value'] / j['rm_time'] if j['rm_time'] else 0
    return f"{j['type']} {j['job_no']} (${j['value']:,.0f}/{j['rm_time']:.2f}h = ${rate:,.0f}/hr)"
```
Always build the full efficiency phrase from the same `j` object.

## Tone and Style
- Professional, executive-level, concise.
- Use 4 bullets only.
- Keep wording close to the existing template’s style.
- Percentages: 1 decimal place inside narrative is acceptable where the existing template uses it, but table/chart percentages elsewhere remain 2 decimals.
- Currency in narrative: usually rounded to nearest dollar unless the template requires cents.
- Hours: 2 decimals where shown in a formula; otherwise concise wording is acceptable.

## Validation Checklist
- Four bullets are present.
- Quote of the Day comes from `active_Q` only.
- Order efficiency or time-heavy statements come from `active_O` only.
- Every job number/value/time/rate phrase is built from one row object.
- Capacity bullet matches Summary actual hours and 40-hour baseline.
- No `Continued` rows are used.


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
