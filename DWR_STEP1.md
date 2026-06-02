# SDL DWR Step 1 Instructions — Operational Tables

## Mission
Generate the Step 1 Daily Work Report `.docx` using `Daily Work Report Step 1 Template - enTop v1.0.0.docx`. This step outputs only the operational production section that appears in the Step 1 template: PC job counts, completed tasks, incomplete tasks, and the resource-utilization table.


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


## Step 1 Template Scope
Populate only the content that exists in this template:
1. `PC Job Counts and Duration (Hours)` heading/date line
2. Quotations summary table
3. Orders summary table
4. Total Jobs summary table
5. `Production Center – Completed Tasks`
6. `Production Center – Incomplete Tasks`
7. Resource Utilization table, if present in the Step 1 template
8. Header/footer DWR number, date, and client/contractor placeholders already present in the template

## Required Calculations

### Quotations table
Use Summary cells `A13:I13`:
- Target Qty, Done Qty, WIP Qty, Pending Qty, Actual Hrs, Overdue Hrs, Parts Done, Parts Left, Value NZD

### Orders table
Use Summary cells:
- `J13` = Target Qty
- `K13:R13` = Done Qty, WIP Qty, Pending Qty, Actual Hrs, Overdue Hrs, Parts Done, Parts Left, Value NZD

### Total Jobs table
Use Summary cells `A19:J19`:
- Target Jobs, Jobs Done, Jobs WIP, Jobs Left, Parts Done, Parts Left, Estimated Hrs, Actual Hrs, Pending Hrs, Value NZD

### Completed Tasks table
Rows come from `completed` jobs only:
- Type
- Quote/Order No.
- Customer
- RM Time, formatted to `0.00`

Footer:
- `Total completed PC jobs [count] Nos.`
- total completed hours from Summary `G115`, formatted to `0.00`

### Incomplete Tasks table
Rows come from `incomplete` jobs only:
- Type
- Quote/Order No.
- Customer
- Time Hr: IP hours as `0.00`; NA/GU should display `-`
- Status display mapping:
  - `IP` → `In Progress`
  - `NA` → `Not attempted`
  - `GU` → `Given up`

Footer:
- `Total incomplete PC jobs [IP + NA/GU count] Nos.`
- incomplete hours from Summary `H115 + I115`, formatted to `0.00 Hours`

### Resource Utilization table
Use Summary `C3:I6`, four rows:
1. Senior members' digital processing service
2. Junior members' digital processing service
3. TOTAL
4. SDL Estimation

Keep all table border, shading, font, and highlight formatting from the template. Load % must display as `xx.xx%`, not a decimal.

## Document XML Update Order
1. Global date and DWR number replacements.
2. Replace the report heading/date line.
3. Patch Quotations table using position-based cells.
4. Patch Orders table using position-based cells.
5. Patch Total Jobs table using position-based cells.
6. Replace Completed Tasks rows and footer.
7. Replace Incomplete Tasks rows and footer.
8. Patch Resource Utilization table if present.
9. Patch header/footer date and DWR number.

## Validation Checklist
- No `Grp Work = Continued` rows appear in Completed or Incomplete tables.
- Completed count and hours match Summary `C115` and `G115`.
- Incomplete count and hours match Summary `D115 + E115` and `H115 + I115`.
- Orders target uses `J13`, not `K13`.
- Currency, hours, and percentages display with correct formatting.


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
