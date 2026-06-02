# -*- coding: utf-8 -*-
"""
sdl_dwr_step1_generator.py  —  v1.1.0  (bar-fix edition)
Generate Step 1 SDL Daily Work Report DOCX from the daily Excel input file.

Changes from v1.0.0
-------------------
* Removed paragraph-anchored floating "Rectangle 3" bars from document.xml
  (they drifted and disappeared on the last page).
* Injected a single persistent full-page purple sidebar into header1.xml so
  Word repeats it on every page, top-to-bottom, page 1 to last.
* ZIP output now preserves original per-item compress_type to prevent Word corruption.

Required packages : openpyxl, lxml
Expected template  : Daily Work Report Step 1 Template - enTop v1.0.0.docx
Public function    : generate_dwr_step1(excel_path, template_path=None, output_dir=None) -> str
"""
from __future__ import annotations
import html, os, re, shutil, zipfile
from copy import deepcopy
from datetime import date, datetime
from typing import Any, Dict, List
from lxml import etree
from openpyxl import load_workbook

BASE_DIR           = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE   = os.path.join(BASE_DIR, "Daily Work Report Step 1 Template - enTop v1.0.0.docx")
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "generated_reports")

STATUS_LABELS = {"IP": "In Progress", "NA": "Not attempted", "GU": "Given up"}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS   = {"w": W_NS}

# ---------------------------------------------------------------------------
# Purple sidebar — replaces the template's paragraph-anchored Rectangle 3 bars.
#   * Lives in header1.xml  → Word repeats it on every page automatically.
#   * positionH relativeFrom="page", posOffset=7620 EMU  → left edge, every page.
#   * positionV relativeFrom="page", posOffset=0          → top of page.
#   * cy = 10 692 130 EMU  (full A4 height)               → top-to-bottom bar.
#   * cx = 228 600 EMU                                    → same width as original.
#   * behindDoc="1"                                       → behind all content.
#   * accent1 + lumMod 50000                              → same dark purple.
# ---------------------------------------------------------------------------
_SIDEBAR_XML = (
    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:rPr><w:noProof/></w:rPr>'
    '<w:drawing>'
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
    'relativeHeight="251658240" behindDoc="1" locked="1" layoutInCell="1" allowOverlap="0" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
    '<wp:simplePos x="0" y="0"/>'
    '<wp:positionH relativeFrom="page"><wp:posOffset>7620</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    '<wp:extent cx="228600" cy="10692130"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:wrapNone/>'
    '<wp:docPr id="200000001" name="SidebarBar"/>'
    '<wp:cNvGraphicFramePr/>'
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    '<wps:cNvSpPr/>'
    '<wps:spPr>'
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="228600" cy="10692130"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'
    '<a:ln><a:noFill/></a:ln>'
    '</wps:spPr>'
    '<wps:style>'
    '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
    '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
    '</wps:style>'
    '<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="ctr" anchorCtr="0">'
    '<a:noAutofit/>'
    '</wps:bodyPr>'
    '</wps:wsp>'
    '</a:graphicData>'
    '</a:graphic>'
    '</wp:anchor>'
    '</w:drawing>'
    '</w:r>'
)


# ── Type helpers ──────────────────────────────────────────────────────────────

def _safe_float(v: Any, d: float = 0.0) -> float:
    if v is None or v == "": return d
    try: return float(v)
    except (TypeError, ValueError): return d

def _safe_int(v: Any, d: int = 0) -> int:
    if v is None or v == "": return d
    try: return int(round(float(v)))
    except (TypeError, ValueError): return d

def _fmt_int(v):      return str(_safe_int(v))
def _fmt_hours(v):    return f"{_safe_float(v):.2f}"
def _fmt_currency(v): return f"${_safe_float(v):,.2f}"

def _fmt_pct(v):
    n = _safe_float(v)
    return f"{n*100:.2f}%" if abs(n) <= 10 else f"{n:.2f}%"

def _norm_job_no(r):
    if r is None: return ""
    if isinstance(r, (int, float)): return str(int(r))
    return str(r).strip()

def _norm_cust(r):   return " ".join(str(r).split()) if r else ""
def _norm_status(r): return str(r).strip().upper() if r is not None else ""
def _norm_grp(r):    return str(r).strip() if r is not None else ""

def _excel_date(v: Any) -> date:
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date):     return v
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try: return datetime.strptime(str(v)[:10], fmt).date()
        except Exception: pass
    return date.today()


# ── Excel reader ──────────────────────────────────────────────────────────────

def _read_excel_data(excel_path: str) -> Dict[str, Any]:
    wb = load_workbook(excel_path, data_only=True)
    if "Summary" not in wb.sheetnames or "Production_Center" not in wb.sheetnames:
        raise ValueError("Workbook must contain 'Summary' and 'Production_Center' sheets.")
    ws_sum, ws_pc = wb["Summary"], wb["Production_Center"]

    dwr_no      = str(ws_sum["B1"].value).strip()
    report_date = _excel_date(ws_sum["C1"].value)

    jobs: List[Dict[str, Any]] = []
    for r in range(8, min(ws_pc.max_row, 108) + 1):
        jtype  = str(ws_pc.cell(r, 4).value or "").strip().upper()
        job_no = _norm_job_no(ws_pc.cell(r, 5).value)
        if not jtype and not job_no: continue
        grp = _norm_grp(ws_pc.cell(r, 15).value)
        if grp.lower() == "continued": continue
        jobs.append({
            "type":     jtype,
            "job_no":   job_no,
            "customer": _norm_cust(ws_pc.cell(r, 12).value),
            "rm_time":  _safe_float(ws_pc.cell(r, 21).value),
            "status":   _norm_status(ws_pc.cell(r, 25).value),
        })

    done = [j for j in jobs if j["status"] == "C"]
    pend = [j for j in jobs if j["status"] in {"IP", "NA", "GU"}]

    c115 = _safe_int(ws_sum["C115"].value)
    d115 = _safe_int(ws_sum["D115"].value)
    e115 = _safe_int(ws_sum["E115"].value)
    g115 = _safe_float(ws_sum["G115"].value)
    h115 = _safe_float(ws_sum["H115"].value)
    i115 = _safe_float(ws_sum["I115"].value)
    inc_hrs = (h115 + i115) or sum(j["rm_time"] for j in pend if j["status"] == "IP")

    return {
        "dwr_no":           dwr_no,
        "date":             report_date,
        "completed":        done,
        "incomplete":       pend,
        "q_summary":        [ws_sum.cell(13, c).value for c in range(1, 10)],
        "o_summary":        [ws_sum["J13"].value] + [ws_sum.cell(13, c).value for c in range(11, 19)],
        "total_summary":    [ws_sum.cell(19, c).value for c in range(1, 11)],
        "resource_rows":    [[ws_sum.cell(r, c).value for c in range(3, 10)] for r in range(3, 7)],
        "completed_count":  c115 or len(done),
        "completed_hours":  g115 or sum(j["rm_time"] for j in done),
        "incomplete_count": (d115 + e115) or len(pend),
        "incomplete_hours": inc_hrs,
    }


# ── XML helpers ───────────────────────────────────────────────────────────────

def _parse_xml(data: bytes) -> etree._Element:
    return etree.fromstring(data, etree.XMLParser(remove_blank_text=False, recover=False))

def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=False)

def _patch_text(s: str, repls: Dict[str, str]) -> str:
    for k, v in repls.items(): s = s.replace(k, v)
    return s

def _make_run(rpr, text: str, bold: bool) -> etree._Element:
    r = etree.Element(f"{{{W_NS}}}r")
    if rpr is not None:
        nr = deepcopy(rpr)
        for tag in (f"{{{W_NS}}}b", f"{{{W_NS}}}bCs"):
            el = nr.find(tag)
            if el is not None: nr.remove(el)
        if bold:
            nr.insert(0, etree.Element(f"{{{W_NS}}}bCs"))
            nr.insert(0, etree.Element(f"{{{W_NS}}}b"))
        r.append(nr)
    t = etree.SubElement(r, f"{{{W_NS}}}t")
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r

def _footer_label(tc, kw: str, n: int):
    p = tc.find(f"{{{W_NS}}}p") or etree.SubElement(tc, f"{{{W_NS}}}p")
    runs = p.xpath("./w:r", namespaces=NS)
    rp = rb = None
    for run in runs:
        rpr = run.find(f"{{{W_NS}}}rPr")
        if rpr is None: continue
        if rpr.find(f"{{{W_NS}}}b") is not None:
            if rb is None: rb = rpr
        else:
            if rp is None: rp = rpr
    rp = rp or rb; rb = rb or rp
    for run in runs: p.remove(run)
    for text, bold in [("Total ",False),(kw,True),(" PC jobs ",False),(f"{n} Nos.",True)]:
        p.append(_make_run(rb if bold else rp, text, bold))

def _set_tc(tc, text: Any):
    val = "" if text is None else str(text)
    ts  = tc.xpath(".//w:t", namespaces=NS)
    if ts:
        ts[0].text = val
        if val.startswith(" ") or val.endswith(" "):
            ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for t in ts[1:]: t.text = ""
        return
    p = tc.find(f"{{{W_NS}}}p") or etree.SubElement(tc, f"{{{W_NS}}}p")
    r = p.find(f"{{{W_NS}}}r")  or etree.SubElement(p, f"{{{W_NS}}}r")
    t = etree.SubElement(r, f"{{{W_NS}}}t")
    t.text = val
    if val.startswith(" ") or val.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

def _set_row(row, vals):
    cells = row.xpath("./w:tc", namespaces=NS)
    for i, v in enumerate(vals):
        if i < len(cells): _set_tc(cells[i], v)

def _rebuild_table(tbl, data_rows, footer_vals, fkw=None, fn=None):
    rows = tbl.xpath("./w:tr", namespaces=NS)
    if len(rows) < 3: return
    dtmpl = deepcopy(rows[1]); ftmpl = deepcopy(rows[-1])
    for r in rows[1:]: tbl.remove(r)
    for vals in data_rows:
        nr = deepcopy(dtmpl); _set_row(nr, vals); tbl.append(nr)
    fr = deepcopy(ftmpl)
    if fkw and fn is not None:
        cells = fr.xpath("./w:tc", namespaces=NS)
        for i, v in enumerate(footer_vals):
            if i != 2 and i < len(cells): _set_tc(cells[i], v)
        if len(cells) > 2: _footer_label(cells[2], fkw, fn)
    else:
        _set_row(fr, footer_vals)
    tbl.append(fr)

def _drop_blank_para_before(tbl):
    parent = tbl.getparent()
    if parent is None: return
    siblings = list(parent); idx = siblings.index(tbl)
    if idx == 0: return
    prev = siblings[idx - 1]
    try:
        if etree.QName(prev.tag).localname == "p":
            if not any((t.text or "").strip() for t in prev.xpath(".//w:t", namespaces=NS)):
                parent.remove(prev)
    except Exception: pass


# ── Bar-fix functions ─────────────────────────────────────────────────────────

def _strip_rectangle3(doc_xml: str) -> str:
    """Remove all paragraph-anchored Rectangle 3 floating bar runs from document.xml."""
    return re.sub(
        r'<w:r\b[^>]*>(?:<w:rPr>.*?</w:rPr>)?<mc:AlternateContent>'
        r'(?:(?!<mc:AlternateContent>).)*?name="Rectangle 3".*?</mc:AlternateContent></w:r>',
        '', doc_xml, flags=re.DOTALL,
    )

def _inject_sidebar(header_xml: str) -> str:
    """Insert persistent full-page sidebar into first paragraph of header1.xml."""
    return header_xml.replace('</w:p>', _SIDEBAR_XML + '</w:p>', 1)


# ── Core patcher ──────────────────────────────────────────────────────────────

def _patch_document_xml(doc_bytes: bytes, data: Dict[str, Any]) -> bytes:
    raw  = _strip_rectangle3(doc_bytes.decode("utf-8"))
    root = _parse_xml(raw.encode("utf-8"))
    tbls = root.xpath("//w:tbl", namespaces=NS)
    if len(tbls) < 6: raise ValueError("Template must have at least 6 tables.")

    _set_row(tbls[0].xpath("./w:tr", namespaces=NS)[3],
             [f(v) for v,f in zip(data["q_summary"],
              [_fmt_int]*4+[_fmt_hours]*2+[_fmt_int]*2+[_fmt_currency])])
    _set_row(tbls[1].xpath("./w:tr", namespaces=NS)[3],
             [f(v) for v,f in zip(data["o_summary"],
              [_fmt_int]*4+[_fmt_hours]*2+[_fmt_int]*2+[_fmt_currency])])
    _set_row(tbls[2].xpath("./w:tr", namespaces=NS)[3],
             [f(v) for v,f in zip(data["total_summary"],
              [_fmt_int]*6+[_fmt_hours]*3+[_fmt_currency])])

    _rebuild_table(tbls[3],
        [[j["type"],j["job_no"],j["customer"],_fmt_hours(j["rm_time"])] for j in data["completed"]],
        ["","","",_fmt_hours(data["completed_hours"])],
        fkw="completed", fn=data["completed_count"])

    _rebuild_table(tbls[4],
        [[j["type"],j["job_no"],j["customer"],
          _fmt_hours(j["rm_time"]) if j["status"]=="IP" else "-",
          STATUS_LABELS.get(j["status"],j["status"])] for j in data["incomplete"]],
        ["","","",_fmt_hours(data["incomplete_hours"]),"Hours"],
        fkw="incomplete", fn=data["incomplete_count"])
    _drop_blank_para_before(tbls[4])

    res_rows = tbls[5].xpath("./w:tr", namespaces=NS)
    for idx, rv in enumerate(data["resource_rows"], 1):
        if idx < len(res_rows):
            _set_row(res_rows[idx],
                     [f(v) for v,f in zip(rv,
                      [str,_fmt_int,_fmt_int,_fmt_hours,_fmt_hours,_fmt_hours,_fmt_pct])])

    d   = data["date"]; dwr = data["dwr_no"]
    dl  = f"{d.day} {d.strftime('%B %Y')}"
    xml = _patch_text(_serialize_xml(root).decode("utf-8"), {
        "Publish Date": dl, "DWR# 2026-XX": f"DWR# {dwr}",
        "DWR# 2026-90": f"DWR# {dwr}", "21/05/26": d.strftime("%d/%m/%y"),
        "21 May 2026":  dl,
    })
    return xml.encode("utf-8")


def _patch_all_xml_parts(docx_path: str, data: Dict[str, Any]) -> None:
    d     = data["date"]; dwr = data["dwr_no"]
    dl    = f"{d.day} {d.strftime('%B %Y')}"; iso = d.strftime("%Y-%m-%dT00:00:00")
    repls = {
        "Publish Date": dl, "DWR# 2026-XX": f"DWR# {dwr}",
        "DWR# 2026-90": f"DWR# {dwr}", "21/05/26": d.strftime("%d/%m/%y"),
        "21 May 2026":  dl,
    }
    tmp = docx_path + ".tmp"

    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)

            if item.filename == "word/document.xml":
                content = _patch_document_xml(content, data)

            elif item.filename == "word/header1.xml":
                try:
                    content = _inject_sidebar(
                        _patch_text(content.decode("utf-8"), repls)
                    ).encode("utf-8")
                except UnicodeDecodeError:
                    pass

            elif item.filename.endswith(".xml") and (
                item.filename.startswith("word/") or
                item.filename.startswith("docProps/") or
                item.filename.startswith("customXml/")):
                try:
                    t = _patch_text(content.decode("utf-8"), repls)
                    if item.filename == "docProps/core.xml":
                        t = re.sub(r"<cp:contentStatus>.*?</cp:contentStatus>",
                                   f"<cp:contentStatus>{html.escape(dwr)}</cp:contentStatus>", t)
                    if item.filename.startswith("customXml/"):
                        t = re.sub(r"<PublishDate>.*?</PublishDate>",
                                   f"<PublishDate>{iso}</PublishDate>", t)
                    t = re.sub(r'<w:date w:fullDate="[^"]*"',
                               f'<w:date w:fullDate="{iso}Z"', t)
                    content = t.encode("utf-8")
                except UnicodeDecodeError:
                    pass

            # Preserve original compress_type — critical for Word compatibility
            zout.writestr(item, content, compress_type=item.compress_type)

    os.replace(tmp, docx_path)


# ── Public API ────────────────────────────────────────────────────────────────

def generate_dwr_step1(
    excel_path:    str,
    template_path: str | None = None,
    output_dir:    str | None = None,
) -> str:
    """Generate Step 1 DWR DOCX and return the output file path."""
    template_path = template_path or DEFAULT_TEMPLATE
    output_dir    = output_dir    or DEFAULT_OUTPUT_DIR
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")
    os.makedirs(output_dir, exist_ok=True)
    data = _read_excel_data(excel_path)
    d    = data["date"]
    out  = os.path.join(output_dir,
                        f"DWR_{data['dwr_no']}_{d.strftime('%d-%b-%Y')}-Step_1.docx")
    shutil.copyfile(template_path, out)
    _patch_all_xml_parts(out, data)
    return out


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Generate SDL DWR Step 1 DOCX.")
    ap.add_argument("excel_path")
    ap.add_argument("--template",   default=None)
    ap.add_argument("--output-dir", default=None)
    args = ap.parse_args()
    print(generate_dwr_step1(args.excel_path,
                              template_path=args.template,
                              output_dir=args.output_dir))