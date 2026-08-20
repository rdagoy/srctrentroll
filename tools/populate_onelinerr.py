#!/usr/bin/env python3
"""
populate_onelinerr.py — DIRECT CAP UW Model: raw rent roll -> OneLineRR

Maps any property-management rent-roll export into the OneLineRR tab of the
DIRECT CAP UW Model template (templates/DIRECT_CAP_UW_Model_TEMPLATE.xlsx),
using a surgical XML write so the workbook's charts, drawings, and data
validations are preserved byte-for-byte (openpyxl would drop them).

Design rules (SRCT "Multifamily Rent Roll" process):
  - Never invent data. Every written value comes from the source file.
  - The raw unit-type string is preserved exactly in OneLineRR col K.
  - Control totals are reconciled against the source before the file ships.
  - Market Rent / Subsidy / Move-in / Move-out are left blank unless the
    source provides them AND the analyst maps them (see CONFIG).

Per-deal setup (edit CONFIG below, or import and pass a config dict):
  1. SOURCE_SHEET / HEADER_ROW / column map  -> where each field lives.
  2. FLOOR_PLANS                              -> the floor-plan lookup table
     (rows 5-61 of OneLineRR): label, raw unit type (join key), BD, BA,
     Affordable, Renovated, Reno Type. This is the human gate — confirm
     BD/BA and Renovated per deal. If left None, the table is auto-derived
     from the distinct source unit types with BD/BA parsed from BD_BA_COL
     and Renovated defaulted to "No" (flagged for review).
  3. PROPERTY / ADDRESS / CITY / ASOF         -> header block (C67:C70).

Usage:
    python tools/populate_onelinerr.py SOURCE.xlsx --out "Deal - UW Model.xlsx"

The mapping below reproduces the Hudson View Park (as-of 10/01/2025) deal.
"""
import argparse, re, os, sys
from datetime import date, datetime

# ----------------------------------------------------------------------------
# CONFIG — adjust per deal. Column keys are OneLineRR fields; values are the
# HEADER TEXT to match in the source (case-insensitive, trimmed). Set a value
# to None to leave that OneLineRR column blank.
# ----------------------------------------------------------------------------
CONFIG = dict(
    SOURCE_SHEET=None,      # None -> first sheet; else sheet name
    HEADER_ROW=None,        # None -> auto-detect the header row; else 1-based row
    COLMAP={
        "unit":     "Unit",             # -> J  Unit #
        "unittype": "Unit Type",        # -> K  Unit Type (verbatim; join key)
        "sqft":     "Sqft",             # -> L  Unit Sq Ft
        "tenant":   "Tenant",           # -> M  Resident/Tenant Name
        "market":   None,               # -> N  Market Rent (left blank by default)
        "rent":     "Rent",             # -> AA Rent (feeds O Contract Rent)
        "other":    "Monthly Charges",  # -> AD Other Income
        "lease_from": "Lease From",     # -> Q  Lease Start
        "lease_to":   "Lease To",       # -> R  Lease End
        "movein":   None,               # -> P  Move in
        "moveout":  None,               # -> S  Move Out
    },
    BD_BA_COL="BD/BA",       # source col used to parse BD/BA when auto-deriving floor plans
    VACANT_TOKENS=("Vacant",),   # tenant values that mark a unit vacant
    FLOOR_PLANS=None,        # None -> auto-derive; else list of dicts (see build_floor_plans)
    PROPERTY="Hudson View Park",
    ADDRESS="29 Hudson View Drive",
    CITY="Beacon, NY 12508",
    ASOF=date(2025, 10, 1),
)

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "templates",
                        "DIRECT_CAP_UW_Model_TEMPLATE.xlsx")
ONELINE_SHEET_XML = "xl/worksheets/sheet10.xml"
WORKBOOK_XML = "xl/workbook.xml"
DATA_ROW0 = 99          # first OneLineRR data row
DATA_ROWMAX = 606       # last scaffolded OneLineRR data row (508-unit capacity)
FP_ROW0, FP_ROWMAX = 5, 61   # floor-plan lookup table rows
EPOCH = date(1899, 12, 30)

# --- xlsx helpers -----------------------------------------------------------
def _serial(d):
    if isinstance(d, datetime): d = d.date()
    return (d - EPOCH).days

def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _set(xml, ref, inner, ttype=None):
    pat = re.compile(r'<c r="%s"(?P<a>[^>]*?)(?:/>|>.*?</c>)' % re.escape(ref), re.S)
    def f(m):
        s = re.search(r' s="\d+"', m.group("a"))
        t = ' t="%s"' % ttype if ttype else ''
        return '<c r="%s"%s%s>%s</c>' % (ref, s.group(0) if s else '', t, inner)
    xml, n = pat.subn(f, xml, count=1)
    if n != 1:
        raise RuntimeError("cell %s not found/edited (n=%d)" % (ref, n))
    return xml

def set_text(xml, ref, text):
    return _set(xml, ref, '<is><t xml:space="preserve">%s</t></is>' % _esc(text), "inlineStr")
def set_num(xml, ref, val):
    return _set(xml, ref, '<v>%s</v>' % val)
def set_date(xml, ref, d):
    return _set(xml, ref, '<v>%d</v>' % _serial(d))
def set_formula(xml, ref, formula):
    # writes a plain (non-shared) formula, clearing any cached value
    return _set(xml, ref, '<f>%s</f>' % _esc(formula), "str")

# --- source parsing ---------------------------------------------------------
def parse_source(path, cfg):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[cfg["SOURCE_SHEET"]] if cfg["SOURCE_SHEET"] else wb[wb.sheetnames[0]]
    wanted = {v.strip().lower() for v in cfg["COLMAP"].values() if v}
    wanted |= {cfg["BD_BA_COL"].strip().lower()} if cfg.get("BD_BA_COL") else set()
    # locate header row
    hr = cfg["HEADER_ROW"]
    if hr is None:
        for r in range(1, min(ws.max_row, 60) + 1):
            vals = {str(ws.cell(r, c).value).strip().lower()
                    for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None}
            if wanted and wanted.issubset(vals):
                hr = r; break
        if hr is None:
            raise RuntimeError("Could not locate a header row containing %s" % sorted(wanted))
    hdr = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hr, c).value
        if v is not None:
            hdr[str(v).strip().lower()] = c
    def col(name):
        return hdr.get(name.strip().lower()) if name else None
    idx = {k: col(v) for k, v in cfg["COLMAP"].items()}
    bdba = col(cfg.get("BD_BA_COL"))
    missing = [cfg["COLMAP"][k] for k, ci in idx.items() if cfg["COLMAP"][k] and ci is None]
    if missing:
        raise RuntimeError("Source is missing expected columns: %s" % missing)
    units = []
    for r in range(hr + 1, ws.max_row + 1):
        u = ws.cell(r, idx["unit"]).value if idx["unit"] else None
        if u is None: continue
        s = str(u).strip()
        sl = s.lower()
        # skip separators / totals / footers
        is_footer = (sl == "" or sl in ("total", "occupied", "vacant", "total units")
                     or sl.startswith(("total", "subtotal")) or sl.endswith("units"))
        is_separator = (idx["unittype"] and ws.cell(r, idx["unittype"]).value is None
                        and idx["sqft"] and ws.cell(r, idx["sqft"]).value is None)
        if is_footer or is_separator:
            continue
        rec = {k: (ws.cell(r, ci).value if ci else None) for k, ci in idx.items()}
        rec["unit"] = s
        rec["_bdba"] = ws.cell(r, bdba).value if bdba else None
        units.append(rec)
    return units

def build_floor_plans(units, cfg):
    if cfg["FLOOR_PLANS"]:
        return cfg["FLOOR_PLANS"]
    seen, fps = set(), []
    for u in units:
        ut = u["unittype"]
        if ut in (None, "") or ut in seen: continue
        seen.add(ut)
        bd = ba = None
        if u.get("_bdba") and "/" in str(u["_bdba"]):
            a, b = str(u["_bdba"]).split("/", 1)
            try: bd = float(a)
            except: pass
            try: ba = float(b)
            except: pass
        fps.append(dict(label=str(ut), unittype=str(ut), bd=bd, ba=ba,
                        affordable="No", renovated="No", reno_type="Classic"))
    return fps

# --- write ------------------------------------------------------------------
def populate(source, out, cfg=CONFIG, template=TEMPLATE):
    import zipfile
    units = parse_source(source, cfg)
    fps = build_floor_plans(units, cfg)
    if len(units) > (DATA_ROWMAX - DATA_ROW0 + 1):
        raise RuntimeError("Source has %d units; template scaffold holds %d. Extend the "
                           "OneLineRR data rows first." % (len(units), DATA_ROWMAX - DATA_ROW0 + 1))
    if len(fps) > (FP_ROWMAX - FP_ROW0 + 1):
        raise RuntimeError("Deal has %d floor plans; table holds %d rows." %
                           (len(fps), FP_ROWMAX - FP_ROW0 + 1))
    z = zipfile.ZipFile(template)
    s = z.read(ONELINE_SHEET_XML).decode("utf-8")
    wbx = z.read(WORKBOOK_XML).decode("utf-8")

    # property header block (C67 name, C68 address, C69 city/state/zip, C70 as-of)
    s = set_text(s, "C67", cfg["PROPERTY"])
    if cfg.get("ADDRESS"): s = set_text(s, "C68", cfg["ADDRESS"])
    if cfg.get("CITY"):    s = set_text(s, "C69", cfg["CITY"])
    if cfg.get("ASOF"):    s = set_date(s, "C70", cfg["ASOF"])

    # floor-plan lookup table (rows 5..)
    for i, fp in enumerate(fps):
        r = FP_ROW0 + i
        # link the summary property-name column to the header cell C67
        s = set_formula(s, "B%d" % r, "$C$67")
        s = set_text(s, "C%d" % r, fp["label"])
        s = set_text(s, "D%d" % r, fp["unittype"])
        if fp.get("bd") is not None: s = set_num(s, "E%d" % r, fp["bd"])
        if fp.get("ba") is not None: s = set_num(s, "F%d" % r, fp["ba"])
        s = set_text(s, "G%d" % r, fp.get("affordable", "No"))
        s = set_text(s, "H%d" % r, fp.get("renovated", "No"))
        s = set_text(s, "I%d" % r, fp.get("reno_type", "Classic"))

    # data rows (99..)
    vac = set(t.strip().lower() for t in cfg["VACANT_TOKENS"])
    for i, u in enumerate(units):
        r = DATA_ROW0 + i
        s = set_text(s, "J%d" % r, u["unit"])
        s = set_text(s, "K%d" % r, u["unittype"])
        if isinstance(u["sqft"], (int, float)): s = set_num(s, "L%d" % r, u["sqft"])
        if u["tenant"] is not None:            s = set_text(s, "M%d" % r, u["tenant"])
        if cfg["COLMAP"]["market"] and isinstance(u["market"], (int, float)):
            s = set_num(s, "N%d" % r, u["market"])
        if isinstance(u["lease_from"], (date, datetime)): s = set_date(s, "Q%d" % r, u["lease_from"])
        if isinstance(u["lease_to"], (date, datetime)):   s = set_date(s, "R%d" % r, u["lease_to"])
        if cfg["COLMAP"]["movein"] and isinstance(u["movein"], (date, datetime)):
            s = set_date(s, "P%d" % r, u["movein"])
        if cfg["COLMAP"]["moveout"] and isinstance(u["moveout"], (date, datetime)):
            s = set_date(s, "S%d" % r, u["moveout"])
        if isinstance(u["rent"], (int, float)):  s = set_num(s, "AA%d" % r, u["rent"])
        if isinstance(u["other"], (int, float)): s = set_num(s, "AD%d" % r, u["other"])

    wbx = wbx.replace('<calcPr calcId="191028" iterate="1"/>',
                      '<calcPr calcId="191028" iterate="1" fullCalcOnLoad="1"/>')
    sub = {ONELINE_SHEET_XML: s.encode("utf-8"), WORKBOOK_XML: wbx.encode("utf-8")}
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(template) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            zo.writestr(it, sub.get(it.filename, zin.read(it.filename)))

    # reconcile source vs what was written back to the workbook
    return verify_written(out, units, cfg, len(fps))


def verify_written(out, units, cfg, floor_plans=None):
    """Re-open the output and confirm the written OneLineRR inputs match the
    parsed source (counts + column totals). Powers the 'does it match the
    source?' confirmation gate. Returns a report with per-field ok flags."""
    import openpyxl
    ws = openpyxl.load_workbook(out, data_only=True)["OneLineRR"]
    COL = dict(unit=10, unittype=11, sqft=12, tenant=13, rent=27, other=30)  # J,K,L,M,AA,AD
    def src_sum(key):
        return round(sum(u[key] for u in units if isinstance(u[key], (int, float))), 2)
    def wb_sum(c):
        return round(sum(ws.cell(r, c).value for r in range(DATA_ROW0, DATA_ROW0 + len(units))
                         if isinstance(ws.cell(r, c).value, (int, float))), 2)
    vac = set(t.strip().lower() for t in cfg["VACANT_TOKENS"])
    checks = {
        "units":       (len(units), sum(1 for r in range(DATA_ROW0, DATA_ROW0 + len(units))
                                        if ws.cell(r, COL["unit"]).value not in (None, ""))),
        "vacant":      (sum(1 for u in units if str(u["tenant"]).strip().lower() in vac),
                        sum(1 for r in range(DATA_ROW0, DATA_ROW0 + len(units))
                            if str(ws.cell(r, COL["tenant"]).value).strip().lower() in vac)),
        "rent_total":  (src_sum("rent"), wb_sum(COL["rent"])),
        "other_total": (src_sum("other"), wb_sum(COL["other"])),
        "sqft_total":  (src_sum("sqft"), wb_sum(COL["sqft"])),
    }
    ok = {k: (a == b) for k, (a, b) in checks.items()}
    return dict(units=len(units), floor_plans=floor_plans,
                checks=checks, ok=ok, all_ok=all(ok.values()))

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("source", help="raw rent-roll export (.xlsx/.xls)")
    ap.add_argument("--out", required=True, help="output UW Model path")
    ap.add_argument("--template", default=TEMPLATE)
    a = ap.parse_args(argv)
    rep = populate(a.source, a.out, CONFIG, a.template)
    print("Wrote", a.out)
    print("  units       ", rep["units"])
    print("  floor_plans ", rep["floor_plans"])
    print("  %-14s %14s %14s  %s" % ("check", "source", "workbook", "match"))
    for k, (src, wb) in rep["checks"].items():
        fmt = lambda x: "{:,.2f}".format(x) if isinstance(x, float) else str(x)
        print("  %-14s %14s %14s  %s" % (k, fmt(src), fmt(wb),
                                         "OK" if rep["ok"][k] else "*** MISMATCH ***"))
    print("\nRECONCILED — confirm every 'source' equals 'workbook' above before hand-off."
          if rep["all_ok"] else "\n*** MISMATCH — do not deliver until resolved. ***")
    return 0 if rep["all_ok"] else 1

if __name__ == "__main__":
    sys.exit(main())
