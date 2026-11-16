#!/usr/bin/env python3
"""Build index.html for the INSII focal contact map.

Usage:
    python3 build.py path/to/workbook.xlsx            # reads sheet "Website"
    python3 build.py path/to/workbook.xlsx --sheet Website

Reads ONLY these columns from the sheet (matched by header name):
    ISO (first column, header may be blank), Country, Lat, Lon, Region,
    Sub Region, INSII Contact, INSII (Institution)
No email addresses or other columns are read or written.

Output: index.html (single file: map library, data and page code inlined).
Requires: Python 3.9+ and openpyxl (pip install openpyxl).
"""
import argparse, json, re, sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required: pip install openpyxl")

ROOT = Path(__file__).resolve().parent
WANTED = {"country": "country", "lat": "lat", "lon": "lon", "region": "region",
          "sub region": "subregion", "insii contact": "contact",
          "insii (institution)": "institution"}
EMPTY = {"", "NA", "N/A", "UNSPECIFIED"}


def clean(v):
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()
    return None if s.upper() in EMPTY else s


def read_rows(xlsx, sheet):
    # data_only=True reads the values Excel last saved; save the workbook in Excel before building.
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        sys.exit(f'Sheet "{sheet}" not found. Sheets: {wb.sheetnames}')
    rows = list(wb[sheet].iter_rows(values_only=True))
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    missing = [k for k in WANTED if k not in header]
    if missing:
        sys.exit(f"Missing columns in header row: {missing}. Found: {header}")
    idx = {WANTED[k]: header.index(k) for k in WANTED}
    out, problems = [], []
    for n, r in enumerate(rows[1:], start=2):
        iso = clean(r[0])
        if not iso or not any(r):
            continue
        rec = {"iso": iso.upper()}
        for key, i in idx.items():
            rec[key] = r[i] if key in ("lat", "lon") else clean(r[i])
        if not isinstance(rec["lat"], (int, float)) or not isinstance(rec["lon"], (int, float)):
            problems.append(f"row {n} ({iso}): lat/lon not numeric - skipped")
            continue
        if not rec["country"]:
            problems.append(f"row {n} ({iso}): no country name - skipped")
            continue
        out.append(rec)
    isos = [o["iso"] for o in out]
    dups = sorted({i for i in isos if isos.count(i) > 1})
    if dups:
        sys.exit(f"Duplicate ISO codes: {dups}")
    return out, problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--sheet", default="Website")
    ap.add_argument("--out", default=str(ROOT / "index.html"))
    a = ap.parse_args()

    data, problems = read_rows(a.xlsx, a.sheet)
    bbox = json.loads((ROOT / "data" / "country_bbox.json").read_text())
    no_bbox = [d["iso"] for d in data if d["iso"] not in bbox]

    blob = json.dumps(data, ensure_ascii=False)
    if re.search(r"[\w.+-]+@[\w-]+\.[A-Za-z]{2,}", blob):
        sys.exit("An email-like string was found in the data. Check the Website sheet; nothing written.")

    v = ROOT / "vendor" / "maplibre-gl-4.7.1"
    html = (ROOT / "src" / "template.html").read_text(encoding="utf-8")
    html = (html.replace("/*MAPLIBRE_CSS*/", (v / "maplibre-gl.css").read_text(encoding="utf-8"))
                .replace("/*MAPLIBRE_JS*/", (v / "maplibre-gl.js").read_text(encoding="utf-8").replace("</script", "<\\/script"))
                .replace("/*DATA*/", blob.replace("</", "<\\/"))
                .replace("/*BBOX*/", json.dumps(bbox, separators=(",", ":"))))
    Path(a.out).write_text(html, encoding="utf-8")

    print(f"Wrote {a.out}: {len(data)} countries")
    missing_inst = [d["country"] for d in data if not d["institution"]]
    missing_cont = [d["country"] for d in data if not d["contact"]]
    if missing_inst: print("No institution:", ", ".join(missing_inst))
    if missing_cont: print("No contact name:", ", ".join(missing_cont))
    if no_bbox: print("No zoom extent (map flies to a fixed view):", ", ".join(no_bbox))
    for p in problems: print("WARNING:", p)


if __name__ == "__main__":
    main()
