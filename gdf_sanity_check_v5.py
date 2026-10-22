#!/usr/bin/env python3
"""
GDF Delivery Sanity Checker - v5
================================
What is new vs v4
  * LONG-format CASEDATA supported (columns: Respondent.Serial, Top Level Field Name,
    First..Fourth Level Iterator Name/Label/Field Name, Response Name, Week number).
    v4 assumed one column per variable, which produced false "missing variable / Rounds" FAILs.
  * CASEDATA is streamed in chunks (multi-GB UTF-16 files are fine).
  * A cell delivered in several period files (e.g. IN HSKF 202301_202413 / 202414_202513 /
    202514_202639) is merged into ONE cell; gaps / overlaps between periods are checked.
  * Same cell + same period delivered twice (different delivery date) -> only the latest is used.
  * Excel row-limit truncation (1,048,575 / 1,048,576 data rows) is flagged.
  * --expected-cells : flags cells that were expected but not delivered (and the reverse).
  * Labels compared per CODE (same code must keep the same label everywhere). Codes that only
    exist in some markets/categories are INFO, not FAIL. Two codes with the same label inside
    one variable, and blank labels, are flagged.
  * Brand / statement labels used inside CASEDATA (iterator labels) checked AS IS against METADATA.
  * T3B is optional (INFO); T2B and Mean are required in check tables.

Usage
  python gdf_sanity_check_v5.py --input <folder_of_cell_zips> --output report.xlsx
        [--expected-cells "ZAHHON,INHSKF,IDHCLF,JPHSOF,BRHSOF,FRFMSN,CNHSOM,THHDHN,ZAHDOM,USFSMB,USFBSN"]
        [--baseline previous_snapshot.json] [--brand-master brands.csv]
  python gdf_sanity_check_v5.py --input <folder_of_METADATA_csvs> --output report.xlsx --metadata-only

Requires: pandas, numpy, openpyxl
"""
import argparse, csv, difflib, io, json, os, re, sys, unicodedata, zipfile
from collections import defaultdict
from datetime import datetime
import numpy as np
import pandas as pd
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# =====================================================================================
# CONFIGURATION
# =====================================================================================
WEIGHT_VARS = ["WEIGHT", "Synthetic", "Weight_human", "human_weights_squared", "boost_factor",
               "weightage", "ESSBase"]
TIME_VARS = ["Rounds", "Yearly", "Quarterly", "Roll4", "Roll8", "Roll12", "Roll24", "Roll26",
             "Roll52", "Disc1", "Disc4", "Disc8", "Disc12"]
FILTER_VARS = ["Country", "Category", "QSex", "MRKAgeBands", "Age", "AgeGroup", "Generation_GDF",
               "HHOLD", "WS", "Chief", "HijabWearer", "TenureCat", "panelSource", "SampleQuota",
               "SegDer", "LFBGSCell", "HStudyType", "QDevice", "BD1", "BD2", "BD1_2", "W1", "W2",
               "S1", "Y1", "MRKupos1Base", "SOCMED1_CORE", "SOCMED2_CORE"]
# allowed to be missing in some cells (category- or market-specific) -> INFO only
OPTIONAL_VARS = ["Ethnicity", "Q10", "HPC2", "HPC3", "HPC8", "HPC9", "HPC12", "HPC14", "HPC16",
                 "HPC18a", "HPC19", "HPC21b", "HPC22", "HPC22Deo", "QDistrict", "qCID", "qCITY",
                 "classeHFCT", "Weight_synthetic"]
# answer lists that legitimately differ by market: only "same code -> same label" is checked
MARKET_SPECIFIC_VARS = ["Ethnicity", "QDistrict", "qCID", "qCITY", "classeHFCT", "Q10",
                        "SampleQuota", "SOCMED1_CORE", "SOCMED2_CORE", "COM19Flip", "AgeGroup"]
BRAND_ITER_METRICS = ["BE6Total", "BE7Total", "BP6", "COM19Flip", "ACT_V1_AVAIL",
                      "BP1_Perceived_Price_EU", "BP6_Worth_Price_EU", "UPOS2_HeardOf"]
BRAND_ITER2_METRICS = ["BP1"]
BRAND_RESPONSE_METRICS = ["Q7aYes", "upos1", "upos2", "SLPAT", "Pack1Core", "QSLP", "UsageTab",
                          "BE6_T2B", "BE7Total_T2B", "BP1_SameVsDifferent_T2B", "BP1_Needs_T2B",
                          "BP1_FollowVsTrends_T2B", "BP1_Perceived_Price_T2B",
                          "BP6_Worth_Price_TB", "ACT_V1_AVAIL_T2B", "BHCAsked"]
STATEMENT_ITER_METRICS = ["upos1", "upos2", "SLPAT", "Pack1Core", "W3"]
STATEMENT_RESPONSE_METRICS = ["UPOS2_HeardOf"]
ITEM_LIST_METRICS = ["BP1", "UsageTab", "MED30", "MED31", "med1", "GENMEDIA"]
DERIVED_SOURCES = {"BE6_T2B": "BE6Total", "BE7Total_T2B": "BE7Total", "BP6_Worth_Price_TB": "BP6",
                   "ACT_V1_AVAIL_T2B": "ACT_V1_AVAIL", "BP6_Worth_Price_EU": "BP6"}
SCALE_VARS = ["BE6Total", "BE7Total", "BP1", "BP1_Perceived_Price_EU", "ACT_V1_AVAIL", "W3"]
REQUIRED_CHECK_STATS = ["T2B", "Mean"]          # T3B reported as INFO only
ADMIN_VARS = ["sampsrc", "ttimes", "Timer", "HPC1", "HPC1_NR", "F1", "F1_NR", "ubx10selected",
              "UBX10", "UBX11", "ube5v"]
WAVE_VAR = "Rounds"
WEEK_COL = "Week number"
SERIAL_COL = "Respondent.Serial"
CHUNK_ROWS = 1_000_000
EXCEL_ROW_LIMITS = {1048575, 1048576}

YES_LABELS = {"yes", "y", "selected", "ticked", "true"}
NO_LABELS = {"no", "n", "notselected", "notticked", "false"}
BRAND_CODE = re.compile(r"^_\d+$")
FUZZY_THRESHOLD = 0.92
EXPECTED_FILES_PER_ZIP = 5
REQUIRED_FILE_TYPES = ["CASEDATA", "METADATA", "EXISTENCE"]
NAME_PATTERN = re.compile(r"^[A-Za-z0-9]{2}[A-Za-z]{2}[A-Za-z0-9]{4}$")
PERIOD_RX = re.compile(r"_(\d{6})_(\d{6})(?:_(\d{8}))?", re.I)
ALLOWED_ENCODINGS = {"utf-8", "utf-8-sig", "utf-16"}
CHECK_TOKENS = {"T2B": re.compile(r"\bT2B\b|top\s*2\s*box", re.I),
                "T3B": re.compile(r"\bT3B\b|top\s*3\s*box", re.I),
                "Mean": re.compile(r"\bmean\b|\baverage\b", re.I)}
TOP = "Top Level Field Name"
LEVELS = [("Iter1", "First Level Iterator Name", "First Level Iterator Label"),
          ("Iter2", "Second Level Iterator Name", "Second Level Iterator Label"),
          ("Iter3", "Third Level Iterator Name", "Third Level Iterator Label"),
          ("Iter4", "Fourth Level Iterator Name", "Fourth Level Iterator Label")]
RESP = ("Response Name", "Response Label")

BRAND_VARS_ALL = BRAND_ITER_METRICS + BRAND_ITER2_METRICS + BRAND_RESPONSE_METRICS
CORE_VARS = list(dict.fromkeys(WEIGHT_VARS + TIME_VARS + FILTER_VARS + BRAND_VARS_ALL
                               + STATEMENT_ITER_METRICS + STATEMENT_RESPONSE_METRICS + ITEM_LIST_METRICS))
# =====================================================================================

FINDINGS = []


def add(check, cell, sev, detail, variable=""):
    FINDINGS.append({"Check": check, "Severity": sev, "Cell": cell, "Variable": variable, "Detail": detail})


def key(label):
    s = unicodedata.normalize("NFKD", str(label)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9+\-]", "", s.replace("&", "and"))


def short(items, n=6):
    c = sorted(set(map(str, items)))
    return ", ".join(c[:n]) + (f" +{len(c) - n} more" if len(c) > n else "")


def show(label):
    return f"'{label}'" + (" [extra spaces]" if label != label.strip() or "  " in label else "")


# ------------------------------------------------------------------ file helpers
def detect_encoding(raw):
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "utf-16"
    if raw[:3] == b"\xef\xbb\xbf":
        return "utf-8-sig"
    try:
        raw.decode("utf-8"); return "utf-8"
    except UnicodeDecodeError as e:
        return "utf-8" if e.start > len(raw) - 4 else "non-UTF"


def detect_delimiter(first_line):
    counts = {d: first_line.count(d) for d in [",", ";", "\t", "|"]}
    return max(counts, key=counts.get) if any(counts.values()) else ","


def classify(name):
    up = os.path.basename(name).upper()
    return next((t for t in REQUIRED_FILE_TYPES if t in up), "CHECK")


def open_text(opener):
    """opener() -> binary stream. Returns (text stream, encoding, delimiter)."""
    with opener() as b:
        head = b.read(65536)
    enc = detect_encoding(head)
    py_enc = enc if enc in ALLOWED_ENCODINGS else "latin-1"
    first = head.decode(py_enc, errors="ignore").lstrip("\ufeff").splitlines()[0] if head else ""
    return io.TextIOWrapper(opener(), encoding=py_enc, errors="replace", newline=""), enc, detect_delimiter(first)


def find_col(cols, name):
    return next((c for c in cols if c.strip().lower() == name.lower()), None)


# ------------------------------------------------------------------ discovery
def discover(folder, metadata_only):
    """Return {cell: [part, ...]} where part = dict(period, date, src, members{type: [opener,name,size]})."""
    files, parts = [], defaultdict(dict)
    paths = sorted(os.path.join(r, f) for r, _, fs in os.walk(folder) for f in fs)
    entries = []                       # (cell, period, date, type, name, size, opener, container)
    if metadata_only:
        for p in paths:
            if "METADATA" in os.path.basename(p).upper():
                entries.append((p, os.path.basename(p), os.path.getsize(p), (lambda p=p: open(p, "rb")), "folder"))
    else:
        for p in [p for p in paths if p.lower().endswith(".zip")]:
            zname = os.path.basename(p)
            try:
                zf = zipfile.ZipFile(p)
            except zipfile.BadZipFile:
                add("F. File quality", zname[:8], "FAIL", f"{zname} is corrupt"); continue
            mem = [m for m in zf.infolist() if not m.is_dir() and not os.path.basename(m.filename).startswith(".")]
            types = defaultdict(list)
            for m in mem:
                types[classify(m.filename)].append(m.filename)
            if len(mem) != EXPECTED_FILES_PER_ZIP:
                add("B. Delivery structure", zname[:8].upper(), "FAIL",
                    f"{zname}: {len(mem)} files, expected {EXPECTED_FILES_PER_ZIP} "
                    f"(CHECK files: {[os.path.basename(x) for x in types['CHECK']]})")
            for t in REQUIRED_FILE_TYPES:
                if not types[t]:
                    add("B. Delivery structure", zname[:8].upper(), "FAIL", f"{zname}: {t} file missing")
            if not types["CHECK"]:
                add("B. Delivery structure", zname[:8].upper(), "FAIL", f"{zname}: no check files/tables")
            for m in mem:
                entries.append((m.filename, os.path.basename(m.filename), m.file_size,
                                (lambda zf=zf, n=m.filename: zf.open(n)), zname))
    for path, base, size, opener, container in entries:
        cell = base[:8].upper()
        if not NAME_PATTERN.match(cell):
            add("B. File naming", cell, "FAIL", f"{base}: not Study(2)+Market(2)+Category(4)")
        m = PERIOD_RX.search(base)
        period = (m[1], m[2]) if m else ("?", "?")
        date = (m[3] or "") if m else ""
        ftype = classify(base)
        files.append({"Cell": cell, "Container": container, "File": base, "Type": ftype,
                      "Period": f"{period[0]}_{period[1]}", "DeliveryDate": date, "SizeBytes": size})
        if size == 0:
            add("F. File quality", cell, "FAIL", f"{base} is empty (0 bytes)")
        slot = parts[cell].setdefault(period, defaultdict(list))
        slot[container if container != "folder" else date].append((ftype, base, size, opener, date))
    cells = {}
    for cell, per in parts.items():
        cells[cell] = []
        for period, versions in sorted(per.items()):
            vdate = {v: max((x[4] for x in fl), default="") for v, fl in versions.items()}
            ordered = sorted(versions, key=lambda v: (vdate[v], v))
            if len(ordered) > 1:
                add("B. Duplicate delivery", cell, "WARN",
                    f"Period {period[0]}_{period[1]} delivered {len(ordered)} times "
                    f"({', '.join(f'{v} [{vdate[v]}]' for v in ordered)}) - only the latest ({ordered[-1]}) is checked")
            latest = ordered[-1]
            members = defaultdict(list)
            for ftype, base, size, opener, d in versions[latest]:
                members[ftype].append((d, base, size, opener))
            cells[cell].append({"period": period, "date": vdate[latest], "version": latest, "members": members})
        check_periods(cell, [p["period"] for p in cells[cell]])
    for f in files:
        f["Used"] = "Yes" if any(f["Container"] == p["version"] or (f["Container"] == "folder" and f["DeliveryDate"] == p["version"])
                                 for p in cells.get(f["Cell"], []) if f["Period"] == "_".join(p["period"])) else "No (superseded)"
    return cells, files


def week_next(yw):
    y, w = int(yw[:4]), int(yw[4:])
    return f"{y + 1}01" if w >= 52 else f"{y}{w + 1:02d}"


def check_periods(cell, periods):
    ps = sorted(p for p in periods if p[0] != "?")
    if len(ps) > 1:
        add("B. Split delivery", cell, "INFO", f"Cell delivered in {len(ps)} period files: "
            + ", ".join(f"{a}_{b}" for a, b in ps) + " - merged and checked as one cell")
    for (a1, b1), (a2, b2) in zip(ps, ps[1:]):
        if a2 <= b1:
            add("B. Split delivery", cell, "FAIL", f"Periods overlap: {a1}_{b1} and {a2}_{b2}")
        elif a2 != week_next(b1) and not (int(b1[4:]) == 53 and a2 == f"{int(b1[:4]) + 1}01"):
            add("B. Split delivery", cell, "WARN", f"Gap between periods: {a1}_{b1} then {a2}_{b2}")


# ------------------------------------------------------------------ metadata
def read_small(member, cell, files_rec):
    d, base, size, opener = member
    txt, enc, delim = open_text(opener)
    files_rec.update(Encoding=enc, Delimiter=repr(delim))
    if enc not in ALLOWED_ENCODINGS:
        add("F. Encoding", cell, "FAIL", f"{base}: {enc} (expected UTF-8/UTF-16)")
    if delim != ",":
        add("F. Delimiter", cell, "FAIL", f"{base}: delimiter {delim!r} (expected comma)")
    with txt:
        df = pd.read_csv(txt, sep=delim, dtype=str, keep_default_na=False)
    df.columns = [c.strip().lstrip("\ufeff") for c in df.columns]
    files_rec["Rows"] = len(df)
    return df


def metadata_long(cell, md):
    rows, order = [], 0
    if TOP not in md.columns or RESP[0] not in md.columns:
        add("D. Metadata format", cell, "FAIL", f"METADATA columns missing (found {list(md.columns)[:6]}...)")
        return rows
    for lvl, n, l in LEVELS:
        if n in md.columns:
            x = md[md[n].str.strip() != ""][[TOP, n, l]].drop_duplicates()
            for v, c, lb in x.values:
                rows.append((cell, v, lvl, c.strip(), lb, order)); order += 1
    x = md[md[RESP[0]].str.strip() != ""][[TOP, *RESP]]
    for v, c, lb in x.values:
        rows.append((cell, v, "Response", c.strip(), lb, order)); order += 1
    for v in md[TOP].unique():
        rows.append((cell, v, "Variable", "", "", order)); order += 1
    return rows


# ------------------------------------------------------------------ casedata (streamed)
def scan_casedata(cell, part, member, meta_resp_vars, files_rec):
    d, base, size, opener = member
    txt, enc, delim = open_text(opener)
    files_rec.update(Encoding=enc, Delimiter=repr(delim))
    if enc not in ALLOWED_ENCODINGS:
        add("F. Encoding", cell, "FAIL", f"{base}: {enc} (expected UTF-8/UTF-16)")
    if delim != ",":
        add("F. Delimiter", cell, "FAIL", f"{base}: delimiter {delim!r} (expected comma)")
    acc = {"rows": 0, "serials": set(), "vars": set(), "iters": set(), "resp": set(),
           "waves": [], "weeks": set(), "mc": defaultdict(set), "hash": [], "sid": [], "layout": None}
    sid_map = {}
    with txt:
        reader = pd.read_csv(txt, sep=delim, dtype=str, keep_default_na=False, chunksize=CHUNK_ROWS)
        for ch in reader:
            ch.columns = [c.strip().lstrip("\ufeff") for c in ch.columns]
            sc = find_col(ch.columns, SERIAL_COL)
            if acc["layout"] is None:
                acc["layout"] = "long" if TOP in ch.columns else "wide"
                if not sc:
                    add("A. Respondent ID", cell, "FAIL", f"{base}: '{SERIAL_COL}' column not found")
            acc["rows"] += len(ch)
            if sc:
                s = ch[sc].str.strip()
                new = [x for x in s.unique() if x not in sid_map]
                sid_map.update({x: len(sid_map) + i for i, x in enumerate(new)})
                acc["sid"].append(s.map(sid_map).to_numpy(np.int64))
                if (s == "").any():
                    add("A. Respondent ID", cell, "FAIL", f"{base}: {(s == '').sum()} blank serials")
            acc["hash"].append(pd.util.hash_pandas_object(ch, index=False).to_numpy())
            if acc["layout"] == "long":
                top = ch[TOP].str.strip()
                acc["vars"].update(top.unique())
                for lvl, n, l in LEVELS:
                    if n in ch.columns:
                        m = ch[n].str.strip() != ""
                        cols = [TOP, n] + ([l] if l in ch.columns else [])
                        for t in ch.loc[m, cols].drop_duplicates().itertuples(index=False):
                            acc["iters"].add((lvl, t[0].strip(), t[1].strip(), t[2] if len(t) > 2 else None))
                if RESP[0] in ch.columns:
                    m = top.isin(meta_resp_vars) & (ch[RESP[0]].str.strip() != "")
                    acc["resp"].update(map(tuple, ch.loc[m, [TOP, RESP[0]]].drop_duplicates().values))
                    w = ch.loc[top == WAVE_VAR, [sc, RESP[0]]] if sc else None
                    if w is not None and len(w):
                        acc["waves"].append(w.drop_duplicates())
                    for v in ("Country", "Category"):
                        acc["mc"][v].update(ch.loc[top == v, RESP[0]].str.strip().unique())
            else:                                              # wide layout
                for c in ch.columns:
                    mm = re.match(r"^([^\[\.]+)(?:\[\{?([^}\]]*)\}?\])?", c.strip())
                    v, it = (mm[1], mm[2] or "") if mm else (c, "")
                    acc["vars"].add(v)
                    if it:
                        acc["iters"].add(("Iter1", v, it, None))
                    if v in meta_resp_vars:
                        for val in ch[c].dropna().unique():
                            for code in re.split(r"[,;]", str(val).strip("{} ")):
                                if code.strip():
                                    acc["resp"].add((v, code.strip()))
                if sc and find_col(ch.columns, WAVE_VAR):
                    wc = find_col(ch.columns, WAVE_VAR)
                    acc["waves"].append(ch[[sc, wc]].rename(columns={wc: RESP[0]}).drop_duplicates())
                for v in ("Country", "Category"):
                    c = find_col(ch.columns, v)
                    if c:
                        acc["mc"][v].update(ch[c].str.strip().unique())
            wk = find_col(ch.columns, WEEK_COL)
            if wk:
                acc["weeks"].update(ch[wk].str.strip().unique())
    files_rec["Rows"] = acc["rows"]
    acc["serials"] = set(sid_map)
    acc["id2serial"] = {i: s for s, i in sid_map.items()} if sid_map else {}
    # ---- per-file checks
    if acc["rows"] == 0:
        add("F. File quality", cell, "FAIL", f"{base}: header only, no rows")
    if acc["rows"] in EXCEL_ROW_LIMITS:
        add("F. Truncation", cell, "FAIL", f"{base}: exactly {acc['rows']:,} data rows = Excel row limit. "
            "File was very likely opened/saved in Excel and cut off - request a fresh export")
    h = np.concatenate(acc["hash"]) if acc["hash"] else np.array([], dtype=np.uint64)
    dup_serials = []
    if len(h):
        u, inv, cnt = np.unique(h, return_inverse=True, return_counts=True)
        dmask = cnt[inv] > 1
        ndup = int(dmask.sum() - (cnt > 1).sum())
        if ndup:
            sids = np.concatenate(acc["sid"])[dmask] if acc["sid"] else []
            dup_serials = sorted({acc["id2serial"][i] for i in np.unique(sids)})
            add("A. Duplicate rows", cell, "FAIL", f"{base}: {ndup:,} fully duplicated rows across "
                f"{len(dup_serials)} respondent(s): {short(dup_serials, 5)}")
    acc["dup_serials"] = dup_serials
    del acc["hash"], acc["sid"]
    weeks = sorted(w for w in acc["weeks"] if re.fullmatch(r"\d{6}", w))
    a, b = part["period"]
    if weeks and a != "?":
        out = [w for w in weeks if w < a or w > b]
        if out:
            add("B. Period vs data", cell, "FAIL", f"{base}: {len(out)} week(s) outside file period {a}_{b}: {short(out, 6)}")
        missing = []
        w = a
        while w <= b and len(missing) < 500:
            if w not in acc["weeks"]:
                missing.append(w)
            w = week_next(w)
            if w[4:] == "53":
                w = week_next(w)
        if missing:
            add("B. Period vs data", cell, "INFO", f"{base}: no data for {len(missing)} week(s) in {a}_{b}: {short(missing, 8)}")
    acc["week_range"] = f"{weeks[0]}-{weeks[-1]}" if weeks else ""
    return acc


def scan_serials(cell, member, files_rec):
    d, base, size, opener = member
    txt, enc, delim = open_text(opener)
    files_rec.update(Encoding=enc, Delimiter=repr(delim))
    if enc not in ALLOWED_ENCODINGS:
        add("F. Encoding", cell, "FAIL", f"{base}: {enc} (expected UTF-8/UTF-16)")
    if delim != ",":
        add("F. Delimiter", cell, "FAIL", f"{base}: delimiter {delim!r} (expected comma)")
    out, rows = set(), 0
    with txt:
        for ch in pd.read_csv(txt, sep=delim, dtype=str, keep_default_na=False, chunksize=CHUNK_ROWS):
            ch.columns = [c.strip().lstrip("\ufeff") for c in ch.columns]
            rows += len(ch)
            sc = find_col(ch.columns, SERIAL_COL)
            if sc:
                out.update(ch[sc].str.strip().unique())
    files_rec["Rows"] = rows
    if rows in EXCEL_ROW_LIMITS:
        add("F. Truncation", cell, "FAIL", f"{base}: exactly {rows:,} rows = Excel row limit (likely truncated)")
    return out


def read_check_text(member):
    d, base, size, opener = member
    with opener() as f:
        raw = f.read()
    if os.path.splitext(base)[1].lower() in (".xlsx", ".xlsm", ".xls"):
        try:
            sh = pd.read_excel(io.BytesIO(raw), sheet_name=None, header=None, dtype=str)
            return " ".join(" ".join(df.fillna("").astype(str).values.ravel()) for df in sh.values())
        except Exception:
            return ""
    enc = detect_encoding(raw[:65536])
    return raw.decode(enc if enc in ALLOWED_ENCODINGS else "latin-1", errors="ignore")


# ------------------------------------------------------------------ C: coverage
def role(v):
    for name, lst in [("Weight", WEIGHT_VARS), ("Time period", TIME_VARS), ("Filter", FILTER_VARS),
                      ("Optional (market/category)", OPTIONAL_VARS), ("Admin", ADMIN_VARS)]:
        if v in lst:
            return name
    if v in DERIVED_SOURCES or re.search(r"_(T2B|TB|EU)$", v):
        return "Metric (derived)"
    if v in BRAND_VARS_ALL + STATEMENT_ITER_METRICS + ITEM_LIST_METRICS:
        return "Metric"
    return "Unclassified"


def check_coverage(a, cells, baseline):
    allc = sorted(cells)
    present = a.groupby("Var")["Cell"].apply(set).to_dict()
    for v in CORE_VARS:
        miss = set(allc) - present.get(v, set())
        alt = next((k for k in present if k.lower() == v.lower() and k != v), None)
        if alt:
            add("C. Variable names", "ALL", "FAIL", f"Name changed: expected '{v}', found '{alt}'", v)
        elif miss:
            add("C. Variable coverage", short(miss), "FAIL", f"Core variable missing in METADATA of {len(miss)} cell(s)", v)
    for v in OPTIONAL_VARS:
        miss = set(allc) - present.get(v, set())
        if miss and present.get(v):
            add("C. Variable coverage", short(miss, 4), "INFO", f"Market/category-specific - present in {len(present[v])} of {len(allc)} cells", v)
    known = {x.lower() for x in CORE_VARS + OPTIONAL_VARS + ADMIN_VARS}
    new = sorted(v for v in present if v.lower() not in known)
    if new:
        add("C. New variables", "ALL", "WARN", "Not in config - confirm with DP and classify: "
            + "; ".join(f"{v} ({short(present[v], 3)})" for v in new))
    if baseline:
        prev = set(baseline.get("variables", []))
        if prev - set(present):
            add("G. Vs baseline", "ALL", "FAIL", f"Variables removed: {sorted(prev - set(present))}")
        if set(present) - prev:
            add("G. Vs baseline", "ALL", "WARN", f"Variables added (notified?): {sorted(set(present) - prev)}")
    return pd.DataFrame([{"Variable": v, "Role": role(v), "Cells": len(present[v]),
                          "Missing in": short(set(allc) - present[v], 25)} for v in sorted(present, key=str.lower)])


# ------------------------------------------------------------------ D1: brands
def brand_rows(a):
    b = pd.concat([a[(a.Level == "Iter1") & a.Var.isin(BRAND_ITER_METRICS)],
                   a[(a.Level == "Iter2") & a.Var.isin(BRAND_ITER2_METRICS)],
                   a[(a.Level == "Response") & a.Var.isin(BRAND_RESPONSE_METRICS)]])
    return b[b.Code.str.match(BRAND_CODE)]


def check_brands(a, brand_master, baseline):
    b = brand_rows(a)
    if b.empty:
        add("D1. Brand labels", "ALL", "WARN", "No brand lists found - check BRAND_* lists in config")
        return pd.DataFrame(), {}
    n0 = len(FINDINGS)
    for code, g in b.groupby("Code"):
        labs = g.groupby("Label")["Cell"].apply(set)
        if len(labs) > 1:
            add("D1. Brand labels", "ALL", "FAIL", f"Brand code {code} label not identical across files: "
                + " | ".join(f"{show(l)} ({short(c, 4)})" for l, c in labs.items()))
    for lab, g in b.groupby("Label"):
        codes = g.groupby("Code")["Cell"].apply(set)
        if len(codes) > 1:
            add("D1. Brand labels", "ALL", "FAIL", f"Brand {show(lab)} uses {len(codes)} codes: "
                + " | ".join(f"{c} ({short(x, 4)})" for c, x in codes.items()))
    for k, g in b.groupby("Key"):
        if g.Label.nunique() > 1 and g.Code.nunique() > 1:
            add("D1. Brand labels", "ALL", "FAIL", "Same brand written differently: "
                + " | ".join(f"{show(l)}={'/'.join(sorted(set(x.Code)))} ({short(x.Cell, 3)})" for l, x in g.groupby("Label")))
    for lab in [l for l in b.Label.unique() if l != l.strip() or "  " in l or not l.strip()]:
        add("D1. Brand labels", short(b[b.Label == lab].Cell), "FAIL", f"Brand label blank or has extra spaces: {show(lab)}")
    keys = sorted(b.Key.unique()); first = b.groupby("Key").Label.first()
    for i, k1 in enumerate(keys):
        for k2 in keys[i + 1:]:
            if k1[:2] == k2[:2] and abs(len(k1) - len(k2)) <= 3 and \
                    difflib.SequenceMatcher(None, k1, k2).ratio() >= FUZZY_THRESHOLD:
                add("D1. Brand labels", "ALL", "WARN", f"Look-alike brands, please confirm: {show(first[k1])} vs {show(first[k2])}")
    if brand_master is not None:
        for lab in sorted(set(b.Label) - set(brand_master.brand_name)):
            add("D1. Brand labels", short(b[b.Label == lab].Cell), "WARN", f"Brand {show(lab)} not in brand master (exact text)")
        if "brand_code" in brand_master.columns:
            m = dict(zip(brand_master.brand_name, brand_master.brand_code.str.strip()))
            for lab, g in b.groupby("Label"):
                if m.get(lab) and m[lab] not in set(g.Code):
                    add("D1. Brand labels", "ALL", "FAIL", f"{show(lab)}: master code {m[lab]} vs delivered {sorted(set(g.Code))}")
    if baseline:
        prev = baseline.get("brands", {})
        for code, g in b.groupby("Code"):
            if code in prev and prev[code] not in set(g.Label):
                add("G. Vs baseline", "ALL", "FAIL", f"Brand code {code} label changed: {show(prev[code])} -> {sorted(set(g.Label))}")
    if not any(f["Severity"] == "FAIL" for f in FINDINGS[n0:]):
        add("D1. Brand labels", "ALL", "PASS", f"{b.Label.nunique()} brand labels identical across files, one code each")
    look = b.groupby("Key").Label.apply(lambda s: sorted(set(s))).to_dict()
    code_labels = b.groupby("Code").Label.nunique()
    reg = b.groupby("Label").agg(Codes=("Code", lambda s: " | ".join(sorted(set(s)))), CodeCount=("Code", "nunique"),
                                 Key=("Key", "first"), Cells=("Cell", "nunique"), CellList=("Cell", lambda s: short(s, 25)),
                                 Metrics=("Var", "nunique")).reset_index()
    reg["Other spellings"] = [" | ".join(x for x in look[k] if x != l) for l, k in zip(reg.Label, reg.Key)]
    reg["Status"] = ["CHECK" if (n > 1 or o or l != l.strip() or any(code_labels[c] > 1 for c in cs.split(" | "))) else "OK"
                     for l, n, o, cs in zip(reg.Label, reg.CodeCount, reg["Other spellings"], reg.Codes)]
    reg = reg.rename(columns={"Label": "Brand label (as is)", "Metrics": "Metric count"}).drop(columns="Key")
    return reg.sort_values(["Status", "Brand label (as is)"]), b.groupby("Code").Label.first().to_dict()


# ------------------------------------------------------------------ D2-D4: labels per code
def is_yes_no(keys):
    ks = set(keys)
    return bool(ks) and ks <= YES_LABELS | NO_LABELS and bool(ks & YES_LABELS)


def check_labels(a, baseline):
    issues, snap = [], {}
    brand_lvls = {(v, "Iter1") for v in BRAND_ITER_METRICS} | {(v, "Iter2") for v in BRAND_ITER2_METRICS} \
        | {(v, "Response") for v in BRAND_RESPONSE_METRICS}
    stmt_lvls = {(v, "Iter1") for v in STATEMENT_ITER_METRICS} | {(v, "Response") for v in STATEMENT_RESPONSE_METRICS}
    lab = a[(a.Level != "Variable") & ~a.Var.isin(ADMIN_VARS) & ~a.Var.isin(TIME_VARS)]
    n0 = len(FINDINGS)
    for (v, lvl), g in lab.groupby(["Var", "Level"]):
        if (v, lvl) in brand_lvls:
            g = g[~g.Code.str.match(BRAND_CODE)]
            if g.empty:
                continue
        stmt = (v, lvl) in stmt_lvls
        chk = "D2. Statement labels" if stmt else ("D3. Response labels" if lvl == "Response" else "D4. Item list labels")
        # 1. same code -> same label in every cell
        for code, gc in g.groupby("Code"):
            w = gc.groupby("Label")["Cell"].apply(set)
            if len(w) > 1:
                if stmt:
                    sev, why = "INFO", "treated as separate statements"
                elif gc.Key.nunique() == 1:
                    sev, why = "WARN", "case/spacing/punctuation only"
                else:
                    sev, why = "FAIL", "different wording"
                add(chk, "ALL", sev, f"{lvl} code '{code}' has {len(w)} labels ({why}): "
                    + " | ".join(f"{show(l)} ({short(c, 4)})" for l, c in w.items()), v)
                issues += [{"Check": chk, "Severity": sev, "Variable": v, "Level": lvl, "Code": code, "Label": l,
                            "Cells": len(c), "CellList": short(c, 25), "Issue": "Same code, different label"} for l, c in w.items()]
        # 2. same label used by two codes inside one cell
        for cell, gcell in g.groupby("Cell"):
            dup = gcell.groupby("Label").Code.apply(lambda s: sorted(set(s)))
            dup = dup[dup.map(len) > 1]
            for l, codes in dup.items():
                add(chk, cell, "WARN", f"Two codes share one label {show(l)}: {codes} (dashboard cannot tell them apart)", v)
                issues.append({"Check": chk, "Severity": "WARN", "Variable": v, "Level": lvl, "Code": " | ".join(codes),
                               "Label": l, "Cells": 1, "CellList": cell, "Issue": "Same label, two codes in one cell"})
            blank = gcell[gcell.Label.str.strip() == ""].Code.tolist()
            if blank:
                add(chk, cell, "FAIL", f"Blank labels for codes {blank[:10]}", v)
        # 3. same label -> different codes across cells (not for Yes/No, not market-specific)
        if v not in MARKET_SPECIFIC_VARS and not is_yes_no(set(g.Key)):
            for k, gk in g.groupby("Key"):
                if k and gk.Code.nunique() > 1 and gk.groupby("Cell").Code.nunique().max() == 1:
                    add(chk, "ALL", "WARN", f"Label {show(gk.Label.iloc[0])} uses different codes across cells: "
                        + " | ".join(f"{c} ({short(x.Cell, 3)})" for c, x in gk.groupby("Code")), v)
        # 4. code list differences -> INFO
        sets = g.groupby("Cell").Code.apply(frozenset)
        if sets.nunique() > 1 and not stmt and v not in MARKET_SPECIFIC_VARS + ["Country", "Category"]:
            allc = set().union(*sets)
            common = frozenset.intersection(*sets)
            extra = {c: sorted(s - common) for c, s in sets.items() if s - common}
            add(chk, "ALL", "INFO", f"Answer list differs by cell ({len(common)} common codes of {len(allc)}). Extra codes: "
                + "; ".join(f"{c}: {x[:5]}" for c, x in list(extra.items())[:6]), v)
        snap[f"{v}|{lvl}"] = g.groupby("Code").Label.first().to_dict()
        if baseline:
            old = baseline.get("labels", {}).get(f"{v}|{lvl}")
            if isinstance(old, dict):
                cur = g.groupby("Code").Label.apply(set).to_dict()
                for code, ol in old.items():
                    if code in cur and ol not in cur[code]:
                        add("G. Vs baseline", "ALL", "FAIL", f"{lvl} code {code} label changed: '{ol}' -> {sorted(cur[code])}", v)
    if not any(f["Severity"] == "FAIL" and f["Check"][:2] in ("D3", "D4") for f in FINDINGS[n0:]):
        add("D3-D4. Response labels", "ALL", "PASS", "Every response code keeps one label across cells")
    return pd.DataFrame(issues), snap


def check_derived(a):
    b = brand_rows(a)
    for der, src in DERIVED_SOURCES.items():
        for cell in sorted(a.Cell.unique()):
            d = set(b[(b.Cell == cell) & (b.Var == der)].Code)
            s = set(b[(b.Cell == cell) & (b.Var == src)].Code)
            if d and s and d != s:
                add("D5. Derived variables", cell, "FAIL",
                    f"Brands differ from {src}: only in {der} {sorted(d - s)[:8]}, only in {src} {sorted(s - d)[:8]}", der)


# ------------------------------------------------------------------ H: CASEDATA vs METADATA
def check_case_vs_meta(cell, acc, a):
    m = a[a.Cell == cell]
    meta_vars = set(m[m.Level == "Variable"].Var)
    case_vars = acc["vars"] - {""}
    rows, n0 = [], len(FINDINGS)
    if WAVE_VAR in meta_vars and WAVE_VAR not in case_vars:
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"'{WAVE_VAR}' defined in METADATA but has no rows in CASEDATA", WAVE_VAR)
    miss = sorted(v for v in meta_vars - case_vars if v not in ADMIN_VARS)
    core = [v for v in miss if v in CORE_VARS]
    if core:
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"Core variables in METADATA with no rows in CASEDATA: {core}")
    if set(miss) - set(core):
        add("H. CASEDATA vs METADATA", cell, "INFO", f"Other METADATA variables with no CASEDATA rows: {sorted(set(miss) - set(core))[:30]}")
    extra = sorted(case_vars - meta_vars)
    if extra:
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"Variables in CASEDATA not defined in METADATA: {extra[:30]}")
    meta_it = {(r.Level, r.Var, r.Code): r.Label for r in m[m.Level.str.startswith("Iter")].itertuples()}
    brand_iter = {("Iter1", v) for v in BRAND_ITER_METRICS} | {("Iter2", v) for v in BRAND_ITER2_METRICS}
    bad_code, bad_label = defaultdict(set), defaultdict(set)
    for lvl, v, code, label in acc["iters"]:
        if v not in meta_vars:
            continue
        mk = (lvl, v, code)
        if mk not in meta_it:
            bad_code[v].add(code); rows.append({"Cell": cell, "Variable": v, "Level": lvl, "Code": code, "CaseData label": label, "Metadata label": "", "Issue": "Iterator not in METADATA"})
        elif label is not None and label != meta_it[mk]:
            bad_label[v].add(code)
            rows.append({"Cell": cell, "Variable": v, "Level": lvl, "Code": code, "CaseData label": label,
                         "Metadata label": meta_it[mk], "Issue": "Brand label differs" if (lvl, v) in brand_iter else "Iterator label differs"})
    for v, codes in bad_code.items():
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"Iterators in CASEDATA not defined in METADATA: {sorted(codes)[:12]}", v)
    for v, codes in bad_label.items():
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"Iterator/brand labels in CASEDATA differ from METADATA (as is) for codes {sorted(codes)[:12]}", v)
    meta_resp = m[m.Level == "Response"].groupby("Var").Code.apply(set).to_dict()
    undefined = defaultdict(set)
    for v, code in acc["resp"]:
        if v in meta_resp and code not in meta_resp[v]:
            undefined[v].add(code)
    for v, codes in undefined.items():
        add("H. CASEDATA vs METADATA", cell, "FAIL", f"Response codes in CASEDATA with no label in METADATA: {sorted(codes)[:15]}", v)
        rows += [{"Cell": cell, "Variable": v, "Level": "Response", "Code": c, "CaseData label": "", "Metadata label": "",
                  "Issue": "Response code not in METADATA"} for c in sorted(codes)[:200]]
    for v, code in (("Country", cell[2:4]), ("Category", cell[4:8])):
        vals = acc["mc"].get(v, set()) - {""}
        if len(vals) > 1:
            add("F. Single market/category", cell, "FAIL", f"{v} has {len(vals)} values in CASEDATA: {short(vals)}")
        elif vals and code.upper() not in next(iter(vals)).upper():
            lab = m[(m.Var == v) & (m.Level == "Response") & (m.Code == next(iter(vals)))].Label
            if not (len(lab) and code.upper() in lab.iloc[0].upper()) and v == "Category":
                add("F. Single market/category", cell, "WARN", f"{v} value {vals} does not contain file-name code {code}")
    if len(FINDINGS) == n0:
        add("H. CASEDATA vs METADATA", cell, "PASS", "All CASEDATA variables, iterators, labels and codes match METADATA")
    return rows


# ------------------------------------------------------------------ W: waves
def build_waves(cell_waves, a, names):
    w = a[(a.Var == WAVE_VAR) & (a.Level == "Response")]
    labels = dict(zip(w.Code, w.Label.str.strip()))
    order = {c: i for i, c in enumerate(w.drop_duplicates("Code").sort_values("Order").Code)}
    present = set().union(*[set(x) for x in cell_waves.values()]) if cell_waves else set()
    for c in present - set(labels):
        add("W. Wave availability", "ALL", "FAIL", f"Wave code '{c}' used in CASEDATA but not defined in METADATA")
        labels[c] = c
    waves = sorted(present, key=lambda c: (order.get(c, 10**6), c))
    seen = defaultdict(int)
    for c in waves:
        seen[labels[c]] += 1
    wl = {c: labels[c] if seen[labels[c]] == 1 else f"{labels[c]} ({c})" for c in waves}
    mat = pd.DataFrame({names[c]["Combo"]: [cell_waves.get(c, {}).get(x, 0) for x in waves] for c in sorted(names)},
                       index=[wl[x] for x in waves])
    mat.index.name = "Wave"
    for c in sorted(names):
        cw = cell_waves.get(c, {})
        idx = [i for i, x in enumerate(waves) if cw.get(x, 0) > 0]
        if not idx:
            add("W. Wave availability", c, "FAIL" if cw is not None else "WARN", "No wave data found in CASEDATA"); continue
        gaps = [wl[waves[i]] for i in range(idx[0], idx[-1] + 1) if cw.get(waves[i], 0) == 0]
        if gaps:
            add("W. Wave availability", c, "WARN", f"{len(gaps)} wave(s) missing inside its range: {gaps[:12]}")
        add("W. Wave availability", c, "INFO", f"{len(idx)} waves with data: {wl[waves[idx[0]]]} to {wl[waves[idx[-1]]]}")
    return mat


def write_wave_sheets(wb, mat):
    hdr, hf = PatternFill("solid", start_color="1F4E78"), Font(name="Arial", bold=True, color="FFFFFF")
    green, red = PatternFill("solid", start_color="C6EFCE"), PatternFill("solid", start_color="FFC7CE")
    wm = wb.create_sheet("Wave_Matrix")
    wm.cell(1, 1, "Wave")
    for j, col in enumerate(mat.columns, 2):
        wm.cell(1, j, col)
    for i, (wave, row) in enumerate(mat.iterrows(), 2):
        wm.cell(i, 1, str(wave))
        for j, v in enumerate(row.values, 2):
            wm.cell(i, j, int(v))
    nW, nC = len(mat), len(mat.columns)
    last = get_column_letter(nC + 1)
    data, wr, cr = f"Wave_Matrix!$B$2:${last}${nW + 1}", f"Wave_Matrix!$A$2:$A${nW + 1}", f"Wave_Matrix!$B$1:${last}$1"
    wm.cell(1, nC + 2, "Combinations with data")
    for i in range(2, nW + 2):
        wm.cell(i, nC + 2, f'=COUNTIF(B{i}:{last}{i},">0")')
    wm.conditional_formatting.add(f"B2:{last}{nW + 1}", CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    wm.conditional_formatting.add(f"B2:{last}{nW + 1}", CellIsRule(operator="equal", formula=["0"], fill=red))
    for c in wm[1]:
        c.fill, c.font = hdr, hf
    wm.column_dimensions["A"].width = 22
    wm.freeze_panes = "B2"
    ws = wb.create_sheet("Wave_Availability", 1)
    ws["A1"] = "Wave availability by Country x Category (respondents per wave, from CASEDATA)"
    ws["A1"].font = Font(name="Arial", bold=True, size=14)
    ws["A3"], ws["A4"] = "Select Country x Category", "Select wave"
    ws["B3"], ws["B4"] = mat.columns[0], str(mat.index[-1])
    for c in ("B3", "B4"):
        ws[c].fill = PatternFill("solid", start_color="FFFF00"); ws[c].font = Font(name="Arial", bold=True, color="0000FF")
    dv1, dv2 = DataValidation(type="list", formula1=f"={cr}"), DataValidation(type="list", formula1=f"={wr}")
    ws.add_data_validation(dv1); ws.add_data_validation(dv2); dv1.add("B3"); dv2.add("B4")
    ws["A6"], ws["B6"] = "Waves with data (selected combination)", f'=COUNTIF(E9:E{nW + 8},">0")'
    ws["A7"], ws["B7"] = "Combinations with data (selected wave)", f'=COUNTIF(B10:B{nC + 9},">0")'
    for j, h in enumerate(["Country x Category", "Respondents", "Status"], 1):
        ws.cell(9, j, h)
    for i, combo in enumerate(mat.columns, 10):
        ws.cell(i, 1, combo)
        ws.cell(i, 2, f"=IFERROR(INDEX({data},MATCH($B$4,{wr},0),MATCH(A{i},{cr},0)),0)")
        ws.cell(i, 3, f'=IF(B{i}>0,"Available","Not available")')
    for j, h in zip((4, 5, 6), ["Wave", "Respondents", "Status"]):
        ws.cell(8, j, h)
    for i, wave in enumerate(mat.index, 9):
        ws.cell(i, 4, str(wave))
        ws.cell(i, 5, f"=IFERROR(INDEX({data},MATCH(D{i},{wr},0),MATCH($B$3,{cr},0)),0)")
        ws.cell(i, 6, f'=IF(E{i}>0,"Available","Not available")')
    for rng in (f"C10:C{nC + 9}", f"F9:F{nW + 8}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Available"'], fill=green))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Not available"'], fill=red))
    ws.conditional_formatting.add(f"D9:F{nW + 8}", FormulaRule(formula=["$D9=$B$4"], font=Font(bold=True)))
    for c in list(ws[9])[:3] + list(ws[8])[3:6]:
        c.fill, c.font = hdr, hf
    for col, wdt in zip("ABCDEF", (46, 13, 14, 22, 13, 14)):
        ws.column_dimensions[col].width = wdt
    ch = LineChart()
    ch.title, ch.y_axis.title, ch.x_axis.title = "Respondents by wave - selected Country x Category", "Respondents", "Wave"
    ch.add_data(Reference(ws, min_col=5, min_row=8, max_row=nW + 8), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=4, min_row=9, max_row=nW + 8))
    ch.height, ch.width, ch.legend = 9, 26, None
    ch.series[0].smooth = False
    ch.series[0].marker.symbol = "circle"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ws.add_chart(ch, "H3")
    ws.freeze_panes = "A10"


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", default="gdf_sanity_report.xlsx")
    ap.add_argument("--metadata-only", action="store_true")
    ap.add_argument("--expected-cells", help="Comma list of MarketCategory codes (e.g. ZAHHON,INHSKF) or a text file")
    ap.add_argument("--baseline")
    ap.add_argument("--brand-master")
    o = ap.parse_args()

    cells, files = discover(o.input, o.metadata_only)
    if not cells:
        sys.exit("No cells found (use --metadata-only for loose METADATA CSVs)")
    print(f"{len(cells)} cells: {', '.join(sorted(cells))}")
    base = json.load(open(o.baseline)) if o.baseline and os.path.exists(o.baseline) else None
    bm = None
    if o.brand_master:
        bm = pd.read_csv(o.brand_master, dtype=str, keep_default_na=False)
        bm.columns = [c.strip().lower() for c in bm.columns]
    if o.expected_cells:
        raw = open(o.expected_cells).read() if os.path.exists(o.expected_cells) else o.expected_cells
        exp = {x.strip().upper()[-6:] for x in re.split(r"[,\s;]+", raw) if x.strip()}
        got = {c[2:8] for c in cells}
        for x in sorted(exp - got):
            add("B. Expected cells", x, "FAIL", f"Expected cell {x[:2]} x {x[2:]} was not delivered")
        for x in sorted(got - exp):
            add("B. Expected cells", x, "INFO", f"Cell {x[:2]} x {x[2:]} delivered but not in the expected list")
    frec = {(f["Cell"], f["File"]): f for f in files}

    # ---- metadata (latest period first; others compared)
    rows, metas = [], {}
    for cell, parts in sorted(cells.items()):
        sigs = {}
        for p in parts:
            for mem in p["members"].get("METADATA", []):
                df = read_small(mem, cell, frec[(cell, mem[1])])
                sigs["_".join(p["period"])] = pd.util.hash_pandas_object(df, index=False).sum()
                metas[cell] = df
        if len(set(sigs.values())) > 1:
            add("D. Metadata consistency", cell, "WARN", f"METADATA differs between period files {list(sigs)} - latest used")
        if cell in metas:
            rows += metadata_long(cell, metas[cell])
    a = pd.DataFrame(rows, columns=["Cell", "Var", "Level", "Code", "Label", "Order"]).drop_duplicates(["Cell", "Var", "Level", "Code", "Label"])
    a["Key"] = a.Label.map(key)
    names = {}
    for c in cells:
        cat = a[(a.Var == "Category") & (a.Level == "Response") & (a.Code == c[4:8])].Label
        cty = a[(a.Var == "Country") & (a.Level == "Response") & (a.Code.str.upper() == c[2:4])].Label
        names[c] = {"Country": c[2:4], "CountryLabel": cty.iloc[0].strip() if len(cty) else c[2:4], "Category": c[4:8],
                    "CategoryLabel": cat.iloc[0].strip() if len(cat) else c[4:8],
                    "Combo": f"{c[2:4]} x {c[4:8]}" + (f" - {cat.iloc[0].strip()}" if len(cat) else "")}

    coverage = check_coverage(a, cells, base)
    brands, bsnap = check_brands(a, bm, base)
    label_issues, lsnap = check_labels(a, base)
    check_derived(a)

    # ---- casedata / existence / check tables
    owner, undefined, dup_rows, part_rows, cell_waves = defaultdict(set), [], [], [], {}
    if not o.metadata_only:
        meta_resp_vars = set(a[(a.Level == "Response")].Var) - set(ADMIN_VARS) - {"Age"}
        for cell, parts in sorted(cells.items()):
            print("  scanning", cell)
            agg = {"vars": set(), "iters": set(), "resp": set(), "mc": defaultdict(set)}
            wave_frames, check_txt, cell_serials = [], "", set()
            for p in parts:
                pname = "_".join(p["period"])
                acc = None
                for mem in p["members"].get("CASEDATA", []):
                    acc = scan_casedata(cell, p, mem, meta_resp_vars, frec[(cell, mem[1])])
                if acc is None:
                    continue
                overlap = cell_serials & acc["serials"]
                if overlap:
                    add("A. Respondent ID", cell, "FAIL", f"{len(overlap)} serials appear in more than one period file (e.g. {short(overlap, 3)})")
                cell_serials |= acc["serials"]
                for k in ("vars", "iters", "resp"):
                    agg[k] |= acc[k]
                for k, v in acc["mc"].items():
                    agg["mc"][k] |= v
                wave_frames += acc["waves"]
                dup_rows += [{"Cell": cell, "Period": pname, "Respondent.Serial": s, "Issue": "Fully duplicated row(s)"} for s in acc["dup_serials"][:500]]
                ex_serials = None
                for mem in p["members"].get("EXISTENCE", []):
                    ex_serials = scan_serials(cell, mem, frec[(cell, mem[1])])
                if ex_serials is not None and ex_serials != acc["serials"]:
                    add("F. Completeness", cell, "WARN", f"{pname}: EXISTENCE vs CASEDATA serials - {len(acc['serials'] - ex_serials):,} only in CASEDATA, {len(ex_serials - acc['serials']):,} only in EXISTENCE")
                for mem in p["members"].get("CHECK", []):
                    check_txt += " " + read_check_text(mem)
                nw = pd.concat(acc["waves"]).iloc[:, 0].nunique() if acc["waves"] else 0
                part_rows.append({"Cell": cell, "Period": pname, "DeliveryDate": p["date"], "CaseData rows": acc["rows"],
                                  "Respondents": len(acc["serials"]), "Respondents with Rounds": nw,
                                  "Week range in data": acc["week_range"], "Layout": acc["layout"]})
            for s in cell_serials:
                owner[s].add(cell)
            if wave_frames:
                wf = pd.concat(wave_frames).drop_duplicates()
                sc = wf.columns[0]
                multi = wf.groupby(sc)[RESP[0]].nunique()
                if (multi > 1).any():
                    add("A. Respondent ID", cell, "WARN", f"{int((multi > 1).sum())} respondents have more than one wave in Rounds")
                cell_waves[cell] = wf.groupby(RESP[0])[sc].nunique().to_dict()
            else:
                cell_waves[cell] = {}
            undefined += check_case_vs_meta(cell, agg, a)
            scales = [v for v in SCALE_VARS if v in agg["vars"]]
            if scales and parts and any(p["members"].get("CHECK") for p in parts):
                miss = [t for t in REQUIRED_CHECK_STATS if not CHECK_TOKENS[t].search(check_txt)]
                if miss:
                    add("F. Check tables", cell, "FAIL", f"Check tables lack {miss} (scale metrics present: {scales})")
                if not CHECK_TOKENS["T3B"].search(check_txt):
                    add("F. Check tables", cell, "INFO", "No T3B in check tables - confirm whether T3B is required")
        cross = {s: c for s, c in owner.items() if len(c) > 1}
        if cross:
            add("A. Respondent ID", "ALL", "FAIL", f"{len(cross):,} Respondent.Serial values appear in more than one cell")
            dup_rows += [{"Cell": " | ".join(sorted(c)), "Period": "", "Respondent.Serial": s, "Issue": "Reused across cells"} for s, c in list(cross.items())[:5000]]
        elif owner:
            add("A. Respondent ID", "ALL", "PASS", f"{len(owner):,} Respondent.Serial values, all unique across cells")
    encs = {f.get("Encoding") for f in files if f.get("Encoding") and f["Type"] != "CHECK"}
    if len(encs) > 1:
        add("F. Encoding", "ALL", "WARN", f"Mixed encodings across files: {sorted(encs)}")
    wave_mat = build_waves(cell_waves, a, names) if cell_waves else pd.DataFrame()

    f = pd.DataFrame(FINDINGS)
    order = {"FAIL": 0, "WARN": 1, "INFO": 2, "PASS": 3}
    f = f.assign(_o=f.Severity.map(order)).sort_values(["_o", "Check", "Cell", "Variable"]).drop(columns="_o")
    summary = f.pivot_table(index="Check", columns="Severity", values="Detail", aggfunc="count", fill_value=0).reset_index()
    matrix = a[a.Level == "Variable"].assign(v="Y").pivot_table(index="Var", columns="Cell", values="v", aggfunc="first", fill_value="-").reset_index()
    sheets = {"Summary": summary, "Findings": f, "Delivery_Parts": pd.DataFrame(part_rows), "Brand_Register": brands,
              "Label_Issues": label_issues, "Variable_Coverage": coverage, "Variable_x_Cell": matrix,
              "Cells": pd.DataFrame([{"Cell": c, **names[c]} for c in sorted(names)]),
              "CaseData_vs_Metadata": pd.DataFrame(undefined), "Duplicate_Rows": pd.DataFrame(dup_rows),
              "File_Inventory": pd.DataFrame(files)}
    with pd.ExcelWriter(o.output, engine="openpyxl") as xw:
        for n, df in sheets.items():
            (df if len(df) else pd.DataFrame({"Result": ["No issues"]})).to_excel(xw, sheet_name=n, index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col[:500]) + 2, 80)
            for c in ws[1]:
                c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", start_color="1F4E78")
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
        if len(wave_mat) and len(wave_mat.columns):
            write_wave_sheets(xw.book, wave_mat)
        xw.book.calculation.fullCalcOnLoad = True
    snap = {"created": datetime.now().isoformat(timespec="seconds"), "variables": sorted(a.Var.unique()),
            "brands": bsnap, "labels": lsnap}
    sp = os.path.splitext(o.output)[0] + "_snapshot.json"
    json.dump(snap, open(sp, "w"), indent=1)
    print(f"Done: {(f.Severity == 'FAIL').sum()} FAIL, {(f.Severity == 'WARN').sum()} WARN\nReport: {o.output}\nSnapshot: {sp}")


if __name__ == "__main__":
    main()
