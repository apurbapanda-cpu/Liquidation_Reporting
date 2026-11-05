"""
Liquidation Reporting Pipeline  (spec: spec_files_Liq_reporting.xlsx)

Step 1 : Resolve input files (latest of each; confirm).
Step 2 : Load IPC NPD-risk raw (chunked, pre-filtered) + mappings, NLC, LP, sales.
Step 3 : Build FSN x DS x City output.
Step 4 : Export CSV to Output_files_Liquidation.

Usage:  python main.py [--yes]      (--yes skips the confirm prompt)
"""

import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------------------
# Paths (Drive shortcut root, same as smnm_reporting)
# ---------------------------------------------------------------------------
ROOT          = Path(r"G:\.shortcut-targets-by-id\15JLkN3qz_KXLVTnsFWiCIXCbhkgPWc4q\SMNM_Minutes")
MAPPING_DIR   = ROOT / "Mapping_Tables"
LIQ_RAW_DIR   = ROOT / "Liquidation_Raw_Data"
SALES_DAY_DIR = ROOT / "Sales_Data" / "day_level_sales_fsn_ds"
NLC_WEEK_DIR  = ROOT / "NLC_Dump" / "nlc_fsn_city_weekly"
LIVE_PRICES_DIR = ROOT / "Live_Prices"
OUTPUT_DIR    = ROOT / "Output_files_Liquidation"

# Risk ATP measure in the IPC NPD-risk file (qty not consumed before NPD/expiry).
RISK_COL = "atp"                                     # ATP (>0) is bucketed by risk_bucket
BUCKETS  = ["0-3", "4-7", "8-15", "16-31"]          # matched against risk_bucket label
IPC_COLS = ["fsn", "ds", "analytic_vertical", "brand", "product_title",
            RISK_COL, "risk_bucket"]

_MONTH = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


CHUNK_ROWS = 1_000_000          # rows per chunk for the multi-GB files (lower = less memory)
_T0 = datetime.now()
_N_STEPS = 6
_step_t = _T0


def log(msg: str) -> None:
    el = int((datetime.now() - _T0).total_seconds())
    print(f"[{datetime.now():%H:%M:%S} | +{el // 60}m{el % 60:02d}s] {msg}", flush=True)


def step(n: int, name: str) -> None:
    """Print a step header ([n/total], % of steps done) and how long the previous step took."""
    global _step_t
    now = datetime.now()
    if n > 1:
        log(f"  step done in {int((now - _step_t).total_seconds())}s")
    _step_t = now
    log(f"[Step {n}/{_N_STEPS}] {name}  ({(n - 1) * 100 // _N_STEPS}% complete)")


def read_chunks(path: Path, label: str, chunksize: int | None = None, **kw):
    """Yield CSV chunks, logging % of file read (by bytes), rows so far and ETA for this file."""
    size, t0, rows = max(path.stat().st_size, 1), datetime.now(), 0
    with open(path, "rb") as fh:          # binary handle -> fh.tell() approximates bytes consumed
        for i, ch in enumerate(pd.read_csv(fh, chunksize=chunksize or CHUNK_ROWS, **kw), 1):
            rows += len(ch)
            pct = min(99, fh.tell() * 100 // size)
            el = (datetime.now() - t0).total_seconds()
            eta = f", ETA ~{int(el * (100 - pct) / pct)}s" if pct > 0 else ""
            log(f"  {label}: {pct}% of file read ({rows:,} rows, chunk {i}{eta})")
            yield ch
    log(f"  {label}: 100% read in {int((datetime.now() - t0).total_seconds())}s")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _file_date_key(p: Path) -> datetime:
    """YYYYMMDD prefix/suffix, else last {D}{Mon}{YY} token (range files), else mtime."""
    m = re.search(r"(20\d{6})", p.name)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y%m%d")
        except ValueError:
            pass
    toks = re.findall(r"(\d{1,2})([A-Za-z]{3})(\d{2})", p.name)
    if toks:
        d, mon, y = toks[-1]
        if mon.lower() in _MONTH:
            return datetime(2000 + int(y), _MONTH[mon.lower()], int(d))
    return datetime.fromtimestamp(p.stat().st_mtime)


def latest(directory: Path, pattern: str) -> Path | None:
    files = sorted(directory.glob(pattern), key=_file_date_key, reverse=True)
    return files[0] if files else None


def _s(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def _read_csv(path: Path, usecols=None, **kw) -> pd.DataFrame:
    header = pd.read_csv(path, nrows=0)
    header.columns = header.columns.str.lstrip("\ufeff").str.strip()
    keep = [c for c in usecols if c in header.columns] if usecols else None
    df = pd.read_csv(path, usecols=keep, low_memory=False, **kw)
    df.columns = df.columns.str.lstrip("\ufeff").str.strip()
    return df


# ---------------------------------------------------------------------------
# Step 1 â€” File selection
# ---------------------------------------------------------------------------
def _pick_file(matches: list[Path], auto_yes: bool) -> Path | None:
    """Liquidation file picker (newest first); Enter or --yes takes the newest."""
    if not matches or auto_yes or len(matches) == 1:
        return matches[0] if matches else None
    print("\n  Liquidation (IPC NPD risk) files found (newest first):")
    for i, p in enumerate(matches, 1):
        print(f"    [{i}] {p.name}{'  <- default' if i == 1 else ''}")
    while True:
        sel = input("  Select file [Enter for default, or number]: ").strip()
        if sel == "":
            return matches[0]
        if sel.isdigit() and 1 <= int(sel) <= len(matches):
            return matches[int(sel) - 1]
        print(f"  Enter a number from 1 to {len(matches)}.")


def select_files(auto_yes: bool) -> dict[str, Path]:
    log("Step 1: Resolving files")
    specs = {
        "DS_City_Map":     (MAPPING_DIR, "*_DS_City_Map.csv"),
        "Minutes_BU_Tags": (MAPPING_DIR, "*_Minutes_BU_Tags.csv"),
        "IPC_Raw":         (LIQ_RAW_DIR, "ipc_DataSense_Dry_NPD_ds_npd_risk_*.csv"),
        "Sales_Day":       (SALES_DAY_DIR, "Sales_FSN_DS_*.csv"),
        "Live_Prices":     (LIVE_PRICES_DIR, "*_Live_Prices_Minutes.csv"),
    }
    files: dict[str, Path] = {}
    missing = []
    for label, (d, pat) in specs.items():
        if label == "IPC_Raw":
            p = _pick_file(sorted(d.glob(pat), key=_file_date_key, reverse=True), auto_yes)
        else:
            p = latest(d, pat)
        if p is None:
            missing.append(label)
        else:
            files[label] = p
            print(f"  {label:<16} -> {p.name}")
    if missing:
        raise SystemExit(f"ERROR: missing inputs: {', '.join(missing)}")
    return files


def _ask_date(prompt: str, lo: datetime, hi: datetime) -> datetime:
    while True:
        try:
            d = datetime.strptime(input(prompt).strip(), "%Y-%m-%d")
        except ValueError:
            print("  Use YYYY-MM-DD.")
            continue
        if lo <= d <= hi:
            return d
        print(f"  Date must be within {lo:%Y-%m-%d} .. {hi:%Y-%m-%d}.")


def select_sales_period(sales_path: Path, auto_yes: bool) -> list[datetime]:
    """Prompt for the sales period; default = last 7 days available in the sales file."""
    log("Scanning sales file for available dates ...")
    chunks = [(c["order_date"].min(), c["order_date"].max()) for c in
              read_chunks(sales_path, "sales dates", usecols=["order_date"])]
    lo = datetime.strptime(str(int(min(c[0] for c in chunks))), "%Y%m%d")
    hi = datetime.strptime(str(int(max(c[1] for c in chunks))), "%Y%m%d")
    start, end = max(lo, hi - timedelta(days=6)), hi
    if not auto_yes:
        print(f"\n  Sales available: {lo:%d-%b-%y} to {hi:%d-%b-%y}")
        print(f"  [1] Last 7 days ({start:%d-%b-%y} to {end:%d-%b-%y})  <- default")
        print("  [2] Custom date range")
        if input("  Select sales period [1/2]: ").strip() == "2":
            while True:
                start = _ask_date("  Date from (YYYY-MM-DD): ", lo, hi)
                end = _ask_date("  Date to   (YYYY-MM-DD): ", lo, hi)
                if start <= end:
                    break
                print("  'From' must not be after 'to'.")
    days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    log(f"Sales period: {days[0]:%d-%b-%y} to {days[-1]:%d-%b-%y} ({len(days)} days)")
    return days


# ---------------------------------------------------------------------------
# NLC (weekly xlsx only, latest week/version first; first file with the FSN wins)
# ---------------------------------------------------------------------------
def _norm(c: str) -> str:
    return re.sub(r"[^a-z0-9%]+", "", str(c).lower())


def _std_nlc(df: pd.DataFrame) -> pd.DataFrame | None:
    norm = {c: _norm(c) for c in df.columns}
    fsn_c = next((c for c, n in norm.items() if n in ("fsn", "productid")), None)
    nlc_cs = [c for c, n in norm.items() if "nlc" in n and "tax%" not in n]
    if fsn_c is None or not nlc_cs:
        return None
    excl = next((c for c in nlc_cs if any(k in norm[c] for k in ("w/o", "wo", "excl", "without"))), None)
    col = excl or nlc_cs[0]          # tax-exclusive preferred; else first NLC column
    out = pd.DataFrame({"fsn": _s(df[fsn_c]), "NLC": pd.to_numeric(df[col], errors="coerce")})
    out = out[out["NLC"] > 0]
    return out.groupby("fsn", as_index=False)["NLC"].mean()


def build_nlc() -> pd.DataFrame:
    log("Building NLC lookup")
    week = sorted(NLC_WEEK_DIR.glob("Category_WEEK*NLC*.xls*"),
                  key=lambda p: (int((re.search(r"WEEK\s+(\d+)", p.name, re.I) or [0, 0])[1]),
                                 int((re.search(r"_V(\d+)", p.name, re.I) or [0, 0])[1])),
                  reverse=True)
    frames, seen = [], set()
    for i, p in enumerate(week, 1):
        log(f"  NLC file {i}/{len(week)}: reading {p.name} ...")
        try:
            # parse only the FSN / NLC columns (Excel can't be chunked, so cut width instead)
            raw = pd.read_excel(p, usecols=lambda c: _norm(c) in ("fsn", "productid")
                                or ("nlc" in _norm(c) and "tax%" not in _norm(c)))
            raw.columns = [str(c).strip() for c in raw.columns]
            df = _std_nlc(raw)
        except Exception as e:
            log(f"    skipped: {e}")
            continue
        if df is None or df.empty:
            log("    skipped: no usable FSN / NLC columns or rows")
            continue
        new = df[~df["fsn"].isin(seen)]
        frames.append(new)
        seen.update(new["fsn"])
        log(f"    {len(df):,} FSNs in file, +{len(new):,} new  (total resolved: {len(seen):,})")
    if not frames:
        return pd.DataFrame(columns=["fsn", "NLC"])
    out = pd.concat(frames, ignore_index=True)
    log(f"  NLC: {len(out):,} FSNs")
    return out


# ---------------------------------------------------------------------------
# Step 2/3 â€” Load + build
# ---------------------------------------------------------------------------
def load_ipc(path: Path) -> pd.DataFrame:
    """Chunked read; keep rows with ATP > 0 and risk > 0; aggregate to FSN x DS."""
    log(f"Loading IPC {path.name}")
    header = pd.read_csv(path, nrows=0).columns.str.strip()
    keep = [c for c in IPC_COLS if c in header]
    parts = []
    for ch in read_chunks(path, "IPC", chunksize=500_000, usecols=keep, low_memory=False):
        ch[RISK_COL] = pd.to_numeric(ch[RISK_COL], errors="coerce").fillna(0)
        ch = ch[ch[RISK_COL] > 0]
        ch["fsn"], ch["ds"] = _s(ch["fsn"]), _s(ch["ds"])
        parts.append(ch)
    raw = pd.concat(parts, ignore_index=True)
    log(f"  rows after {RISK_COL}>0: {len(raw):,}")

    raw["_bkt"] = raw["risk_bucket"].astype(str).str.extract(r"(\d+\s*-\s*\d+)")[0].str.replace(" ", "")
    unknown = raw.loc[~raw["_bkt"].isin(BUCKETS), "risk_bucket"].unique()
    if len(unknown):
        log(f"  NOTE: risk_buckets counted only in Total: {list(unknown)}")
    for b in BUCKETS:
        raw[f"Risk ATP ({b} days)"] = raw[RISK_COL].where(raw["_bkt"] == b, 0)
    raw = raw.rename(columns={RISK_COL: "Total Risk ATP"})

    sums = ["Total Risk ATP"] + [f"Risk ATP ({b} days)" for b in BUCKETS]
    meta = [c for c in ["analytic_vertical", "brand", "product_title"] if c in raw.columns]
    agg = {**{c: "sum" for c in sums}, **{c: "first" for c in meta}}
    return raw.groupby(["fsn", "ds"], as_index=False).agg(agg)


def build(files: dict[str, Path], days: list[datetime]) -> tuple[pd.DataFrame, datetime]:
    ipc_date = _file_date_key(files["IPC_Raw"])
    step(2, "Load Liquidation (IPC) file")
    base = load_ipc(files["IPC_Raw"])
    fsns, dss = set(base["fsn"]), set(base["ds"])
    log(f"  base: {len(base):,} FSN x DS rows ({len(fsns):,} FSNs, {len(dss):,} DSs)")

    step(3, "Map Sell BU, City, NLC")
    # Sell BU
    tags = _read_csv(files["Minutes_BU_Tags"])
    tags["Vertical"] = _s(tags["Vertical"])
    tags = tags[["Vertical", "Sell BU"]].drop_duplicates("Vertical")
    base = base.merge(tags, left_on="analytic_vertical", right_on="Vertical", how="left").drop(columns="Vertical")

    # City (DS_City_Map: Warehouse ID -> City)
    cm = _read_csv(files["DS_City_Map"])
    cm = cm.assign(ds=_s(cm["Warehouse ID"]))[["ds", "City"]].drop_duplicates("ds")
    base = base.merge(cm, on="ds", how="left")

    # NLC + risk value
    base = base.merge(build_nlc(), on="fsn", how="left")
    for c in [c for c in base.columns if c.startswith(("Total Risk ATP", "Risk ATP ("))]:
        base[c.replace("ATP", "Value", 1)] = base[c] * base["NLC"]

    # Live LP: mean final_price per FSN x DS (business_zone)
    step(4, "Load Live Prices (LP, MRP) - largest file, takes a few minutes")
    parts = []                          # chunked: the live-price file is several GB
    for ch in read_chunks(files["Live_Prices"], "live prices", low_memory=False,
                          usecols=["fsn", "business_zone", "final_price", "mrp"]):
        ch = ch[_s(ch["fsn"]).isin(fsns)]               # drop unneeded rows before anything else
        ch = ch.assign(fsn=_s(ch["fsn"]), ds=_s(ch["business_zone"]),
                       final_price=pd.to_numeric(ch["final_price"], errors="coerce"),
                       mrp=pd.to_numeric(ch["mrp"], errors="coerce"))
        ch = ch[ch["ds"].isin(dss)]
        # reduce each chunk to FSN x DS sums/counts so memory stays small
        parts.append(ch.groupby(["fsn", "ds"], as_index=False).agg(
            lp_sum=("final_price", "sum"), lp_n=("final_price", "count"),
            mrp_sum=("mrp", "sum"), mrp_n=("mrp", "count")))
    lp = pd.concat(parts, ignore_index=True).groupby(["fsn", "ds"], as_index=False).sum()
    lp["Live LP"] = lp["lp_sum"] / lp["lp_n"].where(lp["lp_n"] > 0)
    lp["MRP"] = (lp["mrp_sum"] / lp["mrp_n"].where(lp["mrp_n"] > 0)).round(0).astype("Int64")  # 0 decimals
    lp = lp[["fsn", "ds", "Live LP", "MRP"]]
    base = base.merge(lp, on=["fsn", "ds"], how="left")

    # Daily units over the user-selected period
    step(5, "Load daily sales and compute DRR / % Achievement")
    day_ids = {int(d.strftime("%Y%m%d")): d.strftime("%d-%b-%y") for d in days}
    parts = []
    for ch in read_chunks(files["Sales_Day"], "sales", low_memory=False,
                          usecols=["fsn", "store", "order_date", "units"]):
        ch = ch[ch["order_date"].isin(day_ids)]
        ch = ch.assign(fsn=_s(ch["fsn"]), store=_s(ch["store"]),
                       units=pd.to_numeric(ch["units"], errors="coerce").fillna(0))
        ch = ch[ch["fsn"].isin(fsns) & ch["store"].isin(dss)]
        parts.append(ch.groupby(["fsn", "store", "order_date"], as_index=False)["units"].sum())
    sales = pd.concat(parts, ignore_index=True).groupby(["fsn", "store", "order_date"], as_index=False)["units"].sum()
    sales["day"] = sales["order_date"].map(day_ids)
    piv = (sales.pivot_table(index=["fsn", "store"], columns="day", values="units", aggfunc="sum")
                .reindex(columns=list(day_ids.values())).reset_index().rename(columns={"store": "ds"}))
    base = base.merge(piv, on=["fsn", "ds"], how="left")
    day_cols = list(day_ids.values())
    base[day_cols] = base[day_cols].fillna(0)
    base[f"Sales Sum ({len(days)} days)"] = base[day_cols].sum(axis=1)
    base[f"Avg DRR ({len(days)} days)"] = base[day_cols].mean(axis=1)
    base["% Achievement"] = base[f"Sales Sum ({len(days)} days)"] / base["Total Risk ATP"].where(
        base["Total Risk ATP"] > 0)                          # blank where no risk ATP
    return base, ipc_date


# ---------------------------------------------------------------------------
# Step 4 â€” Export
# ---------------------------------------------------------------------------
def export(df: pd.DataFrame, ipc_date: datetime, days: list[datetime]) -> Path:
    step(6, "Export CSV")
    cols = (["fsn", "ds", "City", "analytic_vertical", "Sell BU", "brand", "product_title", "NLC"]
            + ["Total Risk ATP"] + [f"Risk ATP ({b} days)" for b in BUCKETS]
            + ["Total Risk Value"] + [f"Risk Value ({b} days)" for b in BUCKETS]
            + ["Live LP", "MRP"] + [c for c in df.columns if re.fullmatch(r"\d{2}-\w{3}-\d{2}", c)]
            + [c for c in df.columns if c.startswith("Sales Sum (")]
            + [c for c in df.columns if c.startswith("Avg DRR (")]
            + ["% Achievement"])
    out = df[[c for c in cols if c in df.columns]].sort_values("Total Risk ATP", ascending=False).copy()
    out["% Achievement"] = out["% Achievement"].round(4)    # numeric ratio (0.85 = 85%); format as % in Excel
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_DIR / (f"{ipc_date:%Y%m%d}_liquidation_output_sales_{days[0]:%d%b%y}-{days[-1]:%d%b%y}"
                         f"_{datetime.now():%H%M%S}.csv")
    out.to_csv(path, index=False)
    log(f"Exported {len(out):,} rows -> {path}")
    return path


if __name__ == "__main__":
    auto = "--yes" in sys.argv
    step(1, "Select files and sales period")
    f = select_files(auto)
    period = select_sales_period(f["Sales_Day"], auto)
    if not auto and input("\nConfirm files and period? [Y/n]: ").strip().lower() not in ("", "y", "yes"):
        raise SystemExit("Aborted.")
    _T0 = datetime.now()                  # restart the clock after the user prompts
    result, d = build(f, period)
    export(result, d, period)
    log(f"[Done] 100% complete, total run time {int((datetime.now() - _T0).total_seconds())}s")

