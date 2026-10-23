"""
Streamlit app: Excel-style aggregations (max/sum/min/count, sumifs, countifs, xlookup)
between a LEFT table and a RIGHT table.

Install : pip install streamlit pandas openpyxl xlrd
Run     : streamlit run lookup_agg_app.py
"""

import time                                   # used for the elapsed-time messages
import re                                     # used to translate Excel-style date formats
import uuid                                   # used to build random "NA" placeholder tokens
from io import BytesIO                        # in-memory file buffer (for reading uploads / writing Excel)

import pandas as pd                           # all the data work
import streamlit as st                        # the UI

st.set_page_config(page_title="🧮 Formula X - Lookup & Aggregation Tool", layout="wide")  # wide layout so columns fit


def blue_divider():
    """Horizontal line in blue (st.divider() has no colour option, so a coloured <div> bar is used)."""
    st.markdown('<div style="height:3px; width:100%; background-color:#1E88E5; margin:1rem 0;"></div>', unsafe_allow_html=True)  # a 3px tall blue bar (a plain div renders more reliably than <hr>)


BASIC_AGGS = ["max", "sum", "min", "count"]               # plain aggregations (no second table needed)
LOOKUP_AGGS = ["sumifs", "countifs", "xlookup"]           # aggregations that use the right table
ALL_AGGS = BASIC_AGGS + LOOKUP_AGGS                       # full dropdown list
SEP = "\x1f"                                              # invisible separator used to join multi-condition keys
NO_SORT = "(none - keep file order)"                      # dropdown label meaning "do not sort"


# ----------------------------------------------------------------------------
# File reading helpers (headers are read first, the full data only when really needed)
# ----------------------------------------------------------------------------
def file_key(uploaded):
    """Cheap identity of an upload, used as the cache key (avoids hashing the whole file on every rerun)."""
    return (uploaded.name, uploaded.size, getattr(uploaded, "file_id", None))  # name + size + upload id


@st.cache_data(show_spinner=False)                        # cache sheet names per file
def list_sheets(_data, fkey):                             # `_data` (leading underscore) is NOT hashed by streamlit
    if fkey[0].lower().endswith(".csv"):                  # CSV files have no sheets
        return []                                         # empty list = "no sheet choice needed"
    return pd.ExcelFile(BytesIO(_data)).sheet_names       # for Excel, return all sheet names


@st.cache_data(show_spinner=False)                        # cache the header row per file + sheet
def read_headers(_data, fkey, sheet):
    """Read ONLY the header row (0 data rows) - enough to fill the dropdowns."""
    if fkey[0].lower().endswith(".csv"):                  # decide the reader using the file extension
        try:
            df = pd.read_csv(BytesIO(_data), nrows=0)     # header only
        except UnicodeDecodeError:
            df = pd.read_csv(BytesIO(_data), nrows=0, encoding="latin-1")  # fall back for odd encodings
    else:
        df = pd.read_excel(BytesIO(_data), sheet_name=sheet, nrows=0)      # header only
    return [str(c) for c in df.columns]                   # column names as text


@st.cache_resource(show_spinner=False, max_entries=3)     # keep the last few full tables (no copying; we never modify them)
def read_table(_data, fkey, sheet, usecols=None):
    """Read the full table. `usecols` = tuple of column POSITIONS to read (None = all columns)."""
    cols = list(usecols) if usecols is not None else None  # pandas wants a list
    if fkey[0].lower().endswith(".csv"):                  # CSV
        try:
            df = pd.read_csv(BytesIO(_data), usecols=cols)  # try the default utf-8 first
        except UnicodeDecodeError:
            df = pd.read_csv(BytesIO(_data), usecols=cols, encoding="latin-1")  # fall back for odd encodings
    else:                                                 # Excel
        df = pd.read_excel(BytesIO(_data), sheet_name=sheet, usecols=cols)     # read the chosen sheet
    df.columns = df.columns.astype(str)                   # make sure every column name is a string
    return df                                             # return the table


def pick_source(label: str, uploaded, key: str):
    """Show the sheet dropdown (if needed) and return a light-weight handle: file bytes + sheet + headers."""
    if uploaded is None:                                  # nothing uploaded yet
        return None
    data = uploaded.getvalue()                            # raw bytes of the upload
    fkey = file_key(uploaded)                             # cache key for this file
    sheets = list_sheets(data, fkey)                      # sheet names (empty for csv)
    sheet = None                                          # default: no sheet (csv)
    if len(sheets) > 1:                                   # several sheets -> let the user choose
        sheet = st.selectbox(f"{label} - select sheet", sheets, key=key)
    elif len(sheets) == 1:                                # a single sheet -> use it directly
        sheet = sheets[0]
    headers = read_headers(data, fkey, sheet)             # only the column names are read here
    return {"data": data, "fkey": fkey, "name": uploaded.name, "sheet": sheet, "headers": headers}


def load_df(src, idx=None):
    """Read the full table (optionally only the columns at positions `idx`) for a handle from pick_source."""
    return read_table(src["data"], src["fkey"], src["sheet"], idx)


# ----------------------------------------------------------------------------
# Key cleaning helpers (case-insensitive, whitespace-insensitive, "1.0" -> "1")
# ----------------------------------------------------------------------------
def clean_series(s: pd.Series) -> pd.Series:
    """Convert a column to clean lowercase strings; blanks/NaN stay as <NA>."""
    if pd.api.types.is_float_dtype(s):                    # floats need special care (1.0 -> "1")
        whole = s.notna() & (s % 1 == 0)                  # True where the float is a whole number
        as_int = s.where(whole).astype("Int64").astype("string")  # whole floats as integer text (no ".0", no sci-notation)
        s = s.astype("string").mask(whole, as_int)        # use integer text where whole, normal text elsewhere
    else:
        s = s.astype("string")                            # everything else -> pandas string dtype
    s = s.str.strip().str.lower()                         # remove whitespace + lowercase
    s = s.str.replace(r"^(-?\d+)\.0+$", r"\1", regex=True)  # "123.0" stored as text -> "123"
    return s.mask(s.str.len() == 0)                       # empty strings count as missing


def make_token(used: set) -> str:
    """Random text guaranteed not to be present in `used`."""
    while True:                                           # loop until we find an unused token
        token = "na_" + uuid.uuid4().hex                  # random hex string
        if token not in used:                             # make sure it is not already a real value
            return token


def build_keys(left, right, pairs):
    """Build one matching key per row for the left and right tables.
    pairs = [(left_search_col, right_lookup_col), ...]  (multiple pairs = multiple conditions)."""
    lparts, rparts = [], []                               # cleaned key pieces per condition
    for s_col, l_col in pairs:                            # go through each condition
        lc = clean_series(left[s_col])                    # cleaned search-key column (left)
        rc = clean_series(right[l_col])                   # cleaned lookup column (right)
        used = set(lc.dropna()) | set(rc.dropna())        # every real value present on either side
        ltok = make_token(used)                           # NA placeholder for left
        rtok = make_token(used | {ltok})                  # DIFFERENT NA placeholder for right, so NA never matches NA
        lparts.append(lc.fillna(ltok))                    # fill left NAs
        rparts.append(rc.fillna(rtok))                    # fill right NAs
    lkey, rkey = lparts[0], rparts[0]                     # start with the first condition
    for lp, rp in zip(lparts[1:], rparts[1:]):            # join any extra conditions with the separator
        lkey = lkey + SEP + lp
        rkey = rkey + SEP + rp
    return lkey.astype(object), rkey.astype(object)       # object dtype -> safest for map/groupby


# ----------------------------------------------------------------------------
# Aggregation engines (all vectorised)
# ----------------------------------------------------------------------------
def apply_basic(df, cfg):
    """max/sum/min/count either row-wise (axis=1) or column-wise (axis=0). Returns (column, scalar_or_None)."""
    agg = cfg["agg"]                                      # which aggregation
    if cfg["mode"] == "row":                              # axis = 1 : across several columns
        if agg == "count":                                # count = number of non-empty cells in the row
            return df[cfg["cols"]].notna().sum(axis=1), None
        num = df[cfg["cols"]].apply(pd.to_numeric, errors="coerce")  # force numeric, bad values -> NaN
        return getattr(num, agg)(axis=1), None            # num.max(axis=1) / sum / min
    col = df[cfg["col"]]                                  # axis = 0 : one single column
    if agg == "count":                                    # count of non-empty cells in the column
        value = int(col.notna().sum())
    else:
        value = getattr(pd.to_numeric(col, errors="coerce"), agg)()  # max/sum/min of the column
    return pd.Series(value, index=df.index), value        # same value on every row + the scalar itself


def apply_sumifs(left, right, cfg):
    lkey, rkey = build_keys(left, right, cfg["pairs"])    # matching keys for both tables
    vals = pd.to_numeric(right[cfg["ret"]], errors="coerce")  # numbers to be summed
    sums = vals.groupby(rkey).sum()                       # one total per unique key (the "group by")
    return lkey.map(sums).fillna(0)                       # look each left key up; no match -> 0


def apply_countifs(left, right, cfg):
    lkey, rkey = build_keys(left, right, cfg["pairs"])    # matching keys
    counts = rkey.value_counts()                          # how many times each key appears on the right
    return lkey.map(counts).fillna(0).astype(int)         # look up counts; no match -> 0


def smart_sort_key(s: pd.Series) -> pd.Series:
    """Sort numerically if the whole column is numeric, otherwise as-is."""
    num = pd.to_numeric(s, errors="coerce")               # try converting to numbers
    return num if num.notna().sum() == s.notna().sum() else s  # use numbers only if nothing was lost


# ---- Date / datetime support for the xlookup "sort" column ----
CUSTOM_FMT = "Custom..."                                  # dropdown entry that reveals a manual format box
DATE_FORMATS = [                                          # Excel-style presets shown in the dropdown
    "dd/mm/yyyy", "mm/dd/yyyy", "yyyy-mm-dd", "dd-mm-yyyy", "dd-mmm-yyyy", "d-mmm-yy", "dd mmmm yyyy",
    "dd/mm/yyyy hh:mm", "dd/mm/yyyy hh:mm:ss", "mm/dd/yyyy hh:mm AM/PM", "mm/dd/yyyy hh:mm:ss AM/PM",
    "yyyy-mm-dd hh:mm:ss", "yyyy-mm-dd hh:mm:ss.000", "dd-mmm-yyyy hh:mm:ss",
]
EXCEL_TOKEN_RE = re.compile(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM/PM|\.0+", re.I)  # Excel format pieces
TOKEN_MAP = {                                             # Excel token -> Python strptime directive
    "yyyy": "%Y", "yy": "%y", "mmmm": "%B", "mmm": "%b", "dddd": "%A", "ddd": "%a",
    "dd": "%d", "d": "%d", "ss": "%S", "s": "%S", "am/pm": "%p",
}


def excel_to_strptime(fmt: str) -> str:
    """Translate an Excel-style format (dd/mm/yyyy hh:mm:ss) into a Python one (%d/%m/%Y %H:%M:%S).
    A format that already contains '%' is assumed to be Python-style and is returned unchanged."""
    if "%" in fmt:                                        # user typed a Python format directly
        return fmt
    tokens = list(EXCEL_TOKEN_RE.finditer(fmt))           # every date/time piece found in the format
    has_ampm = any(t.group(0).lower() == "am/pm" for t in tokens)  # 12-hour clock if AM/PM is present
    out, last = [], 0                                     # output pieces + end of the previous token
    for idx, m in enumerate(tokens):                      # walk through the tokens in order
        out.append(fmt[last:m.start()])                   # keep the separators between tokens as they are
        low = m.group(0).lower()                          # token in lowercase for comparisons
        if low in ("m", "mm"):                            # Excel uses "mm" for BOTH month and minute
            prev_hour = idx > 0 and tokens[idx - 1].group(0).lower() in ("h", "hh")   # right after hours -> minutes
            next_sec = idx + 1 < len(tokens) and tokens[idx + 1].group(0).lower() in ("s", "ss")  # right before seconds -> minutes
            out.append("%M" if (prev_hour or next_sec) else "%m")  # minute or month
        elif low in ("h", "hh"):                          # hours
            out.append("%I" if has_ampm else "%H")        # 12-hour or 24-hour
        elif low.startswith("."):                         # ".000" -> fractional seconds
            out.append(".%f")
        else:                                             # all remaining tokens come from the lookup table
            out.append(TOKEN_MAP[low])
        last = m.end()                                    # remember where this token ended
    out.append(fmt[last:])                                # trailing literal text
    return "".join(out)                                   # final Python format string


def parse_dates(s: pd.Series, fmt: str, col: str) -> pd.Series:
    """Convert a column to datetimes using `fmt`; raise ValueError if any non-blank value does not fit."""
    if pd.api.types.is_datetime64_any_dtype(s):           # already real datetimes (e.g. from Excel) -> nothing to parse
        return s
    py_fmt = excel_to_strptime(fmt)                       # Excel-style -> Python-style
    txt = s.astype("string").str.strip()                  # work on clean text
    if "%f" in py_fmt:                                    # format has fractional seconds
        txt = txt.str.replace(r"(\.\d{6})\d+", r"\1", regex=True)  # cut nanosecond digits beyond 6 (not an error)
    else:                                                 # format has no fractional seconds
        txt = txt.str.replace(r"(?<=:\d{2})\.\d+", "", regex=True)   # ignore ".123" after seconds (not an error)
    parsed = pd.to_datetime(txt, format=py_fmt, errors="coerce")  # unparseable values become NaT
    bad = parsed.isna() & txt.notna() & (txt.str.len() > 0)       # non-blank values that failed
    if bad.any():                                         # stop and tell the user exactly what failed
        examples = ", ".join(repr(v) for v in txt[bad].unique()[:5])  # first few offending values
        raise ValueError(f"Sort column '{col}': {int(bad.sum())} value(s) do not match the date format "
                         f"'{fmt}' (e.g. {examples}). Pick another format or fix the column.")
    return parsed                                         # clean datetime column


def apply_xlookup(left, right, cfg):
    lkey, rkey = build_keys(left, right, cfg["pairs"])    # matching keys
    r = right.copy()                                      # work on a copy of the right table
    r["_key"] = rkey                                      # attach the key to each right row
    if cfg["sort"] != NO_SORT and cfg.get("is_date"):     # sort column is a date/datetime
        r["_sort"] = parse_dates(r[cfg["sort"]], cfg["date_fmt"], cfg["sort"])  # convert text -> datetime (or raise)
        r = r.sort_values("_sort", ascending=cfg["asc"], kind="stable", na_position="last")  # sort by real dates
    elif cfg["sort"] != NO_SORT:                          # normal (non-date) sort column
        r = r.sort_values(cfg["sort"], ascending=cfg["asc"], key=smart_sort_key,
                          kind="stable", na_position="last")  # best row ends up first
    r = r.drop_duplicates("_key", keep="first")           # keep ONE row per key (the first after sorting)
    lookup = pd.Series(r[cfg["ret"]].values, index=r["_key"].values)  # key -> return value
    result = lkey.map(lookup)                             # fetch the value for every left row
    found = lkey.isin(lookup.index)                       # which left rows actually matched
    return result.where(found, cfg["not_found"])          # unmatched rows get the "not found" text


def right_cols_needed(configs):
    """Right-table columns required by a list of aggregation configs (unique, in order of first use)."""
    needed = []                                           # result list
    for cfg in configs:                                   # go through every aggregation
        if cfg["agg"] in LOOKUP_AGGS:                     # only lookup-type aggregations use the right table
            cols = [p[1] for p in cfg["pairs"]]           # lookup columns
            cols += [cfg["ret"]] if cfg.get("ret") else []  # return column
            cols += [cfg["sort"]] if cfg.get("sort") not in (None, NO_SORT) else []  # sort column
            for c in cols:                                # keep unique, keep order
                if c not in needed:
                    needed.append(c)
    return needed                                         # list of column names


def run_segment(left, right, configs, log=print):
    """Apply the aggregations in order. Returns (output table, right columns used, column-wise scalars)."""
    out = left.copy()                                     # original left table stays untouched
    scalars = []                                          # axis=0 results to show in the UI
    n = len(configs)                                      # how many aggregations in total
    for k, cfg in enumerate(configs, 1):                  # one pass per aggregation block
        t0 = time.perf_counter()                          # start timer for this block
        agg = cfg["agg"]                                  # which aggregation
        log(f"⚙️ [{k}/{n}] {agg} → new column '{cfg['name']}'")  # progress message
        if agg in BASIC_AGGS:                             # plain max/sum/min/count
            col, scalar = apply_basic(out, cfg)           # (uses earlier generated columns too)
            if scalar is not None:
                scalars.append((cfg["name"], scalar))     # remember axis=0 result
        else:                                             # lookup-type aggregations need the right table
            log(f"   matching {len(out):,} left rows against {len(right):,} right rows")
            if agg == "sumifs":
                col = apply_sumifs(out, right, cfg)
            elif agg == "countifs":
                col = apply_countifs(out, right, cfg)
            else:
                col = apply_xlookup(out, right, cfg)
        out[cfg["name"]] = col.values                     # add the new column at the far right
        log(f"   ✔ done in {time.perf_counter() - t0:.2f}s")  # per-block timing
    return out, right_cols_needed(configs), scalars       # table, used right columns, scalars


def write_excel(out_df, stages, log=print):
    """Build the Excel file: left table + new columns on sheet 1, raw used right columns on the next sheets."""
    buf = BytesIO()                                       # in-memory Excel file
    try:
        import xlsxwriter                                 # noqa: F401  (much faster writer, if installed)
        engine = "xlsxwriter"
    except ImportError:
        engine = "openpyxl"                               # slower, but always available
    log(f"   writer engine: {engine}")                    # tell the user which engine is used
    right_stages = [g for g in stages if g.get("right_used") is not None]  # stages that used a right table
    with pd.ExcelWriter(buf, engine=engine) as writer:
        out_df.to_excel(writer, sheet_name="Left_Output", index=False)  # sheet 1: left + new columns
        for n, g in enumerate(right_stages, 1):           # one extra sheet per right table used
            if len(right_stages) == 1:                    # a single right table keeps the simple name
                name = "Right_Used"
            else:                                         # several -> number + file name (max 31 chars, no illegal chars)
                stem = re.sub(r"[\[\]:*?/\\]", "_", g["right_label"].rsplit(".", 1)[0])
                name = f"R{n}_{stem}"[:31]
            g["right_used"].to_excel(writer, sheet_name=name, index=False)  # raw right columns, no aggregation
    return buf.getvalue()                                 # bytes ready for download


# ----------------------------------------------------------------------------
# UI helpers
# ----------------------------------------------------------------------------
def unique_name(base: str, taken: list) -> str:
    """Return base, or base_1, base_2 ... if the name is already taken."""
    if base not in taken:                                 # free -> use as is
        return base
    i = 1
    while f"{base}_{i}" in taken:                         # find the first free suffix
        i += 1
    return f"{base}_{i}"


def name_box(i, base, taken, container):
    """Text box (pre-filled with a suggested name) for the new column; returns the final unique name."""
    default = unique_name(base, taken)                    # suggested name
    typed = container.text_input("New column name", value=default, key=f"name_{i}_{base}")  # user can edit
    final = unique_name(typed.strip() or default, taken)  # guarantee uniqueness
    taken.append(final)                                   # later blocks can now use this new column
    return final


def render_block(i, left_cols, right_cols, taken):
    """Draw the widgets for aggregation #i and return its config dict."""
    st.markdown(f"#### Aggregation #{st.session_state.get('n_saved', 0) + i % 100 + 1}")  # numbering continues after saved steps
    agg = st.selectbox("Select agg", ALL_AGGS, key=f"agg_{i}")  # which aggregation
    cfg = {"agg": agg}                                    # config being built

    if agg in BASIC_AGGS:                                 # ---- max / sum / min / count ----
        mode = st.radio("Apply", ["Row-wise (axis=1, across 2+ columns)", "Column-wise (axis=0, one column)"],
                        key=f"mode_{i}", horizontal=True)
        if mode.startswith("Row"):
            cfg["mode"] = "row"
            cfg["cols"] = st.multiselect("Columns (left table)", left_cols, key=f"cols_{i}")
            base = f"{agg}_" + "_".join(cfg["cols"])      # suggested name
        else:
            cfg["mode"] = "col"
            cfg["col"] = st.selectbox("Column (left table)", left_cols, key=f"col_{i}")
            base = f"{agg}_{cfg['col']}"
        if len(base) > 60:                                # keep suggested names short
            base = base[:60]
        cfg["name"] = name_box(i, base, taken, st)        # name text box
        return cfg

    multi = st.checkbox("Multiple conditions", key=f"multi_{i}")  # ---- sumifs / countifs / xlookup ----
    n = int(st.number_input("Number of conditions", 2, 6, 2, key=f"n_{i}")) if multi else 1
    pairs = []                                            # (search col, lookup col) per condition
    for j in range(n):
        c = st.columns(4)                                 # 4 slots: search | lookup | return | sort
        s_col = c[0].selectbox(f"\"search key\" column{j + 1} (left)", left_cols, key=f"s_{i}_{j}")
        l_col = c[1].selectbox(f"\"lookup key\" column{j + 1} (right)", right_cols, key=f"l_{i}_{j}")
        if j == 0 and agg in ("sumifs", "xlookup"):       # return column only once, on the first row
            cfg["ret"] = c[2].selectbox("\"return key\" column (right)", right_cols, key=f"r_{i}")
        if j == 0 and agg == "xlookup":                   # sort column only for xlookup
            cfg["sort"] = c[3].selectbox("\"sort\" column (right)", [NO_SORT] + right_cols, key=f"sort_{i}")
        pairs.append((s_col, l_col))
    cfg["pairs"] = pairs

    if agg == "xlookup":                                  # extra xlookup options
        o = st.columns(4, vertical_alignment="bottom")    # one row; widgets line up on their bottom edge
        order = o[0].selectbox("Sorting order (first row kept)", ["Ascending", "Descending"], key=f"ord_{i}")  # same height as the other inputs
        cfg["asc"] = order == "Ascending"                 # True = ascending
        cfg["not_found"] = o[1].text_input("If not found, show", "#N/A", key=f"nf_{i}")  # text for unmatched rows
        cfg["is_date"] = False                            # default: not a date column
        if cfg["sort"] != NO_SORT:                        # the checkbox only appears once a sort column is picked
            cfg["is_date"] = o[2].checkbox("Sort column is a date/datetime", key=f"isdate_{i}")  # sits in the same row
        if cfg["is_date"]:                                # format picker appears in the last slot of the SAME row
            choice = o[3].selectbox("Date format (Excel style)", DATE_FORMATS + [CUSTOM_FMT], key=f"dfmt_{i}")
            if choice == CUSTOM_FMT:                      # only when "Custom..." is picked, one extra box appears below it
                fmt = st.columns(4)[3].text_input("Custom format", key=f"cfmt_{i}", placeholder="dd/mm/yyyy hh:mm:ss",
                                                  help="Excel style (yyyy mm dd hh mm ss AM/PM, .000) or Python style (%d/%m/%Y %H:%M:%S).")
            else:
                fmt = choice                              # use the preset
            cfg["date_fmt"] = fmt.strip()                 # store the chosen format

    base = f"{agg}_" + "_".join(f"{s}_{l}" for s, l in pairs)  # suggested name from search + lookup columns
    base = base[:60]                                      # trim very long names
    cfg["name"] = name_box(i, base, taken, st)            # editable name box
    return cfg


def validate(configs):
    """Return a list of problems found in the configuration."""
    errors = []
    for k, cfg in enumerate(configs, 1):
        if cfg["agg"] in BASIC_AGGS and cfg["mode"] == "row" and len(cfg["cols"]) < 2:
            errors.append(f"Aggregation #{k}: pick at least 2 columns for row-wise.")
        if cfg["agg"] == "xlookup" and cfg.get("is_date") and not cfg.get("date_fmt"):
            errors.append(f"Aggregation #{k}: enter a date format for the sort column.")
    return errors


# ----------------------------------------------------------------------------
# Main page
# ----------------------------------------------------------------------------
st.title("🧮 Formula X - Lookup & Aggregation Tool")      # page title
blue_divider()                                           # blue line under the title

ss = st.session_state                                     # short alias for the session state
for _k, _v in {"stages": [], "saved_left": None, "seg": 0, "n_saved": 0, "left_id": None, "flash": None}.items():
    ss.setdefault(_k, _v)                                 # create the state keys on the first run


def reset_steps():
    """Forget all saved steps (also resets the aggregation widgets by changing their key prefix)."""
    ss["stages"] = []                                     # no saved steps
    ss["saved_left"] = None                               # go back to the original left table
    ss["n_saved"] = 0                                     # numbering restarts
    ss["seg"] += 1                                        # new widget keys -> fresh blocks
    ss.pop("result", None)                                # drop any old output


def process(configs, left_src, right_src, commit):
    """Run the current aggregations. commit=True -> 'Save step'; commit=False -> 'Generate' (writes the Excel)."""
    active = [c for c in configs if not (c["agg"] in BASIC_AGGS and c["mode"] == "row" and not c["cols"])]  # ignore untouched empty blocks
    problems = validate(active)                           # basic checks first
    if any(c["agg"] in LOOKUP_AGGS for c in active) and right_src is None:
        problems.append("Upload the right table - it is needed for sumifs/countifs/xlookup.")
    if not active and (commit or not ss["stages"]):       # nothing to save / generate
        problems.append("Nothing to do yet - configure at least one aggregation.")
    if problems:                                          # stop and show every problem
        for msg in problems:
            st.error(msg)
        return False
    t_all = time.perf_counter()                           # total timer
    with st.status("Processing...", expanded=True) as status:  # live progress box
        try:
            if ss["saved_left"] is not None:              # earlier steps were saved -> start from that table
                left_df = ss["saved_left"]
                st.write(f"📋 Using the saved table ({len(left_df):,} rows × {left_df.shape[1]} columns) - left file not re-read.")
            else:                                         # first step -> read the full left file now
                st.write("📥 Reading the full left table...")
                t = time.perf_counter()
                left_df = load_df(left_src)
                st.write(f"   ✔ {len(left_df):,} rows × {left_df.shape[1]} columns in {time.perf_counter() - t:.1f}s")
            right_df = None                               # right table only if a lookup aggregation needs it
            needed = right_cols_needed(active)            # right columns required by this step
            if needed:
                idx = tuple(right_src["headers"].index(c) for c in needed)  # their positions in the file
                st.write(f"📥 Reading the right table - only {len(needed)} column(s): {', '.join(needed)}")
                t = time.perf_counter()
                right_df = load_df(right_src, idx)        # read just those columns
                st.write(f"   ✔ {len(right_df):,} rows in {time.perf_counter() - t:.1f}s")
            if active:                                    # run the aggregations
                out_df, used, scalars = run_segment(left_df, right_df, active, log=st.write)
            else:                                         # Generate with nothing new: just export what was saved
                out_df, used, scalars = left_df, [], []
            stage = None                                  # snapshot of this step (kept for the Excel's right sheets)
            if active:
                label = right_src["name"] + (f" [{right_src['sheet']}]" if right_src["sheet"] else "") if right_src else ""
                stage = {"right_label": label, "right_used": right_df[used].copy() if used else None,
                         "new_cols": [c["name"] for c in active], "scalars": scalars}
            if commit:                                    # ---- Save step ----
                ss["stages"].append(stage)                # remember the step
                ss["saved_left"] = out_df                 # later steps build on this table
                ss["seg"] += 1                            # fresh aggregation blocks
                ss["n_saved"] += len(active)              # keep the numbering going
                ss.pop("result", None)                    # old output is now outdated
                ss["flash"] = (f"💾 Step {len(ss['stages'])} saved in {time.perf_counter() - t_all:.1f}s. "
                               "Swap the right table if needed and add more aggregations. The Excel is written when you click Generate.")
                status.update(label="Step saved", state="complete", expanded=False)
            else:                                         # ---- Generate ----
                all_stages = ss["stages"] + ([stage] if stage else [])  # saved steps + this one
                st.write("📝 Writing the Excel file (can take a while for large tables)...")
                t = time.perf_counter()
                xlsx = write_excel(out_df, all_stages, log=st.write)
                st.write(f"   ✔ Excel ready in {time.perf_counter() - t:.1f}s")
                ss["result"] = {"df": out_df, "xlsx": xlsx, "scalars": [x for g in all_stages for x in g["scalars"]]}
                status.update(label=f"Done in {time.perf_counter() - t_all:.1f}s", state="complete", expanded=False)
        except Exception as e:                            # show any failure inside the status box
            status.update(label="Failed", state="error", expanded=True)
            st.error(str(e) if isinstance(e, ValueError) else f"{type(e).__name__}: {e}")
            ss.pop("result", None)                        # drop a stale result
            return False
    return True


up_l, up_r = st.columns(2)                                # two upload areas side by side
with up_l:
    f_left = st.file_uploader("Upload File for Left Table", type=["csv", "xlsx", "xlsm", "xls"], key="up_left")
    left_src = pick_source("Left table", f_left, "sheet_left")    # sheet dropdown + header row only
with up_r:
    f_right = st.file_uploader("Upload File for Right Table", type=["csv", "xlsx", "xlsm", "xls"], key="up_right")
    right_src = pick_source("Right table", f_right, "sheet_right")

if left_src is None:                                      # left table is mandatory
    st.info("Upload the left table to begin.")
    st.stop()

left_id = (left_src["fkey"], left_src["sheet"])           # identity of the chosen left table
if ss["left_id"] != left_id:                              # a different left file/sheet -> saved steps no longer apply
    reset_steps()
    ss["left_id"] = left_id

if ss["flash"]:                                           # one-time message after "Save step"
    st.success(ss["flash"])
    ss["flash"] = None

if ss["stages"]:                                          # summary of the saved steps
    with st.expander(f"✅ Saved steps ({len(ss['stages'])})"):
        for n, g in enumerate(ss["stages"], 1):
            src_txt = f" · right table: `{g['right_label']}`" if g["right_label"] else ""
            st.markdown(f"**Step {n}**{src_txt} · new columns: {', '.join(g['new_cols'])}")
        if st.button("Reset all saved steps"):            # start over from the original left file
            reset_steps()
            st.rerun()

right_cols = list(right_src["headers"]) if right_src is not None else []  # right-table column list
base_cols = list(ss["saved_left"].columns) if ss["saved_left"] is not None else list(left_src["headers"])  # left columns incl. saved ones

configs, taken, i = [], list(base_cols), 0                # `taken` = left columns + generated ones so far
while True:                                               # keep drawing blocks while "add another" is ticked
    blue_divider()
    uid = ss["seg"] * 100 + i                             # unique widget id (changes after every saved step)
    available_left = list(taken)                          # left columns usable in this block (incl. generated)
    cfg = render_block(uid, available_left, right_cols, taken)
    configs.append(cfg)
    if not st.checkbox("Add another aggregation", key=f"more_{uid}"):
        break                                             # stop when the box is not ticked
    i += 1

blue_divider()
def stretch_button(col, label, **kwargs):
    """Button that fills its column (equal-width buttons); works on old and new Streamlit versions."""
    try:
        return col.button(label, width="stretch", **kwargs)              # newer Streamlit
    except TypeError:
        return col.button(label, use_container_width=True, **kwargs)     # older Streamlit


b1, b2, _ = st.columns([1.4, 1.4, 7], gap="small")        # two equal, narrow button columns placed side by side on the left
save_clicked = stretch_button(b1, "💾 Save step", help="Run the aggregations above and keep the result, so you can swap the right table and continue.")
gen_clicked = stretch_button(b2, "Generate", type="primary", help="Run everything and build the Excel file.")

if save_clicked:                                          # checkpoint: compute now, write the Excel later
    if process(configs, left_src, right_src, commit=True):
        st.rerun()                                        # refresh so the blocks reset for the next step
if gen_clicked:                                           # final output
    process(configs, left_src, right_src, commit=False)

res = ss.get("result")                                    # show last result (survives download-button reruns)
if res:
    st.success(f"Done: {len(res['df'])} rows, {res['df'].shape[1]} columns.")
    for name, value in res["scalars"]:                    # column-wise (axis=0) results
        st.metric(name, value)
    st.dataframe(res["df"].head(100))                     # preview of the output
    st.download_button("Download Excel", res["xlsx"], file_name="output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
