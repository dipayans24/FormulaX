"""
Streamlit app: Excel-style aggregations (max/sum/min/count, sumifs, countifs, xlookup)
between a LEFT table and a RIGHT table.

Install : pip install streamlit pandas openpyxl xlrd
Run     : streamlit run lookup_agg_app.py
"""

import uuid                                   # used to build random "NA" placeholder tokens
from io import BytesIO                        # in-memory file buffer (for reading uploads / writing Excel)

import pandas as pd                           # all the data work
import streamlit as st                        # the UI

st.set_page_config(page_title="Lookup & Aggregation Tool", layout="wide")  # wide layout so columns fit

BASIC_AGGS = ["max", "sum", "min", "count"]               # plain aggregations (no second table needed)
LOOKUP_AGGS = ["sumifs", "countifs", "xlookup"]           # aggregations that use the right table
ALL_AGGS = BASIC_AGGS + LOOKUP_AGGS                       # full dropdown list
SEP = "\x1f"                                              # invisible separator used to join multi-condition keys
NO_SORT = "(none - keep file order)"                      # dropdown label meaning "do not sort"


# ----------------------------------------------------------------------------
# File reading helpers
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)                        # cache so reruns don't re-read the file
def list_sheets(data: bytes, filename: str):
    if filename.lower().endswith(".csv"):                 # CSV files have no sheets
        return []                                         # empty list = "no sheet choice needed"
    return pd.ExcelFile(BytesIO(data)).sheet_names        # for Excel, return all sheet names


@st.cache_data(show_spinner=False)                        # cache the parsed table
def read_table(data: bytes, filename: str, sheet):
    if filename.lower().endswith(".csv"):                 # decide reader using the file extension
        try:
            df = pd.read_csv(BytesIO(data))               # try the default utf-8 first
        except UnicodeDecodeError:
            df = pd.read_csv(BytesIO(data), encoding="latin-1")  # fall back for odd encodings
    else:
        df = pd.read_excel(BytesIO(data), sheet_name=sheet)      # read the chosen sheet
    df.columns = df.columns.astype(str)                   # make sure every column name is a string
    return df                                             # return the table


def pick_table(label: str, uploaded, key: str):
    """Show sheet dropdown (if needed) and return the DataFrame for an uploaded file."""
    if uploaded is None:                                  # nothing uploaded yet
        return None
    data = uploaded.getvalue()                            # raw bytes of the upload
    sheets = list_sheets(data, uploaded.name)             # sheet names (empty for csv)
    sheet = None                                          # default: no sheet (csv)
    if len(sheets) > 1:                                   # several sheets -> let the user choose
        sheet = st.selectbox(f"{label} - select sheet", sheets, key=key)
    elif len(sheets) == 1:                                # a single sheet -> use it directly
        sheet = sheets[0]
    return read_table(data, uploaded.name, sheet)         # parse and return


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


def apply_xlookup(left, right, cfg):
    lkey, rkey = build_keys(left, right, cfg["pairs"])    # matching keys
    r = right.copy()                                      # work on a copy of the right table
    r["_key"] = rkey                                      # attach the key to each right row
    if cfg["sort"] != NO_SORT:                            # sort only if a sort column was chosen
        r = r.sort_values(cfg["sort"], ascending=cfg["asc"], key=smart_sort_key,
                          kind="stable", na_position="last")  # best row ends up first
    r = r.drop_duplicates("_key", keep="first")           # keep ONE row per key (the first after sorting)
    lookup = pd.Series(r[cfg["ret"]].values, index=r["_key"].values)  # key -> return value
    result = lkey.map(lookup)                             # fetch the value for every left row
    found = lkey.isin(lookup.index)                       # which left rows actually matched
    return result.where(found, cfg["not_found"])          # unmatched rows get the "not found" text


def run_all(left, right, configs):
    """Apply every configured aggregation in order and build the output workbook."""
    out = left.copy()                                     # original left table stays untouched
    used_right = []                                       # right-table columns used (for sheet 2)
    scalars = []                                          # axis=0 results to show in the UI
    for cfg in configs:                                   # one pass per aggregation block
        agg = cfg["agg"]
        if agg in BASIC_AGGS:                             # plain max/sum/min/count
            col, scalar = apply_basic(out, cfg)           # (uses earlier generated columns too)
            if scalar is not None:
                scalars.append((cfg["name"], scalar))     # remember axis=0 result
        else:                                             # lookup-type aggregations need the right table
            if agg == "sumifs":
                col = apply_sumifs(out, right, cfg)
            elif agg == "countifs":
                col = apply_countifs(out, right, cfg)
            else:
                col = apply_xlookup(out, right, cfg)
            needed = [p[1] for p in cfg["pairs"]]         # lookup columns
            needed += [cfg["ret"]] if cfg.get("ret") else []  # return column
            needed += [cfg["sort"]] if cfg.get("sort") not in (None, NO_SORT) else []  # sort column
            for c in needed:                              # collect unique, keep order
                if c not in used_right:
                    used_right.append(c)
        out[cfg["name"]] = col.values                     # add the new column at the far right
    buf = BytesIO()                                       # in-memory Excel file
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="Left_Output", index=False)          # sheet 1: left + new columns
        if used_right:                                    # sheet 2: raw right columns that were used
            right[used_right].to_excel(writer, sheet_name="Right_Used", index=False)
    return out, buf.getvalue(), scalars


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
    st.markdown(f"#### Aggregation #{i + 1}")             # block heading
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
        o = st.columns(4)
        order = o[0].radio("Sorting order (first row kept)", ["Ascending", "Descending"], key=f"ord_{i}", horizontal=True)
        cfg["asc"] = order == "Ascending"
        cfg["not_found"] = o[1].text_input("If not found, show", "#N/A", key=f"nf_{i}")

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
    return errors


# ----------------------------------------------------------------------------
# Main page
# ----------------------------------------------------------------------------
st.title("Lookup & Aggregation Tool")                     # page title

up_l, up_r = st.columns(2)                                # two upload areas side by side
with up_l:
    f_left = st.file_uploader("Upload File for Left Table", type=["csv", "xlsx", "xlsm", "xls"], key="up_left")
    left_df = pick_table("Left table", f_left, "sheet_left")   # sheet dropdown + parse
with up_r:
    f_right = st.file_uploader("Upload File for Right Table", type=["csv", "xlsx", "xlsm", "xls"], key="up_right")
    right_df = pick_table("Right table", f_right, "sheet_right")

if left_df is None:                                       # left table is mandatory
    st.info("Upload the left table to begin.")
    st.stop()
needs_right = True                                        # right table is only needed for lookup-type aggs
right_cols = list(right_df.columns) if right_df is not None else []  # right-table column list

configs, taken, i = [], list(left_df.columns), 0          # `taken` = left columns + generated ones so far
while True:                                               # keep drawing blocks while "add another" is ticked
    st.divider()
    available_left = list(taken)                          # left columns usable in this block (incl. generated)
    cfg = render_block(i, available_left, right_cols, taken)
    configs.append(cfg)
    if not st.checkbox("Add another aggregation", key=f"more_{i}"):
        break                                             # stop when the box is not ticked
    i += 1

st.divider()
if st.button("Generate", type="primary"):                 # run everything
    problems = validate(configs)                          # basic checks first
    if any(c["agg"] in LOOKUP_AGGS for c in configs) and right_df is None:
        problems.append("Upload the right table - it is needed for sumifs/countifs/xlookup.")
    if problems:
        for p in problems:
            st.error(p)
    else:
        with st.spinner("Calculating..."):
            out_df, xlsx_bytes, scalars = run_all(left_df, right_df, configs)
        st.session_state["result"] = {"df": out_df, "xlsx": xlsx_bytes, "scalars": scalars}  # keep across reruns

res = st.session_state.get("result")                      # show last result (survives download-button reruns)
if res:
    st.success(f"Done: {len(res['df'])} rows, {res['df'].shape[1]} columns.")
    for name, value in res["scalars"]:                    # column-wise (axis=0) results
        st.metric(name, value)
    st.dataframe(res["df"].head(100))                     # preview of the output
    st.download_button("Download Excel", res["xlsx"], file_name="output.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
