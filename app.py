import streamlit as st
import pandas as pd
import math
from io import BytesIO
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="HOI Flash Quarterly Inventory Projection",
    page_icon="📈",
    layout="wide"
)

# =============================================================================
# STEP 1 — Parse raw "testing data" block-format Excel
# =============================================================================

_ALL_DESC = [
    "POR demand", "PO vs POR adustment", "Backlog", "Build and Hold",
    "Pre-build", "Test Req't", "Strategic Buffer", "DCR",
    "Others (PP,NPI build, etc)", "RMA (QI)",
    "SupplierHP(UnconfirmedOrders)", "SupplierHP(ConfirmedOrders)",
]

_SHEET_CONFIGS = [
    ("FXN 2X",      ["FXN 2X CISS", "FXN2X CISS", "FXN 2X", "FXN2X",
                     "FXN 2X CISS with extra WOS", "FXN 2X CISS (normal WOS)"]),
    ("FXN 4X CISS", ["FXN 4X CISS", "FXN4X CISS", "FXN 4X", "FXN4X"]),
    ("NKG TH",      ["NKG TH", "NKGTH", "NKG"]),
    ("HQ",          ["HQ", "HQ_Trillium"]),
]

_DESC_NORM = {
    s.strip().upper().replace(" ","").replace("_","").replace("-",""): s
    for s in _ALL_DESC
}

def _n(s):
    return str(s).strip().upper().replace(" ","").replace("_","").replace("-","")

def _find_sheet(available, candidates):
    m = {_n(s): s for s in available}
    for c in candidates:
        hit = m.get(_n(c))
        if hit:
            return hit
    return None

def _to_float(v):
    try:
        f = float(v)
        return 0.0 if math.isnan(f) else f
    except:
        return 0.0

def _match_desc(raw):
    nrm = _n(raw)
    if nrm in _DESC_NORM:
        return _DESC_NORM[nrm]
    for k, v in _DESC_NORM.items():
        if k in nrm or nrm in k:
            return v
    return None

def _load_master_lookup(xl):
    sheet = _find_sheet(xl.sheet_names, ["Master", "MASTER", "master"])
    if not sheet:
        return {}
    raw = xl.parse(sheet, header=None)
    hrow = 0
    for r in range(min(6, len(raw))):
        vals = [_n(str(v)) for v in raw.iloc[r] if pd.notna(v)]
        if any("HPPN" in v or "PART" in v for v in vals):
            hrow = r
            break
    raw.columns = [str(v).strip() if pd.notna(v) else f"__c{i}"
                   for i, v in enumerate(raw.iloc[hrow])]
    data = raw.iloc[hrow+1:].reset_index(drop=True)
    def _col(kws):
        for col in data.columns:
            if any(k.upper() in col.upper() for k in kws):
                return col
        return None
    hppn_col  = _col(["HPPN","PART NUMBER","PARTNUMBER"])
    moq_col   = _col(["MOQ"])
    price_col = _col(["COST","PRICE","IPRICE"])
    result = {}
    if not hppn_col:
        return result
    for _, row in data.iterrows():
        pn = str(row[hppn_col]).strip() if pd.notna(row[hppn_col]) else ""
        if not pn or pn.lower() in ("nan","hppn","part number"):
            continue
        result[pn] = {
            "MOQ":    _to_float(row[moq_col])   if moq_col   else 0.0,
            "Iprice": _to_float(row[price_col]) if price_col else 0.0,
        }
    return result

def _load_sdos(xl):
    sheet = _find_sheet(xl.sheet_names, ["SDOS","sdos"])
    if not sheet:
        return {}
    raw  = xl.parse(sheet, header=None)
    PRIO = {"SG5HVC":0,"01EMVL":1,"02AMVC":2}
    best = {}
    for _, row in raw.iterrows():
        loc = _n(str(row.iloc[0])) if pd.notna(row.iloc[0]) else ""
        if loc not in PRIO:
            continue
        pn = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if not pn or pn.lower() in ("nan","product id"):
            continue
        raw_val = row.iloc[9] if len(row) > 9 else None
        if not pd.notna(raw_val):
            continue
        try:
            tdos = int(float(raw_val))
        except:
            continue
        cur_prio = PRIO.get(best.get(pn,{}).get("_loc","ZZZ"),99)
        if PRIO[loc] < cur_prio:
            best[pn] = {"tdos":tdos,"_loc":loc}
    return {pn: v["tdos"] for pn,v in best.items()}

def _parse_one_sheet(xl, sheet_name, master_lut, sdos_lut):
    raw = xl.parse(sheet_name, header=None)
    header_rows = [r for r in range(len(raw))
                   if _n(str(raw.iloc[r,1])) == "PARTNUMBER"]
    if not header_rows:
        return pd.DataFrame()
    records = []
    for hr in header_rows:
        date_cols = []
        for c in range(11, raw.shape[1]):
            v = raw.iloc[hr, c]
            if pd.notna(v) and hasattr(v,"year") and v.year >= 2020:
                date_cols.append((c, pd.Timestamp(v)))
        if not date_cols:
            continue
        mpa = pn = None
        onhand = 0.0
        data_map = {d:{} for d in _ALL_DESC}
        for r in range(hr+1, min(hr+14, len(raw))):
            row = raw.iloc[r]
            c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if mpa is None and c0 and c0.lower() not in ("nan","mpa",""):
                mpa = c0
            if pn is None and c1 and c1.lower() not in ("nan","part number",""):
                pn = c1
            desc_raw = row.iloc[9] if 9 < len(row) else None
            if not pd.notna(desc_raw):
                continue
            desc_str = str(desc_raw).strip()
            if _n(desc_str) == _n("Balance (Include UNCONFIRMED ORDERS)"):
                v10 = row.iloc[10] if 10 < len(row) else None
                if pd.notna(v10):
                    onhand = _to_float(v10)
                continue
            if _n(desc_str) == "WOS":
                continue
            canon = _match_desc(desc_str)
            if canon is None:
                continue
            for col_idx, date in date_cols:
                cell = row.iloc[col_idx] if col_idx < len(row) else None
                data_map[canon][date] = _to_float(cell) if pd.notna(cell) and cell != " " else 0.0
        if pn is None:
            continue
        if onhand == 0.0:
            inv = raw.iloc[hr+1, 2] if hr+1 < len(raw) else None
            if pd.notna(inv):
                onhand = _to_float(inv)
        pn_n = _n(pn)
        if pn in master_lut:
            moq = master_lut[pn]["MOQ"]
        else:
            hit = next((v for k,v in master_lut.items() if _n(k)==pn_n), None)
            moq = hit["MOQ"] if hit else 0.0
        if pn in sdos_lut:
            tdos = sdos_lut[pn]
        else:
            hit_t = next((v for k,v in sdos_lut.items() if _n(k)==pn_n), None)
            tdos = hit_t if hit_t else 0
        n_val = int(tdos//7+1) if tdos > 0 else 1
        for _, date in date_cols:
            rec = {"MPA":mpa,"Part Number":pn,"Date":date,
                   "MOQ":moq,"TDOS":tdos,"n":n_val,"Onhand":onhand}
            for desc in _ALL_DESC:
                rec[desc] = data_map[desc].get(date, 0.0)
            records.append(rec)
    return pd.DataFrame(records)

# =============================================================================
# STEP 2 — Flash logic
# =============================================================================

def _run_flash(df):
    df = df.fillna(0).copy()
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.sort_values(by=['Part Number','Date']).reset_index(drop=True)
    unconf_col   = 'SupplierHP(UnconfirmedOrders)'
    conf_col     = 'SupplierHP(ConfirmedOrders)'
    deductions   = ['POR demand','PO vs POR adustment','Backlog','Build and Hold',
                    'Pre-build',"Test Req't",'Strategic Buffer','DCR','Others (PP,NPI build, etc)']
    forward_cols = ['POR demand','PO vs POR adustment','Backlog','Build and Hold',
                    'Pre-build',"Test Req't"]
    final_results = []
    for part, group in df.groupby('Part Number'):
        records      = group.to_dict('records')
        num_rows     = len(records)
        prev_balance = None
        for i in range(num_rows):
            row         = records[i]
            moq         = max(float(row.get('MOQ',1)), 1)
            n_weeks     = int(row.get('n',5))
            start_val   = float(row.get('Onhand',0)) if i == 0 else prev_balance
            current_ded = sum(float(row.get(c,0)) for c in deductions if c in row)
            c_val       = float(row[conf_col])
            if i + n_weeks < num_rows:
                target_sum = sum(
                    sum(float(records[j].get(c,0)) for c in forward_cols if c in records[j])
                    for j in range(i+1, i+1+n_weeks)
                )
                base_bal        = start_val + c_val - current_ded
                k               = math.floor((target_sum - base_bal) / moq) + 1
                row[unconf_col] = k * moq
            final_supply          = float(row[unconf_col]) + c_val
            this_week_balance     = start_val + final_supply - current_ded
            row['Calculated_Balance'] = this_week_balance
            prev_balance          = this_week_balance
            if i + n_weeks < num_rows:
                future_sum = sum(
                    sum(float(records[j].get(c,0)) for c in forward_cols if c in records[j])
                    for j in range(i+1, i+1+n_weeks)
                )
                row['WOS'] = (this_week_balance / future_sum) * n_weeks if future_sum > 0 else 999
        final_results.extend(records)
    return pd.DataFrame(final_results)

# =============================================================================
# STEP 3 — Full pipeline: parse + flash + All MPA
# =============================================================================

def build_all_results(file_bytes):
    xl         = pd.ExcelFile(BytesIO(file_bytes))
    master_lut = _load_master_lookup(xl)
    sdos_lut   = _load_sdos(xl)
    sheet_dict = {}
    stats_dict = {}
    for label, candidates in _SHEET_CONFIGS:
        src = _find_sheet(xl.sheet_names, candidates)
        if src is None:
            continue
        raw_df = _parse_one_sheet(xl, src, master_lut, sdos_lut)
        if raw_df.empty:
            continue
        processed = _run_flash(raw_df)
        sheet_dict[label] = processed
        valid_wos = processed['WOS'].replace(999, pd.NA).dropna() if 'WOS' in processed.columns else pd.Series(dtype=float)
        stats_dict[label] = {
            "total_parts": processed['Part Number'].nunique(),
            "avg_wos":     float(valid_wos.mean()) if len(valid_wos) > 0 else 0.0,
        }
    if sheet_dict:
        all_mpa = pd.concat(list(sheet_dict.values()), ignore_index=True)
        sheet_dict["All MPA"] = all_mpa
        valid_all = all_mpa['WOS'].replace(999, pd.NA).dropna() if 'WOS' in all_mpa.columns else pd.Series(dtype=float)
        stats_dict["All MPA"] = {
            "total_parts": all_mpa['Part Number'].nunique(),
            "avg_wos":     float(valid_all.mean()) if len(valid_all) > 0 else 0.0,
        }
    return sheet_dict, stats_dict

# =============================================================================
# UI  (original app.py UI structure, preserved exactly)
# =============================================================================

def get_last_monday_of_month(year, month):
    if month == 12:
        last_day = datetime(year+1,1,1).toordinal()-1
    else:
        last_day = datetime(year,month+1,1).toordinal()-1
    last_date = datetime.fromordinal(last_day)
    offset = last_date.weekday()
    return (last_date - timedelta(days=offset)).date()


st.title("📈 HOI Flash Quarterly Inventory Projection")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Re-run pipeline only when a new file is uploaded
    file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("_file_key") != file_key:
        with st.spinner("⚙️ Processing all sheets…"):
            file_bytes = uploaded_file.read()
            sheet_dict, stats_dict = build_all_results(file_bytes)
        st.session_state["_file_key"]   = file_key
        st.session_state["_sheet_dict"] = sheet_dict
        st.session_state["_stats_dict"] = stats_dict

    sheet_dict = st.session_state["_sheet_dict"]
    stats_dict = st.session_state["_stats_dict"]

    if not sheet_dict:
        st.error(
            "❌ No recognisable MPA sheets found. "
            "Please upload a testing data file containing "
            "FXN 2X / FXN 4X CISS / NKG TH / HQ, plus Master and SDOS sheets."
        )
        st.stop()

    with st.sidebar:
        st.header("Control Panel")
        selected_sheet = st.selectbox("Target Sheet", options=list(sheet_dict.keys()))

    cur_stats = stats_dict.get(selected_sheet, {"total_parts":0,"avg_wos":0})

    tab1, tab2 = st.tabs(["📊 Planning Preview", "💰 Monthly Summary"])

    with tab1:
        st.subheader("Global Metrics")
        m_col1, m_col2 = st.columns(2)
        m_col1.metric("Unique Parts", cur_stats["total_parts"])
        m_col2.metric("Avg WOS", f"{cur_stats['avg_wos']:.2f}")
        st.dataframe(sheet_dict[selected_sheet].head(100), use_container_width=True)

    with tab2:
        xl_check = pd.ExcelFile(BytesIO(uploaded_file.getvalue()))
        master_sheet = _find_sheet(xl_check.sheet_names, ["Master","MASTER","master"])
        if master_sheet:
            master_df     = xl_check.parse(master_sheet)
            price_col_list = [c for c in master_df.columns if 'Cost' in str(c)]
            if price_col_list:
                price_col    = price_col_list[0]
                price_lookup = master_df.set_index('HPPN')[price_col].to_dict()
                opt_df = sheet_dict[selected_sheet].copy()
                opt_df['Date'] = pd.to_datetime(opt_df['Date'])
                opt_df['Date_Only'] = opt_df['Date'].dt.date
                years_months = opt_df['Date'].dt.to_period('M').unique()
                target_dates = [get_last_monday_of_month(ym.year, ym.month) for ym in years_months]
                summary_rows = []
                for part, group in opt_df.groupby('Part Number'):
                    unit_price = price_lookup.get(part, 0)
                    for t_date in target_dates:
                        match = group[group['Date_Only'] == t_date]
                        bal   = match['Calculated_Balance'].iloc[0] if not match.empty else 0
                        summary_rows.append({
                            'Part Number':   part,
                            'Snapshot Date': t_date,
                            'Month':         t_date.strftime('%Y-%m'),
                            'Balance':       bal,
                            'Unit Price':    unit_price,
                            'Amount':        bal * unit_price,
                        })
                summary_df = pd.DataFrame(summary_rows)

                st.markdown("### 🔍 Financial Analysis Filters")
                available_months = sorted(summary_df['Month'].unique())
                selected_months  = st.multiselect(
                    "Filter by Month(s):", options=available_months, default=available_months
                )
                filtered_df = summary_df[summary_df['Month'].isin(selected_months)]

                st.markdown("---")
                if selected_months:
                    st.write("**Amount Sum per Snapshot Date:**")
                    date_totals = filtered_df.groupby('Snapshot Date')['Amount'].sum()
                    cols = st.columns(min(len(date_totals), 4))
                    for idx, (d, amt) in enumerate(date_totals.items()):
                        cols[idx % 4].info(f"**{d}**\n\n ${amt:,.2f}")
                st.markdown("---")
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("No 'Cost' column found in Master sheet.")
        else:
            st.warning("Master sheet not found.")

    st.divider()
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df_to_save in sheet_dict.items():
            df_to_save.to_excel(writer, index=False, sheet_name=name)
    st.download_button(
        "📥 Download Final Excel",
        data=output.getvalue(),
        file_name=f"WOS_Audited_{uploaded_file.name}",
        use_container_width=True,
    )

else:
    st.info("Please upload an Excel file and click 'Process Data' in the sidebar.")
