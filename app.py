import streamlit as st
import pandas as pd
import math
from io import BytesIO
from datetime import datetime, timedelta

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="HOI Flash Quarterly Inventory Projection",
    page_icon="📈",
    layout="wide"
)


def get_last_monday_of_month(year, month):
    """Finds the last Monday of any given month and year."""
    if month == 12:
        last_day = datetime(year + 1, 1, 1).toordinal() - 1
    else:
        last_day = datetime(year, month + 1, 1).toordinal() - 1
    last_date = datetime.fromordinal(last_day)
    offset = last_date.weekday()  # Monday=0
    return (last_date - timedelta(days=offset)).date()


def process_data(all_sheets_dict, selected_sheet):
    try:
        df = all_sheets_dict[selected_sheet].copy()
        df = df.fillna(0)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'])

        df = df.sort_values(by=['Part Number', 'Date']).reset_index(drop=True)

        unconf_col = 'SupplierHP(UnconfirmedOrders)'
        conf_col = 'SupplierHP(ConfirmedOrders)'

        deductions = ['POR demand', 'PO vs POR adustment', 'Backlog', 'Build and Hold',
                      'Pre-build', "Test Req't", 'Strategic Buffer', 'DCR', 'Others (PP,NPI build, etc)']
        forward_cols = ['POR demand', 'PO vs POR adustment', 'Backlog', 'Build and Hold',
                        'Pre-build', "Test Req't"]

        final_results = []
        for part, group in df.groupby('Part Number'):
            records = group.to_dict('records')
            num_rows = len(records)
            prev_balance = None

            for i in range(num_rows):
                row = records[i]
                moq = max(float(row.get('MOQ', 1)), 1)
                n_weeks = int(row.get('n', 5))
                start_val = float(row.get('Onhand', 0)) if i == 0 else prev_balance
                current_deductions = sum(float(row.get(c, 0)) for c in deductions if c in row)
                c_val = float(row[conf_col])

                if i + n_weeks < num_rows:
                    target_sum = sum(sum(float(records[j].get(c, 0)) for c in forward_cols if c in records[j]) for j in
                                     range(i + 1, i + 1 + n_weeks))
                    base_bal = start_val + c_val - current_deductions
                    gap = target_sum - base_bal
                    k = math.floor(gap / moq) + 1
                    row[unconf_col] = k * moq

                final_supply = float(row[unconf_col]) + c_val
                this_week_balance = start_val + final_supply - current_deductions
                row['Calculated_Balance'] = this_week_balance
                prev_balance = this_week_balance

                if i + n_weeks < num_rows:
                    future_sum = sum(sum(float(records[j].get(c, 0)) for c in forward_cols if c in records[j]) for j in
                                     range(i + 1, i + 1 + n_weeks))
                    row['WOS_Check'] = (this_week_balance / future_sum) * n_weeks if future_sum > 0 else 999
            final_results.extend(records)

        optimized_df = pd.DataFrame(final_results)
        optimized_df['Date_Only'] = optimized_df['Date'].dt.date
        all_sheets_dict[selected_sheet] = optimized_df.drop(columns=['Date_Only'])

        # Summary Report
        summary_df = pd.DataFrame()
        if 'Master' in all_sheets_dict:
            master_df = all_sheets_dict['Master']
            price_col_list = [c for c in master_df.columns if 'Cost' in str(c)]
            if price_col_list:
                price_col = price_col_list[0]
                price_lookup = master_df.set_index('HPPN')[price_col].to_dict()
                years_months = optimized_df['Date'].dt.to_period('M').unique()
                target_dates = [get_last_monday_of_month(ym.year, ym.month) for ym in years_months]

                summary_rows = []
                for part, group in optimized_df.groupby('Part Number'):
                    unit_price = price_lookup.get(part, 0)
                    for t_date in target_dates:
                        match = group[group['Date_Only'] == t_date]
                        bal = match['Calculated_Balance'].iloc[0] if not match.empty else 0
                        summary_rows.append({
                            'Part Number': part,
                            'Snapshot Date': t_date,
                            'Month': t_date.strftime('%Y-%m'),
                            'Balance': bal,
                            'Unit Price': unit_price,
                            'Amount': bal * unit_price
                        })
                summary_df = pd.DataFrame(summary_rows)
                all_sheets_dict['Summary Report'] = summary_df

        stats = {
            "total_parts": optimized_df['Part Number'].nunique(),
            "avg_wos": optimized_df['WOS_Check'].mean() if 'WOS_Check' in optimized_df else 0
        }
        return all_sheets_dict, stats

    except Exception as e:
        st.error(f"Logic Error: {e}")
        return None, None


# --- UI LOGIC ---
st.title("📈 HOI Flash Quarterly Inventory Projection")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Use session_state to keep results alive across re-renders
    if "results" not in st.session_state:
        st.session_state.results = None
    if "stats" not in st.session_state:
        st.session_state.stats = None

    xl_file = pd.ExcelFile(uploaded_file)
    sheet_names = xl_file.sheet_names

    with st.sidebar:
        st.header("Control Panel")
        selected_sheet = st.selectbox("Target Sheet", options=sheet_names)
        if st.button("🚀 Process Data", use_container_width=True):
            with st.spinner("Calculating..."):
                all_sheets = {name: xl_file.parse(name) for name in sheet_names}
                updated_dict, stats = process_data(all_sheets, selected_sheet)
                st.session_state.results = updated_dict
                st.session_state.stats = stats

    # Display results if they exist in session_state
    if st.session_state.results:
        res_dict = st.session_state.results
        res_stats = st.session_state.stats

        tab1, tab2 = st.tabs(["📊 Planning Preview", "💰 Monthly Summary"])

        with tab1:
            st.subheader("Global Metrics")
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("Unique Parts", res_stats["total_parts"])
            m_col2.metric("Avg WOS", f"{res_stats['avg_wos']:.2f}")
            st.dataframe(res_dict[selected_sheet].head(100), use_container_width=True)

        with tab2:
            if 'Summary Report' in res_dict:
                sum_df = res_dict['Summary Report']

                # --- PRETTY UI FOR FILTERS ---
                st.markdown("### 🔍 Financial Analysis Filters")
                available_months = sorted(sum_df['Month'].unique())
                selected_months = st.multiselect(
                    "Filter by Month(s):",
                    options=available_months,
                    default=available_months
                )

                filtered_df = sum_df[sum_df['Month'].isin(selected_months)]

                # --- METRICS BAR ---
                st.markdown("---")
                total_amt = filtered_df['Amount'].sum()

                # Subtotals in a clean grid
                if selected_months:
                    st.write("**Amount Sum per Snapshot Date:**")
                    # Dynamically adjust columns for wrapping if many months are selected
                    date_totals = filtered_df.groupby('Snapshot Date')['Amount'].sum()
                    cols = st.columns(min(len(date_totals), 4))  # Max 4 per row
                    for idx, (d, amt) in enumerate(date_totals.items()):
                        cols[idx % 4].info(f"**{d}**\n\n ${amt:,.2f}")

                st.markdown("---")
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("Master sheet not found.")

        # Download Section (always visible after processing)
        st.divider()
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for name, df_to_save in res_dict.items():
                df_to_save.to_excel(writer, index=False, sheet_name=name)

        st.download_button("📥 Download Final Excel", data=output.getvalue(),
                           file_name=f"WOS_Audited_{uploaded_file.name}", use_container_width=True)
else:
    st.info("Please upload an Excel file and click 'Process Data' in the sidebar.")