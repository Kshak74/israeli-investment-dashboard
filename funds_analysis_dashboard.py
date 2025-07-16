import streamlit as st
import pandas as pd
import plotly.express as px
import collections

st.set_page_config(page_title="Quarterly Investment Analysis", page_icon="📊", layout="wide")

PROFESSIONAL_COLORS = [
    '#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087',
    '#f95d6a', '#ff7c43', '#ffa600', '#90be6d', '#43aa8b'
]

st.markdown("""
<style>
    .main { padding: 1rem 1rem; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%); color: #F3F6FB;}
    .stApp { max-width: 1200px; margin: 0 auto; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%);}
    .metric-value { font-size: 2rem; font-weight: bold; color: #67B7DC;}
    .metric-label { font-size: 1rem; color: #CED2D9;}
    .sub-header { font-size: 1.5rem; font-weight: bold; margin: 1rem 0; padding-bottom: 0.5rem; border-bottom: 1px solid #3A4055;}
    .insight-box { background-color: rgba(40, 45, 60, 0.8); border-left: 4px solid #67B7DC; padding: 1rem; margin: 1rem 0; border-radius: 0.25rem;}
    .insight-box h3 { margin-top: 0; color: #8AD6CC;}
    .stDataFrame { margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)

def format_number(num):
    if pd.isnull(num): return "-"
    if num >= 1_000_000_000: return f"{num/1_000_000_000:.2f}B"
    elif num >= 1_000_000: return f"{num/1_000_000:.2f}M"
    elif num >= 1_000: return f"{num/1_000:.2f}K"
    else: return f"{num:.2f}"

def ensure_unique_columns(df):
    if df is None or df.empty: return df
    if len(df.columns) != len(set(df.columns)):
        seen = collections.defaultdict(int)
        new_cols = []
        for col in df.columns:
            if col in seen:
                seen[col] += 1
                new_cols.append(f"{col}_{seen[col]}")
            else:
                seen[col] = 0
                new_cols.append(col)
        df.columns = new_cols
    return df

def normalize_columns(df):
    col_map = {
        'investment name': 'Investment Name',
        'שם קרן השקעה': 'Investment Name',
        'geography': 'Geography',
        'מדינה לפי חשיפה כלכלית': 'Geography',
        'strategy': 'Strategy',
        'אסטרטגיה': 'Strategy',
        'nav (ils)': 'NAV (ILS)',
        "שווי הוגן (באלפי ש\"ח)": 'NAV (ILS)',
        "שווי הוגן (באלפי ש''ח)": 'NAV (ILS)',
        "nav (במטבע הדיווח של קרן ההשקעה)": "NAV (OC)",
    }
    df = df.rename(columns={c: col_map.get(c.strip().lower(), c) for c in df.columns})
    return df

def load_data(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()
        df = normalize_columns(df)
        df = ensure_unique_columns(df)
        df = df.dropna(how='all').dropna(axis=1, how='all')
        return df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def merge_quarters(dfs, period_names):
    for i, (df, pname) in enumerate(zip(dfs, period_names)):
        df['Period'] = pname
    common_cols = set(dfs[0].columns)
    for df in dfs[1:]:
        common_cols &= set(df.columns)
    dfs_common = [df[list(common_cols) + ['Period']] for df in dfs]
    return pd.concat(dfs_common, ignore_index=True)

def create_dashboard(df):
    st.sidebar.title("Quarterly Analysis Filters")
    period_col = 'Period'
    geo_col = 'Geography'
    strat_col = 'Strategy'
    nav_col = 'NAV (ILS)'

    period_options = sorted(df[period_col].dropna().astype(str).unique())
    selected_periods = st.sidebar.multiselect("Select Period(s)", period_options, period_options)
    geo_options = sorted(df[geo_col].dropna().unique())
    strat_options = sorted(df[strat_col].dropna().unique())

    selected_geo = st.sidebar.multiselect("Geography", ['All']+geo_options, ['All'])
    selected_strat = st.sidebar.multiselect("Strategy", ['All']+strat_options, ['All'])

    filtered = df.copy()
    if selected_periods:
        filtered = filtered[filtered[period_col].astype(str).isin(selected_periods)]
    if selected_geo and 'All' not in selected_geo:
        filtered = filtered[filtered[geo_col].isin(selected_geo)]
    if selected_strat and 'All' not in selected_strat:
        filtered = filtered[filtered[strat_col].isin(selected_strat)]

    total_nav = filtered[nav_col].sum()
    total_inv = len(filtered)
    avg_inv = total_nav / total_inv if total_inv else 0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"<div class='metric-value'>{format_number(total_nav)} ILS</div><div class='metric-label'>Total NAV</div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='metric-value'>{total_inv}</div><div class='metric-label'>Total Investments</div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='metric-value'>{format_number(avg_inv)} ILS</div><div class='metric-label'>Average Investment Size</div>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["Quarterly Trends", "Geography", "Strategy"])

    with tab1:
        st.subheader("Trend of NAV (ILS) by Strategy Over Periods")
        trend = filtered.groupby(['Period', strat_col])[nav_col].sum().reset_index()
        fig = px.line(trend, x='Period', y=nav_col, color=strat_col, markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("NAV by Geography Over Periods")
        geo_trend = filtered.groupby(['Period', geo_col])[nav_col].sum().reset_index()
        fig = px.line(geo_trend, x='Period', y=nav_col, color=geo_col, markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.subheader("Strategy NAV in Selected Periods")
        data = filtered.groupby([strat_col, 'Period'])[nav_col].sum().reset_index()
        fig = px.bar(data, x='Period', y=nav_col, color=strat_col, barmode='group', color_discrete_sequence=PROFESSIONAL_COLORS)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Raw Data")
    st.dataframe(filtered, use_container_width=True)
    st.download_button("Download filtered data", filtered.to_csv(index=False).encode("utf-8-sig"), file_name="quarterly_filtered.csv", mime="text/csv")

def main():
    st.title("Investment Funds Quarterly Comparison")
    uploaded_files = st.sidebar.file_uploader("Upload Excel Files (quarterly)", type=["xlsx", "xls"], accept_multiple_files=True)
    period_names = []
    dfs = []

    REQUIRED_COLS = ['Investment Name', 'Geography', 'Strategy', 'NAV (ILS)']

    if uploaded_files:
        for upf in uploaded_files:
            df = load_data(upf)
            if df is not None:
                missing = [col for col in REQUIRED_COLS if col not in df.columns]
                if missing:
                    st.warning(f"File {upf.name} missing columns: {missing}. Please rename or map columns in the Excel file.")
                    continue
                # בחירת עמודת תקופה (Period)
                period_guess = None
                for col in df.columns:
                    if "רבעון" in col or "שנה" in col or "period" in col.lower() or "date" in col.lower():
                        period_guess = col
                        break
                if not period_guess:
                    period_guess = st.sidebar.text_input(f"Enter period/quarter column for file {upf.name}")
                period_value = df[period_guess].iloc[0] if period_guess in df.columns else st.sidebar.text_input(f"Enter period label for {upf.name}")
                period_names.append(str(period_value))
                dfs.append(df)
        if len(dfs) >= 2:
            combined = merge_quarters(dfs, period_names)
            create_dashboard(combined)
        else:
            st.info("Please upload at least two quarterly files for comparison.")
    else:
        st.info("Upload at least two Excel files (with columns: 'Investment Name', 'Geography', 'Strategy', 'NAV (ILS)', etc.)")

if __name__ == "__main__":
    main()
