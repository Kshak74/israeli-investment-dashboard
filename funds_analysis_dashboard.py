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
        'main characteristic': 'Main Characteristic',
        'מאפיין עיקרי': 'Main Characteristic'
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

def parse_period(x):
    # Handle formats like "Q1 2024"
    try:
        x = str(x).strip()
        if x.startswith("Q"):
            q, y = x.replace("Q", "").split()
            return int(y) * 10 + int(q)
        return int(x)
    except Exception:
        return 0

def merge_quarters(dfs, period_names):
    for i, (df, pname) in enumerate(zip(dfs, period_names)):
        df['Period'] = pname
    # השתמש בעמודות המשותפות בלבד, בסדר אחיד
    common_cols = list(set.intersection(*[set(df.columns) for df in dfs]))
    dfs_common = [df[common_cols + ['Period']] if 'Period' not in common_cols else df[common_cols] for df in dfs]
    combined = pd.concat(dfs_common, ignore_index=True)
    # סדר כרונולוגי לעמודת Period
    unique_periods = list(sorted(combined['Period'].unique(), key=parse_period))
    combined['Period'] = pd.Categorical(combined['Period'], categories=unique_periods, ordered=True)
    return combined

def create_dashboard(df):
    st.sidebar.title("Quarterly Analysis Filters")
    period_col = 'Period'
    geo_col = 'Geography'
    strat_col = 'Strategy'
    nav_col = 'NAV (ILS)'
    char_col = 'Main Characteristic'

    period_options = list(df[period_col].dropna().unique())
    selected_periods = st.sidebar.multiselect("Select Period(s)", period_options, period_options)
    geo_options = sorted(df[geo_col].dropna().unique())
    strat_options = sorted(df[strat_col].dropna().unique())
    char_options = sorted(df[char_col].dropna().unique()) if char_col in df.columns else []

    selected_geo = st.sidebar.multiselect("Geography", ['All']+geo_options, ['All'])
    selected_strat = st.sidebar.multiselect("Strategy", ['All']+strat_options, ['All'])
    selected_char = st.sidebar.multiselect("Main Characteristic", ['All']+char_options, ['All']) if char_options else []

    filtered = df.copy()
    if selected_periods:
        filtered = filtered[filtered[period_col].astype(str).isin(selected_periods)]
    if selected_geo and 'All' not in selected_geo:
        filtered = filtered[filtered[geo_col].isin(selected_geo)]
    if selected_strat and 'All' not in selected_strat:
        filtered = filtered[filtered[strat_col].isin(selected_strat)]
    if selected_char and selected_char and 'All' not in selected_char and char_col in filtered.columns:
        filtered = filtered[filtered[char_col].isin(selected_char)]

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

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Quarterly Comparison", "Geography", "Strategy", "Main Characteristic", "Additional Insights"
    ])

    # 1. טאב השוואה רב רבעונית
    with tab1:
        st.subheader("Quarterly Comparison - NAV by Strategy & Geography")
        # Strategy Trend
        trend = filtered.groupby([period_col, strat_col])[nav_col].sum().reset_index()
        fig = px.line(trend, x=period_col, y=nav_col, color=strat_col, markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
        fig.update_xaxes(tickfont=dict(size=18))
        fig.update_yaxes(tickfont=dict(size=18))
        st.plotly_chart(fig, use_container_width=True)
        # Geography Trend
        geo_trend = filtered.groupby([period_col, geo_col])[nav_col].sum().reset_index()
        fig_geo = px.line(geo_trend, x=period_col, y=nav_col, color=geo_col, markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
        fig_geo.update_xaxes(tickfont=dict(size=18))
        fig_geo.update_yaxes(tickfont=dict(size=18))
        st.plotly_chart(fig_geo, use_container_width=True)

    # 2. טאב גיאוגרפיה
    with tab2:
        st.subheader("NAV Distribution by Geography")
        geo_data = filtered.groupby(geo_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_geo_bar = px.bar(
                geo_data, x=nav_col, y=geo_col, orientation='h', color=geo_col,
                color_discrete_sequence=PROFESSIONAL_COLORS, width=550, height=500
            )
            fig_geo_bar.update_traces(marker_line_width=0, width=0.65)
            fig_geo_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_geo_bar.update_xaxes(tickfont=dict(size=18), title="NAV (ILS)")
            fig_geo_bar.update_yaxes(tickfont=dict(size=18), title="Geography")
            st.plotly_chart(fig_geo_bar, use_container_width=True)
        with col2:
            fig_geo_pie = px.pie(
                geo_data, values=nav_col, names=geo_col,
                hole=0.45, color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_geo_pie, use_container_width=True)
        # Insights
        if not geo_data.empty:
            st.markdown(f"""
            <div class="insight-box">
                <h3>Geography Insights</h3>
                <p>The largest exposure is <b>{geo_data.iloc[0][geo_col]}</b> ({geo_data.iloc[0][nav_col]/total_nav*100:.1f}%)</p>
            </div>
            """, unsafe_allow_html=True)

    # 3. טאב אסטרטגיה
    with tab3:
        st.subheader("NAV Distribution by Strategy")
        strat_data = filtered.groupby(strat_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_strat_bar = px.bar(
                strat_data, x=nav_col, y=strat_col, orientation='h', color=strat_col,
                color_discrete_sequence=PROFESSIONAL_COLORS, width=550, height=500
            )
            fig_strat_bar.update_traces(marker_line_width=0, width=0.65)
            fig_strat_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_strat_bar.update_xaxes(tickfont=dict(size=18), title="NAV (ILS)")
            fig_strat_bar.update_yaxes(tickfont=dict(size=18), title="Strategy")
            st.plotly_chart(fig_strat_bar, use_container_width=True)
        with col2:
            fig_strat_pie = px.pie(
                strat_data, values=nav_col, names=strat_col,
                hole=0.45, color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strat_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_strat_pie, use_container_width=True)
        if not strat_data.empty:
            st.markdown(f"""
            <div class="insight-box">
                <h3>Strategy Insights</h3>
                <p>Dominant strategy: <b>{strat_data.iloc[0][strat_col]}</b> ({strat_data.iloc[0][nav_col]/total_nav*100:.1f}%)</p>
            </div>
            """, unsafe_allow_html=True)

    # 4. טאב Main Characteristic (אם קיים)
    with tab4:
        st.subheader("NAV by Main Characteristic")
        if char_col in filtered.columns:
            char_data = filtered.groupby(char_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
            if not char_data.empty:
                fig_char_bar = px.bar(
                    char_data, x=nav_col, y=char_col, orientation='h', color=char_col,
                    color_discrete_sequence=PROFESSIONAL_COLORS, width=550, height=500
                )
                fig_char_bar.update_traces(marker_line_width=0, width=0.65)
                fig_char_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
                fig_char_bar.update_xaxes(tickfont=dict(size=18), title="NAV (ILS)")
                fig_char_bar.update_yaxes(tickfont=dict(size=18), title="Main Characteristic")
                st.plotly_chart(fig_char_bar, use_container_width=True)
            else:
                st.info("No data for Main Characteristic.")
        else:
            st.info("Main Characteristic column not found in uploaded files.")

    # 5. טאב Additional Insights (חלקי/מותאם לפי קובץ)
    with tab5:
        st.markdown('<div class="sub-header">Additional Insights</div>', unsafe_allow_html=True)
        # Example: Israel vs. International
        if geo_col in df.columns:
            df['Israel_Flag'] = df[geo_col].apply(lambda x: 'Israel' if str(x).strip().lower() in ['israel', 'ישראל', 'il'] else 'International')
            israel_data = df.groupby('Israel_Flag')[nav_col].sum().reset_index()
            col1, col2 = st.columns(2)
            with col1:
                fig = px.bar(israel_data, x='Israel_Flag', y=nav_col, color='Israel_Flag',
                             color_discrete_sequence=PROFESSIONAL_COLORS)
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                pie = px.pie(israel_data, values=nav_col, names='Israel_Flag', hole=0.45,
                             color_discrete_sequence=PROFESSIONAL_COLORS)
                st.plotly_chart(pie, use_container_width=True)
        else:
            st.info("No geography data for Israel/International split.")

    # --- הצגת טבלת הנתונים להורדה ---
    st.subheader("Raw Data")
    st.dataframe(filtered, use_container_width=True)
    st.download_button("Download filtered data", filtered.to_csv(index=False).encode("utf-8-sig"),
                       file_name="quarterly_filtered.csv", mime="text/csv")

def main():
    st.title("Investment Funds Quarterly Comparison")
    uploaded_files = st.sidebar.file_uploader(
        "Upload Excel Files (quarterly)", type=["xlsx", "xls"], accept_multiple_files=True)
    period_names = []
    dfs = []
    REQUIRED_COLS = ['Investment Name', 'Geography', 'Strategy', 'NAV (ILS)']
    if uploaded_files:
        for upf in uploaded_files:
            df = load_data(upf)
            if df is not None:
                missing = [col for col in REQUIRED_COLS if col not in df.columns]
                if missing:
                    st.warning(f"File {upf.name} missing columns: {missing}. Please rename/map columns in the Excel file.")
                    continue
                period_col_guess = None
                for col in df.columns:
                    if "רבעון" in col or "שנה" in col or "period" in col.lower() or "date" in col.lower():
                        period_col_guess = col
                        break
                if period_col_guess:
                    period_value = str(df[period_col_guess].iloc[0])
                else:
                    period_value = st.sidebar.text_input(f"Enter period label for {upf.name}")
                period_names.append(period_value)
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
