import streamlit as st
import pandas as pd
import plotly.express as px
import collections

# Page config
st.set_page_config(
    page_title="Investment Funds Analysis Dashboard",
    page_icon="📊",
    layout="wide"
)

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
    if pd.isnull(num):
        return "-"
    if num >= 1_000_000_000:
        return f"{num/1_000_000_000:.2f}B"
    elif num >= 1_000_000:
        return f"{num/1_000_000:.2f}M"
    elif num >= 1_000:
        return f"{num/1_000:.2f}K"
    else:
        return f"{num:.2f}"

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
        'מאפיין עיקרי': 'Main Characteristic',
        'characteristic': 'Main Characteristic',
        'feature': 'Main Characteristic'
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

def create_dashboard(df):
    st.sidebar.title("Investment Funds Analysis")
    nav_col = 'NAV (ILS)'
    geo_col = 'Geography'
    strat_col = 'Strategy'
    char_col = 'Main Characteristic' if 'Main Characteristic' in df.columns else None

    st.sidebar.subheader("Filters")
    geo_options = sorted(df[geo_col].dropna().unique())
    strat_options = sorted(df[strat_col].dropna().unique())
    char_options = sorted(df[char_col].dropna().unique()) if char_col else []

    selected_geo = st.sidebar.multiselect("Geography", ['All']+geo_options, ['All'])
    selected_strat = st.sidebar.multiselect("Strategy", ['All']+strat_options, ['All'])
    selected_char = st.sidebar.multiselect("Main Characteristic", ['All']+char_options, ['All']) if char_col else []

    filtered = df.copy()
    if selected_geo and 'All' not in selected_geo:
        filtered = filtered[filtered[geo_col].isin(selected_geo)]
    if selected_strat and 'All' not in selected_strat:
        filtered = filtered[filtered[strat_col].isin(selected_strat)]
    if char_col and selected_char and 'All' not in selected_char:
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

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Geography Analysis", "Strategy Analysis", "Main Characteristic Analysis", "Detailed Data", "Additional Insights", "Quarterly Comparison"
    ])

    with tab1:
        st.markdown('<div class="sub-header">NAV Distribution by Geography</div>', unsafe_allow_html=True)
        geo_data = filtered.groupby(geo_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_geo_bar = px.bar(
                geo_data, x=nav_col, y=geo_col, orientation='h', color=geo_col,
                color_discrete_sequence=PROFESSIONAL_COLORS, width=550, height=500
            )
            fig_geo_bar.update_traces(marker_line_width=0, marker_line_color="black", width=0.8) # עיבוי ברים
            fig_geo_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_geo_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
            fig_geo_bar.update_yaxes(tickfont=dict(size=16), title="Geography")
            st.plotly_chart(fig_geo_bar, use_container_width=True)
        with col2:
            fig_geo_pie = px.pie(
                geo_data, values=nav_col, names=geo_col,
                hole=0.45, color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_geo_pie, use_container_width=True)

    with tab2:
        st.markdown('<div class="sub-header">NAV Distribution by Strategy</div>', unsafe_allow_html=True)
        strat_data = filtered.groupby(strat_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_strat_bar = px.bar(
                strat_data, x=nav_col, y=strat_col, orientation='h', color=strat_col,
                color_discrete_sequence=PROFESSIONAL_COLORS, width=700, height=500
            )
            fig_strat_bar.update_traces(marker_line_width=0, marker_line_color="black", width=0.8)
            fig_strat_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_strat_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
            fig_strat_bar.update_yaxes(tickfont=dict(size=16), title="Strategy")
            st.plotly_chart(fig_strat_bar, use_container_width=True)
        with col2:
            fig_strat_pie = px.pie(
                strat_data, values=nav_col, names=strat_col,
                hole=0.45, color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strat_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_strat_pie, use_container_width=True)

    with tab3:
        st.markdown('<div class="sub-header">NAV Distribution by Main Characteristic</div>', unsafe_allow_html=True)
        if char_col and char_col in filtered.columns:
            char_data = filtered.groupby(char_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)
            if not char_data.empty:
                col1, col2 = st.columns(2)
                with col1:
                    fig_char_bar = px.bar(
                        char_data, x=nav_col, y=char_col, orientation='h', color=char_col,
                        color_discrete_sequence=PROFESSIONAL_COLORS, width=550, height=500
                    )
                    fig_char_bar.update_traces(marker_line_width=0, marker_line_color="black", width=0.8)
                    fig_char_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
                    fig_char_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
                    fig_char_bar.update_yaxes(tickfont=dict(size=16), title="Main Characteristic")
                    st.plotly_chart(fig_char_bar, use_container_width=True)
                with col2:
                    fig_char_pie = px.pie(
                        char_data, values=nav_col, names=char_col,
                        hole=0.45, color_discrete_sequence=PROFESSIONAL_COLORS
                    )
                    fig_char_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_char_pie, use_container_width=True)
            else:
                st.info("No data to display for Main Characteristic.")
        else:
            st.info("Column 'Main Characteristic' is missing from data.")

    with tab4:
        st.markdown('<div class="sub-header">Detailed Investment Data</div>', unsafe_allow_html=True)
        st.dataframe(filtered, use_container_width=True)
        st.download_button(
            "Download filtered data", filtered.to_csv(index=False).encode("utf-8-sig"),
            file_name="filtered_investments.csv", mime="text/csv"
        )

    with tab5:
        st.markdown('<div class="sub-header">Additional Insights</div>', unsafe_allow_html=True)
        try:
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
                st.info("No Geography data for Additional Insights.")

        except Exception as e:
            st.warning(f"Could not generate Additional Insights: {e}")

    # --------------------------
    #   טאב חדש: השוואה בין רבעונים
    # --------------------------
    with tab6:
        st.header("Quarterly Comparison: Trends by Geography and Strategy")
        uploaded_files = st.file_uploader("Upload quarterly Excel files to compare trends", type=["xlsx", "xls"], accept_multiple_files=True, key="multi_upload")
        if uploaded_files and len(uploaded_files) > 1:
            dfs = []
            period_list = []
            for file in uploaded_files:
                df2 = pd.read_excel(file, engine='openpyxl')
                df2 = normalize_columns(df2)
                # תשאל את המשתמש את שם הרבעון (אבל תציע כברירת מחדל את שם הקובץ)
                default_period = file.name.split('.')[0]
                period = st.text_input(f"תווית רבעון עבור {file.name}", value=default_period, key=f'label_{file.name}')
                df2['Period'] = period.strip()
                dfs.append(df2)
                period_list.append(period.strip())
            all_data = pd.concat(dfs, ignore_index=True)
            # תקן לפי סדר כרונולוגי לפי ה-Period (אם אפשר)
            all_data['Period'] = pd.Categorical(all_data['Period'], categories=sorted(period_list), ordered=True)

            st.markdown("### שינוי חשיפה לפי Geography")
            if 'Geography' in all_data.columns and 'NAV (ILS)' in all_data.columns:
                geo_trend = all_data.groupby(['Period', 'Geography'])['NAV (ILS)'].sum().reset_index()
                geo_trend = geo_trend.sort_values('Period')
                fig_geo = px.line(geo_trend, x='Period', y='NAV (ILS)', color='Geography', markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
                st.plotly_chart(fig_geo, use_container_width=True)

            st.markdown("### שינוי חשיפה לפי Strategy")
            if 'Strategy' in all_data.columns and 'NAV (ILS)' in all_data.columns:
                strat_trend = all_data.groupby(['Period', 'Strategy'])['NAV (ILS)'].sum().reset_index()
                strat_trend = strat_trend.sort_values('Period')
                fig_strat = px.line(strat_trend, x='Period', y='NAV (ILS)', color='Strategy', markers=True, color_discrete_sequence=PROFESSIONAL_COLORS)
                st.plotly_chart(fig_strat, use_container_width=True)

            st.markdown("### All Data")
            st.dataframe(all_data, use_container_width=True)
            st.download_button("Download combined data", all_data.to_csv(index=False).encode("utf-8-sig"), file_name="combined_quarters.csv", mime="text/csv")
        else:
            st.info("Please upload at least two quarterly files to view trends.")

def main():
    st.title("Investment Funds Analysis Dashboard")
    uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="main_upload")
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        if df is not None:
            create_dashboard(df)
    else:
        st.markdown('<div class="main-header">Investment Funds Analysis Dashboard</div>', unsafe_allow_html=True)
        st.write("Please upload an Excel file containing investment funds data.")

if __name__ == "__main__":
    main()
