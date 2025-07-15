import streamlit as st
import pandas as pd
import plotly.express as px
import collections

# הגדרות עיצוב
PROFESSIONAL_COLORS = [
    '#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087',
    '#f95d6a', '#ff7c43', '#ffa600', '#90be6d', '#43aa8b'
]

st.set_page_config(page_title="Investment Funds Analysis Dashboard", layout="wide")

st.markdown("""
<style>
    .main { padding: 1rem 1rem; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%); color: #F3F6FB;}
    .stApp { max-width: 1200px; margin: 0 auto; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%);}
    .metric-value { font-size: 2rem; font-weight: bold; color: #67B7DC;}
    .metric-label { font-size: 1rem; color: #CED2D9;}
    .sub-header { font-size: 1.5rem; font-weight: bold; margin: 1rem 0; padding-bottom: 0.5rem; border-bottom: 1px solid #3A4055;}
    .insight-box { background-color: rgba(40, 45, 60, 0.8); border-left: 4px solid #67B7DC; padding: 1rem; margin: 1rem 0; border-radius: 0.25rem;}
    .insight-box h3 { margin-top: 0; color: #8AD6CC;}
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

def load_data(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df.columns = df.columns.astype(str).str.strip()
        df.columns = [f'Column_{i}' if pd.isna(col) or col == '' else col for i, col in enumerate(df.columns)]
        df = ensure_unique_columns(df)
        df = df.dropna(how='all').dropna(axis=1, how='all')
        for col in df.columns:
            if 'NAV' in col or 'nav' in col.lower() or 'value' in col.lower() or 'amount' in col.lower():
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def create_dashboard(df):
    if df is None or df.empty:
        st.warning("No data available. Please upload a file.")
        return

    # זיהוי עמודות
    nav_col = [col for col in df.columns if 'NAV' in col or 'nav' in col.lower()][0]
    geo_col = [col for col in df.columns if 'Geo' in col or 'מדינה' in col or 'אזור' in col][0]
    char_col = [col for col in df.columns if 'מאפיין' in col or 'Characteristic' in col][0]
    strat_col = [col for col in df.columns if 'Strategy' in col or 'אסטרטגיה' in col or 'סוג' in col][0]
    fund_col = [col for col in df.columns if 'Fund' in col or 'קרן' in col or 'שם' in col][0]

    # מחשבים נתונים עיקריים
    total_nav = df[nav_col].sum()
    total_inv = len(df)
    avg_inv = total_nav / total_inv if total_inv > 0 else 0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"<div class='metric-value'>{format_number(total_nav)} ILS</div><div class='metric-label'>Total NAV</div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='metric-value'>{total_inv}</div><div class='metric-label'>Total Investments</div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='metric-value'>{format_number(avg_inv)} ILS</div><div class='metric-label'>Average Investment Size</div>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Geography Analysis", "Strategy Analysis", "Main Characteristic Analysis", "Detailed Data", "Additional Insights"])

    with tab1:
        st.markdown('<div class="sub-header">NAV Distribution by Geography</div>', unsafe_allow_html=True)
        geo_data = df.groupby(geo_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            # עיבוי ברים באמצעות width
            fig_geo_bar = px.bar(
                geo_data,
                x=nav_col,
                y=geo_col,
                orientation='h',
                color=geo_col,
                color_discrete_sequence=PROFESSIONAL_COLORS,
                width=550, height=500
            )
            fig_geo_bar.update_traces(marker_line_width=0, width=0.65)  # עובי ברים!
            fig_geo_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_geo_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
            fig_geo_bar.update_yaxes(tickfont=dict(size=16), title="Geography")
            st.plotly_chart(fig_geo_bar, use_container_width=True)
        with col2:
            fig_geo_pie = px.pie(
                geo_data,
                values=nav_col,
                names=geo_col,
                title=f"NAV Distribution by {geo_col}",
                hole=0.45,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_geo_pie, use_container_width=True)

    with tab2:
        st.markdown('<div class="sub-header">NAV Distribution by Strategy</div>', unsafe_allow_html=True)
        strat_data = df.groupby(strat_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig_strat_bar = px.bar(
                strat_data,
                x=nav_col,
                y=strat_col,
                orientation='h',
                color=strat_col,
                color_discrete_sequence=PROFESSIONAL_COLORS,
                width=550, height=500
            )
            fig_strat_bar.update_traces(marker_line_width=0, width=0.65)
            fig_strat_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_strat_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
            fig_strat_bar.update_yaxes(tickfont=dict(size=16), title="Strategy")
            st.plotly_chart(fig_strat_bar, use_container_width=True)
        with col2:
            fig_strat_pie = px.pie(
                strat_data,
                values=nav_col,
                names=strat_col,
                title=f"NAV Distribution by {strat_col}",
                hole=0.45,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strat_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_strat_pie, use_container_width=True)

        # תובנות
        if not strat_data.empty:
            top_strat = strat_data.iloc[0][strat_col]
            top_pct = strat_data.iloc[0][nav_col] / total_nav * 100
            st.markdown(f"""
            <div class="insight-box">
                <h3>Strategy Insights</h3>
                <p>האסטרטגיה המובילה: <b>{top_strat}</b>, מהווה <b>{top_pct:.1f}%</b> מסך הנכסים.</p>
            </div>
            """, unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="sub-header">NAV Distribution by Main Characteristic</div>', unsafe_allow_html=True)
        char_data = df.groupby(char_col)[nav_col].sum().reset_index().sort_values(by=nav_col, ascending=False)

        col1, col2 = st.columns(2)
        with col1:
            fig_char_bar = px.bar(
                char_data,
                x=nav_col,
                y=char_col,
                orientation='h',
                color=char_col,
                color_discrete_sequence=PROFESSIONAL_COLORS,
                width=550, height=500
            )
            fig_char_bar.update_traces(marker_line_width=0, width=0.65)
            fig_char_bar.update_layout(showlegend=True, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(40,45,60,0.8)')
            fig_char_bar.update_xaxes(tickfont=dict(size=16), title="NAV (ILS)")
            fig_char_bar.update_yaxes(tickfont=dict(size=16), title="Main Characteristic")
            st.plotly_chart(fig_char_bar, use_container_width=True)
        with col2:
            fig_char_pie = px.pie(
                char_data,
                values=nav_col,
                names=char_col,
                title=f"NAV Distribution by {char_col}",
                hole=0.45,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_char_pie.update_layout(height=500, showlegend=True, paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_char_pie, use_container_width=True)

    with tab4:
        st.markdown('<div class="sub-header">Detailed Investment Data</div>', unsafe_allow_html=True)
        st.dataframe(df, use_container_width=True)

    with tab5:
        st.markdown('<div class="sub-header">Additional Insights</div>', unsafe_allow_html=True)
        # דוגמה לתובנה נוספת
        st.markdown(f"""
        <div class="insight-box">
        <h3>Portfolio Concentration</h3>
        <p>Top Geography: <b>{geo_data.iloc[0][geo_col]}</b> - {geo_data.iloc[0][nav_col]/total_nav*100:.1f}%</p>
        <p>Top Strategy: <b>{strat_data.iloc[0][strat_col]}</b> - {strat_data.iloc[0][nav_col]/total_nav*100:.1f}%</p>
        <p>Number of Unique Main Characteristics: <b>{df[char_col].nunique()}</b></p>
        </div>
        """, unsafe_allow_html=True)

def main():
    st.sidebar.title("Investment Funds Analysis")
    uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        if df is not None:
            create_dashboard(df)
    else:
        st.markdown('<div class="main-header">Investment Funds Analysis Dashboard</div>', unsafe_allow_html=True)
        st.write("Please upload an Excel file containing investment funds data to begin analysis.")

if __name__ == "__main__":
    main()
