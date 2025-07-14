import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import collections

# Set page configuration
st.set_page_config(
    page_title="Investment Funds Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "Investment Funds Analysis Dashboard with Dark Theme"
    }
)

# Professional color palette
PROFESSIONAL_COLORS = [
    '#003f5c', '#2f4b7c', '#665191', '#a05195',
    '#d45087', '#f95d6a', '#ff7c43', '#ffa600'
]

# Add custom CSS
st.markdown("""
<style>
    .main { padding: 1rem 1rem; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%); color: #F3F6FB;}
    .stApp { max-width: 1200px; margin: 0 auto; background: linear-gradient(135deg, #20242f 0%, #282D3C 100%);}
    .metric-container { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 1rem; margin-bottom: 1rem;}
    .metric-box { background-color: rgba(40, 45, 60, 0.8); border-radius: 0.5rem; padding: 1rem; flex: 1; min-width: 200px;
        box-shadow: 0 0.15rem 0.3rem rgba(0, 0, 0, 0.3); text-align: center; border: 1px solid rgba(165, 180, 196, 0.3);}
    .metric-value { font-size: 2rem; font-weight: bold; color: #67B7DC; margin-bottom: 0.5rem;}
    .metric-label { font-size: 1rem; color: #CED2D9;}
    .sub-header { font-size: 1.5rem; font-weight: bold; margin: 1rem 0; padding-bottom: 0.5rem; border-bottom: 1px solid #3A4055; color: #F3F6FB;}
    .insight-box { background-color: rgba(40, 45, 60, 0.8); border-left: 4px solid #67B7DC; padding: 1rem; margin: 1rem 0; border-radius: 0.25rem; color: #F3F6FB;}
    .insight-box h3 { margin-top: 0; color: #8AD6CC;}
    .stDataFrame { margin: 1rem 0;}
    .stTabs [data-baseweb="tab-list"] { gap: 8px;}
    .stTabs [data-baseweb="tab"] { background-color: rgba(40, 45, 60, 0.8); border-radius: 4px 4px 0 0; color: #CED2D9; padding: 10px 16px;}
    .stTabs [aria-selected="true"] { background-color: rgba(103, 183, 220, 0.2) !important; color: #F3F6FB !important;}
</style>
""", unsafe_allow_html=True)

def format_number(num):
    """Format large numbers with commas and abbreviate if very large."""
    if num >= 1_000_000_000:
        return f"{num/1_000_000_000:.2f}B"
    elif num >= 1_000_000:
        return f"{num/1_000_000:.2f}M"
    elif num >= 1_000:
        return f"{num/1_000:.2f}K"
    else:
        return f"{num:.2f}"

def ensure_unique_columns(df):
    if df is None or df.empty:
        return df
    try:
        if len(df.columns) != len(set(df.columns)):
            duplicates = [item for item, count in collections.Counter(df.columns).items() if count > 1]
            new_columns = []
            seen = collections.defaultdict(int)
            for col in df.columns:
                if col in seen:
                    seen[col] += 1
                    new_columns.append(f"{col}_{seen[col]}")
                else:
                    seen[col] = 0
                    new_columns.append(col)
            df.columns = new_columns
    except Exception:
        pass
    return df

def load_data(uploaded_file):
    try:
        engines = ['openpyxl', 'xlrd']
        df = None
        for engine in engines:
            try:
                df = pd.read_excel(uploaded_file, engine=engine)
                break
            except Exception:
                continue
        if df is None:
            st.error("Failed to read the Excel file with any available engine.")
            return None
        df.columns = df.columns.astype(str).str.strip()
        df.columns = [f'Column_{i}' if pd.isna(col) or col == '' else col for i, col in enumerate(df.columns)]
        df = ensure_unique_columns(df)
        df = df.dropna(how='all').dropna(axis=1, how='all')
        for col in df.columns:
            if 'NAV' in col or 'nav' in col.lower() or 'value' in col.lower() or 'amount' in col.lower():
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except:
                    pass
        return df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def create_dashboard(df):
    if df is None or df.empty:
        st.warning("No data available. Please upload a file.")
        return

    # Identify key columns
    nav_column, geography_column, strategy_column, main_characteristic_column, fund_name_column = None, None, None, None, None
    for col in df.columns:
        col_lower = col.lower()
        if nav_column is None and (('nav' in col_lower or 'שווי' in col_lower) and ('ils' in col_lower or 'nis' in col_lower or 'שקל' in col_lower)):
            nav_column = col
        if geography_column is None and ('country' in col_lower or 'geography' in col_lower or 'region' in col_lower or 'מדינה' in col_lower or 'אזור' in col_lower):
            geography_column = col
        if strategy_column is None and ('strategy' in col_lower or 'type' in col_lower or 'category' in col_lower or 'אסטרטגיה' in col_lower or 'סוג' in col_lower):
            strategy_column = col
        if main_characteristic_column is None and ('מאפיין' in col_lower or 'characteristic' in col_lower or 'feature' in col_lower):
            main_characteristic_column = col
        if fund_name_column is None and ('fund' in col_lower or 'name' in col_lower or 'קרן' in col_lower or 'שם' in col_lower):
            fund_name_column = col
    # Fallbacks
    for col in df.columns:
        if nav_column is None and pd.api.types.is_numeric_dtype(df[col]) and df[col].mean() > 1000:
            nav_column = col
        if strategy_column is None and ('אסטרטגיה' in col or 'סוג' in col):
            strategy_column = col
        if main_characteristic_column is None and 'מאפיין עיקרי' in col:
            main_characteristic_column = col

    # Allow manual column selection
    st.sidebar.subheader("Column Selection")
    numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
    nav_column = st.sidebar.selectbox("Select NAV Column", options=numeric_columns, index=numeric_columns.index(nav_column) if nav_column in numeric_columns else 0)
    geography_column = st.sidebar.selectbox("Select Geography Column", options=df.columns, index=df.columns.get_loc(geography_column) if geography_column in df.columns else 0)
    strategy_column = st.sidebar.selectbox("Select Strategy Column", options=df.columns, index=df.columns.get_loc(strategy_column) if strategy_column in df.columns else 0)
    main_characteristic_column = st.sidebar.selectbox("Select Main Characteristic Column", options=df.columns, index=df.columns.get_loc(main_characteristic_column) if main_characteristic_column in df.columns else 0)

    # Filters
    geographies = ['All'] + sorted(df[geography_column].dropna().unique().tolist())
    selected_geography = st.sidebar.multiselect("Geography", options=geographies, default=['All'])
    strategies = ['All'] + sorted(df[strategy_column].dropna().unique().tolist())
    selected_strategy = st.sidebar.multiselect("Strategy", options=strategies, default=['All'])
    characteristics = ['All'] + sorted(df[main_characteristic_column].dropna().unique().tolist())
    selected_characteristic = st.sidebar.multiselect("Main Characteristic", options=characteristics, default=['All'])

    filtered_df = df.copy()
    if selected_geography and 'All' not in selected_geography:
        filtered_df = filtered_df[filtered_df[geography_column].isin(selected_geography)]
    if selected_strategy and 'All' not in selected_strategy:
        filtered_df = filtered_df[filtered_df[strategy_column].isin(selected_strategy)]
    if selected_characteristic and 'All' not in selected_characteristic:
        filtered_df = filtered_df[filtered_df[main_characteristic_column].isin(selected_characteristic)]
    st.sidebar.info(f"Displaying {len(filtered_df)} of {len(df)} investments ({len(filtered_df)/len(df)*100:.1f}%)")
    if filtered_df.empty:
        st.warning("No data matches the selected filters. Please adjust your filter criteria.")
        filtered_df = df.copy()
    total_nav = filtered_df[nav_column].sum()
    total_investments = len(filtered_df)

    st.markdown('<div class="main-header">Investment Funds Analysis Dashboard</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""<div class="metric-card"><div class="metric-value">{format_number(total_nav)} ILS</div>
        <div class="metric-label">Total NAV</div></div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class="metric-card"><div class="metric-value">{total_investments}</div>
        <div class="metric-label">Total Investments</div></div>""", unsafe_allow_html=True)
    with col3:
        avg_investment = total_nav / total_investments if total_investments > 0 else 0
        st.markdown(f"""<div class="metric-card"><div class="metric-value">{format_number(avg_investment)} ILS</div>
        <div class="metric-label">Average Investment Size</div></div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["Geography Analysis", "Strategy Analysis", "Main Characteristic Analysis"])

    with tab1:
        st.markdown('<div class="sub-header">NAV Distribution by Geography</div>', unsafe_allow_html=True)
        geo_data = filtered_df.groupby(geography_column)[nav_column].sum().reset_index().sort_values(by=nav_column, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_geo_bar = px.bar(
                geo_data, x=nav_column, y=geography_column, orientation='h',
                title=f"NAV by {geography_column}",
                labels={nav_column: "NAV (ILS)", geography_column: "Geography"},
                color=geography_column,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_bar.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                     plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                     margin=dict(l=50, r=50, t=70, b=80))
            st.plotly_chart(fig_geo_bar, use_container_width=True)
        with col2:
            fig_geo_pie = px.pie(
                geo_data, values=nav_column, names=geography_column,
                title=f"NAV Distribution by {geography_column}", hole=0.4,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_pie.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                     plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                     margin=dict(l=40, r=40, t=50, b=40))
            st.plotly_chart(fig_geo_pie, use_container_width=True)

    with tab2:
        st.markdown('<div class="sub-header">NAV Distribution by Strategy</div>', unsafe_allow_html=True)
        strategy_data = filtered_df.groupby(strategy_column)[nav_column].sum().reset_index().sort_values(by=nav_column, ascending=False)
        col1, col2 = st.columns(2)
        with col1:
            fig_strategy_bar = px.bar(
                strategy_data, x=nav_column, y=strategy_column, orientation='h',
                title=f"NAV by {strategy_column}",
                labels={nav_column: "NAV (ILS)", strategy_column: "Strategy"},
                color=strategy_column,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strategy_bar.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                          plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                          margin=dict(l=50, r=50, t=70, b=80))
            st.plotly_chart(fig_strategy_bar, use_container_width=True)
        with col2:
            fig_strategy_pie = px.pie(
                strategy_data, values=nav_column, names=strategy_column,
                title=f"NAV Distribution by {strategy_column}", hole=0.4,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strategy_pie.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                          plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                          margin=dict(l=40, r=40, t=50, b=40))
            st.plotly_chart(fig_strategy_pie, use_container_width=True)

    with tab3:
        st.markdown('<div class="sub-header">NAV Distribution by Main Characteristic</div>', unsafe_allow_html=True)
        try:
            calc_df = filtered_df.copy()
            calc_df[main_characteristic_column] = calc_df[main_characteristic_column].astype(str).fillna('Unknown')
            calc_df[nav_column] = pd.to_numeric(calc_df[nav_column], errors='coerce')
            calc_df = calc_df.dropna(subset=[nav_column])
            char_data = calc_df.groupby(main_characteristic_column)[nav_column].sum().reset_index().sort_values(by=nav_column, ascending=False)
            char_count = calc_df.groupby(main_characteristic_column).size().reset_index(name='Count').sort_values(by='Count', ascending=False)
            if not char_data.empty:
                col1, col2 = st.columns(2)
                with col1:
                    fig_char_bar = px.bar(
                        char_data, x=nav_column, y=main_characteristic_column, orientation='h',
                        title=f"NAV by {main_characteristic_column}",
                        labels={nav_column: "NAV (ILS)", main_characteristic_column: "Main Characteristic"},
                        color=main_characteristic_column,
                        color_discrete_sequence=PROFESSIONAL_COLORS
                    )
                    fig_char_bar.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                              plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                              margin=dict(l=50, r=50, t=70, b=80))
                    st.plotly_chart(fig_char_bar, use_container_width=True)
                with col2:
                    fig_char_pie = px.pie(
                        char_data, values=nav_column, names=main_characteristic_column,
                        title=f"NAV Distribution by {main_characteristic_column}", hole=0.4,
                        color_discrete_sequence=PROFESSIONAL_COLORS
                    )
                    fig_char_pie.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)',
                                              plot_bgcolor='rgba(40,45,60,0.8)', font=dict(color='#F3F6FB'),
                                              margin=dict(l=40, r=40, t=50, b=40))
                    st.plotly_chart(fig_char_pie, use_container_width=True)
                # טבלת מאפיינים
                merged_data = pd.merge(char_data, char_count, on=main_characteristic_column)
                merged_data[nav_column] = pd.to_numeric(merged_data[nav_column], errors='coerce')
                merged_data['Count'] = pd.to_numeric(merged_data['Count'], errors='coerce')
                merged_data = merged_data.dropna(subset=[nav_column, 'Count'])
                merged_data['Average NAV'] = merged_data.apply(lambda row: row[nav_column] / row['Count'] if row['Count'] > 0 else 0, axis=1)
                merged_data = merged_data.sort_values(by=nav_column, ascending=False)
                if total_nav > 0:
                    merged_data['% of Total NAV'] = (merged_data[nav_column] / total_nav * 100).apply(lambda x: f"{x:.1f}%")
                else:
                    merged_data['% of Total NAV'] = "0.0%"
                merged_data['NAV (ILS)'] = merged_data[nav_column].apply(format_number)
                merged_data['Average NAV (ILS)'] = merged_data['Average NAV'].apply(format_number)
                display_cols = [main_characteristic_column, 'Count', 'NAV (ILS)', 'Average NAV (ILS)', '% of Total NAV']
                st.dataframe(
                    merged_data[display_cols],
                    use_container_width=True,
                    hide_index=True
                )
        except Exception as e:
            st.error(f"Error analyzing main characteristic data: {str(e)}")

def main():
    st.sidebar.title("Investment Funds Analysis")
    st.sidebar.write("Upload your funds data file to analyze and visualize the investment portfolio.")
    uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        if df is not None:
            create_dashboard(df)
    else:
        st.markdown('<div class="main-header">Investment Funds Analysis Dashboard</div>', unsafe_allow_html=True)
        st.write("Please upload an Excel file containing investment funds data to begin analysis.")
        st.write("The file should contain columns for:")
        st.write("- Fund names")
        st.write("- NAV (ILS) values")
        st.write("- Geography information (country, region)")
        st.write("- Strategy classification")
        st.write("- Additional metadata (currency, valuation agent, etc.)")

if __name__ == "__main__":
    main()
