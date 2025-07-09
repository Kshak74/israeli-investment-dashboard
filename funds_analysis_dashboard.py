import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import collections
from datetime import datetime

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

# Define the professional D3 color palette
PROFESSIONAL_COLORS = [
    '#003f5c',  # deep blue
    '#2f4b7c',  # navy blue
    '#665191',  # purple
    '#a05195',  # magenta
    '#d45087',  # pink
    '#f95d6a',  # coral
    '#ff7c43',  # orange
    '#ffa600',  # amber
]

# Add custom CSS for better styling with dark theme
st.markdown("""
<style>
    .main {
        padding: 1rem 1rem;
        background: linear-gradient(135deg, #20242f 0%, #282D3C 100%);
        color: #F3F6FB;
    }
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
        background: linear-gradient(135deg, #20242f 0%, #282D3C 100%);
    }
    .metric-container {
        display: flex;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 1rem;
        margin-bottom: 1rem;
    }
    .metric-box {
        background-color: rgba(40, 45, 60, 0.8);
        border-radius: 0.5rem;
        padding: 1rem;
        flex: 1;
        min-width: 200px;
        box-shadow: 0 0.15rem 0.3rem rgba(0, 0, 0, 0.3);
        text-align: center;
        border: 1px solid rgba(165, 180, 196, 0.3);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #67B7DC;
        margin-bottom: 0.5rem;
    }
    .metric-label {
        font-size: 1rem;
        color: #CED2D9;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        margin: 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid #3A4055;
        color: #F3F6FB;
    }
    .insight-box {
        background-color: rgba(40, 45, 60, 0.8);
        border-left: 4px solid #67B7DC;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 0.25rem;
        color: #F3F6FB;
    }
    .insight-box h3 {
        margin-top: 0;
        color: #8AD6CC;
    }
    .stDataFrame {
        margin: 1rem 0;
    }
    /* Style for tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(40, 45, 60, 0.8);
        border-radius: 4px 4px 0 0;
        color: #CED2D9;
        padding: 10px 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: rgba(103, 183, 220, 0.2) !important;
        color: #F3F6FB !important;
    }
    /* Style for sidebar */
    .css-1d391kg, .css-12oz5g7 {
        background-color: #20242f;
    }
    /* Style for text inputs */
    .stTextInput>div>div>input {
        background-color: #282D3C;
        color: #F3F6FB;
    }
    /* Style for selectbox */
    .stSelectbox>div>div>div {
        background-color: #282D3C;
        color: #F3F6FB;
    }
    /* Style for multiselect */
    .stMultiSelect>div>div>div {
        background-color: #282D3C;
        color: #F3F6FB;
    }
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
    """Ensure dataframe has unique column names to prevent rendering errors."""
    if df is None or df.empty:
        return df
        
    try:
        # Check for duplicate columns
        if len(df.columns) != len(set(df.columns)):
            # Find duplicates
            duplicates = [item for item, count in collections.Counter(df.columns).items() if count > 1]
            st.warning(f"Found duplicate columns that need to be renamed: {', '.join(duplicates)}")
            
            # Create unique column names
            new_columns = []
            seen = collections.defaultdict(int)
            
            for col in df.columns:
                if col in seen:
                    seen[col] += 1
                    new_columns.append(f"{col}_{seen[col]}")
                else:
                    seen[col] = 0
                    new_columns.append(col)
            
            # Set the new column names
            df.columns = new_columns
            st.success("Column names have been made unique for proper display")
    except Exception as e:
        st.error(f"Error ensuring unique columns: {str(e)}")
    
    return df

def load_data(uploaded_file):
    """Load and process the uploaded Excel file."""
    try:
        # Try different engines
        engines = ['openpyxl', 'xlrd']
        df = None
        error_messages = []
        
        for engine in engines:
            try:
                with st.spinner(f"Loading data with {engine}..."):
                    df = pd.read_excel(uploaded_file, engine=engine)
                    break  # Successfully loaded
            except Exception as e:
                error_messages.append(f"Error with {engine}: {str(e)}")
                continue
        
        if df is None:
            for msg in error_messages:
                st.error(msg)
            st.error("Failed to read the Excel file with any available engine.")
            return None
        
        # Clean up the data
        # Convert column names to strings and strip whitespace
        df.columns = df.columns.astype(str).str.strip()
        
        # Replace NaN or empty column names with descriptive names
        df.columns = [f'Column_{i}' if pd.isna(col) or col == '' else col for i, col in enumerate(df.columns)]
        
        # Ensure unique column names
        df = ensure_unique_columns(df)
        
        # Drop completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Convert numeric columns
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
    """Create the dashboard with the processed data."""
    if df is None or df.empty:
        st.warning("No data available. Please upload a file.")
        return
        
    # Convert all potential groupby columns to strings to avoid type comparison issues
    def safely_convert_to_string(df, columns_to_check):
        """Convert columns that might be used for grouping to strings to avoid type comparison errors."""
        for col in columns_to_check:
            if col in df.columns:
                try:
                    # Check if column contains mixed types
                    if df[col].dtype == 'object' or pd.api.types.is_datetime64_any_dtype(df[col]):
                        df[col] = df[col].astype(str)
                        st.sidebar.info(f"Column '{col}' converted to string type to ensure compatibility.")
                except Exception as e:
                    st.sidebar.warning(f"Could not convert column '{col}' to string: {str(e)}")
        return df
    
    # List of columns that might be used for grouping
    potential_groupby_columns = [
        # Common categorical columns
        'Country', 'Region', 'Geography', 'Strategy', 'Type', 'Category', 'Fund', 'Manager',
        'Currency', 'Valuation Agent', 'Year', 'GP', 'מאפיין עיקרי', 'מדינה', 'אזור', 'אסטרטגיה',
        'סוג', 'קרן', 'מנהל', 'מטבע', 'שנה'
    ]
    
    # Add all object columns as potential groupby columns
    for col in df.columns:
        if df[col].dtype == 'object' or pd.api.types.is_datetime64_any_dtype(df[col]):
            potential_groupby_columns.append(col)
    
    # Convert columns to strings
    df = safely_convert_to_string(df, potential_groupby_columns)
    
    # Identify key columns
    nav_column = None
    geography_column = None
    strategy_column = None
    fund_name_column = None
    currency_column = None
    valuation_agent_column = None
    year_column = None
    gp_column = None
    
    # Look for NAV column with expanded search terms
    for col in df.columns:
        col_lower = col.lower()
        # First priority: NAV in ILS/NIS/Shekels
        if ('nav' in col_lower or 'שווי' in col_lower) and ('ils' in col_lower or 'nis' in col_lower or 'שקל' in col_lower):
            nav_column = col
            break
    
    if nav_column is None:
        # Second priority: Any column with NAV and numeric
        for col in df.columns:
            col_lower = col.lower()
            if ('nav' in col_lower or 'שווי' in col_lower) and pd.api.types.is_numeric_dtype(df[col]):
                nav_column = col
                break
    
    if nav_column is None:
        # Third priority: Any value/amount column that's numeric
        for col in df.columns:
            col_lower = col.lower()
            if ('value' in col_lower or 'amount' in col_lower or 'סכום' in col_lower or 'ערך' in col_lower) and pd.api.types.is_numeric_dtype(df[col]):
                nav_column = col
                break
                
    if nav_column is None:
        # Last resort: Any numeric column with large values
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                # Check if this column has values that look like monetary amounts (larger numbers)
                if df[col].mean() > 1000:  # Assuming NAV values are typically large
                    nav_column = col
                    break
    
    # Look for geography column
    for col in df.columns:
        col_lower = col.lower()
        if 'country' in col_lower or 'geography' in col_lower or 'region' in col_lower or 'מדינה' in col_lower or 'אזור' in col_lower:
            geography_column = col
            break
    
    # Look for strategy column - prioritize exact match for 'Strategy'
    if 'Strategy' in df.columns:
        strategy_column = 'Strategy'
    else:
        for col in df.columns:
            col_lower = col.lower()
            if 'strategy' in col_lower or 'type' in col_lower or 'category' in col_lower or 'אסטרטגיה' in col_lower or 'סוג' in col_lower:
                strategy_column = col
                break
                
    # Look for 'מאפיין עיקרי' (main characteristic) column
    main_characteristic_column = None
    if 'מאפיין עיקרי' in df.columns:
        main_characteristic_column = 'מאפיין עיקרי'
    else:
        for col in df.columns:
            col_lower = col.lower()
            if 'מאפיין' in col_lower or 'characteristic' in col_lower or 'feature' in col_lower or 'attribute' in col_lower:
                main_characteristic_column = col
                break
    
    # Look for fund name column
    for col in df.columns:
        col_lower = col.lower()
        if 'fund' in col_lower or 'name' in col_lower or 'קרן' in col_lower or 'שם' in col_lower:
            fund_name_column = col
            break
    
    # Look for currency column
    for col in df.columns:
        col_lower = col.lower()
        if 'currency' in col_lower or 'מטבע' in col_lower:
            currency_column = col
            break
    
    # Look for valuation agent column
    for col in df.columns:
        col_lower = col.lower()
        if 'valuation' in col_lower or 'agent' in col_lower or 'שערוך' in col_lower:
            valuation_agent_column = col
            break
    
    # Look for year column
    for col in df.columns:
        col_lower = col.lower()
        if 'year' in col_lower or 'date' in col_lower or 'שנה' in col_lower or 'תאריך' in col_lower:
            year_column = col
            break
    
    # Look for GP column
    for col in df.columns:
        col_lower = col.lower()
        if 'gp' in col_lower or 'general partner' in col_lower or 'manager' in col_lower or 'מנהל' in col_lower:
            gp_column = col
            break
    
    # Display column detection results in sidebar
    st.sidebar.subheader("Detected Columns")
    st.sidebar.write(f"NAV Column: {nav_column}")
    st.sidebar.write(f"Geography Column: {geography_column}")
    st.sidebar.write(f"Strategy Column: {strategy_column}")
    st.sidebar.write(f"Main Characteristic Column: {main_characteristic_column}")
    st.sidebar.write(f"Fund Name Column: {fund_name_column}")
    
    # Allow manual column selection if needed
    st.sidebar.subheader("Column Selection")
    
    numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
    if not numeric_columns:
        st.error("No numeric columns found in the data. Please check your file.")
        return
        
    # Find the index of nav_column in numeric_columns if it exists
    default_index = 0
    if nav_column in numeric_columns:
        default_index = numeric_columns.index(nav_column)
    
    nav_column = st.sidebar.selectbox(
        "Select NAV Column",
        options=numeric_columns,
        index=default_index
    )
    
    # Geography column selection with error handling
    geo_index = 0
    if geography_column in df.columns:
        geo_index = df.columns.get_loc(geography_column)
    
    geography_column = st.sidebar.selectbox(
        "Select Geography Column",
        options=df.columns,
        index=geo_index
    )
    
    # Strategy column selection with error handling
    strat_index = 0
    if strategy_column in df.columns:
        strat_index = df.columns.get_loc(strategy_column)
    
    strategy_column = st.sidebar.selectbox(
        "Select Strategy Column",
        options=df.columns,
        index=strat_index
    )
    
    # Main characteristic column selection with error handling
    char_index = 0
    if main_characteristic_column in df.columns:
        char_index = df.columns.get_loc(main_characteristic_column)
    
    main_characteristic_column = st.sidebar.selectbox(
        "Select Main Characteristic Column",
        options=df.columns,
        index=char_index
    )
    
    # Create filters with error handling
    st.sidebar.subheader("Filters")
    
    # Filter by geography with error handling
    try:
        geographies = ['All'] + sorted(df[geography_column].dropna().unique().tolist())
        selected_geography = st.sidebar.multiselect(
            "Geography",
            options=geographies,
            default=['All']
        )
    except Exception as e:
        st.sidebar.warning(f"Could not create geography filter: {str(e)}")
        selected_geography = ['All']
    
    # Filter by strategy with error handling
    try:
        strategies = ['All'] + sorted(df[strategy_column].dropna().unique().tolist())
        selected_strategy = st.sidebar.multiselect(
            "Strategy",
            options=strategies,
            default=['All']
        )
    except Exception as e:
        st.sidebar.warning(f"Could not create strategy filter: {str(e)}")
        selected_strategy = ['All']
        
    # Filter by main characteristic with error handling
    try:
        characteristics = ['All'] + sorted(df[main_characteristic_column].dropna().unique().tolist())
        selected_characteristic = st.sidebar.multiselect(
            "Main Characteristic",
            options=characteristics,
            default=['All']
        )
    except Exception as e:
        st.sidebar.warning(f"Could not create main characteristic filter: {str(e)}")
        selected_characteristic = ['All']
    
    # Apply filters with error handling
    filtered_df = df.copy()
    
    try:
        if selected_geography and 'All' not in selected_geography:
            filtered_df = filtered_df[filtered_df[geography_column].isin(selected_geography)]
    except Exception as e:
        st.sidebar.error(f"Error applying geography filter: {str(e)}")
    
    try:
        if selected_strategy and 'All' not in selected_strategy:
            filtered_df = filtered_df[filtered_df[strategy_column].isin(selected_strategy)]
    except Exception as e:
        st.sidebar.error(f"Error applying strategy filter: {str(e)}")
        
    # Apply main characteristic filter
    try:
        if selected_characteristic and 'All' not in selected_characteristic:
            filtered_df = filtered_df[filtered_df[main_characteristic_column].isin(selected_characteristic)]
    except Exception as e:
        st.sidebar.error(f"Error applying main characteristic filter: {str(e)}")
        
    # Show how many records are displayed after filtering
    st.sidebar.info(f"Displaying {len(filtered_df)} of {len(df)} investments ({len(filtered_df)/len(df)*100:.1f}%)")
    
    # Check if we have any data after filtering
    if filtered_df.empty:
        st.warning("No data matches the selected filters. Please adjust your filter criteria.")
        # Reset to original data
        filtered_df = df.copy()
    
    # Calculate KPIs
    total_nav = filtered_df[nav_column].sum()
    total_investments = len(filtered_df)
    
    # Display dashboard header
    st.markdown('<div class="main-header">Investment Funds Analysis Dashboard</div>', unsafe_allow_html=True)
    
    # Display KPIs
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{format_number(total_nav)} ILS</div>
            <div class="metric-label">Total NAV</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_investments}</div>
            <div class="metric-label">Total Investments</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        avg_investment = total_nav / total_investments if total_investments > 0 else 0
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{format_number(avg_investment)} ILS</div>
            <div class="metric-label">Average Investment Size</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Create tabs for different visualizations
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Geography Analysis", "Strategy Analysis", "Main Characteristic Analysis", "Detailed Data", "Additional Insights"])
    
    with tab1:
        st.markdown('<div class="sub-header">NAV Distribution by Geography</div>', unsafe_allow_html=True)
        
        # Group by geography
        geo_data = filtered_df.groupby(geography_column)[nav_column].sum().reset_index()
        geo_data = geo_data.sort_values(by=nav_column, ascending=False)
        
        # Create charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Bar chart
            fig_geo_bar = px.bar(
                geo_data,
                x=geography_column,
                y=nav_column,
                title=f"NAV by {geography_column}",
                labels={nav_column: "NAV (ILS)", geography_column: "Geography"},
                color=geography_column,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_bar.update_layout(
                height=500,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(40,45,60,0.8)',
                font=dict(color='#F3F6FB'),
                margin=dict(l=40, r=40, t=50, b=40)
            )
            st.plotly_chart(fig_geo_bar, use_container_width=True)
        
        with col2:
            # Pie chart
            fig_geo_pie = px.pie(
                geo_data,
                values=nav_column,
                names=geography_column,
                title=f"NAV Distribution by {geography_column}",
                hole=0.4,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_geo_pie.update_layout(
                height=500,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(40,45,60,0.8)',
                font=dict(color='#F3F6FB'),
                margin=dict(l=40, r=40, t=50, b=40)
            )
            st.plotly_chart(fig_geo_pie, use_container_width=True)
        
        # Geography insights
        top_geography = geo_data.iloc[0][geography_column]
        top_geography_pct = (geo_data.iloc[0][nav_column] / total_nav) * 100
        
        st.markdown(f"""
        <div class="insight-box">
            <h3>Geography Insights</h3>
            <p>The largest geographic exposure is to <b>{top_geography}</b>, representing <b>{top_geography_pct:.1f}%</b> of the total NAV.</p>
            <p>This concentration may indicate a strategic focus or home market bias in the investment portfolio.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown('<div class="sub-header">NAV Distribution by Strategy</div>', unsafe_allow_html=True)
        
        # Group by strategy
        strategy_data = filtered_df.groupby(strategy_column)[nav_column].sum().reset_index()
        strategy_data = strategy_data.sort_values(by=nav_column, ascending=False)
        
        # Create charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Bar chart
            fig_strategy_bar = px.bar(
                strategy_data,
                x=strategy_column,
                y=nav_column,
                title=f"NAV by {strategy_column}",
                labels={nav_column: "NAV (ILS)", strategy_column: "Strategy"},
                color=strategy_column,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strategy_bar.update_layout(
                height=500,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(40,45,60,0.8)',
                font=dict(color='#F3F6FB'),
                margin=dict(l=40, r=40, t=50, b=40)
            )
            st.plotly_chart(fig_strategy_bar, use_container_width=True)
        
        with col2:
            # Pie chart
            fig_strategy_pie = px.pie(
                strategy_data,
                values=nav_column,
                names=strategy_column,
                title=f"NAV Distribution by {strategy_column}",
                hole=0.4,
                color_discrete_sequence=PROFESSIONAL_COLORS
            )
            fig_strategy_pie.update_layout(
                height=500,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(40,45,60,0.8)',
                font=dict(color='#F3F6FB'),
                margin=dict(l=40, r=40, t=50, b=40)
            )
            st.plotly_chart(fig_strategy_pie, use_container_width=True)
        
        # Strategy insights
        top_strategy = strategy_data.iloc[0][strategy_column] if not strategy_data.empty else "N/A"
        top_strategy_pct = (strategy_data.iloc[0][nav_column] / total_nav) * 100 if not strategy_data.empty and total_nav > 0 else 0
        
        st.markdown(f"""
        <div class="insight-box">
            <h3>Strategy Insights</h3>
            <p>The dominant investment strategy is <b>{top_strategy}</b>, accounting for <b>{top_strategy_pct:.1f}%</b> of the total NAV.</p>
            <p>This concentration suggests a strategic focus on this particular investment approach, which may reflect the portfolio's risk-return objectives.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with tab3:
        st.markdown('<div class="sub-header">NAV Distribution by Main Characteristic</div>', unsafe_allow_html=True)
        
        try:
            # Group by main characteristic
            char_data = filtered_df.groupby(main_characteristic_column)[nav_column].sum().reset_index()
            char_data = char_data.sort_values(by=nav_column, ascending=False)
            
            # Count investments per characteristic
            char_count = filtered_df.groupby(main_characteristic_column).size().reset_index(name='Count')
            char_count = char_count.sort_values(by='Count', ascending=False)
            
            # Create charts
            col1, col2 = st.columns(2)
            
            with col1:
                # Bar chart for NAV
                fig_char_bar = px.bar(
                    char_data,
                    x=main_characteristic_column,
                    y=nav_column,
                    title=f"NAV by {main_characteristic_column}",
                    labels={nav_column: "NAV (ILS)", main_characteristic_column: "Main Characteristic"},
                    color=main_characteristic_column,
                    color_discrete_sequence=PROFESSIONAL_COLORS
                )
                fig_char_bar.update_layout(
                    height=500,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(40,45,60,0.8)',
                    font=dict(color='#F3F6FB'),
                    margin=dict(l=40, r=40, t=50, b=40)
                )
                st.plotly_chart(fig_char_bar, use_container_width=True)
            
            with col2:
                # Pie chart for NAV
                fig_char_pie = px.pie(
                    char_data,
                    values=nav_column,
                    names=main_characteristic_column,
                    title=f"NAV Distribution by {main_characteristic_column}",
                    hole=0.4,
                    color_discrete_sequence=PROFESSIONAL_COLORS
                )
                fig_char_pie.update_layout(
                    height=500,
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(40,45,60,0.8)',
                    font=dict(color='#F3F6FB'),
                    margin=dict(l=40, r=40, t=50, b=40)
                )
                st.plotly_chart(fig_char_pie, use_container_width=True)
            
            # Table showing count of investments per characteristic
            st.subheader("Number of Investments by Main Characteristic")
            
            # Merge NAV and count data
            merged_data = pd.merge(char_data, char_count, on=main_characteristic_column)
            merged_data['Average NAV'] = merged_data[nav_column] / merged_data['Count']
            merged_data = merged_data.sort_values(by=nav_column, ascending=False)
            
            # Format the table columns
            merged_data['NAV (ILS)'] = merged_data[nav_column].apply(lambda x: format_number(x))
            merged_data['Average NAV (ILS)'] = merged_data['Average NAV'].apply(lambda x: format_number(x))
            merged_data['% of Total NAV'] = (merged_data[nav_column] / total_nav * 100).apply(lambda x: f"{x:.1f}%")
            
            # Display the table
            display_cols = [main_characteristic_column, 'Count', 'NAV (ILS)', 'Average NAV (ILS)', '% of Total NAV']
            st.dataframe(
                merged_data[display_cols],
                use_container_width=True,
                hide_index=True,
                column_config={
                    main_characteristic_column: st.column_config.TextColumn(
                        "Main Characteristic",
                        help="Investment main characteristic",
                        width="medium"
                    ),
                    'Count': st.column_config.NumberColumn(
                        "Count",
                        help="Number of investments",
                        format="%d"
                    ),
                    'NAV (ILS)': st.column_config.TextColumn(
                        "NAV (ILS)",
                        help="Total NAV in ILS"
                    ),
                    'Average NAV (ILS)': st.column_config.TextColumn(
                        "Average NAV (ILS)",
                        help="Average NAV per investment"
                    ),
                    '% of Total NAV': st.column_config.TextColumn(
                        "% of Total NAV",
                        help="Percentage of total NAV"
                    )
                }
            )
            
            # Main characteristic insights
            top_char = char_data.iloc[0][main_characteristic_column] if not char_data.empty else "N/A"
            top_char_pct = (char_data.iloc[0][nav_column] / total_nav) * 100 if not char_data.empty and total_nav > 0 else 0
            most_common_char = char_count.iloc[0][main_characteristic_column] if not char_count.empty else "N/A"
            most_common_count = char_count.iloc[0]['Count'] if not char_count.empty else 0
            
            st.markdown(f"""
            <div class="insight-box">
                <h3>Main Characteristic Insights</h3>
                <p>The dominant main characteristic by NAV is <b>{top_char}</b>, accounting for <b>{top_char_pct:.1f}%</b> of the total NAV.</p>
                <p>The most common main characteristic is <b>{most_common_char}</b> with <b>{most_common_count}</b> investments.</p>
                <p>This analysis reveals {'a strong alignment' if top_char == most_common_char else 'a potential mismatch'} between the number of investments and their total value across different main characteristics.</p>
            </div>
            """, unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"Error analyzing main characteristic data: {str(e)}")
            st.info("Please ensure the main characteristic column is correctly selected and contains valid data.")

    
    with tab4:
        st.markdown('<div class="sub-header">Detailed Investment Data</div>', unsafe_allow_html=True)
        
        # Add drilldown functionality with error handling
        try:
            drilldown_type = st.radio(
                "Drilldown By:",
                options=["Geography", "Strategy", "Main Characteristic"],
                horizontal=True
            )
            
            if drilldown_type == "Geography":
                # Get unique geography values, handle potential errors
                try:
                    geo_options = sorted(df[geography_column].dropna().unique().tolist())
                    if not geo_options:  # If list is empty
                        st.warning(f"No valid values found in the {geography_column} column.")
                        geo_options = ["No Data"]
                except Exception as e:
                    st.error(f"Error getting geography values: {str(e)}")
                    geo_options = ["Error"]
                    
                drilldown_value = st.selectbox(
                    "Select Geography:",
                    options=geo_options
                )
                
                # Filter data based on selection
                if drilldown_value not in ["No Data", "Error"]:
                    drilldown_df = df[df[geography_column] == drilldown_value]
                else:
                    drilldown_df = pd.DataFrame()
            elif drilldown_type == "Strategy":  # Strategy drilldown
                # Get unique strategy values, handle potential errors
                try:
                    strategy_options = sorted(df[strategy_column].dropna().unique().tolist())
                    if not strategy_options:  # If list is empty
                        st.warning(f"No valid values found in the {strategy_column} column.")
                        strategy_options = ["No Data"]
                except Exception as e:
                    st.error(f"Error getting strategy values: {str(e)}")
                    strategy_options = ["Error"]
                    
                drilldown_value = st.selectbox(
                    "Select Strategy:",
                    options=strategy_options
                )
                
                # Filter data based on selection
                if drilldown_value not in ["No Data", "Error"]:
                    drilldown_df = df[df[strategy_column] == drilldown_value]
                else:
                    drilldown_df = pd.DataFrame()
            else:  # Main Characteristic drilldown
                # Get unique main characteristic values, handle potential errors
                try:
                    char_options = sorted(df[main_characteristic_column].dropna().unique().tolist())
                    if not char_options:  # If list is empty
                        st.warning(f"No valid values found in the {main_characteristic_column} column.")
                        char_options = ["No Data"]
                except Exception as e:
                    st.error(f"Error getting main characteristic values: {str(e)}")
                    char_options = ["Error"]
                    
                drilldown_value = st.selectbox(
                    "Select Main Characteristic:",
                    options=char_options
                )
                
                # Filter data based on selection
                if drilldown_value not in ["No Data", "Error"]:
                    drilldown_df = df[df[main_characteristic_column] == drilldown_value]
                else:
                    drilldown_df = pd.DataFrame()
        except Exception as e:
            st.error(f"Error in drilldown functionality: {str(e)}")
            drilldown_df = pd.DataFrame()
            drilldown_value = "Error"
        
        # Display drilldown data with error handling
        st.subheader(f"Investments in {drilldown_value}")
        
        try:
            if not drilldown_df.empty:
                # Calculate subtotal with error handling
                try:
                    subtotal_nav = drilldown_df[nav_column].sum()
                    subtotal_pct = (subtotal_nav / total_nav) * 100 if total_nav > 0 else 0
                    
                    st.write(f"Total NAV: {format_number(subtotal_nav)} ILS ({subtotal_pct:.1f}% of portfolio)")
                    st.write(f"Number of investments: {len(drilldown_df)}")
                except Exception as e:
                    st.error(f"Error calculating subtotals: {str(e)}")
                    st.write(f"Number of investments: {len(drilldown_df)}")
                
                # Format numeric columns for display
                display_df = drilldown_df.copy()
                
                # Try to format the NAV column if it exists
                try:
                    if nav_column in display_df.columns:
                        display_df[nav_column] = display_df[nav_column].apply(lambda x: format_number(x))
                except Exception as e:
                    st.warning(f"Could not format NAV column: {str(e)}")
                
                # Display the data with styled columns
                column_configs = {}
                
                # Add column configs for common columns
                if nav_column in display_df.columns:
                    column_configs[nav_column] = st.column_config.TextColumn(
                        "NAV (ILS)",
                        help="Net Asset Value in ILS"
                    )
                
                if fund_name_column in display_df.columns:
                    column_configs[fund_name_column] = st.column_config.TextColumn(
                        "Fund Name",
                        help="Name of the investment fund",
                        width="large"
                    )
                    
                if strategy_column in display_df.columns:
                    column_configs[strategy_column] = st.column_config.TextColumn(
                        "Strategy",
                        help="Investment strategy"
                    )
                    
                if geography_column in display_df.columns:
                    column_configs[geography_column] = st.column_config.TextColumn(
                        "Geography",
                        help="Geographic location"
                    )
                    
                if main_characteristic_column in display_df.columns:
                    column_configs[main_characteristic_column] = st.column_config.TextColumn(
                        "Main Characteristic",
                        help="Main investment characteristic"
                    )
                
                try:
                    st.dataframe(
                        display_df, 
                        use_container_width=True,
                        column_config=column_configs,
                        hide_index=True
                    )
                except Exception as e:
                    st.error(f"Error displaying data table: {str(e)}")
                    st.write("Raw data preview:")
                    st.write(drilldown_df.head().to_dict())
                
                # Add download button for filtered data
                try:
                    csv = drilldown_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="Download Data as CSV",
                        data=csv,
                        file_name=f"investments_{drilldown_value}.csv",
                        mime="text/csv",
                        key="download-csv"
                    )
                except Exception as e:
                    st.error(f"Error creating download button: {str(e)}")
            else:
                st.warning("No data available for the selected filter.")
        except Exception as e:
            st.error(f"Unexpected error in drilldown section: {str(e)}")
            st.write("Please try selecting different filter options.")

    
    with tab5:
        st.markdown('<div class="sub-header">Additional Insights</div>', unsafe_allow_html=True)
        
        # Create additional analysis based on available columns
        additional_analyses = []
        
        # Israel vs. International analysis
        if geography_column:
            df['Israel_Flag'] = df[geography_column].apply(lambda x: 'Israel' if str(x).lower() in ['israel', 'ישראל', 'isr', 'il'] else 'International')
            israel_data = df.groupby('Israel_Flag')[nav_column].sum().reset_index()
            
            additional_analyses.append({
                "title": "Israel vs. International Exposure",
                "data": israel_data,
                "x": "Israel_Flag",
                "y": nav_column,
                "color": "Israel_Flag"
            })
        
        # Currency analysis
        if currency_column:
            currency_data = df.groupby(currency_column)[nav_column].sum().reset_index()
            currency_data = currency_data.sort_values(by=nav_column, ascending=False)
            
            additional_analyses.append({
                "title": f"NAV by {currency_column}",
                "data": currency_data,
                "x": currency_column,
                "y": nav_column,
                "color": currency_column
            })
        
        # Valuation agent analysis
        if valuation_agent_column:
            agent_data = df.groupby(valuation_agent_column)[nav_column].sum().reset_index()
            agent_data = agent_data.sort_values(by=nav_column, ascending=False)
            
            additional_analyses.append({
                "title": f"NAV by {valuation_agent_column}",
                "data": agent_data,
                "x": valuation_agent_column,
                "y": nav_column,
                "color": valuation_agent_column
            })
        
        # Year analysis
        if year_column:
            # Try to convert to year if it's a date
            try:
                if pd.api.types.is_datetime64_dtype(df[year_column]) or isinstance(df[year_column].iloc[0], str):
                    df['Investment_Year'] = pd.to_datetime(df[year_column], errors='coerce').dt.year
                    year_data = df.groupby('Investment_Year')[nav_column].sum().reset_index()
                    year_data = year_data.sort_values(by='Investment_Year')
                    
                    additional_analyses.append({
                        "title": "NAV by Investment Year",
                        "data": year_data,
                        "x": "Investment_Year",
                        "y": nav_column,
                        "color": "Investment_Year"
                    })
            except:
                pass
        
        # GP analysis
        if gp_column:
            gp_data = df.groupby(gp_column)[nav_column].sum().reset_index()
            gp_data = gp_data.sort_values(by=nav_column, ascending=False).head(10)  # Top 10 GPs
            
            additional_analyses.append({
                "title": f"Top 10 {gp_column} by NAV",
                "data": gp_data,
                "x": gp_column,
                "y": nav_column,
                "color": gp_column
            })
        
        # Display additional analyses
        for i, analysis in enumerate(additional_analyses):
            st.subheader(analysis["title"])
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Bar chart
                fig_bar = px.bar(
                    analysis["data"],
                    x=analysis["x"],
                    y=analysis["y"],
                    labels={analysis["y"]: "NAV (ILS)", analysis["x"]: analysis["x"].replace("_", " ")},
                    color=analysis["color"],
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_bar.update_layout(height=400)
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with col2:
                # Pie chart
                fig_pie = px.pie(
                    analysis["data"],
                    values=analysis["y"],
                    names=analysis["x"],
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            # Add insights
            if analysis["title"] == "Israel vs. International Exposure":
                israel_nav = analysis["data"][analysis["data"]["Israel_Flag"] == "Israel"][nav_column].sum()
                israel_pct = (israel_nav / total_nav) * 100
                
                st.markdown(f"""
                <div class="insight-box">
                    <h3>Israel vs. International Insight</h3>
                    <p>Israeli investments represent <b>{israel_pct:.1f}%</b> of the total NAV.</p>
                    <p>This {israel_pct > 50 and 'high' or 'low'} concentration in the Israeli market may reflect {israel_pct > 50 and 'a strategic focus on the local market' or 'a diversification strategy beyond the local market'}.</p>
                </div>
                """, unsafe_allow_html=True)
        
        # Overall portfolio insights
        st.subheader("Portfolio Overview Insights")
        
        # Calculate concentration metrics
        top_geo_concentration = (geo_data.iloc[0][nav_column] / total_nav) * 100 if len(geo_data) > 0 else 0
        top_strategy_concentration = (strategy_data.iloc[0][nav_column] / total_nav) * 100 if len(strategy_data) > 0 else 0
        
        # HHI Index (Herfindahl-Hirschman Index) for concentration
        geo_hhi = ((geo_data[nav_column] / total_nav) ** 2).sum() * 10000 if len(geo_data) > 0 else 0
        strategy_hhi = ((strategy_data[nav_column] / total_nav) ** 2).sum() * 10000 if len(strategy_data) > 0 else 0
        
        st.markdown(f"""
        <div class="insight-box">
            <h3>Portfolio Concentration Analysis</h3>
            <p>The top geography represents <b>{top_geo_concentration:.1f}%</b> of the portfolio (HHI: {geo_hhi:.0f}/10000).</p>
            <p>The top strategy represents <b>{top_strategy_concentration:.1f}%</b> of the portfolio (HHI: {strategy_hhi:.0f}/10000).</p>
            <p>This indicates a {'highly concentrated' if (geo_hhi > 2500 or strategy_hhi > 2500) else 'moderately concentrated' if (geo_hhi > 1500 or strategy_hhi > 1500) else 'diversified'} portfolio.</p>
        </div>
        """, unsafe_allow_html=True)

def main():
    """Main function to run the Streamlit app."""
    # App title and description
    st.sidebar.title("Investment Funds Analysis")
    st.sidebar.write("Upload your funds data file to analyze and visualize the investment portfolio.")
    
    # File uploader
    uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Load and process data
        df = load_data(uploaded_file)
        
        if df is not None:
            # Create dashboard
            create_dashboard(df)
    else:
        # Display instructions
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
