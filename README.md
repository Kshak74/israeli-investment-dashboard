# Israeli Investment Dashboard

An interactive dashboard for analyzing Israeli insurance companies' investment fund data.

## Features

- **File Upload**: Easily upload Excel files with investment data
- **Automatic Classification**: Categorizes investments into sub-segments (Buyout, VC, etc.)
- **Interactive Visualizations**: Charts for investment distribution, regional analysis, and fund types
- **Filtering**: Filter by sub-segment, region, and fund type
- **Data Table**: Searchable and sortable data table with export options
- **Hebrew Support**: Handles Hebrew text in Excel files
- **Visualizations**:
  - Pie chart showing investment distribution by sub-segment
  - Bar chart showing investments by region
- **Data Table**: Detailed view of filtered investments with sorting and pagination
- **Responsive Design**: Works on different screen sizes

## Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

## Installation

1. Clone this repository or download the files to your local machine
2. Navigate to the project directory:
   ```
   cd path/to/israeli-investment-dashboard
   ```
3. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Place your Excel file containing the "קרנות השקעה" sheet in the project directory
2. Update the `file_path` variable in `app.py` to point to your Excel file
3. Run the application:
   ```
   python app.py
   ```
4. Open your web browser and navigate to `http://127.0.0.1:8050/`

## Data Requirements

The Excel file should contain a sheet named "קרנות השקעה" with the following columns (column names should be in Hebrew):

- שם הקרן (Fund Name)
- תיאור (Description)
- איזור/אזור (Region) - optional
- סוג הקרן (Fund Type) - optional
- גודל ההשקעה (Investment Size) - should be numeric

## Dashboard Features

1. **Filters**:
   - Investment Sub-Segment (auto-classified)
   - Region (if available in data)
   - Fund Type (if available in data)

2. **Visualizations**:
   - Investment distribution by sub-segment (Pie Chart)
   - Investments by region (Bar Chart)

3. **Data Table**:
   - View all investment details
   - Sort by any column
   - Pagination for large datasets

## Customization

You can customize the dashboard by modifying the `app.py` file:

- Update the classification logic in the `classify_investment` function
- Add more visualizations using Plotly Express
- Modify the layout and styling using Dash Bootstrap Components

## License

This project is open source and available under the MIT License.

## Support

For any issues or feature requests, please open an issue in the repository.
