# CSV/Excel PPT Generator - Usage Guide

## Overview
The PPT Generator now fully supports both **CSV** and **Excel** files with advanced features for data analysis and presentation generation.

## Supported File Formats
- **CSV files**: `.csv`
- **Excel files**: `.xlsx`, `.xls`

## Key Features

### 🔍 Automatic File Detection
The system automatically detects whether your file is CSV or Excel format and handles them appropriately.

### 📋 Excel Advanced Features
- **Multiple Sheet Support**: Automatically lists all sheets and their data status
- **Smart Sheet Selection**: Auto-selects the sheet with the most data
- **Manual Sheet Selection**: Choose specific sheets by name or number
- **Named Range Support**: Load specific named ranges from Excel files
- **Data Validation**: Removes empty rows/columns and validates data quality

## Usage Examples

### 1. Command Line Usage

#### Basic Usage (Auto-detect format)
```bash
python3 advanced_ppt_generator.py your_data.csv
python3 advanced_ppt_generator.py your_data.xlsx
```

#### Excel-Specific Options
```bash
# List all sheets in Excel file
python3 advanced_ppt_generator.py your_data.xlsx --list-sheets

# Generate from specific sheet
python3 advanced_ppt_generator.py your_data.xlsx --sheet "Sales_Data"

# Generate from named range
python3 advanced_ppt_generator.py your_data.xlsx --sheet "Sheet1" --range "A1:D10"

# Specify output filename
python3 advanced_ppt_generator.py your_data.xlsx --output "my_presentation.pptx"
```

### 2. Interactive Usage

#### Run the examples script
```bash
python3 examples.py
```

Choose option **1. 📊 CSV/Excel Data Presentation** for guided Excel/CSV processing:
- Enter your file path
- For Excel files, select from available sheets
- Specify output filename (optional)

### 3. Programmatic Usage

```python
from advanced_ppt_generator import CSVPPTGenerator

# Initialize generator
generator = CSVPPTGenerator()

# CSV file
generator.create_presentation_from_csv("data.csv")

# Excel file (auto-select best sheet)
generator.create_presentation_from_csv("data.xlsx")

# Excel file (specific sheet)
generator.create_presentation_from_csv("data.xlsx", sheet_name="Sales_Data")

# Excel file (named range)
generator.create_presentation_from_csv("data.xlsx", sheet_name="Sheet1", named_range="A1:D10")
```

### 4. Advanced Excel Operations

#### Get Excel File Information
```python
# Get detailed information about Excel file
excel_info = generator.load_excel_info("data.xlsx")
print(f"Total sheets: {excel_info['total_sheets']}")

for sheet_name, info in excel_info['sheets'].items():
    print(f"{sheet_name}: {info['estimated_records']} rows")
```

#### Smart Sheet Selection
```python
# Let the system choose the best sheet automatically
analysis = generator.load_and_analyze_data("data.xlsx")
print(f"Selected sheet: {analysis['source_metadata']['source_sheet']}")
```

## Excel Sheet Selection Logic

The system uses intelligent logic to select the best sheet:
1. **Most Data**: Prioritizes sheets with the most rows
2. **Avoids Metadata**: Skips sheets named 'summary', 'metadata', 'info', etc.
3. **Data Quality**: Only considers sheets that actually contain data

## Error Handling

The system provides comprehensive error handling:
- **File Format Validation**: Checks for supported formats
- **Sheet Existence**: Validates sheet names exist
- **Data Validation**: Ensures sheets contain actual data
- **Graceful Fallbacks**: Provides alternatives when specific requests fail

## Data Analysis Features

Both CSV and Excel files get the same comprehensive analysis:
- **Statistical Analysis**: Mean, median, std deviation, outliers
- **Data Quality Assessment**: Missing values, duplicates, completeness
- **Correlation Analysis**: Relationships between numeric variables
- **Pattern Detection**: Time series, high cardinality categories
- **Chart Recommendations**: 5+ different visualization types

## Example Excel File Structure

For best results, structure your Excel files like this:

### Sheet: "Sales_Data"
| Product | Sales | Region | Quarter |
|---------|-------|--------|---------|
| Laptop  | 1200  | North  | Q1      |
| Mouse   | 300   | South  | Q2      |

### Sheet: "Employee_Data"
| Name    | Department | Salary | Experience |
|---------|------------|--------|------------|
| Alice   | Sales      | 50000  | 3          |
| Bob     | IT         | 60000  | 5          |

## Chart Types Generated

The system automatically generates multiple chart types:
- **Bar Charts**: Category comparisons, distributions
- **Pie Charts**: Proportional breakdowns
- **Line Charts**: Trends and time series
- **Scatter Plots**: Variable relationships
- **Heatmaps**: Correlation matrices

## Tips for Best Results

### Excel Files
1. **Use Clear Headers**: First row should contain column names
2. **Consistent Data Types**: Keep data types consistent within columns
3. **Avoid Empty Sheets**: Remove or rename empty sheets
4. **Meaningful Sheet Names**: Use descriptive names like "Sales_2023" not "Sheet1"

### CSV Files
1. **UTF-8 Encoding**: Ensure proper character encoding
2. **Consistent Delimiters**: Use standard comma separators
3. **Clean Data**: Remove extra spaces and formatting

## Troubleshooting

### Common Issues

#### "Sheet not found"
- Check sheet name spelling (case-sensitive)
- Use `--list-sheets` to see available sheets

#### "No data found"
- Ensure sheet contains actual data (not just headers)
- Check for hidden characters or formatting issues

#### "Module not found" errors
Install required dependencies:
```bash
pip3 install openpyxl xlrd pandas matplotlib seaborn python-pptx openai python-dotenv
```

### Performance Notes
- Large Excel files (>10MB) may take longer to process
- Files with many sheets will show selection options
- Complex formulas in Excel are not evaluated (only values are used)

## Configuration

Create a `.env` file with your OpenAI API key:
```
OPENAI_API_KEY=your_api_key_here
```

## Examples Included

Run the test script to see the system in action:
```bash
python3 test_excel_support.py
```

This demonstrates:
- File type detection
- Excel sheet information extraction
- Data loading from multiple sheets
- Presentation generation
- CSV compatibility testing

---

## Need Help?

The system provides extensive logging and error messages to help debug issues. All operations include progress indicators and detailed feedback about what's happening behind the scenes.
