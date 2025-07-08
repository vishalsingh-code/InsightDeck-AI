#!/usr/bin/env python3
"""
Test script to demonstrate CSV and Excel file support
"""

import pandas as pd
import os
from advanced_ppt_generator import CSVPPTGenerator

def create_sample_excel_file():
    """Create a sample Excel file with multiple sheets for testing"""
    
    # Sample data for different sheets
    sales_data = {
        'Product': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Tablet'],
        'Sales': [1200, 300, 150, 800, 900],
        'Region': ['North', 'South', 'East', 'West', 'North'],
        'Quarter': ['Q1', 'Q2', 'Q1', 'Q3', 'Q2']
    }
    
    employee_data = {
        'Name': ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve'],
        'Department': ['Sales', 'IT', 'Marketing', 'Sales', 'IT'],
        'Salary': [50000, 60000, 45000, 52000, 58000],
        'Experience': [3, 5, 2, 4, 6]
    }
    
    time_series_data = {
        'Date': pd.date_range('2023-01-01', periods=12, freq='M'),
        'Revenue': [100000, 110000, 105000, 120000, 115000, 130000, 
                   125000, 140000, 135000, 150000, 145000, 160000],
        'Expenses': [80000, 85000, 82000, 90000, 87000, 95000,
                    92000, 100000, 98000, 105000, 102000, 110000]
    }
    
    # Create Excel file with multiple sheets
    excel_filename = 'sample_data.xlsx'
    with pd.ExcelWriter(excel_filename) as writer:
        pd.DataFrame(sales_data).to_excel(writer, sheet_name='Sales_Data', index=False)
        pd.DataFrame(employee_data).to_excel(writer, sheet_name='Employee_Data', index=False)
        pd.DataFrame(time_series_data).to_excel(writer, sheet_name='Time_Series', index=False)
        # Add an empty sheet to test sheet detection
        pd.DataFrame().to_excel(writer, sheet_name='Empty_Sheet', index=False)
    
    print(f"✅ Created sample Excel file: {excel_filename}")
    return excel_filename

def test_file_detection():
    """Test file type detection"""
    print("\n🔍 Testing File Type Detection...")
    generator = CSVPPTGenerator()
    
    # Test CSV detection
    if os.path.exists('test_data.csv'):
        file_type = generator.detect_file_type('test_data.csv')
        print(f"  CSV file detected as: {file_type}")
    
    # Test Excel detection
    excel_file = create_sample_excel_file()
    file_type = generator.detect_file_type(excel_file)
    print(f"  Excel file detected as: {file_type}")
    
    return excel_file

def test_excel_info(excel_file):
    """Test Excel file information extraction"""
    print(f"\n📋 Testing Excel File Information...")
    generator = CSVPPTGenerator()
    
    try:
        excel_info = generator.load_excel_info(excel_file)
        print(f"  Total sheets: {excel_info['total_sheets']}")
        print(f"  Sheets with data: {excel_info['sheets_with_data']}")
        
        print("\n  Sheet Details:")
        for sheet_name, info in excel_info['sheets'].items():
            status = "✅" if info['has_data'] else "❌"
            print(f"    {status} {sheet_name}: {info['estimated_records']} rows")
            
    except Exception as e:
        print(f"  ❌ Error: {e}")

def test_data_loading(excel_file):
    """Test loading data from different sheets"""
    print(f"\n📊 Testing Data Loading...")
    generator = CSVPPTGenerator()
    
    try:
        # Test auto-selecting best sheet
        print("  Auto-selecting best sheet...")
        analysis = generator.load_and_analyze_data(excel_file)
        print(f"    Loaded: {analysis['shape'][0]} rows × {analysis['shape'][1]} columns")
        print(f"    Source: {analysis['source_metadata']['source_sheet']}")
        
        # Test specific sheet selection
        print("\n  Loading specific sheet 'Employee_Data'...")
        analysis = generator.load_and_analyze_data(excel_file, sheet_name='Employee_Data')
        print(f"    Loaded: {analysis['shape'][0]} rows × {analysis['shape'][1]} columns")
        print(f"    Columns: {', '.join(analysis['columns'])}")
        
    except Exception as e:
        print(f"  ❌ Error: {e}")

def test_presentation_generation(excel_file):
    """Test presentation generation from Excel file"""
    print(f"\n🎯 Testing Presentation Generation...")
    generator = CSVPPTGenerator()
    
    try:
        # Generate presentation from Excel file
        output_file = generator.create_presentation_from_csv(
            excel_file, 
            output_filename='excel_test_presentation.pptx',
            sheet_name='Sales_Data'
        )
        print(f"  ✅ Presentation generated: {output_file}")
        
    except Exception as e:
        print(f"  ❌ Error: {e}")

def test_csv_compatibility():
    """Test CSV file compatibility"""
    print(f"\n📄 Testing CSV Compatibility...")
    
    # Check if existing CSV files work
    csv_files = ['test_data.csv', 'time_series_data.csv']
    generator = CSVPPTGenerator()
    
    for csv_file in csv_files:
        if os.path.exists(csv_file):
            try:
                print(f"  Testing {csv_file}...")
                analysis = generator.load_and_analyze_data(csv_file)
                print(f"    ✅ Loaded: {analysis['shape'][0]} rows × {analysis['shape'][1]} columns")
            except Exception as e:
                print(f"    ❌ Error with {csv_file}: {e}")
        else:
            print(f"  ⚠️  {csv_file} not found, skipping...")

def main():
    """Main test function"""
    print("🧪 CSV/Excel PPT Generator - Compatibility Test")
    print("=" * 50)
    
    try:
        # Test file detection
        excel_file = test_file_detection()
        
        # Test Excel info extraction
        test_excel_info(excel_file)
        
        # Test data loading
        test_data_loading(excel_file)
        
        # Test presentation generation
        test_presentation_generation(excel_file)
        
        # Test CSV compatibility
        test_csv_compatibility()
        
        print("\n" + "=" * 50)
        print("🎉 All tests completed! The generator supports both CSV and Excel files.")
        print("\n💡 Key Features:")
        print("  • Automatic file type detection (.csv, .xlsx, .xls)")
        print("  • Excel sheet information and auto-selection")
        print("  • Multiple sheet support with manual selection")
        print("  • Consistent data analysis for both formats")
        print("  • Enhanced error handling and validation")
        
        # Cleanup
        if os.path.exists(excel_file):
            os.remove(excel_file)
            print(f"\n🧹 Cleaned up test file: {excel_file}")
            
    except Exception as e:
        print(f"\n❌ Test failed: {e}")

if __name__ == "__main__":
    main()
