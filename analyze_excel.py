import sys
import pandas as pd

def analyze_excel_file(file_path):
    """Analyze an Excel file and print its structure and sample data"""
    try:
        print(f"\n--- ANALYZING EXCEL FILE: {file_path} ---\n")
        
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Basic information
        print(f"Shape: {df.shape} (rows, columns)")
        print(f"Columns: {list(df.columns)}")
        print("\nData Types:")
        for col, dtype in df.dtypes.items():
            print(f"  - {col}: {dtype}")
        
        # Sample data (first 5 rows)
        print("\nSample Data (first 5 rows):")
        print(df.head(5).to_string())
        
        # Summary statistics for numeric columns
        print("\nSummary Statistics:")
        for col in df.select_dtypes(include=['number']).columns:
            print(f"  - {col}:")
            print(f"    Min: {df[col].min()}")
            print(f"    Max: {df[col].max()}")
            print(f"    Mean: {df[col].mean()}")
            
        # Unique values for categorical columns (limited to 10)
        print("\nUnique Values in Categorical Columns:")
        for col in df.select_dtypes(include=['object']).columns:
            unique_values = df[col].unique()
            print(f"  - {col}: {len(unique_values)} unique values")
            if len(unique_values) <= 10:
                print(f"    Values: {list(unique_values)}")
            else:
                print(f"    Sample values: {list(unique_values[:5])}")
        
        # Check for missing values
        missing = df.isnull().sum()
        if missing.sum() > 0:
            print("\nMissing Values:")
            for col, count in missing.items():
                if count > 0:
                    print(f"  - {col}: {count} missing values")
        else:
            print("\nNo missing values found.")
            
        print("\n--- ANALYSIS COMPLETE ---\n")
        
    except Exception as e:
        print(f"Error analyzing file: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        for file_path in sys.argv[1:]:
            analyze_excel_file(file_path)
    else:
        print("Usage: python analyze_excel.py <excel_file1> <excel_file2> ...")
        print("Example: python analyze_excel.py Clockify_Time_Report.xlsx projects.xlsx hr.xlsx") 