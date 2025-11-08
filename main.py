# -*- coding: utf-8 -*-
import pandas as pd
import sys
import os

def csv_to_excel(csv_file_path, excel_file_path=None):
    """
    Convert CSV file to Excel format
    
    Args:
        csv_file_path (str): Path to the input CSV file
        excel_file_path (str): Path to the output Excel file (optional)
    """
    try:
        # Read and preprocess the CSV file to handle malformed quotes
        with open(csv_file_path, 'r', encoding='utf-8-sig') as file:
            lines = file.readlines()
        
        # Fix the malformed CSV by removing outer quotes from data rows
        fixed_lines = []
        for i, line in enumerate(lines):
            line = line.strip()
            if i == 0:  # Header row
                fixed_lines.append(line)
            else:  # Data rows - remove outer quotes and fix inner double quotes
                if line.startswith('"') and line.endswith('"'):
                    # Remove outer quotes
                    line = line[1:-1]
                    # Fix double quotes inside (convert "" to ")
                    line = line.replace('""', '"')
                fixed_lines.append(line)
        
        # Write fixed CSV to temporary string and read with pandas
        import io
        csv_string = '\n'.join(fixed_lines)
        df = pd.read_csv(io.StringIO(csv_string), sep=',')
        
        # If no output path specified, create one based on input filename
        if excel_file_path is None:
            base_name = os.path.splitext(csv_file_path)[0]
            excel_file_path = f"{base_name}.xlsx"
        
        # Write to Excel file
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Amazon Data', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Amazon Data']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"Successfully converted {csv_file_path} to {excel_file_path}")
        return excel_file_path
        
    except FileNotFoundError:
        print(f"Error: File {csv_file_path} not found")
        return None
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return None

def main():
    # Default CSV file
    default_csv = "report-octobre.csv"
    
    # Check if CSV file argument is provided
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]
    else:
        csv_file = default_csv
    
    # Check if output file argument is provided
    excel_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Check if file exists
    if not os.path.exists(csv_file):
        print(f"Error: File {csv_file} does not exist")
        print(f"Usage: python main.py [csv_file] [excel_file]")
        print(f"Example: python main.py report-octobre.csv report-octobre.xlsx")
        return
    
    # Convert CSV to Excel
    result = csv_to_excel(csv_file, excel_file)
    
    if result:
        print(f"Excel file created: {result}")

if __name__ == "__main__":
    main()