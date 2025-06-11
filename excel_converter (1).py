import pandas as pd
import json
import sys
from pathlib import Path

class SalesDataConverter:
    """
    Converts Excel sales data to the format required by the dashboard.
    Handles various Excel formats and data inconsistencies.
    """
    
    def __init__(self):
        self.unit_categories = ['0-4', '5-9', '10-17', '18-25', '26+']
        self.required_columns = [
            'Sales Rep',
            'Issued Appts', 
            'Overall Close %',
            'Units Captured on Sold Jobs %'
        ]
        
    def detect_column_format(self, df):
        """Detect the format of columns in the Excel file"""
        columns = df.columns.tolist()
        
        # Check for standard format
        category_pattern_found = any(f'({cat}) Issued Appts' in str(col) for col in columns for cat in self.unit_categories)
        
        if category_pattern_found:
            return "standard"
        else:
            return "custom"
    
    def safe_convert_numeric(self, value, is_percentage=False, default=0):
        """Safely convert values to numeric, handling various formats"""
        if pd.isna(value) or value == '-' or value == '' or value is None:
            return None if is_percentage and default == 0 else default
            
        try:
            # Handle string values
            if isinstance(value, str):
                # Remove percentage signs and convert
                value = value.replace('%', '').replace(',', '').strip()
                if value == '' or value == '-':
                    return None if is_percentage and default == 0 else default
            
            numeric_val = float(value)
            
            # Convert percentage format (0.0-1.0) to percentage (0-100)
            if is_percentage and numeric_val <= 1.0:
                numeric_val *= 100
                
            return numeric_val
            
        except (ValueError, TypeError):
            print(f"Warning: Could not convert '{value}' to numeric. Using default: {default}")
            return None if is_percentage and default == 0 else default
    
    def process_standard_format(self, df):
        """Process Excel data in standard format (matching your current structure)"""
        sales_data = []
        
        for idx, row in df.iterrows():
            try:
                # Basic rep information
                rep_name = str(row.get('Sales Rep', f'Rep_{idx}')).strip()
                if rep_name == '' or rep_name.lower() == 'nan':
                    continue
                    
                rep_data = {
                    'name': rep_name,
                    'totalAppts': self.safe_convert_numeric(row.get('Issued Appts', 0)),
                    'overallClose': self.safe_convert_numeric(row.get('Overall Close %', 0), is_percentage=True),
                    'overallCapture': self.safe_convert_numeric(row.get('Units Captured on Sold Jobs %', 0), is_percentage=True),
                    'categories': {}
                }
                
                # Process each unit category
                for cat in self.unit_categories:
                    appts_col = f'({cat}) Issued Appts'
                    close_col = f'({cat}) Overall Close %'
                    capture_col = f'({cat}) Units Captured on Sold Jobs %'
                    
                    appointments = self.safe_convert_numeric(row.get(appts_col, 0))
                    close_rate = self.safe_convert_numeric(row.get(close_col), is_percentage=True)
                    capture_rate = self.safe_convert_numeric(row.get(capture_col), is_percentage=True)
                    
                    rep_data['categories'][cat] = {
                        'appointments': appointments,
                        'closeRate': close_rate,
                        'captureRate': capture_rate
                    }
                
                sales_data.append(rep_data)
                
            except Exception as e:
                print(f"Warning: Error processing row {idx} for {rep_name}: {str(e)}")
                continue
        
        return sales_data
    
    def process_custom_format(self, df):
        """Process Excel data in custom/different format"""
        print("Detected custom format. Attempting intelligent mapping...")
        
        # Try to map columns intelligently
        column_mapping = {}
        columns = df.columns.tolist()
        
        # Map common column variations
        for col in columns:
            col_lower = str(col).lower()
            if 'sales rep' in col_lower or 'rep name' in col_lower or 'name' in col_lower:
                column_mapping['Sales Rep'] = col
            elif 'total appt' in col_lower or 'issued appt' in col_lower or 'appointments' in col_lower:
                column_mapping['Issued Appts'] = col
            elif 'overall close' in col_lower or 'close rate' in col_lower or 'close %' in col_lower:
                column_mapping['Overall Close %'] = col
            elif 'capture' in col_lower and ('overall' in col_lower or 'total' in col_lower):
                column_mapping['Units Captured on Sold Jobs %'] = col
        
        print(f"Column mapping: {column_mapping}")
        
        # Create a standardized DataFrame
        standardized_data = []
        for idx, row in df.iterrows():
            try:
                rep_data = {
                    'Sales Rep': row.get(column_mapping.get('Sales Rep', 'Sales Rep'), f'Rep_{idx}'),
                    'Issued Appts': row.get(column_mapping.get('Issued Appts', 'Issued Appts'), 0),
                    'Overall Close %': row.get(column_mapping.get('Overall Close %', 'Overall Close %'), 0),
                    'Units Captured on Sold Jobs %': row.get(column_mapping.get('Units Captured on Sold Jobs %', 'Units Captured on Sold Jobs %'), 0)
                }
                
                # Try to find category columns
                for cat in self.unit_categories:
                    for col in columns:
                        col_str = str(col).lower()
                        if cat.replace('-', '') in col_str.replace('-', '').replace(' ', ''):
                            if 'appt' in col_str:
                                rep_data[f'({cat}) Issued Appts'] = row.get(col, 0)
                            elif 'close' in col_str:
                                rep_data[f'({cat}) Overall Close %'] = row.get(col, 0)
                            elif 'capture' in col_str:
                                rep_data[f'({cat}) Units Captured on Sold Jobs %'] = row.get(col, 0)
                
                standardized_data.append(rep_data)
            except Exception as e:
                print(f"Warning: Error processing custom row {idx}: {str(e)}")
                continue
        
        # Convert to DataFrame and process with standard method
        standardized_df = pd.DataFrame(standardized_data)
        return self.process_standard_format(standardized_df)
    
    def validate_data(self, sales_data):
        """Validate the converted data and provide statistics"""
        if not sales_data:
            raise ValueError("No valid sales data found")
        
        stats = {
            'total_reps': len(sales_data),
            'reps_with_appointments': len([rep for rep in sales_data if rep['totalAppts'] > 0]),
            'avg_appointments': sum(rep['totalAppts'] for rep in sales_data) / len(sales_data),
            'categories_with_data': {}
        }
        
        # Check category data availability
        for cat in self.unit_categories:
            reps_with_cat_data = len([
                rep for rep in sales_data 
                if rep['categories'][cat]['appointments'] > 0 and 
                   rep['categories'][cat]['closeRate'] is not None
            ])
            stats['categories_with_data'][cat] = reps_with_cat_data
        
        return stats
    
    def convert_excel_file(self, file_path, output_format='json'):
        """Main method to convert Excel file to dashboard format"""
        try:
            # Read Excel file
            if isinstance(file_path, str):
                df = pd.read_excel(file_path)
            else:
                # Handle uploaded file objects (for Streamlit)
                df = pd.read_excel(file_path)
            
            print(f"Loaded Excel file with {len(df)} rows and {len(df.columns)} columns")
            print(f"Columns: {list(df.columns)}")
            
            # Detect format and process
            format_type = self.detect_column_format(df)
            print(f"Detected format: {format_type}")
            
            if format_type == "standard":
                sales_data = self.process_standard_format(df)
            else:
                sales_data = self.process_custom_format(df)
            
            # Validate data
            stats = self.validate_data(sales_data)
            print(f"Conversion successful: {stats}")
            
            if output_format == 'json':
                return json.dumps(sales_data, indent=2)
            else:
                return sales_data
                
        except Exception as e:
            print(f"Error converting Excel file: {str(e)}")
            raise e
    
    def create_sample_excel(self, filename="sample_sales_data.xlsx"):
        """Create a sample Excel file in the correct format"""
        sample_data = {
            'Sales Rep': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'Issued Appts': [25, 30, 20],
            'RpA 
                    ': [15000, 18000, 12000],
            'RpU 
                    ': [1500, 1800, 1200],
            'Overall Close %': [0.35, 0.42, 0.28],
            'Avg Sale Price 
                    ': [42857, 42857, 42857],
            'Units Captured on Sold Jobs %': [1.05, 0.98, 1.12],
        }
        
        # Add category columns
        categories = ['0-4', '5-9', '10-17', '18-25', '26+']
        for cat in categories:
            sample_data[f'({cat}) Issued Appts'] = [8, 6, 4, 4, 3]
            sample_data[f'({cat}) Mix %'] = [0.32, 0.20, 0.16, 0.16, 0.12]
            sample_data[f'({cat}) Overall Close %'] = [0.38, 0.50, 0.25, 0.25, 0.33]
            sample_data[f'({cat}) Units Captured on Sold Jobs %'] = [1.2, 0.9, 1.1, 0.8, 1.0]
        
        df = pd.DataFrame(sample_data)
        df.to_excel(filename, index=False)
        print(f"Sample Excel file created: {filename}")
        
        return filename

def main():
    """Command line interface for the converter"""
    if len(sys.argv) < 2:
        print("Usage: python converter.py <excel_file_path> [output_file]")
        print("Example: python converter.py sales_data.xlsx converted_data.json")
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    converter = SalesDataConverter()
    
    try:
        # Convert the file
        result = converter.convert_excel_file(input_file)
        
        if output_file:
            with open(output_file, 'w') as f:
                f.write(result)
            print(f"Conversion complete. Output saved to: {output_file}")
        else:
            print("Conversion complete. JSON output:")
            print(result[:500] + "..." if len(result) > 500 else result)
            
    except Exception as e:
        print(f"Conversion failed: {str(e)}")
        
        # Offer to create sample file
        create_sample = input("Would you like to create a sample Excel file? (y/n): ")
        if create_sample.lower() == 'y':
            sample_file = converter.create_sample_excel()
            print(f"Sample file created: {sample_file}")

if __name__ == "__main__":
    main()
                    '