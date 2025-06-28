import os
import pandas as pd

def init_dict(file_path="./Excels/sales_purchase_data.xlsx", broker_names=[]): 
    sales_purchase_dict = {}
    
    # Check if the file path exists
    if os.path.exists(file_path):
        # Load the Excel file with explicit engine
        print("file_path", file_path)
        xls = pd.ExcelFile(file_path)
        print("xls", xls)
        
        # Initialize sales_purchase_dict with sheet names as keys and the last rows of each sheet as DataFrames
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            sales_purchase_dict[sheet_name] = df.tail(1)  # Get the last row of the sheet
        
    else:
        # File path does not exist, initialize with empty DataFrames
        print("sales_purchase_data.xlsx not found")
        sales_purchase_dict['Overall'] = pd.DataFrame(columns=['Date', 'Value', 'Purchase', 'Sales', 'Net Fund', 'Units', 'NAV'])
        
        # Initialize with given broker names
        for broker_name in broker_names:
            sales_purchase_dict[broker_name] = pd.DataFrame(columns=['Date', 'Value', 'Purchase', 'Sales', 'Net Fund', 'Units', 'NAV'])
    
    return sales_purchase_dict

def save_sales_purchase_dict(sales_purchase_dict, filename='./Excels/sales_purchase_data.xlsx'):
    print("save sales_purchase_dict")
    print(sales_purchase_dict)
    # Check if the file already exists
    if not os.path.exists(filename):
        # Save the dictionary to a new Excel file
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for sheet_name, df in sales_purchase_dict.items():
                if not sheet_name:
                    sheet_name = "unknown"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"File '{filename}' created successfully.")
    else:
        # File exists; read existing sheets
        existing_data = pd.read_excel(filename, sheet_name=None, engine='openpyxl')
        
        # Create a writer to overwrite the file after checking overlaps
        with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
            for sheet_name, new_df in sales_purchase_dict.items():
                if sheet_name in existing_data:
                    # Merge existing and new data, ensuring no overlap
                    existing_df = existing_data[sheet_name]
                    
                    # Find the maximum date in the existing DataFrame and convert it to Timestamp
                    max_date = pd.to_datetime(existing_df['Date'].max())
                    
                    # Ensure new_df['Date'] is also of Timestamp type
                    new_df['Date'] = pd.to_datetime(new_df['Date'])
                    
                    # Filter new_df to only include entries after the maximum date
                    non_overlapping_df = new_df[new_df['Date'] > max_date]
                    
                    # Concatenate existing data with non-overlapping new data
                    combined_df = pd.concat([existing_df, non_overlapping_df], ignore_index=True)
                else:
                    # No existing data for this sheet; just use new data
                    combined_df = new_df
                
                # Save the combined DataFrame to the Excel file
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"File '{filename}' updated with new data without duplicating overlapping days.")
