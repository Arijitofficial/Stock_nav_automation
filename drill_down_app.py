import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry

def filter_and_save_csv(start_date, end_date):
    """
    Filters the master track CSV based on start_date and end_date, then saves the filtered data.
    """
    # Ensure the master CSV exists
    input_file = './Excels/drill_down_track.csv'
    output_dir = './drill_downs'
    if not os.path.exists(input_file):
        messagebox.showerror("Error", "Master CSV not found in './Excels'. Ensure the file exists.")
        return
    
    # Read the CSV file
    df = pd.read_csv(input_file)
    df['date'] = pd.to_datetime(df['date'])
    
    # Convert input dates to datetime
    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD.")
        return
    
    # Validate date range
    if start_date > end_date:
        messagebox.showerror("Error", "Start date must be before or equal to end date.")
        return
    
    # Filter data based on the date range
    df_range = df[(df['date'] >= start_date) & (df['date'] <= end_date)]
    filtered_df = df_range[df_range['date'].isin([df_range['date'].min(), df_range['date'].max()])]

    
    if filtered_df.empty:
        messagebox.showinfo("No Data", "No data found in the given date range.")
        return
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Save the filtered data
    output_file = os.path.join(output_dir, f'drill_down_{start_date.date()}_to_{end_date.date()}.csv')
    filtered_df.to_csv(output_file, index=False)
    
    messagebox.showinfo("Success", f"CSV file saved: {output_file}")

def create_gui():
    """
    Creates a simple Tkinter GUI for date input and CSV generation.
    """
    root = tk.Tk()
    root.title("Drill Down CSV Generator")
    root.geometry("400x300")
    
    tk.Label(root, text="Select Start Date:").pack(pady=5)
    start_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', date_pattern='yyyy-MM-dd')
    start_date_entry.pack(pady=5)
    
    tk.Label(root, text="Select End Date:").pack(pady=5)
    end_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', date_pattern='yyyy-MM-dd')
    end_date_entry.pack(pady=5)
    
    def on_calculate():
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()
        filter_and_save_csv(start_date, end_date)
    
    calculate_button = tk.Button(root, text="Generate CSV", command=on_calculate)
    calculate_button.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()
