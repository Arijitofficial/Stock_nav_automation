import pandas as pd
from datetime import datetime, timedelta
import os
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry


class DrillDownPivotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drill Down & Pivot Analysis")
        self.root.geometry("1400x800")
        
        # Data storage
        self.drill_down_df = None
        self.sell_purchase_df = None
        self.current_pivot_data = {}
        self.all_brokers = []
        
        # Load data
        self.load_data()
        
        # Create UI
        self.create_ui()
    
    def load_data(self):
        """Load drill down and sell/purchase tracking data"""
        drill_down_file = './Excels/drill_down_track.csv'
        sell_purchase_file = './Excels/sell_purchase_track.csv'
        
        if os.path.exists(drill_down_file):
            self.drill_down_df = pd.read_csv(drill_down_file)
            self.drill_down_df['date'] = pd.to_datetime(self.drill_down_df['date'])
            self.all_brokers = sorted(self.drill_down_df['broker'].unique().tolist())
        else:
            messagebox.showerror("Error", f"Drill down file not found: {drill_down_file}")
            self.root.destroy()
            return
        
        if os.path.exists(sell_purchase_file):
            self.sell_purchase_df = pd.read_csv(sell_purchase_file)
            self.sell_purchase_df['Purchase Date'] = pd.to_datetime(
                self.sell_purchase_df['Purchase Date'], errors='coerce'
            )
            self.sell_purchase_df['Sell Date'] = pd.to_datetime(
                self.sell_purchase_df['Sell Date'], errors='coerce'
            )
        else:
            messagebox.showwarning("Warning", "Sell/Purchase track file not found. Purchase/Sell values may be inaccurate.")
            self.sell_purchase_df = pd.DataFrame()
    
    def create_ui(self):
        """Create the main user interface"""
        # Create main container
        main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel for controls
        left_panel = ttk.Frame(main_container, width=300)
        main_container.add(left_panel, weight=0)
        
        # Right panel for table display
        right_panel = ttk.Frame(main_container)
        main_container.add(right_panel, weight=1)
        
        # Build left panel
        self.create_control_panel(left_panel)
        
        # Build right panel
        self.create_table_panel(right_panel)
    
    def create_control_panel(self, parent):
        """Create the left control panel"""
        # Title
        title_label = ttk.Label(parent, text="Pivot Analysis", font=('Arial', 14, 'bold'))
        title_label.pack(pady=10)
        
        # Date Selection Frame
        date_frame = ttk.LabelFrame(parent, text="Select End Date", padding=10)
        date_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.end_date_entry = DateEntry(
            date_frame, 
            width=20, 
            background='darkblue',
            foreground='white', 
            date_pattern='yyyy-MM-dd'
        )
        self.end_date_entry.pack(pady=5)
        
        # Broker and Duration Selection Frame
        selection_frame = ttk.LabelFrame(parent, text="Select Broker & Duration", padding=10)
        selection_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Tree view for broker and duration
        self.tree = ttk.Treeview(selection_frame, selectmode='browse', height=15)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar for tree
        tree_scroll = ttk.Scrollbar(selection_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        # Populate tree
        self.populate_tree()
        
        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # Buttons Frame
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame,
            text="Generate Pivot",
            command=self.generate_pivot
        ).pack(fill=tk.X, pady=2)
        
        ttk.Button(
            button_frame,
            text="Download Current Table",
            command=self.download_current_table
        ).pack(fill=tk.X, pady=2)
        
        ttk.Button(
            button_frame,
            text="Download All Pivot Tables",
            command=self.download_all_pivot_tables
        ).pack(fill=tk.X, pady=2)
        
        ttk.Button(
            button_frame,
            text="Generate Drill Down CSV",
            command=self.generate_drill_down_csv
        ).pack(fill=tk.X, pady=2)
    
    def populate_tree(self):
        """Populate the tree with brokers and durations"""
        durations = [
            ('1M', '1 Month'),
            ('3M', '3 Months'),
            ('6M', '6 Months'),
            ('9M', '9 Months'),
            ('12M', '12 Months')
        ]
        
        for broker in self.all_brokers:
            broker_node = self.tree.insert('', 'end', text=broker, values=(broker, ''))
            
            for duration_code, duration_name in durations:
                self.tree.insert(
                    broker_node, 
                    'end', 
                    text=duration_name,
                    values=(broker, duration_code)
                )
    
    def create_table_panel(self, parent):
        """Create the right panel for table display"""
        # Title
        self.table_title = ttk.Label(
            parent, 
            text="Select a broker and duration to view pivot table",
            font=('Arial', 12, 'bold')
        )
        self.table_title.pack(pady=10)
        
        # Table frame
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview for table
        columns = (
            'Stock Name',
            'Qty (Start)',
            'Value (Start)',
            'Qty (End)',
            'Value (End)',
            'Purchase Value',
            'Sell Value',
            'Net Change'
        )
        
        self.table_tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show='headings',
            height=30
        )
        
        # Define column headings and widths
        column_widths = {
            'Stock Name': 150,
            'Qty (Start)': 100,
            'Value (Start)': 120,
            'Qty (End)': 100,
            'Value (End)': 120,
            'Purchase Value': 130,
            'Sell Value': 120,
            'Net Change': 120
        }
        
        for col in columns:
            self.table_tree.heading(col, text=col)
            self.table_tree.column(col, width=column_widths.get(col, 100), anchor='center')
        
        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.table_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table_tree.xview)
        self.table_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.table_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
    
    def on_tree_select(self, event):
        """Handle tree selection"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = self.tree.item(selection[0])
        values = item['values']
        
        if len(values) == 2 and values[1]:  # Has duration
            broker, duration = values
            self.current_selection = {'broker': broker, 'duration': duration}
    
    def generate_pivot(self):
        """Generate pivot table based on selection"""
        if not hasattr(self, 'current_selection'):
            messagebox.showwarning("Warning", "Please select a broker and duration from the tree")
            return
        
        broker = self.current_selection['broker']
        duration = self.current_selection['duration']
        end_date = datetime.strptime(self.end_date_entry.get(), '%Y-%m-%d')
        
        # Calculate start date based on duration
        duration_map = {
            '1M': 1,
            '3M': 3,
            '6M': 6,
            '9M': 9,
            '12M': 12
        }
        
        months = duration_map.get(duration, 1)
        start_date = end_date - timedelta(days=months * 30)  # Approximate
        
        # Generate pivot data
        pivot_df = self.calculate_pivot_table(broker, start_date, end_date)
        
        if pivot_df is not None and not pivot_df.empty:
            # Store current pivot
            key = f"{broker}_{duration}_{end_date.date()}"
            self.current_pivot_data[key] = {
                'broker': broker,
                'duration': duration,
                'start_date': start_date,
                'end_date': end_date,
                'data': pivot_df
            }
            
            # Display pivot
            self.display_pivot_table(pivot_df, broker, duration, start_date, end_date)
        else:
            messagebox.showinfo("No Data", "No data found for the selected criteria")
    
    def calculate_pivot_table(self, broker, start_date, end_date):
        """Calculate pivot table data"""
        # Filter drill down data for the broker
        broker_df = self.drill_down_df[self.drill_down_df['broker'] == broker].copy()
        
        if broker_df.empty:
            return None
        
        # Get unique stocks
        stocks = broker_df['share name'].unique()
        
        pivot_data = []
        
        for stock in stocks:
            stock_df = broker_df[broker_df['share name'] == stock]
            
            # Data at start date (closest date <= start_date)
            start_data = stock_df[stock_df['date'] <= start_date]
            if not start_data.empty:
                start_data = start_data.sort_values('date').iloc[-1]
                qty_start = start_data['quantity']
                value_start = start_data['total market value']
            else:
                qty_start = 0
                value_start = 0
            
            # Data at end date (closest date <= end_date)
            end_data = stock_df[stock_df['date'] <= end_date]
            if not end_data.empty:
                end_data = end_data.sort_values('date').iloc[-1]
                qty_end = end_data['quantity']
                value_end = end_data['total market value']
            else:
                qty_end = 0
                value_end = 0
            
            # Calculate purchase and sell values from sell_purchase_track
            purchase_value = 0
            sell_value = 0
            
            if not self.sell_purchase_df.empty:
                stock_transactions = self.sell_purchase_df[
                    (self.sell_purchase_df['Stock Symbol'] == stock) &
                    (self.sell_purchase_df['Broker'] == broker)
                ]
                
                # Purchases in range
                purchases = stock_transactions[
                    (stock_transactions['Purchase Date'] >= start_date) &
                    (stock_transactions['Purchase Date'] <= end_date) &
                    (stock_transactions['Purchase Price'].notna())
                ]
                purchase_value = (purchases['Purchase Price'] * purchases['Quantity']).sum()
                
                # Sells in range
                sells = stock_transactions[
                    (stock_transactions['Sell Date'] >= start_date) &
                    (stock_transactions['Sell Date'] <= end_date) &
                    (stock_transactions['Sell Price'].notna())
                ]
                sell_value = (sells['Sell Price'] * sells['Quantity']).sum()
            
            # Calculate net change
            net_change = value_end - value_start + sell_value - purchase_value
            
            # Only include if there's any activity
            if qty_start > 0 or qty_end > 0 or purchase_value > 0 or sell_value > 0:
                pivot_data.append({
                    'Stock Name': stock,
                    'Qty (Start)': qty_start,
                    'Value (Start)': round(value_start, 2),
                    'Qty (End)': qty_end,
                    'Value (End)': round(value_end, 2),
                    'Purchase Value': round(purchase_value, 2),
                    'Sell Value': round(sell_value, 2),
                    'Net Change': round(net_change, 2)
                })
        
        return pd.DataFrame(pivot_data)
    
    def display_pivot_table(self, pivot_df, broker, duration, start_date, end_date):
        """Display pivot table in the treeview"""
        # Update title
        self.table_title.config(
            text=f"{broker} - {duration} | {start_date.date()} to {end_date.date()}"
        )
        
        # Clear existing data
        for item in self.table_tree.get_children():
            self.table_tree.delete(item)
        
        # Insert data
        for _, row in pivot_df.iterrows():
            values = (
                row['Stock Name'],
                f"{row['Qty (Start)']:.2f}",
                f"₹{row['Value (Start)']:,.2f}",
                f"{row['Qty (End)']:.2f}",
                f"₹{row['Value (End)']:,.2f}",
                f"₹{row['Purchase Value']:,.2f}",
                f"₹{row['Sell Value']:,.2f}",
                f"₹{row['Net Change']:,.2f}"
            )
            self.table_tree.insert('', 'end', values=values)
        
        # Add total row
        totals = (
            'TOTAL',
            '',
            f"₹{pivot_df['Value (Start)'].sum():,.2f}",
            '',
            f"₹{pivot_df['Value (End)'].sum():,.2f}",
            f"₹{pivot_df['Purchase Value'].sum():,.2f}",
            f"₹{pivot_df['Sell Value'].sum():,.2f}",
            f"₹{pivot_df['Net Change'].sum():,.2f}"
        )
        total_item = self.table_tree.insert('', 'end', values=totals, tags=('total',))
        self.table_tree.tag_configure('total', background='lightgray', font=('Arial', 9, 'bold'))
    
    def download_current_table(self):
        """Download currently displayed pivot table"""
        if not self.current_pivot_data:
            messagebox.showwarning("Warning", "No pivot table generated yet")
            return
        
        # Get the last generated pivot
        last_key = list(self.current_pivot_data.keys())[-1]
        pivot_info = self.current_pivot_data[last_key]
        
        # Create output directory
        output_dir = './pivot_tables'
        os.makedirs(output_dir, exist_ok=True)
        
        # Create filename
        filename = f"pivot_{pivot_info['broker']}_{pivot_info['duration']}_{pivot_info['end_date'].date()}.csv"
        filepath = os.path.join(output_dir, filename)
        
        # Save
        pivot_info['data'].to_csv(filepath, index=False)
        
        messagebox.showinfo("Success", f"Pivot table saved to:\n{filepath}")
    
    def download_all_pivot_tables(self):
        """Generate and download pivot tables for all brokers and durations"""
        end_date = datetime.strptime(self.end_date_entry.get(), '%Y-%m-%d')
        
        output_dir = './pivot_tables'
        os.makedirs(output_dir, exist_ok=True)
        
        durations = ['1M', '3M', '6M', '9M', '12M']
        duration_map = {'1M': 1, '3M': 3, '6M': 6, '9M': 9, '12M': 12}
        
        total_files = len(self.all_brokers) * len(durations)
        count = 0
        
        # Create progress window
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Generating Pivot Tables")
        progress_window.geometry("400x100")
        
        progress_label = ttk.Label(progress_window, text="Generating pivot tables...")
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, length=350, mode='determinate', maximum=total_files)
        progress_bar.pack(pady=10)
        
        for broker in self.all_brokers:
            for duration in durations:
                months = duration_map[duration]
                start_date = end_date - timedelta(days=months * 30)
                
                pivot_df = self.calculate_pivot_table(broker, start_date, end_date)
                
                if pivot_df is not None and not pivot_df.empty:
                    filename = f"pivot_{broker}_{duration}_{end_date.date()}.csv"
                    filepath = os.path.join(output_dir, filename)
                    pivot_df.to_csv(filepath, index=False)
                
                count += 1
                progress_bar['value'] = count
                progress_label.config(text=f"Generating: {broker} - {duration} ({count}/{total_files})")
                progress_window.update()
        
        progress_window.destroy()
        messagebox.showinfo("Success", f"All pivot tables saved to:\n{output_dir}/")
    
    def generate_drill_down_csv(self):
        """Generate filtered drill down CSV (original functionality)"""
        # Create date selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Generate Drill Down CSV")
        dialog.geometry("400x250")
        
        ttk.Label(dialog, text="Select Date Range for Drill Down", font=('Arial', 12, 'bold')).pack(pady=10)
        
        ttk.Label(dialog, text="Start Date:").pack(pady=5)
        start_date_entry = DateEntry(
            dialog,
            width=20,
            background='darkblue',
            foreground='white',
            date_pattern='yyyy-MM-dd'
        )
        start_date_entry.pack(pady=5)
        
        ttk.Label(dialog, text="End Date:").pack(pady=5)
        end_date_entry = DateEntry(
            dialog,
            width=20,
            background='darkblue',
            foreground='white',
            date_pattern='yyyy-MM-dd'
        )
        end_date_entry.pack(pady=5)
        
        def on_generate():
            start_date = datetime.strptime(start_date_entry.get(), '%Y-%m-%d')
            end_date = datetime.strptime(end_date_entry.get(), '%Y-%m-%d')
            
            if start_date > end_date:
                messagebox.showerror("Error", "Start date must be before end date")
                return
            
            # Filter data
            df_range = self.drill_down_df[
                (self.drill_down_df['date'] >= start_date) &
                (self.drill_down_df['date'] <= end_date)
            ]
            filtered_df = df_range[
                df_range['date'].isin([df_range['date'].min(), df_range['date'].max()])
            ]
            
            if filtered_df.empty:
                messagebox.showinfo("No Data", "No data found in the given date range")
                return
            
            # Save
            output_dir = './drill_downs'
            os.makedirs(output_dir, exist_ok=True)
            
            output_file = os.path.join(
                output_dir,
                f'drill_down_{start_date.date()}_to_{end_date.date()}.csv'
            )
            filtered_df.to_csv(output_file, index=False)
            
            messagebox.showinfo("Success", f"Drill down CSV saved to:\n{output_file}")
            dialog.destroy()
        
        ttk.Button(dialog, text="Generate", command=on_generate).pack(pady=20)


def main():
    root = tk.Tk()
    app = DrillDownPivotApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()