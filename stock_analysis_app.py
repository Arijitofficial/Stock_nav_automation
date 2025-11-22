"""
Stock Analysis Application - Refactored for Performance and Maintainability
"""
from datetime import timedelta, date
from typing import Dict, Optional, List, Tuple
from dataclasses import dataclass
import json
import os
import sys
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dateutil.parser import parse
import threading
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class AppConfig:
    """Application configuration"""
    config_file: str = "defaults.json"
    main_file_path: str = ""
    sales_purchase_file_path: str = ""
    sheet_name: str = ""
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    
    @classmethod
    def load(cls, config_file: str = "defaults.json") -> 'AppConfig':
        """Load configuration from file"""
        if os.path.exists(config_file):
            try:
                with open(config_file, "r") as f:
                    data = json.load(f)
                    return cls(
                        config_file=config_file,
                        main_file_path=data.get("main_file_path", ""),
                        sales_purchase_file_path=data.get("sales_purchase_file_path", ""),
                        sheet_name=data.get("sheet_name", "")
                    )
            except Exception as e:
                logger.error(f"Failed to load config: {e}")
        return cls(config_file=config_file)
    
    def save(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, "w") as f:
                json.dump({
                    "main_file_path": self.main_file_path,
                    "sales_purchase_file_path": self.sales_purchase_file_path,
                    "sheet_name": self.sheet_name
                }, f, indent=2)
        except Exception as e:
            logger.error(f"Failed to save config: {e}")


class DataLoader:
    """Handles loading and preprocessing of data"""
    
    @staticmethod
    def load_main_dataframe(file_path: str, sheet_name: str) -> pd.DataFrame:
        """Load and preprocess main dataframe"""
        try:
            df = pd.read_excel(file_path, sheet_name, engine='openpyxl')
            
            # Filter by category
            df = df[df["Cat"].str.lower().isin(['normal', 'others'])]
            
            # Fill NSE Name forward and backward within groups
            df["NSE Name "] = df.groupby("Name of Shares")["NSE Name "].transform(
                lambda x: x.fillna(method='ffill').fillna(method='bfill')
            )
            
            # Handle broker names
            
            df["Broker"].fillna("unknown", inplace=True)
            df["Broker"] = df["Broker"].str.replace(r"(?i)\bone\b", "IIFL", regex=True)
            
            # Convert dates
            df['DOP'] = pd.to_datetime(df['DOP'], errors='coerce')
            df['S. Date'] = pd.to_datetime(df['S. Date'], errors='coerce')
            
            logger.info(f"Loaded {len(df)} records from {sheet_name}")
            return df
            
        except Exception as e:
            logger.error(f"Failed to load dataframe: {e}")
            raise
    
    @staticmethod
    def get_unique_brokers(df: pd.DataFrame) -> List[str]:
        logger.info(f"Loaded unique broker records")
        """Extract unique broker names"""
        return df["Broker"].unique().tolist()
    


class PriceDataManager:
    """Manages stock price data fetching and caching"""
    
    def __init__(self, start_date: date, end_date: date):
        self.start_date = start_date
        self.end_date = end_date
        self.price_df = None
    
    def prepare_ticker_mapping(self, df: pd.DataFrame) -> Tuple[List[str], pd.DataFrame]:
        """Prepare ticker symbols and mapping dataframe"""
        filtered_df = df[
            df['S. Date'].isna() | (df['S. Date'].dt.date > self.start_date)
        ] if self.start_date else df
        
        unique_df = filtered_df[['NSE Name ', 'Symbol', 'Name of Shares']].drop_duplicates()
        
        # Create ticker symbols
        unique_df['ticker'] = unique_df.apply(
            lambda row: f"{row['NSE Name ']}.NS" if row['Symbol'] == "NSE" 
                       else f"{row['NSE Name ']}.BO",
            axis=1
        )
        
        return list(set(unique_df['ticker'])), unique_df
    
    def fetch_prices(self, tickers: List[str]) -> pd.DataFrame:
        """Fetch price data for given tickers"""
        from Utils.symbol_change_handler import map_symbols
        from Utils.down_close_price_data import create_stock_price_df
        
        logger.info(f"Fetching prices for {len(tickers)} tickers")
        
        symbols_dict = map_symbols(
            tickers, 
            start_date=str(self.start_date), 
            end_date=str(self.end_date)
        )
        
        self.price_df = create_stock_price_df(
            start_date=str(self.start_date),
            end_date=str(self.end_date),
            keys=tickers,
            symbols_dict=symbols_dict
        )
        
        # Cache to disk
        self.price_df.to_csv("close_prices_dataframe.csv", index=False)
        logger.info("Price data fetched and cached")
        
        return self.price_df
    
    def get_price(self, symbol: str, date_str: str) -> Optional[float]:
        """Get closing price for a symbol on a specific date"""
        if self.price_df is None:
            return None
        
        try:
            if symbol in self.price_df['ticker'].values:
                if date_str in self.price_df.columns:
                    price = self.price_df.loc[
                        self.price_df['ticker'] == symbol, date_str
                    ].values[0]
                    
                    return None if pd.isna(price) else float(price)
        except Exception as e:
            logger.warning(f"Error getting price for {symbol} on {date_str}: {e}")
        
        return None


class PortfolioCalculator:
    """Handles portfolio calculations and NAV computation"""
    
    def __init__(self, brokers: List[str], cfca_handler):
        self.brokers = ["Overall"] + list(brokers)
        self.cfca_handler = cfca_handler
        
    def initialize_metrics(self) -> Dict[str, Dict[str, float]]:
        """Initialize broker-wise metrics"""
        return {
            'values': {broker: 0.0 for broker in self.brokers},
            'sales': {broker: 0.0 for broker in self.brokers},
            'purchases': {broker: 0.0 for broker in self.brokers},
            'net_fund': {broker: 0.0 for broker in self.brokers},
            'units': {broker: 0.0 for broker in self.brokers},
            'nav': {broker: 0.0 for broker in self.brokers}
        }
    
    def calculate_holding_value(
        self, 
        row: pd.Series, 
        closing_price: float,
        current_date: date
    ) -> Tuple[float, bool]:
        """Calculate value for a single holding"""
        market_value = closing_price * row['No. ']
        is_sold = (pd.notna(row['S. Date']) and 
                   row['S. Date'].date() == current_date)
        
        return market_value, is_sold
    
    def update_nav(
        self,
        broker: str,
        current_value: float,
        purchase: float,
        sale: float,
        prev_unit: float,
        prev_nav: float,
        prev_value: float
    ) -> Tuple[float, float]:
        """Calculate NAV and units for a broker"""
        net_fund = purchase - sale
        
        # Handle zero previous NAV
        if prev_nav == 0:
            prev_nav = 1000.0
        
        # Calculate units
        if prev_value != 0:
            units = prev_unit + (net_fund / prev_nav)
        else:
            units = current_value / prev_nav
        
        # Calculate NAV
        nav = current_value / units if units != 0 else 0.0
        
        # Handle zero value case
        if current_value == 0:
            units = 0.0
            nav = 0.0
        
        return units, nav


class PortfolioProcessor:
    """Main processor for portfolio analysis"""
    
    def __init__(
        self,
        df: pd.DataFrame,
        config: AppConfig,
        brokers: List[str],
        price_manager: PriceDataManager,
        cfca_handler
    ):
        self.df = df
        self.config = config
        self.brokers = brokers
        self.price_manager = price_manager
        self.cfca_handler = cfca_handler
        self.calculator = PortfolioCalculator(brokers, cfca_handler)
        
        # Initialize data structures
        self.sales_purchase_dict = self._init_sales_purchase_dict()
        self.drill_down_df = self._init_drill_down_df()
    
    def _init_sales_purchase_dict(self) -> Dict[str, pd.DataFrame]:
        """Initialize sales/purchase tracking dictionary"""
        from Utils.sales_purchase_util import init_dict
        return init_dict(
            file_path=self.config.sales_purchase_file_path,
            broker_names=self.brokers
        )
    
    def _init_drill_down_df(self) -> pd.DataFrame:
        """Initialize drill-down dataframe"""
        from Utils.drill_down_util import init_drill_down_df
        return init_drill_down_df()
    
    def process(self, progress_callback=None) -> bool:
        """Process portfolio data day by day"""
        try:
            total_days = (self.config.end_date - self.config.start_date).days + 1
            current_date = self.config.start_date
            
            # Work on a copy
            temp_df = self.df.copy()
            original_volume = temp_df['No. '].copy()
            
            # Reverse corporate actions to start state
            temp_df = self.cfca_handler.reverse_actions(
                df=temp_df, 
                start_date=str(current_date)
            )
            
            days_processed = 0
            update_frequency = 5
            
            while current_date <= self.config.end_date:
                # Apply corporate actions for current day
                temp_df = self.cfca_handler.apply_tday_actions(
                    temp_df, 
                    current_date=str(current_date)
                )
                
                # Process day
                success = self._process_single_day(temp_df, current_date)
                
                if not success:
                    logger.warning(f"Skipping day {current_date} due to missing data")
                
                # Update progress
                days_processed += 1
                if progress_callback and (days_processed % update_frequency == 0 or 
                                         current_date == self.config.end_date):
                    progress_callback(days_processed, total_days)
                
                current_date += timedelta(days=1)
            
            # Restore original volumes and save
            temp_df["No. "] = original_volume
            self._save_results(temp_df)
            
            return True
            
        except Exception as e:
            logger.error(f"Processing failed: {e}", exc_info=True)
            return False
    
    def _process_single_day(self, df: pd.DataFrame, current_date: date) -> bool:
        """Process portfolio for a single day"""
        metrics = self.calculator.initialize_metrics()
        holdings_processed = 0
        
        current_date_dt = pd.Timestamp(current_date)
        
        for index, row in df.iterrows():
            # Skip if not owned on this date
            if row['DOP'] > current_date_dt:
                continue
            
            if pd.notna(row['S. Date']) and row['S. Date'] < current_date_dt:
                continue
            
            # Get closing price
            symbol = row['NSE Name ']
            if pd.isna(symbol):
                continue
            
            symbol_ns = symbol + (".NS" if row['Symbol'] == "NSE" else ".BO")
            date_str = current_date.strftime("%Y-%m-%d")
            
            closing_price = self.price_manager.get_price(symbol_ns, date_str)
            
            if closing_price is None:
                # Use cost price as fallback
                closing_price = row['Cost/Sh']
            
            # Calculate values
            market_value, is_sold = self.calculator.calculate_holding_value(
                row, closing_price, current_date
            )
            
            broker = row['Broker']
            
            # Update metrics
            if is_sold:
                metrics['sales']['Overall'] += row['Net Sale']
                metrics['sales'][broker] += row['Net Sale']
            else:
                metrics['values']['Overall'] += market_value
                metrics['values'][broker] += market_value
            
            # Track purchases on DOP
            if row['DOP'].date() == current_date:
                metrics['purchases']['Overall'] += row['Net Cost']
                metrics['purchases'][broker] += row['Net Cost']
            
            # Update drill-down
            self._update_drill_down(
                current_date, symbol, broker, row, closing_price, market_value
            )
            
            holdings_processed += 1
        
        # Update sales/purchase tracking
        self._update_sales_purchase_tracking(current_date, metrics)
        
        return holdings_processed > 0
    
    def _update_drill_down(
        self, 
        current_date: date, 
        symbol: str, 
        broker: str, 
        row: pd.Series,
        closing_price: float,
        market_value: float
    ):
        """Update drill-down tracking"""
        from Utils.drill_down_util import enter_track
        
        self.drill_down_df = enter_track(
            self.drill_down_df,
            current_date,
            symbol,
            broker,
            row['File'],
            row['No. '],
            row['Cost/Sh'],
            closing_price,
            market_value
        )
    
    def _update_sales_purchase_tracking(self, current_date: date, metrics: Dict):
        """Update sales/purchase tracking for all brokers"""
        for broker in self.calculator.brokers:
            # Skip if no activity and no holdings
            has_activity = (metrics['sales'][broker] != 0 or 
                          metrics['purchases'][broker] != 0)
            
            if not has_activity and metrics['values'][broker] == 0:
                continue
            
            # Get previous values
            prev_data = self._get_previous_broker_data(broker)
            
            # Calculate NAV
            units, nav = self.calculator.update_nav(
                broker,
                metrics['values'][broker],
                metrics['purchases'][broker],
                metrics['sales'][broker],
                prev_data['units'],
                prev_data['nav'],
                prev_data['value']
            )
            
            # Append new row
            new_row = {
                'Date': current_date,
                'Value': metrics['values'][broker],
                'Purchase': metrics['purchases'][broker],
                'Sales': metrics['sales'][broker],
                'Net Fund': metrics['purchases'][broker] - metrics['sales'][broker],
                'Units': units,
                'NAV': nav
            }
            
            self.sales_purchase_dict[broker] = pd.concat([
                self.sales_purchase_dict[broker],
                pd.DataFrame([new_row])
            ], ignore_index=True)
    
    def _get_previous_broker_data(self, broker: str) -> Dict[str, float]:
        """Get previous day's data for a broker"""
        if self.sales_purchase_dict[broker].empty:
            return {'units': 0.0, 'nav': 1000.0, 'value': 0.0}
        
        last_row = self.sales_purchase_dict[broker].iloc[-1]
        return {
            'units': last_row['Units'],
            'nav': last_row['NAV'],
            'value': last_row['Value']
        }
    
    def _save_results(self, df: pd.DataFrame):
        """Save all results to disk"""
        from Utils.sales_purchase_util import save_sales_purchase_dict
        from Utils.drill_down_util import save_drill_down_df
        
        try:
            # Save updated main file
            with pd.ExcelWriter(
                self.config.main_file_path, 
                engine='openpyxl', 
                mode='a', 
                if_sheet_exists='replace'
            ) as writer:
                df.to_excel(writer, sheet_name=self.config.sheet_name, index=False)
            
            # Save tracking data
            save_sales_purchase_dict(self.sales_purchase_dict)
            save_drill_down_df(self.drill_down_df)
            
            logger.info("Results saved successfully")
            
        except Exception as e:
            logger.error(f"Failed to save results: {e}")
            raise


class StockAnalysisApp:
    """Main application UI"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Stock Analysis Tool")
        self.root.geometry("700x550")
        
        # Load configuration
        self.config = AppConfig.load()
        
        # Initialize components
        self.price_manager = None
        self.processor = None
        self.cfca_handler = None
        
        # Start UI
        self.show_file_input_screen()
    
    def clear_screen(self):
        """Clear all widgets from screen"""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def show_file_input_screen(self):
        """Show file selection screen"""
        self.clear_screen()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Main file
        ttk.Label(frame, text="Main Excel File:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, pady=10, padx=10, sticky='w'
        )
        self.main_file_label = ttk.Label(
            frame, 
            text=self.config.main_file_path or "No file selected",
            width=50,
            relief="sunken"
        )
        self.main_file_label.grid(row=0, column=1, pady=10, padx=10)
        ttk.Button(
            frame, 
            text="Browse", 
            command=self.browse_main_file
        ).grid(row=0, column=2, pady=10, padx=10)
        
        # Sales/Purchase file (optional)
        ttk.Label(frame, text="Sales/Purchase File:", font=('Arial', 10)).grid(
            row=1, column=0, pady=10, padx=10, sticky='w'
        )
        self.sp_file_label = ttk.Label(
            frame,
            text=self.config.sales_purchase_file_path or "No file selected",
            width=50,
            relief="sunken"
        )
        self.sp_file_label.grid(row=1, column=1, pady=10, padx=10)
        ttk.Button(
            frame,
            text="Browse",
            command=self.browse_sp_file
        ).grid(row=1, column=2, pady=10, padx=10)
        
        # Sheet name
        ttk.Label(frame, text="Sheet Name:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, pady=10, padx=10, sticky='w'
        )
        self.sheet_name_var = tk.StringVar(value=self.config.sheet_name)
        ttk.Entry(frame, textvariable=self.sheet_name_var, width=30).grid(
            row=2, column=1, pady=10, padx=10, sticky='w'
        )
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20, sticky='e')
        
        ttk.Button(
            button_frame,
            text="Clear",
            command=self.clear_config
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame,
            text="Next",
            command=self.show_date_input_screen
        ).pack(side=tk.RIGHT, padx=5)
    
    def show_date_input_screen(self):
        """Show date input screen"""
        # Validate and save config
        self.config.main_file_path = self.main_file_label.cget("text")
        self.config.sales_purchase_file_path = self.sp_file_label.cget("text")
        self.config.sheet_name = self.sheet_name_var.get()
        
        if not self.config.main_file_path or self.config.main_file_path == "No file selected":
            messagebox.showerror("Error", "Please select the main Excel file")
            return
        
        if not self.config.sheet_name:
            messagebox.showerror("Error", "Please enter a sheet name")
            return
        
        # Load data
        if not self.load_data():
            return
        
        self.config.save()
        self.clear_screen()
        
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(expand=True)
        
        # Start date
        ttk.Label(
            frame, 
            text="Start Date (optional):", 
            font=('Arial', 10)
        ).pack(pady=10)
        
        self.start_date_var = tk.StringVar()
        
        # Pre-fill if available
        if (self.processor.sales_purchase_dict and 
            "Overall" in self.processor.sales_purchase_dict and
            not self.processor.sales_purchase_dict["Overall"].empty):
            last_date = self.processor.sales_purchase_dict["Overall"]["Date"].iloc[-1]
            self.start_date_var.set(str(last_date))
        
        ttk.Entry(frame, textvariable=self.start_date_var, width=25).pack(pady=5)
        
        # End date
        ttk.Label(
            frame,
            text="End Date (required):",
            font=('Arial', 10, 'bold')
        ).pack(pady=10)
        
        self.end_date_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.end_date_var, width=25).pack(pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=20)
        
        ttk.Button(
            button_frame,
            text="Previous",
            command=self.show_file_input_screen
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(
            button_frame,
            text="Start Processing",
            command=self.start_processing
        ).pack(side=tk.LEFT, padx=10)
    
    def browse_main_file(self):
        """Browse for main Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Main Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.main_file_label.config(text=file_path)
    
    def browse_sp_file(self):
        """Browse for sales/purchase file"""
        file_path = filedialog.askopenfilename(
            title="Select Sales/Purchase File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.sp_file_label.config(text=file_path)
    
    def clear_config(self):
        """Clear all configuration"""
        self.config = AppConfig()
        self.config.save()
        self.show_file_input_screen()
    
    def load_data(self) -> bool:
        """Load and initialize data"""
        try:
            from Utils.corporate_actions_handler import CorporateActionsHandler
            
            # Load main dataframe
            df = DataLoader.load_main_dataframe(
                self.config.main_file_path,
                self.config.sheet_name
            )
            
            # Get brokers
            brokers = DataLoader.get_unique_brokers(df)
            logger.info(f"Loaded cfca brokers")
            # Initialize corporate actions handler
            self.cfca_handler = CorporateActionsHandler("Excels")
            logger.info(f"Loaded cfca handler")

            
            # Initialize processor (will initialize dicts)
            self.processor = PortfolioProcessor(
                df=df,
                config=self.config,
                brokers=brokers,
                price_manager=None,  # Will be set later
                cfca_handler=self.cfca_handler
            )
            logger.info(f"Loaded processor handler")
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to load data: {e}")
            messagebox.showerror("Error", f"Failed to load data:\n{str(e)}")
            return False
    
    def start_processing(self):
        """Start the processing workflow"""
        # Parse dates
        try:
            start_date_str = self.start_date_var.get()
            end_date_str = self.end_date_var.get()
            
            if not end_date_str:
                messagebox.showerror("Error", "End date is required")
                return
            
            self.config.start_date = parse(start_date_str).date() if start_date_str else None
            self.config.end_date = parse(end_date_str).date()
            
        except Exception as e:
            messagebox.showerror("Error", f"Invalid date format: {e}")
            return
        
        # Start threaded processing
        threading.Thread(target=self.run_processing, daemon=True).start()
    
    def run_processing(self):
        """Run the complete processing pipeline"""
        try:
            # Show loading popup
            self.show_loading_popup("Fetching price data...")
            
            # Initialize price manager
            self.price_manager = PriceDataManager(
                self.config.start_date or self.config.end_date,
                self.config.end_date
            )
            
            # Prepare tickers
            tickers, _ = self.price_manager.prepare_ticker_mapping(self.processor.df)
            
            # Fetch prices
            self.price_manager.fetch_prices(tickers)
            
            # Close loading popup
            self.loading_popup.destroy()
            
            # Update processor with price manager
            self.processor.price_manager = self.price_manager
            
            # Show progress window
            self.show_progress_window()
            
            # Process data
            success = self.processor.process(
                progress_callback=self.update_progress
            )
            
            # Close progress window
            self.progress_window.destroy()
            
            if success:
                messagebox.showinfo("Success", "Processing completed successfully!")
                self.root.after(2000, self.terminate_app)
            else:
                messagebox.showerror("Error", "Processing failed. Check logs.")
                
        except Exception as e:
            logger.error(f"Processing error: {e}", exc_info=True)
            if hasattr(self, 'loading_popup'):
                self.loading_popup.destroy()
            if hasattr(self, 'progress_window'):
                self.progress_window.destroy()
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
    
    def show_loading_popup(self, message: str):
        """Show loading popup"""
        self.loading_popup = tk.Toplevel(self.root)
        self.loading_popup.title("Loading")
        self.loading_popup.geometry("300x100")
        self.loading_popup.resizable(False, False)
        
        tk.Label(
            self.loading_popup,
            text=message,
            font=("Arial", 11)
        ).pack(pady=30)
        
        self.loading_popup.grab_set()
        self.root.update_idletasks()
    
    def show_progress_window(self):
        """Show progress window"""
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Processing")
        self.progress_window.geometry("400x120")
        self.progress_window.resizable(False, False)
        
        ttk.Label(
            self.progress_window,
            text="Processing portfolio data...",
            font=("Arial", 11)
        ).pack(pady=15)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_window,
            orient="horizontal",
            length=350,
            mode="determinate"
        )
        self.progress_bar.pack(pady=10)
        
        self.progress_label = ttk.Label(
            self.progress_window,
            text="0%",
            font=("Arial", 9)
        )
        self.progress_label.pack()
        
        self.progress_window.grab_set()
    
    def update_progress(self, current: int, total: int):
        """Update progress bar"""
        if hasattr(self, 'progress_bar'):
            percentage = int((current / total) * 100)
            self.progress_bar["maximum"] = total
            self.progress_bar["value"] = current
            self.progress_label.config(text=f"{percentage}% ({current}/{total} days)")
            self.progress_window.update_idletasks()
    
    def terminate_app(self):
        """Terminate the application"""
        self.root.destroy()
        sys.exit(0)


def main():
    """Application entry point"""
    root = tk.Tk()
    app = StockAnalysisApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()