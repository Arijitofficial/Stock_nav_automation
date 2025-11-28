# Stock Analysis & Drill Down Automation

This project provides a desktop application for analyzing stock portfolios, tracking sales/purchases, and generating drill-down reports. It is built with Python, Tkinter, and Pandas, and is designed for use with Indian stock market data.

## Features

- Import and process Excel files with portfolio data
- Track sales and purchases by broker
- Download and merge historical close price data
- Handle symbol changes automatically
- Generate drill-down CSVs for custom date ranges
- User-friendly GUI for all operations

## Requirements

- Python 3.10+
- See [`requirements.txt`](requirements.txt) for all dependencies

## Installation

1. Clone this repository.
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

To run the main stock analysis app:
```
python stock_analysis_app.py
```

To run the drill down CSV generator:
```
python drill_down_app.py
```

## Building a Desktop App

You can create a standalone Windows desktop app using PyInstaller.  
**Run this command from the project root:**

```
pyinstaller --noconsole --windowed --onefile --name "DrillDownApp" --add-data "Utils;Utils" --distpath Portfolio_analyzers drill_down_app.py

pyinstaller --noconsole --windowed --onefile --name "PivotAnalysisApp" --add-data "Utils;Utils" --distpath Portfolio_analyzers pivot_analysis_app.py

pyinstaller --noconsole --windowed --onefile --name "StockAnalysisApp" --add-data "Utils;Utils" --distpath Portfolio_analyzers stock_analysis_app.py
```

*On Windows Command Prompt, use `^` for line continuation. On PowerShell or Linux, use `\`.*

The executable will be created in the `app` directory.

## Directory Structure

- `Utils/` — Utility modules for symbol changes, drill down, and sales/purchase tracking
- `down_close_price_data.py` — Module for fetching close price data
- `stock_analysis_app.py` — Main GUI application
- `drill_down_app.py` — Drill down CSV generator GUI
- `Excels/` — (Not tracked) Place your Excel/CSV data files here

## Notes

- All `.zip` files, files in `bhavcopies/`, and most Excel/CSV outputs are ignored by git (see [.gitignore](.gitignore)).
- Make sure your data files are in the correct format as expected by the app.

## License

MIT License (add your license details