# ============================================================================
# 1. CREATE NEW FILE: Utils/sell_purchase_track_util.py
# ============================================================================

"""
Utility functions for tracking sell and purchase transactions
"""
import pandas as pd
import os
import logging

logger = logging.getLogger(__name__)


def check_and_create_sell_purchase_track_df():
    """Create initialized sell_purchase_track DataFrame"""
    return pd.DataFrame(columns=[
        'Purchase Date',
        'Sell Date', 
        'Broker',
        'File',
        'Stock Symbol',
        'Purchase Price',
        'Sell Price',
        'Quantity'
    ])


def init_sell_purchase_track_df(filename="Excels/sell_purchase_track.csv"):
    """Initialize sell_purchase_track dataframe from file or create new"""
    if os.path.exists(filename):
        return pd.read_csv(filename)
    else:
        return check_and_create_sell_purchase_track_df()


def enter_purchase_track(
    sell_purchase_track_df,
    purchase_date,
    broker,
    file,
    stock_symbol,
    purchase_price,
    quantity
):
    """
    Record a purchase transaction
    
    Args:
        sell_purchase_track_df: sell_purchase_track dataframe
        purchase_date: Date of purchase
        broker: Broker name
        file: File reference
        stock_symbol: Stock symbol
        purchase_price: Purchase price per share
        quantity: Number of shares
        
    Returns:
        Updated dataframe
    """
    # Initialize if None
    if sell_purchase_track_df is None:
        sell_purchase_track_df = check_and_create_sell_purchase_track_df()
    
    # Create new purchase row
    new_row = pd.DataFrame({
        'Purchase Date': [purchase_date],
        'Sell Date': [None],
        'Broker': [broker],
        'File': [file],
        'Stock Symbol': [stock_symbol],
        'Purchase Price': [purchase_price],
        'Sell Price': [None],
        'Quantity': [quantity]
    })
    
    sell_purchase_track_df = pd.concat([sell_purchase_track_df, new_row], ignore_index=True)
    
    return sell_purchase_track_df


def enter_sell_track(
    sell_purchase_track_df,
    purchase_date,
    sell_date,
    broker,
    file,
    stock_symbol,
    sell_price,
    quantity
):
    """
    Record a sell transaction
    
    Args:
        sell_purchase_track_df: sell_purchase_track dataframe
        purchase_date: Original purchase date
        sell_date: Date of sale
        broker: Broker name
        file: File reference
        stock_symbol: Stock symbol
        sell_price: Sell price per share
        quantity: Number of shares sold
        
    Returns:
        Updated dataframe
    """
    # Initialize if None
    if sell_purchase_track_df is None:
        sell_purchase_track_df = check_and_create_sell_purchase_track_df()
    
    # Create new sell row
    new_row = pd.DataFrame({
        'Purchase Date': [purchase_date],
        'Sell Date': [sell_date],
        'Broker': [broker],
        'File': [file],
        'Stock Symbol': [stock_symbol],
        'Purchase Price': [None],
        'Sell Price': [sell_price],
        'Quantity': [quantity]
    })
    
    sell_purchase_track_df = pd.concat([sell_purchase_track_df, new_row], ignore_index=True)
    
    return sell_purchase_track_df


def save_sell_purchase_track_df(sell_purchase_track_df, filename="Excels/sell_purchase_track.csv"):
    """Save sell_purchase_track dataframe to CSV"""
    try:
        # Sort by Purchase Date, then Sell Date
        sell_purchase_track_df = sell_purchase_track_df.sort_values(
            by=["Purchase Date", "Sell Date"],
            key=lambda col: pd.to_datetime(col, errors="coerce", dayfirst=True)
        )
        
        sell_purchase_track_df.to_csv(filename, index=False)
        logger.info(f"Sell/Purchase track saved to {filename}")
    except Exception as e:
        logger.error(f"Failed to save sell_purchase_track: {e}")
        raise