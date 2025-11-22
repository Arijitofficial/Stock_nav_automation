import math
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, Tuple, List

import pandas as pd
import numpy as np

class CorporateActionsHandler:
    """Handles corporate actions (splits/consolidations) for accurate volume tracking"""
    
    def __init__(self, cfca_path: str):
        """
        Initialize with CFCA data file
        
        Args:
            cfca_path: Path to the CFCA CSV file
        """
        self.cfca_df = self._load_and_process_cfca(cfca_path)
        
    def _load_and_process_cfca(self, cfca_dir: str) -> pd.DataFrame:
        """Load and process CFCA data"""
        from Utils.split_n_merge_handler import is_face_value_action, extract_face_values, get_latest_CFCA_file
        cfca_path = get_latest_CFCA_file(cfca_dir)
    
        cfca_df = pd.read_csv(cfca_path)
        
        # Filter for face value actions only
        cfca_df = cfca_df[cfca_df["PURPOSE"].apply(lambda x: is_face_value_action(x))]
        
        # Extract from and to values
        cfca_df[["from_value", "to_value"]] = cfca_df["PURPOSE"].apply(
            lambda x: pd.Series(extract_face_values(x))
        )
        
        # Remove rows without valid from_value
        cfca_df = cfca_df[~cfca_df["from_value"].isna()]
        
        # Convert EX-DATE to datetime
        cfca_df['EX-DATE'] = pd.to_datetime(cfca_df['EX-DATE'])
        
        # Calculate adjustment ratio (to_value / from_value)
        cfca_df['price_adjustment_ratio'] = cfca_df['to_value'] / cfca_df['from_value']
        cfca_df['volume_adjustment_ratio'] = cfca_df['from_value'] / cfca_df['to_value']
        
        # Sort by date for chronological processing
        cfca_df = cfca_df.sort_values('EX-DATE')
        
        return cfca_df

    def reverse_actions(self, df: pd.DataFrame, start_date: str) -> pd.DataFrame:
        """
        Reverse corporate actions on volumes to reflect state at given start_date (or DOP if later).
        
        Args:
            df: Portfolio dataframe (with columns ["NSE Name ", 'No.', 'DOP'])
            start_date: date (str or datetime) from which to roll back
        
        Returns:
            DataFrame with adjusted volumes as of start_date
        """
        start_date = pd.to_datetime(start_date)

        # Work on a copy
        df = df.copy()
        
        # Ensure date column present
        df['DOP'] = pd.to_datetime(df['DOP'])

        # For each corporate action after start_date, undo the effect
        for _, action in self.cfca_df.iterrows():
            action_date = action['EX-DATE']
            symbol = action['SYMBOL']
            ratio = action['volume_adjustment_ratio']

            if action_date >= start_date:
                # Affected rows: matching symbol, purchased before action date
                mask = (df["NSE Name "] == symbol) & (df['DOP'] <= action_date)
                if mask.any():
                    # Reverse the adjustment: divide by ratio
                    df.loc[mask, 'No. '] = df.loc[mask, 'No. '].apply(
                        lambda x: self._reverse_volume(x, ratio)
                    )
        return df

    

    def apply_tday_actions(self, df: pd.DataFrame, current_date: str) -> pd.DataFrame:
        """
        Apply ONLY the corporate actions happening on the given current_date.
        
        Args:
            df: Portfolio dataframe (with columns ["NSE Name ", 'No.', 'DOP'])
            current_date: date (str or datetime) for which to apply forward adjustments
        
        Returns:
            DataFrame with adjusted volumes for that date
        """
        current_date = pd.to_datetime(current_date)
        df = df.copy()
        df['DOP'] = pd.to_datetime(df['DOP'])

        # filter actions that happen exactly on current_date
        actions_today = self.cfca_df[self.cfca_df['EX-DATE'] == current_date]

        for _, action in actions_today.iterrows():
            symbol = action['SYMBOL']
            ratio = action['to_value'] / action['from_value']   # forward multiplier

            mask = (df["NSE Name "] == symbol) & (df['DOP'] <= current_date)
            if mask.any():
                df.loc[mask, 'No. '] = df.loc[mask, 'No. '].apply(
                    lambda x: self._apply_forward_volume(x, ratio)
                )
        return df
    
    @staticmethod
    def _reverse_volume(units: float, ratio: float) -> int:
        """
        Reverse a split/consolidation effect while handling fractional units.
        
        Example: 3 units (2→5 split, ratio=2.5) → 7 units (forward)
                 7 units reversed → 3 units (not 2).
        """
        reversed_units = units / ratio
        # Round to nearest int, but ensure we don’t lose rightful units
        return math.ceil(reversed_units)

    @staticmethod
    def _apply_forward_volume(units: float, ratio: float) -> int:
        """Apply corporate action effect (fractions floored)."""
        adjusted_units = units * ratio
        return math.floor(adjusted_units)

