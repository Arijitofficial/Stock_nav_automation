import re
import os
from typing import Optional

file_initial = "CF-CA"

import re


def get_latest_CFCA_file( directory: Optional[str] = None) -> Optional[str]:
    """
    Find the latest modified file in a directory that starts with the given initial string.
    
    Args:
        initial (str): The prefix the file should start with.
        directory (str, optional): Directory path. Defaults to current directory.
    
    Returns:
        str or None: Path to the latest file, or None if no match is found.
    """
    if directory is None:
        directory = os.getcwd()

    # List files starting with initial
    matching_files = [
        os.path.join(directory, f)
        for f in os.listdir(directory)
        if f.startswith(file_initial) and os.path.isfile(os.path.join(directory, f))
    ]

    if not matching_files:
        return None

    # Get latest by modification time
    latest_file = max(matching_files, key=os.path.getmtime)
    return latest_file


def is_face_value_action(purpose: str) -> bool:
    """
    Returns True if the PURPOSE text indicates a face value
    split or consolidation (handles abbreviations too).
    """
    text = purpose.lower().replace(".", " ")
    
    # Check for full keywords
    if "consolidation" in text or "split" in text:
        return True
    
    # Check for abbreviations like "fv splt"
    if "fv" in text and ("splt" in text or "split" in text):
        return True
    
    return False


def extract_face_values(text: str):
    """
    Extracts the two face values (before and after) from corporate action text.
    Works with multiple formats:
      - "Face Value Split From Rs. 10 To Rs. 2/-" -> (10, 2)
      - "Face Value Split (Sub Division) - From Rs 10/- Per Share To Re 1/- Per Share" -> (10, 1)
      - "Consolidation Of Equity Shares From Re 1 Per Share To Rs 10 Per Share" -> (1, 10)
      - "Fv Splt Frm Rs 10 To Re 1" -> (10, 1)
      - "Fv Splt Frm Rs 10 To Rs 2" -> (10, 2)
      - "Fv Split Rs.10/- To Rs.2/" -> (10, 2)
    """
    # Normalize
    clean_text = text.replace("/-", "").replace("/", "").replace("-", " ")
    clean_text = clean_text.replace(".", " ")
    
    # Pattern 1: with "From ... To ..."
    match = re.search(
        r"(?:from|frm)\s+[^0-9]*([0-9]+)[^0-9]+to\s+[^0-9]*([0-9]+)",
        clean_text,
        flags=re.IGNORECASE
    )
    if match:
        return int(match.group(1)), int(match.group(2))
    
    # Pattern 2: direct "Rs 10 ... To Rs 2"
    match = re.search(
        r"rs?\s*\.?\s*([0-9]+)[^0-9]+to\s+rs?|re\s*\.?\s*([0-9]+)",
        clean_text,
        flags=re.IGNORECASE
    )
    if match:
        nums = [g for g in match.groups() if g]  # filter None
        if len(nums) == 2:
            return int(nums[0]), int(nums[1])
    
    match = re.search(
        r"(?:rs|re)\s*([0-9]+)\s*.*?\s*to\s*(?:rs|re)\s*([0-9]+)",
        clean_text,
        flags=re.IGNORECASE
    )
    if match:
        return int(match.group(1)), int(match.group(2))
    
    return None, None







