import pandas as pd
from datetime import datetime, timedelta
import os

def create_drill_down_csv():
    """
    Reads the master track CSV if it exists, takes user input for pivot date and duration,
    filters data for the given duration, and saves the filtered data as a new CSV file.
    """
    # Check if the master track CSV exists
    if not os.path.exists('./Excels/drill_down_track.csv'):
        print("Error: './Excels/drill_down_track.csv' not found. Please ensure the file exists in the current directory.")
        return

    # Read the master track CSV
    drill_down_df = pd.read_csv('./Excels/drill_down_track.csv')

    # Ensure 'date' column is in datetime format
    drill_down_df['date'] = pd.to_datetime(drill_down_df['date'])

    # Take user input for pivot date and duration
    pivot_date = input("Enter the pivot date (YYYY-MM-DD): ")
    duration = input("Enter the duration (1m, 3m, or 6m): ")

    # Convert pivot_date to a datetime object
    try:
        pivot_date = datetime.strptime(pivot_date, '%Y-%m-%d')
    except ValueError:
        print("Error: Invalid date format. Please use 'YYYY-MM-DD'.")
        return

    # Determine the start date based on the duration
    if duration.lower() == '1d':
        start_date = pivot_date - timedelta(days=1)
    elif duration.lower() == '7d':
        start_date = pivot_date - timedelta(days=6)
    elif duration.lower() == '1m':
        start_date = pivot_date - timedelta(days=29)
    elif duration.lower() == '3m':
        start_date = pivot_date - timedelta(days=89)
    elif duration.lower() == '6m':
        start_date = pivot_date - timedelta(days=179)
    elif duration.lower() == '9m':
        start_date = pivot_date - timedelta(days=269)
    elif duration.lower() == '12m':
        start_date = pivot_date - timedelta(days=359)
    else:
        print("Error: Invalid duration. Please use '1m', '3m', or '6m'.")
        return
    
    start_date = drill_down_df[drill_down_df['date'] < start_date]['date'].max()

    # Filter the DataFrame for the given time period
    filtered_df = drill_down_df[(drill_down_df['date'] == start_date) | (drill_down_df['date'] == pivot_date)]

    # Check if the filtered DataFrame is empty
    if filtered_df.empty:
        print(f"No data found for the specified period ({duration}) up to {pivot_date.date()}.")
        return

    # Save the filtered DataFrame as a new CSV file
    output_filename = f'drill_down_{duration}_{str(pivot_date)[:10]}.csv'

    filtered_df.to_csv(output_filename, index=False)

    print(f"CSV file '{output_filename}' has been created with data for the past {duration} from {pivot_date.date()}.")


create_drill_down_csv()