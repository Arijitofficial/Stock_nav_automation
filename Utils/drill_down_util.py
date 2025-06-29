import pandas as pd
import os
def check_and_create_drill_down_track_df():
    # This function should return the initialized drill_down_df DataFrame
    return pd.DataFrame(columns=["date", "share name", "broker", "File", "purchase cost", "quantity", "t_day mkt price", "total market value"])

def enter_track(drill_down_df, date, share_name, broker, file, qty, purchase_cost, current_market_price, current_mv):

    # Initialize drill_down_df if it does not exist
    if drill_down_df is None:
        drill_down_df = check_and_create_drill_down_track_df()

    # Find existing row based on "share name," "broker," and "File"
    existing_row = drill_down_df[
        (drill_down_df["date"] == date) &
        (drill_down_df["share name"] == share_name) &
        (drill_down_df["broker"] == broker) &
        (drill_down_df["File"] == file)
    ]

    if not existing_row.empty:
        # Calculate weighted average purchase cost
        old_qty = existing_row["quantity"].iloc[0]
        old_purchase_cost = existing_row["purchase cost"].iloc[0]
        new_avg_purchase_cost = ((old_qty * old_purchase_cost) + (qty * purchase_cost)) / (old_qty + qty)
        new_market_val = existing_row["total market value"].iloc[0] + current_mv
        new_cmp = current_market_price if current_market_price == existing_row["t_day mkt price"].iloc[0] else new_avg_purchase_cost

        # Update the row with new values
        drill_down_df.loc[
            (drill_down_df["date"] == date) &
            (drill_down_df["share name"] == share_name) &
            (drill_down_df["broker"] == broker) &
            (drill_down_df["File"] == file),
            ["quantity", "purchase cost", "t_day mkt price", "total market value"]
        ] = [old_qty + qty, new_avg_purchase_cost, new_cmp, new_market_val]
    else:
        # Append a new row
        new_row = pd.DataFrame({
            "date": [date],
            "share name": [share_name],
            "broker": [broker],
            "File": [file],
            "purchase cost": [purchase_cost],
            "quantity": [qty],
            "t_day mkt price": [current_market_price],
            "total market value": [qty * current_market_price]
        })

        drill_down_df = pd.concat([drill_down_df, new_row], ignore_index=True)

    return drill_down_df


def init_drill_down_df(filename = "Excels/drill_down_track.csv"):
    if os.path.exists(filename):
        return pd.read_csv(filename)
    else:
        return check_and_create_drill_down_track_df()
    
def save_drill_down_df(drill_down_df, filename="Excels/drill_down_track.csv"):

    drill_down_df = drill_down_df.sort_values(
        by="date",
        key=lambda col: pd.to_datetime(col, errors="coerce", dayfirst=True)
    )
    drill_down_df.to_csv(filename, index=False)