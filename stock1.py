import yfinance as yf
import pandas as pd
from datetime import datetime

# Load the Excel file
file_path = 'dtock.xlsx'
current_date = datetime.now().strftime('%d-%m-%Y')

# Read the Excel sheet into a DataFrame
df = pd.read_excel(file_path, engine='openpyxl')

# Check the initial DataFrame
print("Initial DataFrame:")
print(df)

if df.shape[1] >= 2:
    # Loop through each row
    for index, row in df.iterrows():
        # Store the value from the second column (1st index) in a variable
        stock = row.iloc[1]
        print(f"Fetching data for stock: {stock}")  # Print the stock being processed
        
        tick = yf.Ticker(stock)
        try:
            # Get the stock info
            info = tick.info
            
            # Get current price and previous close
            current_price = info.get('currentPrice')
            prev_close = info.get('previousClose')

            # Debugging: Print current price and previous close types and values
            print(f"Current Price: {current_price}, Type: {type(current_price)}")
            print(f"Previous Close: {prev_close}, Type: {type(prev_close)}")

            # Check if prices are available and valid
            if current_price is None or prev_close is None:
                print(f"Data not available for {stock}. Skipping...")
                # Update the current date column if there's no data
                df.at[index, current_date] = None
                continue

            # Type checking to ensure calculations are valid
            if isinstance(current_price, str):
                current_price = float(current_price.replace(',', ''))
            if isinstance(prev_close, str):
                prev_close = float(prev_close.replace(',', ''))

            # Check the types after conversion
            print(f"Converted - Current Price: {current_price}, Type: {type(current_price)}")
            print(f"Converted - Previous Close: {prev_close}, Type: {type(prev_close)}")

            # Calculate change percentage
            if isinstance(current_price, (int, float)) and isinstance(prev_close, (int, float)):
                change = ((current_price - prev_close) / prev_close) * 100
                c1 = f"{change:.2f}"  # Store as string for display
                print(f"Current price: {current_price}, Change: {c1}%")
            else:
                print(f"Invalid types for calculation for {stock}. Skipping...")
                # Update the current date column if there's no data
                df.at[index, current_date] = None
                continue
            
            # Check if the current date column exists
            if current_date in df.columns:
                # Update the existing entry for the current date
                df.at[index, current_date] = current_price
                df.at[index, "change"] = c1
                print(f"Updated row {index} with current price and change for date: {current_date}.")
            else:
                # Add the new columns if they don't exist
                df[current_date] = None  # Add a new column with the current date as the name
                df["change"] = None  # Add a new column named "change"

                # Write the new values in the appropriate columns
                df.at[index, current_date] = current_price  # Set value in the current date column
                df.at[index, "change"] = c1  # Set the change percentage
                print(f"Added new row {index} with current price and change for date: {current_date}.")

        except Exception as e:
            print(f"An error occurred while fetching data for {stock}: {e}")
            # Update the current date column if there's an error
            df.at[index, current_date] = None
            df.at[index,"Change"]=c1

    # Save the modified DataFrame back to the same Excel file
    df.to_excel(file_path, index=False)  # Overwrite the existing file

    # Final DataFrame print
    print("Final DataFrame saved to the same file:")
    print(df)

else:
    print("The sheet doesn't have a second column.")
