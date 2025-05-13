import pandas as pd
from datetime import datetime
import argparse

def parse_schedule(excel_file):
    try:
        phone_df = pd.read_excel(excel_file, sheet_name='telefoonnummers')
        phone_map = dict(zip(phone_df['Name'], phone_df['Phone']))
    except Exception as e:
        print(f"Error reading phone numbers: {e}")
        return []

    schedule = []
    xl = pd.ExcelFile(excel_file)
    
    for sheet_name in xl.sheet_names:
        if sheet_name == 'telefoonnummers':
            continue
        
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        if df.empty:
            continue
        
        try:
            year = pd.to_datetime(df.iloc[0,0]).year
        except:
            print(f"Skipping sheet '{sheet_name}': invalid date format.")
            continue
        
        for col in range(1, df.shape[1]):
            day = df.iloc[0, col]
            if pd.isna(day) or day == '':
                continue
            
            for row in range(2, df.shape[0]):
                name = df.iloc[row, col]
                if pd.notna(name) and name.strip() != '':
                    try:
                        date_str = f"{year}-{sheet_name}-{int(day)}"
                        date = datetime.strptime(date_str, "%Y-%B-%d").strftime("%Y-%m-%d")
                    except ValueError:
                        continue
                    
                    phone = phone_map.get(name, 'N/A')
                    schedule.append({
                        'Date': date,
                        'Name': name,
                        'Phone': phone
                    })
                    break  # Assume one person per day
    
    return schedule

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Parse filled standby schedule Excel file.')
    parser.add_argument('input', help='Path to the filled Excel file')
    parser.add_argument('output', help='Path to the output CSV file')
    args = parser.parse_args()

    data = parse_schedule(args.input)
    if data:
        pd.DataFrame(data).to_csv(args.output, index=False)
        print(f"Success! Parsed data saved to {args.output}")
    else:
        print("No data parsed. Check if the input file is formatted correctly.")