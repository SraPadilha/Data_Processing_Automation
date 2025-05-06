import sys
import pandas as pd
import calendar
from pathlib import Path
from datetime import datetime, timedelta

base_path = Path(sys.argv[1])

INPUT_PATH = base_path / 'processedData'
OUTPUT_PATH = base_path / 'reportData'
COMFORT_MIN = 18
COMFORT_MAX = 29
INTERVAL_MINUTES = 15
HOURS_PER_INTERVAL = INTERVAL_MINUTES / 60

def parse_datetime(dataframe):
    real_dates = []

    for idx, row in dataframe.iterrows():
        try:
            date_str = row['Date']
            time_str = row['Time']

            base_date = datetime.strptime(date_str, '%d/%m/%Y')

            if time_str.startswith('24:'):
                base_date += timedelta(days=1)
                correct_date = datetime(
                    year=base_date.year,
                    month=base_date.month,
                    day=base_date.day,
                    hour=0,
                    minute=0,
                    second=0
                )
            else:
                hour, minute, second = map(int, time_str.split(':'))
                correct_date = datetime(
                    year=base_date.year,
                    month=base_date.month,
                    day=base_date.day,
                    hour=hour,
                    minute=minute,
                    second=second
                )

            real_dates.append(pd.Timestamp(correct_date))

        except Exception as e:
            print(f"Error converting row {idx}: {e}")
            real_dates.append(pd.NaT)

    return real_dates

def load_data():
    files = [file for file in INPUT_PATH.glob('*.xlsx') if not file.name.startswith('~$')]
    if not files:
        raise FileNotFoundError("No .xlsx files found in the processed data folder.")

    file = max(files, key=lambda x: x.stat().st_mtime)
    print(f"Selected file: {file.name}")

    df = pd.read_excel(file)
    print(f"Data loaded. Total records: {len(df):,}")

    required_columns = [
        'Date',
        'Time',
        'Room',
        'Roof'
    ]

    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column not found: {col}")

    for col in ['Room', 'Roof']:
        df[col] = df[col].astype(str).str.replace(',', '.').astype(float)

    df['DateTime'] = parse_datetime(df)
    df = df[df['DateTime'].notna()].copy()

    df['Month'] = df['DateTime'].dt.month
    df['Month_Name'] = df['DateTime'].dt.month_name()
    df['Hour'] = df['DateTime'].dt.hour
    df['Day'] = df['DateTime'].dt.day

    print(f" Data period: {df['DateTime'].min()} to {df['DateTime'].max()}")
    return df

def calculate_comfort(data, temp_column):
    total_hours = len(data) * HOURS_PER_INTERVAL

    day_period = (data['Hour'] >= 6) & (data['Hour'] < 18)
    in_comfort = data[temp_column].between(COMFORT_MIN, COMFORT_MAX)

    comfort_day = (day_period & in_comfort).sum() * HOURS_PER_INTERVAL
    comfort_night = (~day_period & in_comfort).sum() * HOURS_PER_INTERVAL
    total_comfort = comfort_day + comfort_night

    return {
        'Total Hours': round(total_hours, 2),
        'Comfort (Day)': round(comfort_day, 2),
        'Comfort (Night)': round(comfort_night, 2),
        'Total Comfort': round(total_comfort, 2),
        'No Comfort': round(total_hours - total_comfort, 2),
        '% Comfort': round((total_comfort / total_hours) * 100, 1),
        '% No Comfort': round(100 - ((total_comfort / total_hours) * 100), 1)
    }

def generate_report(df):
    results = []

    for month_num in range(1, 13):
        month_name = calendar.month_name[month_num]
        month_data = df[df['Month'] == month_num]

        if month_data.empty:
            print(f"No data for {month_name}")
            for location in ['Room', 'Roof']:
                results.append({
                    'Month': month_name,
                    'Location': location,
                    'Total Hours': 0,
                    'Comfort (Day)': 0,
                    'Comfort (Night)': 0,
                    'Total Comfort': 0,
                    'No Comfort': 0,
                    '% Comfort': 0,
                    '% No Comfort': 0
                })
            continue

        for location, temp_column in [
            ('Room', 'Room'),
            ('Roof', 'Roof')
        ]:
            metrics = calculate_comfort(month_data.copy(), temp_column)
            results.append({
                'Month': month_name,
                'Location': location,
                **metrics
            })

    return pd.DataFrame(results)

if __name__ == "__main__":
    try:
        data = load_data()

        report = generate_report(data)

        base_name = 'final_thermal_comfort_analysis'
        extension = '.xlsx'
        output_file = OUTPUT_PATH / f'{base_name}{extension}'

        counter = 1
        while output_file.exists():
            output_file = OUTPUT_PATH / f'{base_name} ({counter}){extension}'
            counter += 1

        report.to_excel(output_file, index=False)
        print(f"\n Report generated successfully: {output_file}")

        print("\n Report sample:")
        print(report.head(8))

        print("\n Monthly hours check:")
        for month in calendar.month_name[1:]:
            if month in report['Month'].values:
                total = report[report['Month'] == month]['Total Hours'].iloc[0]
                days = int(total / 24)
                print(f"- {month}: {total}h ({days} days)")

    except Exception as e:
        print(f"\n Error during execution: {str(e)}")
