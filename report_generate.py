import sys
import pandas as pd
import calendar
from pathlib import Path
from datetime import datetime, timedelta

base_path = Path(sys.argv[1])

INPUT_PATH = base_path / 'processedData'
OUTPUT_PATH = base_path / 'reportData'
RANGE_MIN = 18
RANGE_MAX = 29
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
        'Sensor1',
        'Sensor2'
    ]

    df.rename(columns={'Room': 'Sensor1', 'Roof': 'Sensor2'}, inplace=True)

    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column not found: {col}")

    for col in ['Sensor1', 'Sensor2']:
        df[col] = df[col].astype(str).str.replace(',', '.').astype(float)

    df['DateTime'] = parse_datetime(df)
    df = df[df['DateTime'].notna()].copy()

    df['Month'] = df['DateTime'].dt.month
    df['Month_Name'] = df['DateTime'].dt.month_name()
    df['Hour'] = df['DateTime'].dt.hour
    df['Day'] = df['DateTime'].dt.day

    print(f" Data period: {df['DateTime'].min()} to {df['DateTime'].max()}")
    return df

def calculate_range_stats(data, temp_column):
    total_hours = len(data) * HOURS_PER_INTERVAL

    day_period = (data['Hour'] >= 6) & (data['Hour'] < 18)
    in_range = data[temp_column].between(RANGE_MIN, RANGE_MAX)

    range_day = (day_period & in_range).sum() * HOURS_PER_INTERVAL
    range_night = (~day_period & in_range).sum() * HOURS_PER_INTERVAL
    total_range = range_day + range_night

    return {
        'Total Hours': round(total_hours, 2),
        'In Range (Day)': round(range_day, 2),
        'In Range (Night)': round(range_night, 2),
        'Total In Range': round(total_range, 2),
        'Out of Range': round(total_hours - total_range, 2),
        '% In Range': round((total_range / total_hours) * 100, 1),
        '% Out of Range': round(100 - ((total_range / total_hours) * 100), 1)
    }

def generate_report(df):
    results = []

    for month_num in range(1, 13):
        month_name = calendar.month_name[month_num]
        month_data = df[df['Month'] == month_num]

        if month_data.empty:
            print(f"No data for {month_name}")
            for location in ['Sensor1', 'Sensor2']:
                results.append({
                    'Month': month_name,
                    'Location': location,
                    'Total Hours': 0,
                    'In Range (Day)': 0,
                    'In Range (Night)': 0,
                    'Total In Range': 0,
                    'Out of Range': 0,
                    '% In Range': 0,
                    '% Out of Range': 0
                })
            continue

        for location, temp_column in [
            ('Sensor1', 'Sensor1'),
            ('Sensor2', 'Sensor2')
        ]:
            metrics = calculate_range_stats(month_data.copy(), temp_column)
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

        base_name = 'final_data_analysis'
        extension = '.xlsx'
        output_file = OUTPUT_PATH / f'{base_name}{extension}'

        counter = 1
        while output_file.exists():
            output_file = OUTPUT_PATH / f'{base_name} ({counter}){extension}'
            counter += 1

        report.to_excel(output_file, index=False)
        print(f"\n Report generated successfully: {output_file}")

    except Exception as e:
        print(f"\n Error during execution: {str(e)}")

