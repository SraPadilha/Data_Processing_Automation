import pandas as pd
from pathlib import Path
import sys

base_folder = Path(sys.argv[1])

input_folder = base_folder / 'inputData'
output_folder = base_folder / 'processedData'

output_folder.mkdir(exist_ok=True)

files = [file for file in input_folder.glob('*.xlsx') if not file.name.startswith('~$')]
if not files:
    raise FileNotFoundError("No .xlsx files found in the input folder.")

latest_file = max(files, key=lambda x: x.stat().st_mtime)
print(f"Selected file: {latest_file.name}")

df = pd.read_excel(latest_file)

columns_map = {
    'Environment:Site Outdoor Air Drybulb Temperature [C](TimeStep)': 'External',
    'QUARTO:Zone Mean Air Temperature [C](TimeStep)': 'Sensor1',
    'TELHADO:Zone Mean Air Temperature [C](TimeStep) ': 'Sensor2'
}
df.rename(columns=columns_map, inplace=True)

pure_date = []
pure_time = []

for value in df['Date/Time'].astype(str):
    parts = value.strip().split()
    if len(parts) == 2:
        date_str, time_str = parts
    else:
        date_str = parts[0]
        time_str = '00:00:00'

    day, month = map(int, date_str.split('/'))
    year = 2025  

    formatted_date = f"{day:02d}/{month:02d}/{year}"
    formatted_time = time_str

    pure_date.append(formatted_date)
    pure_time.append(formatted_time)

df['Date'] = pure_date
df['Time'] = pure_time
df.drop(columns=['Date/Time'], inplace=True)

base_name = latest_file.stem
extension = latest_file.suffix
new_file_name = output_folder / f'{base_name}{extension}'

counter = 1
while new_file_name.exists():
    new_file_name = output_folder / f'{base_name} ({counter}){extension}'
    counter += 1

df.to_excel(new_file_name, index=False)
print(f"New file saved at: {new_file_name}")
