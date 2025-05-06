import subprocess
from pathlib import Path
import shutil

MAIN_PATH = Path('C:/Users/Build/Desktop/folder')
SCRIPTS_PATH = Path('C:/Users/Build/Desktop/folder/scripts')

def process_folder(target_folder):
    print(f"\nProcessing folder: {target_folder.name}")

    input_data_folder = target_folder / 'inputData'
    processed_data_folder = target_folder / 'processedData'
    report_folder = target_folder / 'reportData'

    input_data_folder.mkdir(exist_ok=True)
    processed_data_folder.mkdir(exist_ok=True)
    report_folder.mkdir(exist_ok=True)

    comb_files = list(target_folder.glob('*_COMB.xlsx'))
    if not comb_files:
        print(f"No *_COMB.xlsx files found in {target_folder.name}. Skipping...")
        return

    for file in comb_files:
        destination = input_data_folder / file.name
        if not destination.exists():
            shutil.move(str(file), str(destination))
            print(f"Moved file to inputData/: {file.name}")
        else:
            print(f"File already exists in inputData/: {file.name}")

    comb_files_energyplus = list(input_data_folder.glob('*_COMB.xlsx'))
    if not comb_files_energyplus:
        print(f"No file found to modify in {input_data_folder}. Skipping folder.")
        return

    try:
        print("Running preprocess_data.py...")
        subprocess.run(["python", str(SCRIPTS_PATH / "preprocess_data.py"), str(target_folder)], check=True)

        print("Running report_generate.py...")
        subprocess.run(["python", str(SCRIPTS_PATH / "report_generate.py"), str(target_folder)], check=True)

        print(f"Folder {target_folder.name} processed successfully!")

    except subprocess.CalledProcessError as e:
        print(f"Error processing {target_folder.name}: {e}")

if __name__ == "__main__":
    for path in MAIN_PATH.iterdir():
        if path.is_dir() and not path.name.lower().startswith('scripts'):
            process_folder(path)

    print("\nAll processing completed!")
