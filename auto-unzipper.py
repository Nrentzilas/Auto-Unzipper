import os
import time
import subprocess


DOWNLOADS_FOLDER = os.path.expanduser("~\\Downloads")
EXTRACT_FOLDER = os.path.join(DOWNLOADS_FOLDER, "Extracted")
SEVEN_ZIP_PATH = r"C:\Program Files\7-Zip\7z.exe"


os.makedirs(EXTRACT_FOLDER, exist_ok=True)

def extract_zip(zip_file):
    zip_name = os.path.splitext(os.path.basename(zip_file))[0]  
    extract_path = os.path.join(EXTRACT_FOLDER, zip_name)


    os.makedirs(extract_path, exist_ok=True)

    print(f"Extracting: {zip_file} -> {extract_path}")
    
    subprocess.run([SEVEN_ZIP_PATH, "x", zip_file, f"-o{extract_path}", "-y"], shell=True)

    os.remove(zip_file)

def monitor_downloads():
    processed_files = set()  

    while True:
        zip_files = [f for f in os.listdir(DOWNLOADS_FOLDER) if f.endswith(".zip")]

        for zip_file in zip_files:
            zip_path = os.path.join(DOWNLOADS_FOLDER, zip_file)
            if zip_path not in processed_files:
                extract_zip(zip_path)
                processed_files.add(zip_path)

        time.sleep(10)  

if __name__ == "__main__":
    print("Auto Unzip Script Started...")
    monitor_downloads()
