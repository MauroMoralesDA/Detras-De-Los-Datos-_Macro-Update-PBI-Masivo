import subprocess
import os
import time
import pyautogui
# Define the path to the Power BI Desktop executable
pbi_path = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
# Define the directory containing the Power BI files
directory_path = r'C:\Users\DELL\Documents\Data Target\OneDrive\LinkedIn_DataStrikes\Solutions\Python Solutions\Data'
# Function to open a Power BI file, wait for 5 minutes, and then close it
def open_and_close_pbix(file_path):
    try:
        # Start Power BI Desktop with the specified .pbix file
        process = subprocess.Popen([pbi_path, file_path])
        print(f"Opened {file_path} in Power BI Desktop.")
        # Wait for 5 minutes (300 seconds)
        time.sleep(120)
        pyautogui.moveTo(700, 130)
        pyautogui.click()   
        time.sleep(120) 
        pyautogui.moveTo(40, 10) 
        pyautogui.click()  
        time.sleep(120)
        #Updating the files
        # Terminate the Power BI Desktop process
        process.terminate()
        print(f"Closed {file_path} after 5 minutes.")
    except Exception as e:
        print(f"Failed to manage {file_path}: {e}")
# Loop through all files in the directory
for file_name in os.listdir(directory_path):
    if file_name.endswith(".pbix"):
        file_path = os.path.join(directory_path, file_name)
        open_and_close_pbix(file_path)
