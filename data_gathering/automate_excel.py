import re
import time
import os
import pyautogui

# --- 1. Configuration ---
SOURCE_DIR = r'C:\Users\adamf\PycharmProjects\FedContracts_AlgoTrading\datsets\prices\price_formulas'
OUTPUT_DIR = r'C:\Users\adamf\PycharmProjects\FedContracts_AlgoTrading\datsets\prices\price_loaded'
SHORTCUT = "^+S"    # ^ is Ctrl, + is Shift, S is S
#

DOWNLOAD_WAIT = 260

# Ensure the output directory exists
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def process_with_keys(file_name):
    save_file_name = file_name.replace(".csv", "_loaded.csv")
    full_path = os.path.join(SOURCE_DIR, file_name)
    save_path = os.path.join(OUTPUT_DIR, save_file_name)
        
    print(f"Opening {file_name}...")
    os.startfile(full_path)

    print('Opened file...')
    
    # 1. Wait for Excel to open and be the active window
    time.sleep(40)
    
    # 2. Trigger the Capital IQ Refresh (go to ribon and press refresh)
    print("Triggering Refresh...")
    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('g')
    time.sleep(1)
    pyautogui.press('r')
    time.sleep(0.5)
    pyautogui.press('s')
    
    # 3. Wait for the server to send data
    print(f"Waiting {DOWNLOAD_WAIT} seconds for download...")
    time.sleep(DOWNLOAD_WAIT)

    # --- NEW: CONVERT FORMULAS TO VALUES ---
    print("Converting formulas to values...")
    # Select all data
    pyautogui.hotkey('ctrl', 'a') 
    time.sleep(0.5)
    # Copy all data
    pyautogui.hotkey('ctrl', 'c') 
    time.sleep(0.5)
    # Paste Values: Alt, then H (Home), then V (Paste), then V (Values)
    pyautogui.press('alt')
    time.sleep(0.3)
    pyautogui.press('h')
    time.sleep(0.3)
    pyautogui.press('v') 
    time.sleep(0.3)
    pyautogui.press('v')
    time.sleep(5) # Give Excel a second to process the heavy paste
    
    # 4. Save As Sequence (Alt, F, A) - Standard Excel Save As
    print("Saving...")
    pyautogui.press('alt')
    time.sleep(0.5)
    pyautogui.press('f')
    time.sleep(0.5)
    pyautogui.press('a')
    time.sleep(0.5)
    pyautogui.press('o') # For 'Browse' in modern Excel
    time.sleep(1.5)
    
    # 5. Type the new path into the Save Dialog
    pyautogui.write(save_path)
    time.sleep(2)
    pyautogui.press('enter')
    

    time.sleep(1)
    pyautogui.press('enter') # accept saving one sheet at a time for csv

    time.sleep(10)
    
    print(f"Finished {file_name}\n")



def main():
    # open excel
    pyautogui.press("win")

    time.sleep(1)
    pyautogui.write("excel")

    time.sleep(0.5)
    pyautogui.press('enter')

    time.sleep(10)

    # --- Execution ---
    files = [f for f in os.listdir(SOURCE_DIR) if f.endswith('.csv')]
    completed = [c for c in os.listdir(OUTPUT_DIR) if c.endswith('.csv')]
    # This extracts the digits and treats them as an integer for comparison
    files.sort(key=lambda f: int(re.sub('\D', '', f)))

    # count how many sheets are open
    count = 0
    times = 0

    for i, f in enumerate(files):
        print(f'{times=}')
        if times == 20:
            break
        print(f'Batch {i}')
        if f.replace(".csv", "_loaded.csv") in completed:
            print('Already done')
        else:
            times += 1
            print(f'Processing...')
            process_with_keys(f)
            count += 1
            while count > 1:
                time.sleep(1)
                pyautogui.hotkey('alt', 'f4')
                time.sleep(1)
                pyautogui.press('right')
                pyautogui.press('enter')
                count -= 1

if __name__ == "__main__":
    main()
