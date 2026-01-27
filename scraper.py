import pandas as pd
import os
import time
import sys
import argparse
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURATION ---
# We stop the script 20 mins before the 6h limit to allow Git commit/push time
# 5 hours 40 minutes = 20400 seconds
MAX_RUN_TIME_SECONDS = 20000 

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome_options.add_argument("--window-size=1920,1080")
    # Using a generic user agent so we don't look like total npc-bot-човци
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def get_processed_ids(output_path, sheet_name):
    """ Checks which ID-гащници are already in the database to avoid duplicates. """
    if not os.path.exists(output_path):
        return set()
    try:
        # Check if file is valid zip first
        try:
            df = pd.read_excel(output_path, sheet_name=sheet_name)
        except Exception:
            # If file is corrupted (brainrot), start fresh
            return set()
            
        if df.empty: return set()
        return set(df["Код"].astype(str).str.strip().tolist())
    except Exception as e:
        print(f"⚠️ Warning: Could not read old file. Starting fresh or appending. Error: {e}", flush=True)
        return set()

def save_to_excel(output_path, sheet_name, row_data):
    """ Saves data incrementally. 'Append' mode is sketchy in pandas, so we handle it carefully. """
    row_df = pd.DataFrame([row_data], columns=["Код", "X", "Y"])
    
    # Retry mechanism for file locking issues
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if not os.path.exists(output_path):
                row_df.to_excel(output_path, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    try:
                        # Find the last row
                        if sheet_name in writer.book.sheetnames:
                            startrow = writer.book[sheet_name].max_row
                        else:
                            startrow = 0
                    except KeyError:
                        startrow = 0
                    
                    row_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=(startrow == 0))
            break 
        except Exception as e:
            time.sleep(1)
            if attempt == max_retries - 1:
                print(f"❌ Failed to save row after retries: {e}", flush=True)

# --- MAIN LOGIC ---
if __name__ == "__main__":
    input_path = 'All_Sofia_IDs.xlsx'
    output_path = 'Gathered_Sofia_Coords.xlsx'
    sheet_name = 'Ids List'
    
    # Start the timer for the Sigma Grindset
    start_time = time.time()
    print(f"🚀 Launching scraping operation at: {datetime.now().strftime('%H:%M:%S')}", flush=True)

    try:
        if not os.path.exists(input_path):
            # Create a dummy file for testing if it doesn't exist (for the user to replace)
            print(f"⚠️ Мамка му човече, input file {input_path} missing. Creating dummy for test.", flush=True)
            pd.DataFrame({"IDs": ["68134.4083.606", "68134.4083.607"]}).to_excel(input_path, sheet_name=sheet_name, header=None, index=False)

        # 1. Load all Target IDs
        try:
            all_ids_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
            all_ids = all_ids_df[0].astype(str).str.strip().tolist()
        except ValueError:
             # Try default sheet if name fails
             all_ids_df = pd.read_excel(input_path, header=None)
             all_ids = all_ids_df[0].astype(str).str.strip().tolist()

        # 2. Filter processed ones
        processed_ids = get_processed_ids(output_path, sheet_name)
        to_process = [i for i in all_ids if i not in processed_ids and i != "nan" and i != "КИ"]

        print(f"📊 Total IDs: {len(all_ids)} | Collected: {len(processed_ids)} | Remaining: {len(to_process)}", flush=True)

        if not to_process:
            print("🎉 Mission Passed! All parcel-скандалчовци collected. Respect +.", flush=True)
            sys.exit(0) # Exit code 0 means "DONE"

        # 3. Init Browser
        driver = setup_driver()
        driver.get('https://kais.cadastre.bg/bg/Map')
        wait = WebDriverWait(driver, 20)
        time.sleep(5) 

        # Click Search
        try:
            print("🖱️ Clicking the spyglass...", flush=True)
            # Sometimes XPATH changes, using a more robust wait
            search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_wrap"]/div[2]/div[1]/div[1]/a[1]')))
            search_btn.click()
        except Exception:
            print("⚠️ Failed to click search. Maybe mobile view? Attempting direct interaction.", flush=True)

        # 4. The Grind
        items_processed_this_session = 0
        
        for idx, building_id in enumerate(to_process):
            # CHECK TIME LIMIT - CRITICAL FOR GITHUB
            current_duration = time.time() - start_time
            if current_duration > MAX_RUN_TIME_SECONDS:
                print(f"🛑 Time limit reached ({int(current_duration/60)} mins). Stopping gracefully to commit data.", flush=True)
                print("⚠️ Returning exit code 1 to trigger next workflow run.", flush=True)
                driver.quit()
                sys.exit(1) # Exit code 1 signals "NOT DONE YET, RESTART ME"

            try:
                # Find input
                input_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-search-tabs-1"]//input')))
                input_field.clear()
                # Type slowly like a boomer to avoid detection? Nah, just send keys.
                input_field.send_keys(building_id)
                input_field.send_keys(Keys.RETURN)
                
                # Dynamic wait - KAIS is slow as hell
                time.sleep(2) 

                x_coord, y_coord = '-', '-'
                
                try:
                    # Wait for coords box
                    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-coordinates"]')))
                    
                    x_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[2]/span/span/input[1]')
                    x_coord = x_el.get_attribute("title") or x_el.get_attribute("value") or "-"
                    
                    y_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[3]/span/span/input[1]')
                    y_coord = y_el.get_attribute("title") or y_el.get_attribute("value") or "-"
                except:
                    # If coords don't pop up, maybe ID is invalid or site lagged
                    x_coord, y_coord = "Not Found", "Not Found"

                print(f"🎯 [{idx+1}/{len(to_process)}] ID: {building_id} -> X: {x_coord} | Y: {y_coord}", flush=True)
                
                save_to_excel(output_path, sheet_name, [building_id, x_coord, y_coord])
                items_processed_this_session += 1

            except Exception as e:
                print(f"⚠️ Error on {building_id}: {str(e)[:50]}", flush=True)
                save_to_excel(output_path, sheet_name, [building_id, "Error", "Error"])

            # Anti-ban pause
            if (idx + 1) % 50 == 0:
                print(f"☕ Brainrot break... Progress: {idx+1}/{len(to_process)}", flush=True)
                time.sleep(3)

        driver.quit()
        print("🏁 Session finished naturally. All items in queue processed.", flush=True)
        sys.exit(0) # Done

    except Exception as global_e:
        print(f"💥 Critical Failure (Андибул Морков Error): {global_e}", flush=True)
        if 'driver' in locals(): driver.quit()
        sys.exit(1) # Fail/Retry
