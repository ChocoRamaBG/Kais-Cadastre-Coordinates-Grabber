import pandas as pd
import os
import time
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURATION ---
# Спираме на 5 часа и 30 мин, за да има време за commit и push
# Github убива на 6-тия час.
MAX_RUN_TIME_SECONDS = 19800 

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def get_processed_ids(output_path, sheet_name):
    """ Проверяваме кои ID-гащници вече са минали. """
    if not os.path.exists(output_path):
        return set()
    try:
        df = pd.read_excel(output_path, sheet_name=sheet_name)
        if df.empty: return set()
        # Превръщаме в string, за да не гърми brainrot кода
        return set(df["Код"].astype(str).str.strip().tolist())
    except Exception as e:
        print(f"⚠️ Warning: Could not read old file. Starting fresh. Error: {e}", flush=True)
        return set()

def save_to_excel(output_path, sheet_name, row_data):
    """ Save logic - append mode """
    row_df = pd.DataFrame([row_data], columns=["Код", "X", "Y"])
    
    for attempt in range(3):
        try:
            if not os.path.exists(output_path):
                row_df.to_excel(output_path, sheet_name=sheet_name, index=False)
            else:
                with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    try:
                        startrow = writer.book[sheet_name].max_row
                    except KeyError:
                        startrow = 0
                    row_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=(startrow == 0))
            break 
        except Exception as e:
            time.sleep(1)
            if attempt == 2:
                print(f"❌ Failed to save row: {e}", flush=True)

# --- MAIN LOGIC ---
if __name__ == "__main__":
    # Увери се, че пътищата са точни!
    input_path = 'All_Sofia_IDs.xlsx'
    output_path = 'Gathered_Sofia_Coords.xlsx'
    sheet_name = 'Ids List'
    
    start_time = time.time()
    print(f"🚀 Bootleg Chat Scraper v2.0 started at: {datetime.now().strftime('%H:%M:%S')}", flush=True)

    try:
        # 1. Load Data
        if not os.path.exists(input_path):
            print(f"☠️ Critical: Input file {input_path} not found. Мамка му човече, няма данни!", flush=True)
            sys.exit(1) # Error

        try:
            all_ids_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
        except:
             # Fallback ако sheet name е грешен
             all_ids_df = pd.read_excel(input_path, header=None)
             
        all_ids = all_ids_df[0].astype(str).str.strip().tolist()

        # 2. Filter Processed
        processed_ids = get_processed_ids(output_path, sheet_name)
        to_process = [i for i in all_ids if i not in processed_ids and str(i).lower() not in ["nan", "ки", "код"]]

        print(f"📊 Stats: Total: {len(all_ids)} | Done: {len(processed_ids)} | Left: {len(to_process)}", flush=True)

        if not to_process:
            print("🎉 Mission Passed! Всички имотчовци са събрани.", flush=True)
            sys.exit(0) # SUCCESS - STOP LOOP

        # 3. Init Browser
        driver = setup_driver()
        driver.get('https://kais.cadastre.bg/bg/Map')
        wait = WebDriverWait(driver, 20)
        time.sleep(5) 

        # Click Search
        try:
            search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_wrap"]/div[2]/div[1]/div[1]/a[1]')))
            search_btn.click()
        except:
            print("⚠️ Search button click failed, continuing anyway...", flush=True)

        # 4. The Grind
        for idx, building_id in enumerate(to_process):
            # CHECK TIME LIMIT
            if (time.time() - start_time) > MAX_RUN_TIME_SECONDS:
                print(f"🛑 Time limit reached. Stopping to save progress.", flush=True)
                driver.quit()
                sys.exit(1) # CONTINUE LOOP CODE

            try:
                input_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-search-tabs-1"]//input')))
                input_field.clear()
                input_field.send_keys(building_id)
                input_field.send_keys(Keys.RETURN)
                
                # Sleep is for the weak, but KAIS is weaker
                time.sleep(2.5) 

                x_coord, y_coord = '-', '-'
                try:
                    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-coordinates"]')))
                    x_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[2]/span/span/input[1]')
                    x_coord = x_el.get_attribute("title") or x_el.get_attribute("value") or "-"
                    y_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[3]/span/span/input[1]')
                    y_coord = y_el.get_attribute("title") or y_el.get_attribute("value") or "-"
                except:
                    x_coord, y_coord = "Not Found", "Not Found"

                print(f"🎯 [{idx+1}/{len(to_process)}] {building_id} -> {x_coord} | {y_coord}", flush=True)
                save_to_excel(output_path, sheet_name, [building_id, x_coord, y_coord])

            except Exception as e:
                print(f"⚠️ Glitch on {building_id}: {str(e)[:50]}", flush=True)
                save_to_excel(output_path, sheet_name, [building_id, "Error", "Error"])

            # Anti-ban chill
            if (idx + 1) % 50 == 0:
                print("☕ Brainrot break...", flush=True)
                time.sleep(3)

        driver.quit()
        print("🏁 Queue finished naturally.", flush=True)
        sys.exit(0) # Done

    except Exception as global_e:
        print(f"💥 Critical Failure (Андибул морков): {global_e}", flush=True)
        if 'driver' in locals(): driver.quit()
        sys.exit(2) # Real Error (Don't restart blindly)
