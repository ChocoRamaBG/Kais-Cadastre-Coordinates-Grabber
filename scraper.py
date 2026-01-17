import pandas as pd
import os
import time
import openpyxl
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome_options.add_argument("--window-size=1920,1080")
    # Добавяме User-Agent, за да не ни хванат веднага какви сме гащници
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def get_processed_ids(output_path, sheet_name):
    """ Проверява кои идентификаторчовци вече са в кюпа """
    if not os.path.exists(output_path):
        return set()
    try:
        df = pd.read_excel(output_path, sheet_name=sheet_name)
        if df.empty: return set()
        # Превръщаме в string, за да нямаме скандалчовци с типовете данни
        return set(df["Код"].astype(str).str.strip().tolist())
    except Exception as e:
        print(f"⚠️ Бележка: Не успях да прочета стария файл (може да е празен). Грешка: {e}", flush=True)
        return set()

def save_to_excel(output_path, sheet_name, row_data):
    row_df = pd.DataFrame([row_data])
    if not os.path.exists(output_path):
        row_df.to_excel(output_path, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            try:
                startrow = writer.book[sheet_name].max_row
            except KeyError:
                startrow = 0
            row_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=(startrow == 0))

# --- ОСНОВНА ЛОГИКА ---
input_path = 'All_Sofia_IDs.xlsx'
output_path = 'Gathered_Sofia_Coords.xlsx'
sheet_name = 'Ids List'
headers = ["Код", "X", "Y"]

start_time = time.time()
print(f"🚀 Старт на операцията: {datetime.now().strftime('%H:%M:%S')}", flush=True)

try:
    # 1. Зареждане на всички ID-та
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"⚠️ Мамка му човече, няма го входния файл {input_path}!")
    
    all_ids_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
    all_ids = all_ids_df[0].astype(str).str.strip().tolist()
    
    # 2. Проверка за вече обработени
    processed_ids = get_processed_ids(output_path, sheet_name)
    to_process = [i for i in all_ids if i not in processed_ids and i != "КИ"]
    
    print(f"📊 Общо ID-та: {len(all_ids)} | Вече събрани: {len(processed_ids)} | Остават: {len(to_process)}", flush=True)

    if not to_process:
        print("🎉 Всичко е готово бе, шефе! Никакви бачкаторчовци не останаха.", flush=True)
        sys.exit(0)

    # 3. Стартиране на браузъра
    driver = setup_driver()
    driver.get('https://kais.cadastre.bg/bg/Map')
    wait = WebDriverWait(driver, 20)
    time.sleep(5) # Дай му време да загрее

    # Отваряне на панела за търсене
    print("🖱️ Кликам на лупата...", flush=True)
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_wrap"]/div[2]/div[1]/div[1]/a[1]')))
    search_btn.click()

    # 4. Същинското копаене
    for idx, building_id in enumerate(to_process):
        try:
            input_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-search-tabs-1"]//input')))
            input_field.clear()
            input_field.send_keys(building_id)
            input_field.send_keys(Keys.RETURN)
            
            # Сайтът е бавен като държавен служител пред пенсия
            time.sleep(2.5) 

            x_coord, y_coord = '-', '-'
            # Чакаме да се появят координатчовци
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-coordinates"]')))
            
            x_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[2]/span/span/input[1]')
            x_coord = x_el.get_attribute("title") or x_el.get_attribute("value") or "-"
            
            y_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[3]/span/span/input[1]')
            y_coord = y_el.get_attribute("title") or y_el.get_attribute("value") or "-"
            
            # OUTPUT-А, КОЙТО ИСКАШЕ:
            print(f"🎯 [{idx+1}/{len(to_process)}] ID: {building_id} -> X: {x_coord} | Y: {y_coord}", flush=True)
            
            save_to_excel(output_path, sheet_name, [building_id, x_coord, y_coord])

        except Exception as e:
            print(f"⚠️ Грешка при {building_id}: Скипвам. (Грешка: {str(e)[:50]})", flush=True)
            save_to_excel(output_path, sheet_name, [building_id, "Error", "Error"])

        # На всеки 100 записа, правим малка почивка да не ни баннат
        if (idx + 1) % 100 == 0:
            print(f"☕ Вземам глътка въздух... Прогрес: {idx+1}/{len(to_process)}", flush=True)
            time.sleep(5)

    driver.quit()

except Exception as global_e:
    print(f"💥 Hell no! Критичен скандалчовци: {global_e}", flush=True)

print(f"🏁 Секцията приключи. Време: {int((time.time() - start_time)//60)} мин.", flush=True)
