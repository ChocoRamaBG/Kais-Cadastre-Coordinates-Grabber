import pandas as pd
import os
import time
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # ЗАДЪЛЖИТЕЛНО ЗА CLOUD
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-search-engine-choice-screen")
    chrome_options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def init_excel(output_path, sheet_name, headers):
    if not os.path.exists(output_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        for col_num, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_num, value=header)
        wb.save(output_path)
        print(f"✅ Файлчовци създадени: {output_path}")

def input_excel(input_path, sheet_name):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"⚠️ Мамка му човече, няма го файла {input_path}!")
    ids_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
    building_ids = ids_df[0].tolist()
    return building_ids, len(building_ids)

def save_to_excel(output_path, sheet_name, row_data):
    row_df = pd.DataFrame([row_data])
    if not os.path.exists(output_path):
        row_df.to_excel(output_path, index=False)
    else:
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            try:
                startrow = writer.book[sheet_name].max_row
            except KeyError:
                startrow = 0
            row_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=(startrow == 0))

# --- ЛОГИКА ---
# БЕЗ D:\! Само имената на файлчовци!
input_path = 'All_Sofia_IDs.xlsx'
output_path = 'Gathered_Sofia_Coords.xlsx'
sheet_name = 'Ids List'
headers = ["Код", "X", "Y"]

start_time = time.time()
print(f"🚀 Старт: {datetime.now().strftime('%H:%M:%S')}")

try:
    init_excel(output_path, sheet_name, headers)
    building_ids, count_ids = input_excel(input_path, sheet_name)

    driver = setup_driver()
    driver.get('https://kais.cadastre.bg/bg/Map')
    time.sleep(3)

    # Кликваме на лупата
    wait = WebDriverWait(driver, 15)
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_wrap"]/div[2]/div[1]/div[1]/a[1]')))
    search_btn.click()

    for idx, building_id in enumerate(building_ids):
        if str(building_id) == "КИ": continue
        
        print(f"🔍 Търся: {building_id} ({idx+1}/{count_ids})")
        
        try:
            input_field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-search-tabs-1"]//input')))
            input_field.clear()
            input_field.send_keys(str(building_id))
            input_field.send_keys(Keys.RETURN)
            
            time.sleep(2) # Дай му време на тоя бавен сайт

            x_coord, y_coord = '-', '-'
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="map-coordinates"]')))
            
            x_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[2]/span/span/input[1]')
            x_coord = x_el.get_attribute("title") or x_el.get_attribute("value")
            
            y_el = driver.find_element(By.XPATH, '//*[@id="map-coordinates"]/div/span[3]/span/span/input[1]')
            y_coord = y_el.get_attribute("title") or y_el.get_attribute("value")
            
            save_to_excel(output_path, sheet_name, [building_id, x_coord, y_coord])
        except Exception as e:
            print(f"⚠️ Грешка при {building_id}: {str(e)[:50]}")
            save_to_excel(output_path, sheet_name, [building_id, "Грешка", "Грешка"])

    driver.quit()
except Exception as global_e:
    print(f"💥 Критичен батак: {global_e}")

elapsed = time.time() - start_time
print(f"🏁 Готово за {int(elapsed//60)} мин.!")