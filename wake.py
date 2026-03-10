import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

STREAMLIT_URL = os.environ.get("STREAMLIT_URL", "https://your-app.streamlit.app/")

options = Options()
options.add_argument('--headless=new')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    driver.get(STREAMLIT_URL)
    wait = WebDriverWait(driver, 15)
    try:
        button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Yes, get this app back up')]")))
        button.click()
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//button[contains(text(),'Yes, get this app back up')]")))
        print("✅ App woken up")
    except TimeoutException:
        print("✅ Already awake")
except Exception as e:
    print(f"❌ Error: {e}")
finally:
    driver.quit()
