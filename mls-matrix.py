from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
from time import sleep
import os

URL = "https://sef.clareityiam.net/idp/login"

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
driver.maximize_window()

IDS = {
    "address_input": "Fm11_Ctrl7_TB",
    "bedrooms_input": "Fm11_Ctrl9_TB",
    "baths_input": "Fm11_Ctrl13_TB",
    "sqft_living_area_input": "Fm11_Ctrl57_TB",
    "select_within": "Fm11_Ctrl7_Radius",
    "results_tab": "m_ucResultsPageTabs_m_pnlResultsTab",
    "display": "m_ucDisplayPicker_m_ddlDisplayFormats"
}

XPATHS = {
    "username_input": '//div[@id="clareity"]',
    "password_input": '//div[@id="security"]',
    "login_button": '//button[@id="loginbtn"]',
    "end_tour_button": '//button[@data-role="end"]',
    "matrix_app": '//div[@id="appColumn115"]',
    "search_option": '//li[@class="sf-top"]/a[@href="/Matrix/Search"]/..',
    "options_for_search": '//table[@class="min"]//table[@class="bottom"]//a',
    "dialog_address_search": '//div[@class="mapSearchDialog"]',
    "select_within_option_0.5": f'//select[@id="{IDS["select_within"]}"]/option[@value="0.80467200"]',
    "display_closed_comp": f"//select[@id='{IDS['display']}']/option[@value='U73530']",
    "results_table": "//div[@class='css_container']",
    "sp_tab": "//*[@id='m_pnlDisplay']/table/thead/tr/th[12]"#'''th[@data-mlheader=1\bSP$\a2\bSP$]'''
}


def screenshot_and_crop(folder, location, size, address):
    verify_folder_exists(folder)
    file_path = f"{folder}/{address.lower().replace(' ', '-').replace(',', '')}.png"
    driver.save_screenshot(file_path)

    # crop image
    x = location['x']
    y = location['y']
    width = location['x'] + size['width']
    height = location['y'] + size['height']
    im = Image.open(file_path)
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(file_path)
    return file_path

def verify_folder_exists(folder):
    if not os.path.exists(folder):
        os.makedirs(folder)

# Login
def login():
    # Login
    username_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['username_input'])))

    # Username
    username_input.click()
    username_input.send_keys('3457717')

    # Password
    password_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['password_input'])))
    password_input.click()
    password_input.send_keys('1234armando')

    # Login
    login_button = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['login_button'])))
    login_button.click()

# Select app
def select_matrix_app():
    # Quit dialog
    try:
        end_tour_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, XPATHS['end_tour_button']))).click()
    except:
        pass
    # Select matrix app
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['matrix_app']))).click()

    driver.close()
    new_window = driver.window_handles[0]
    driver.switch_to_window(new_window)



## Searchs ##
# Select type of search
def select_search(search: int):
    # Click on search
    search_option = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['search_option'])))
    search_option.click()
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Select search type
    options_for_search = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['options_for_search'])))
    options_for_search = driver.find_elements_by_xpath(XPATHS['options_for_search'])
    options_for_search[search].click()

def single_family_search(address, baths, rooms, sqft_to):
    # Search
    select_search(0)
    # Filter
    single_family_filter(address, baths, rooms, sqft_to)
    # Results
    return results_family_search(address)

# Filters
def single_family_filter(address, baths, rooms, sqft_to):
    # RE1/RE2 Single Family/Condo Filter

    # Within open
    select_within = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['select_within'])))
    select_within.click()

    # Select option
    select_within_option_05 = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS["select_within_option_0.5"])))

    # Within close
    select_within_option_05.click()

    # Address input
    address_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['address_input'])))
    address_input.send_keys(address)
    # Select option
    dialog_address_search = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['dialog_address_search'])))
    dialog_address_search.click()

    # Bedrooms
    bedrooms_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['bedrooms_input'])))
    if rooms:
        bedrooms_input.send_keys(f"{rooms - 1}-{rooms + 1}")

    # Baths
    baths_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['baths_input'])))
    if baths:
        baths_input.send_keys(f"{baths - 1}-{baths + 1}")

    # SQFT
    sqft_living_area_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['sqft_living_area_input'])))
    if sqft_to:
        sqft_living_area_input.send_keys(f"0-{sqft_to + 700}")

    # Go to Results
    results_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['results_tab'])))
    results_tab.click()

## Results actions ##
# Order
def order_by_sp():
    sp_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['sp_tab'])))
    sp_tab.click()

# Table actions
def set_display(mode='display_closed_com'):
    # Open dropdown
    display = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['display'])))
    display.click()
    # Select option
    display_option = WebDriverWait(driver, 15)\
        .until(EC.element_to_be_clickable((By.XPATH, XPATHS[mode])))
    display_option.click()
    # Click again to close dropdown
    display.click()

# Results
def results_family_search(address):
    set_display()
    order_by_sp()
    sleep(5)

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))
    location = results_table.location_once_scrolled_into_view
    size = results_table.size
    results_path = screenshot_and_crop('results', location, size, address)
    return results_path


def start(address, baths, rooms, sqft_to):
    driver.get(URL)
    login()
    select_matrix_app()

    single_family_search(address, baths, rooms, sqft_to)
