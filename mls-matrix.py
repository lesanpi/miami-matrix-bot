from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.action_chains import ActionChains
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
driver.get(URL)

IDS = {
    "address_input_1": "Fm11_Ctrl7_TB",
    "address_input_2": "Fm16_Ctrl7_TB",
    "address_input_3": "Fm15_Ctrl7_TB",
    "bedrooms_input_1": "Fm11_Ctrl9_TB",
    "bedrooms_input_2": "Fm16_Ctrl9_TB",
    "bedrooms_input_3": "Fm15_Ctrl9_TB",
    "baths_input_1": "Fm11_Ctrl13_TB",
    "baths_input_3": "Fm15_Ctrl13_TB",
    "sqft_living_area_input_1": "Fm11_Ctrl57_TB",
    "sqft_living_area_input_2": "Fm16_Ctrl57_TB",
    "sqft_living_area_input_3": "Fm15_Ctrl57_TB",
    "select_within_1": "Fm11_Ctrl7_Radius",
    "select_within_2": "Fm16_Ctrl7_Radius",
    "select_within_3": "Fm15_Ctrl7_Radius",
    "results_tab": "m_ucResultsPageTabs_m_pnlResultsTab",
    "display": "m_ucDisplayPicker_m_ddlDisplayFormats",
    "read_later": "NewsDetailPostpone",
    "rented_input_3": "FmFm15_Ctrl596_21510_Ctrl596_TB"
}

within_type = 1

XPATHS = {
    "username_input": '//div[@id="clareity"]',
    "password_input": '//div[@id="security"]',
    "login_button": '//button[@id="loginbtn"]',
    "end_tour_button": '//button[@data-role="end"]',
    "matrix_app": '//div[@id="appColumn115"]',
    "search_option": '//a[@href="/Matrix/Search"]/..',
    "options_for_search": '//table[@class="min"]//table[@class="bottom"]//a',
    "dialog_address_search": '//div[@class="mapSearchDialog"]',
    "select_within_option_0.5": f'//select[@id="{IDS[f"select_within_{within_type}"]}"]/option[@value="0.80467200"]',
    "select_within_option_1": f'//select[@id="{IDS[f"select_within_{within_type}"]}"]/option[@value="1.60934400"]',
    "display_closed_comp": f"//select[@id='{IDS['display']}']/option[@value='U73530']",
    "display_marketing_to_realtors": f"//select[@id='{IDS['display']}']/option[@value='U71952']",
    "display_for_sale": f"//select[@id='{IDS['display']}']/option[@value='U74267']",
    "results_table": "//div[@class='css_container']",
    "sp_tab": r"//th[@data-mlheader='1\bSP$\a2\bSP$']",
    "current_price_tab": r"//th[@data-mlheader='1\bCurrent Price\a2\bCurrent Price']",
    "distance_tab": r"//th[@data-mlheader='1\bDistance\a2\bDistance']",
    "active_checkbox_1": '//input[@type="checkbox" and @class="checkbox" and @value="101"]',
    "active_checkbox_2": '//input[@type="checkbox" and @class="checkbox" and @value="101"]',
    "active_checkbox_3": '//input[@type="checkbox" and @class="checkbox" and @value="101"]',
    "rented_checkbox_3": '//input[@type="checkbox" and @class="checkbox" and @value="21510"]',
    "expired_checkbox_3": '//input[@type="checkbox" and @class="checkbox" and @value="106"]'
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
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, XPATHS['end_tour_button']))).click()
    except Exception as e:
        print(e)
    # Select matrix app
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['matrix_app']))).click()

    driver.close()
    new_window = driver.window_handles[0]
    driver.switch_to.window(new_window)
    # driver.switch_to_window(new_window)

    try:
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['read_later']))).click()
    except Exception as e:
        print(e)
        pass


# Searchs
# Select type of search
def select_search(search: int):
    # Click on search
    search_option = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['search_option'])))
    search_option.click()
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Select search type
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['options_for_search'])))
    options_for_search = driver.find_elements_by_xpath(XPATHS['options_for_search'])
    options_for_search[search].click()


def single_family_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=0)

    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    # Filter
    single_family_filter(address, baths, rooms, sqft_to)
    # Results
    return results_family_search(address)


def res_income_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    # Filter
    single_family_filter(address, baths, rooms, sqft_to, miles=1, search_type=2)

    return results_res_income(address)


def res_rental_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=3)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    # Filter
    single_family_filter(address, baths, rooms, sqft_to, miles=1, search_type=3)
    results_res_rental(address)
    results_res_rental(address, display="marketing")


# Filters
def single_family_filter(address, baths, rooms, sqft_to, miles=0.5, search_type=1):
    # RE1/RE2 Single Family/Condo Filter

    # Change within options
    XPATHS["select_within_option_0.5"] = f'//select[@id="{IDS[f"select_within_{search_type}"]}"]' \
                                         f'/option[@value="0.80467200"]'
    XPATHS[
        "select_within_option_1"] = f'//select[@id="{IDS[f"select_within_{search_type}"]}"]/option[@value="1.60934400"]'

    # Within open
    select_within = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, IDS[f'select_within_{search_type}'])))
    select_within.click()

    # Select option
    select_within_option = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS[f"select_within_option_{miles}"])))

    # Within close
    select_within_option.click()

    # Address input
    address_input = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, IDS[f'address_input_{search_type}'])))
    address_input.click()
    address_input.send_keys(address)
    # Select option
    dialog_address_search = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['dialog_address_search'])))
    dialog_address_search.click()

    if search_type != 2:
        # Bedrooms
        bedrooms_input = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, IDS[f'bedrooms_input_{search_type}'])))
        if rooms:
            bedrooms_input.send_keys(f"{rooms - 1}-{rooms + 1}")

        # Baths
        baths_input = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, IDS[f'baths_input_{search_type}'])))
        if baths:
            baths_input.send_keys(f"{baths - 1}-{baths + 1}")

    # SQFT
    sqft_living_area_input = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, IDS[f'sqft_living_area_input_{search_type}'])))
    if sqft_to:
        sqft_living_area_input.send_keys(f"0-{sqft_to + 700}")

    if search_type == 3:
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['active_checkbox_3']))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['rented_checkbox_3']))).click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['expired_checkbox_3']))).click()

        rented_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['rented_input_3'])))
        rented_input.clear()
        rented_input.send_keys(f"0-365")

    # Go to Results
    results_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['results_tab'])))
    results_tab.click()


# Results actions
# Order
def order_by_sp():
    sp_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['sp_tab'])))
    sp_tab.click()


def order_by_current_price():
    current_price_tab = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['current_price_tab'])))
    current_price_tab.click()


def order_by_distance():
    distance_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['distance_tab'])))
    distance_tab.click()


# Table actions
def set_display(mode='display_closed_comp'):
    # Open dropdown
    display = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['display'])))
    display.click()
    # Select option
    display_option = WebDriverWait(driver, 15) \
        .until(EC.element_to_be_clickable((By.XPATH, XPATHS[mode])))
    display_option.click()
    # Click again to close dropdown
    sleep(1)
    # display.click()


# Results
def results_family_search(address):
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    set_display()
    sleep(1)

    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    order_by_sp()
    order_by_sp()
    sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))
    location = results_table.location_once_scrolled_into_view
    size = results_table.size
    results_path = screenshot_and_crop('results/single_family', location, size, address)
    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def results_res_income(address):
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    set_display()
    sleep(1)

    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    order_by_sp()
    order_by_sp()
    sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))
    location = results_table.location_once_scrolled_into_view
    size = results_table.size
    results_path = screenshot_and_crop('results/res_income', location, size, address)
    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def results_res_rental(address, display="for_sale"):
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    if display == "for_sale":
        set_display(mode="display_for_sale")
    else:
        set_display(mode="display_marketing_to_realtors")
    sleep(1)

    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    try:
        if display == "for_sale":
            order_by_distance()
        else:
            order_by_current_price()
    except Exception as e:
        print(e)
        print("No results")
    sleep(1)
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))
    location = results_table.location_once_scrolled_into_view
    size = results_table.size
    results_path = screenshot_and_crop('results/res_rental_' + display, location, size, address)
    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def start(address, baths, rooms, sqft_to):
    login()
    sleep(3)

    select_matrix_app()

    single_family_search(address, baths, rooms, sqft_to)
    sleep(5)
    res_income_search(address, baths, rooms, sqft_to)
    sleep(5)
    res_rental_search(address, baths, rooms, sqft_to)


if __name__ == "__main__":
    verify_folder_exists('results')
    verify_folder_exists('criteria')

    start("416 SW 24th", 2, 2, 500)
