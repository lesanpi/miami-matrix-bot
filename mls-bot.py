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
import pprint
from pptx import Presentation
from pptx.util import Inches, Pt

printer = pprint.PrettyPrinter(indent=1)

URL = "https://sef.clareityiam.net/idp/login"
COUNTY_URL = "https://www.miamidade.gov/Apps/PA/propertysearch/#/"
GOOGLE_URL = "https://www.google.com"

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
#options.add_argument('--headless')
options.add_argument('--log-level=3')
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
driver.maximize_window()

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
    "display_ac/p/a_review": f"//select[@id='{IDS['display']}']/option[@value='U84429']",
    "results_table": "//div[@class='css_container']",
    "sp_tab": r"//th[contains(@data-mlheader, '1\bSP$\a2\b')]",
    "current_price_tab": r"//th[@data-mlheader='1\bCurrent Price\a2\bCurrent Price']",
    "distance_tab": r"//th[@data-mlheader='1\bDistance\a2\bDistance']",
    "active_checkbox": '//input[@type="checkbox" and @class="checkbox" and @value="101"]',
    "rented_checkbox": '//input[@type="checkbox" and @class="checkbox" and @value="21510"]',
    "expired_checkbox": '//input[@type="checkbox" and @class="checkbox" and @value="106"]',
    "pending_checkbox": '//input[@type="checkbox" and @class="checkbox" and @value="21508"]',
    "active_with_contract_checkbox": '//input[@type="checkbox" and @class="checkbox" and @value="21505"]',
    "filter_container": "//div[@class='css_container']",
    "folio_number_link": "//div[@class='results_record ng-scope']/div/span"
}


def screenshot_and_crop(folder, location, size, filename):
    verify_folder_exists(folder)
    file_path = f"{folder}/{filename.lower().replace(' ', '-').replace(',', '')}.png"
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


def criteria_screenshot(address, folder):
    # Screenshot
    filter_container = WebDriverWait(driver, 15) \
        .until(EC.element_to_be_clickable((By.XPATH, XPATHS['filter_container'])))

    filter_path = screenshot_of_element(filter_container, 'criteria/' + folder, address, width=1000)
    driver.execute_script("window.scrollTo(0, 0)")
    return filter_path


def screenshot_of_element(element, folder, filename, width = None):
    location = element.location_once_scrolled_into_view
    size = element.size
    if width:
        size['width'] = width
    file_path = screenshot_and_crop(folder, location, size, filename)
    return file_path


def verify_folder_exists(folder):
    if not os.path.exists(folder):
        os.makedirs(folder)


# Login
def login():
    # Login
    driver.get(URL)
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
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, XPATHS['end_tour_button']))).click()
    except Exception as e:
        print("Quit tour error \n")
    # Select matrix app
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['matrix_app']))).click()

    driver.close()
    new_window = driver.window_handles[0]
    driver.switch_to.window(new_window)
    # driver.switch_to_window(new_window)

    try:
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['read_later']))).click()
    except Exception as e:
        print("Modal Error \n")
        pass


def escape():
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()


# Searchs
# Select type of search
def select_search(search: int):
    # Click on search
    search_option = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['search_option'])))
    search_option.click()
    escape()

    # Select search type
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS['options_for_search'])))
    options_for_search = driver.find_elements_by_xpath(XPATHS['options_for_search'])
    options_for_search[search].click()


def single_family_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=0)

    escape()
    # Filter
    criteria_path = single_family_filter(address, baths, rooms, sqft_to, folder="single_family")
    # Results
    results_path = results_family_search(address)

    return {
        "criteria": criteria_path,
        "results": results_path
    }


def res_income_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=1)
    escape()
    # Filter
    criteria_path = single_family_filter(address, baths, rooms, sqft_to, miles=1, search_type=2, folder="res_income")

    results_path = results_res_income(address)

    return {
        "criteria": criteria_path,
        "results": results_path
    }


def res_rental_search(address, baths, rooms, sqft_to):
    # Search
    select_search(search=3)
    escape()
    # Filter
    criteria_path = single_family_filter(address, baths, rooms, sqft_to, miles=1, search_type=3, folder="res_rental")
    results_path_1 = results_res_rental(address)
    results_path_2 = results_res_rental(address, display_mode="marketing")

    return {
        "criteria": criteria_path,
        "results": [results_path_1, results_path_2]
    }


def single_family_search_2(address, baths, rooms, sqft_to):
    # Search
    select_search(search=0)

    escape()
    # Filter
    criteria_path = single_family_filter(address, baths, rooms, sqft_to, search_type=4, folder="single_family_2")
    # Results
    results_path = results_family_search(address, index=2, display_mode='display_ac/p/a_review')

    return {
        "criteria": criteria_path,
        "results": results_path
    }


def res_income_search_2(address, baths, rooms, sqft_to):
    # Search
    select_search(search=1)
    escape()
    # Filter
    criteria_path = single_family_filter(address, baths, rooms, sqft_to, miles=1, search_type=5, folder="res_income_2")
    results_path = results_res_income(address, index=2, display_mode='display_ac/p/a_review')

    return {
        "criteria": criteria_path,
        "results": results_path
    }


# Filters
def single_family_filter(address, baths, rooms, sqft_to, miles=0.5, search_type=1, folder="single_family"):
    # RE1/RE2 Single Family/Condo Filter

    # Change within options
    XPATHS[
        "select_within_option_0.5"] = f'//select[@id="{IDS[f"select_within_{search_type - 3 if search_type > 3 else search_type}"]}"]' \
                                      f'/option[@value="0.80467200"]'
    XPATHS[
        "select_within_option_1"] = f'//select[@id="{IDS[f"select_within_{search_type - 3 if search_type > 3 else search_type}"]}"]/option[@value="1.60934400"]'

    escape()
    # Within open
    select_within = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (By.ID, IDS[f'select_within_{search_type - 3 if search_type > 3 else search_type}'])
        )
    )
    select_within.click()

    escape()
    # Select option
    select_within_option = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, XPATHS[f"select_within_option_{miles}"])))

    # Within close
    select_within_option.click()

    # Address input
    address_input = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (By.ID, IDS[f'address_input_{search_type - 3 if search_type > 3 else search_type}'])
        )
    )
    escape()
    address_input.click()
    address_input.send_keys(address)

    # Select option
    dialog_address_search = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (By.XPATH, XPATHS['dialog_address_search'])
        )
    )
    escape()
    dialog_address_search.click()

    if search_type not in [2, 5]:
        # Bedrooms
        bedrooms_input = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.ID, IDS[f'bedrooms_input_{search_type - 3 if search_type > 3 else search_type}'])
            )
        )
        if rooms:
            bedrooms_input.send_keys(f"{rooms - 1}-{rooms + 1}")

        # Baths
        baths_input = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.ID, IDS[f'baths_input_{search_type - 3 if search_type > 3 else search_type}'])
            )
        )
        if baths:
            baths_input.send_keys(f"{baths - 1}-{baths + 1}")

    # SQFT
    sqft_living_area_input = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable(
            (By.ID, IDS[f'sqft_living_area_input_{search_type - 3 if search_type > 3 else search_type}'])
        )
    )
    if sqft_to:
        sqft_living_area_input.send_keys(f"0-{sqft_to + 700}")

    # Res rental
    if search_type == 3:
        # Active uncheck
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['active_checkbox']))).click()
        # Rented check
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['rented_checkbox']))).click()
        # Expired check
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['expired_checkbox']))).click()

        rented_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['rented_input_3'])))
        rented_input.clear()
        rented_input.send_keys(f"0-365")
    elif search_type == 4 or search_type == 5:
        # Active
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['active_checkbox']))).click()
        # Pending
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['pending_checkbox']))).click()
        # Active with contract
        WebDriverWait(driver, 15) \
            .until(EC.element_to_be_clickable((By.XPATH, XPATHS['active_with_contract_checkbox']))).click()

    criteria_path = criteria_screenshot(address, folder=folder)

    # Go to Results
    results_tab = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['results_tab'])))
    results_tab.click()

    return criteria_path


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
def set_display(display_mode='display_closed_comp'):
    # Open dropdown
    display = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, IDS['display'])))
    display.click()
    # Select option
    display_option = WebDriverWait(driver, 15) \
        .until(EC.element_to_be_clickable((By.XPATH, XPATHS[display_mode])))
    display_option.click()
    # Click again to close dropdown
    sleep(1)
    # display.click()


# Results
def results_family_search(address, index=None, display_mode='display_closed_comp'):
    escape()

    set_display(display_mode)
    sleep(1)
    escape()
    order_by_sp()
    order_by_sp()

    sleep(1)
    escape()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))

    if not index:
        results_path = screenshot_of_element(results_table, 'results/single_family', address)
    else:
        results_path = screenshot_of_element(results_table, f'results/single_family_{index}', address)

    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def results_res_income(address, index=None, display_mode='display_closed_comp'):
    escape()

    set_display(display_mode)
    sleep(1)

    escape()
    order_by_sp()
    order_by_sp()
    sleep(1)
    escape()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))

    if not index:
        results_path = screenshot_of_element(results_table, 'results/res_income', address)
    else:
        results_path = screenshot_of_element(results_table, f'results/res_income_{index}', address)

    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def results_res_rental(address, display_mode="for_sale"):
    escape()
    if display_mode == "for_sale":
        set_display(display_mode="display_for_sale")
    else:
        set_display(display_mode="display_marketing_to_realtors")
    sleep(1)

    escape()
    try:
        if display_mode == "for_sale":
            order_by_distance()
        else:
            order_by_current_price()
    except Exception as e:
        print("Order error \n")
    sleep(1)
    escape()

    # Screenshot
    results_table = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['results_table'])))
    results_path = screenshot_of_element(results_table, 'results/res_rental_' + display_mode, address)
    driver.execute_script("window.scrollTo(0, 0)")
    return results_path


def top_ten_links_on_google(address):
    driver.get(GOOGLE_URL)
    search_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
    search_input.send_keys(address)
    webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()

    links = driver.find_elements_by_xpath("//div[@class='yuRUbf']/a")[:10]
    links = [link.get_attribute("href") for link in links]
    return links


def extract_county_info(address):
    search_address_on_county(address)
    select_folio_number()
    return screenshots_of_county_info(address)


def search_address_on_county(address):
    driver.get(COUNTY_URL)

    search_input_id = "search_box"
    search_input = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, search_input_id)))
    search_input.clear()
    search_input.send_keys(address)

    search_submit_id = "search_submit"
    search_submit = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, search_submit_id)))
    search_submit.click()


def select_folio_number(index=1):
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, XPATHS['folio_number_link'])))
    folios_number = driver.find_elements_by_xpath(XPATHS['folio_number_link'])
    folio_number_link = folios_number[0]

    folio_number_link.click()

def screenshots_of_county_info(address):
    subfolder = address.lower().replace(' ', '-').replace(',', '')

    driver.execute_script("document.body.style.zoom='80%'")
    sleep(1)
    property_info_id = "property_info"
    property_info = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, property_info_id)))
    location = property_info.location_once_scrolled_into_view
    size = property_info.size
    size['height'] = size['height'] * 0.8
    size['width'] = size['width'] * 0.8
    location['x'] = location['x'] * 0.8
    location['y'] = location['y'] * 10

    property_info_path = screenshot_and_crop('county/' + subfolder, location, size, "property_info")

    driver.execute_script("document.body.style.zoom='100%'")

    full_legal_description_xpath = "//div[@class='col-md-6' and position() = 2]/div[2]"
    full_legal_description = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, full_legal_description_xpath))
    )
    legal_info_path = screenshot_of_element(full_legal_description, 'county/' + subfolder, "full_legal_info")

    taxable_xpath = "//div[@class='col-md-6' and position() = 1]/div[2]"
    taxable = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, taxable_xpath)))
    taxable_info_path = screenshot_of_element(taxable, 'county/' + subfolder, "taxable")

    sales_info_xpath = "//div[@class='row tabular_data' and not(@ng-show)]/div[@class='col-md-12']"
    sales_info = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, sales_info_xpath)))
    sales_info_path = screenshot_of_element(sales_info, 'county/' + subfolder, "sales_info")

    return {
        "property_info": property_info_path,
        "legal_info": legal_info_path,
        "taxable_info": taxable_info_path,
        "sales_info": sales_info_path
    }


def mls_extraction(address, baths, rooms, sqft_to):
    login()
    sleep(1)

    select_matrix_app()

    single_family_1 = single_family_search(address, baths, rooms, sqft_to)
    printer.pprint(single_family_1)
    sleep(1)

    res_income_1 = res_income_search(address, baths, rooms, sqft_to)
    printer.pprint(res_income_1)
    sleep(1)

    res_rental = res_rental_search(address, baths, rooms, sqft_to)
    printer.pprint(res_rental)
    sleep(1)

    single_family_2 = single_family_search_2(address, baths, rooms, sqft_to)
    printer.pprint(single_family_2)
    sleep(1)

    res_income_2 = res_income_search_2(address, baths, rooms, sqft_to)
    printer.pprint(res_income_2)
    sleep(1)

    data = {
        "single_family_1": single_family_1,
        "res_income_1": res_income_1,
        "res_rental": res_rental,
        "single_family_2": single_family_2,
        "res_income_2": res_income_2
    }
    print("Data:")
    printer.pprint(data)
    return data

def extract(address, baths, rooms, sqft_to):
    # Google the address, obtain the top 10 links
    top_google_links = top_ten_links_on_google(address)
    printer.pprint(top_google_links)

    # MLS Extraction
    criterias_results_paths = mls_extraction(address, baths, rooms, sqft_to)
    printer.pprint(criterias_results_paths)

    # Extract the county info: property info, sales info, taxable info, legal info with screenshots
    county_info_paths = extract_county_info(address)

    return top_google_links, criterias_results_paths, county_info_paths

def transform(address, mls_data, county_data):
    titles = [
        address.upper(),
        "PRIMEROS 10 ENLACES DE GOOGLE",
        "CRITERIO PARA COMPS FOR SALE SINGLE FAMILY",
        "COMPS FOR SALE SINGLE FAMILY",
        "CRITERIO PARA COMPS FOR SALE MULTI FAMILY",
        "COMPS FOR SALE MULTI FAMILY",
        "CRITERIO PARA COMPS FOR RENT",
        "COMPS FOR RENT (by Distance and Display For Sale)",
        "COMPS FOR RENT (by Higher Priced and Display MarketingToRealtor)",
        "COUNTY INFO: PROPERTY INFO",
        "COUNTY INFO: FULL LEGAL DESCRIPTION",
        "COUNTY INFO: TAXABLE",
        "COUNTY INFO: SALES INFO",
        "CRITERIO PARA COMPS FOR SALE SINGLE FAMILY (ACTIVE, PENDING Y ACTIVE WITH CONTRACT)",
        "COMPS FOR SALE SINGLE FAMILY (ACTIVE, PENDING Y ACTIVE WITH CONTRACT)",
        "CRITERIO PARA COMPS FOR SALE MULTI FAMILY (ACTIVE, PENDING Y ACTIVE WITH CONTRACT)",
        "COMPS FOR SALE MULTI FAMILY (ACTIVE, PENDING Y ACTIVE WITH CONTRACT)",
    ]
    slides_data = []
    slide_index = 2

    slides_data.append({'image_path': None, 'slide_index': 0, 'title': str(address).upper()})
    slides_data.append({'image_path': None, 'slide_index': 1, 'title': 'PRIMEROS 10 ENLACES DE GOOGLE'}, )
    county_keys = list(county_data.keys())
    mls_keys = list(mls_data.keys())

    similar_keys = mls_keys[-2:]
    mls_keys = mls_keys[:3]

    slide_index = 2
    for i in range(2, 16):
        if slide_index > 16:
            break

        if slide_index in [9, 10, 11, 12]:
            for county_key in county_keys:
                slide = {
                    "title": titles[slide_index],
                    "image_path": county_data[county_key],
                    "slide_index": slide_index
                }
                slides_data.append(slide)
                slide_index += 1
        else:
            keys_for_mls = similar_keys if slide_index > 9 else mls_keys

            for mls_key in keys_for_mls:
                for image_key in ['criteria', 'results']:
                    if mls_key == 'res_rental' and image_key == 'results':
                        for results_images in mls_data[mls_key][image_key]:
                            slide = {
                                "title": titles[slide_index],
                                "image_path": results_images,
                                "slide_index": slide_index
                            }
                            slides_data.append(slide)
                            slide_index += 1
                    else:
                        slide = {
                            "title": titles[slide_index],
                            "image_path": mls_data[mls_key][image_key],
                            "slide_index": slide_index
                        }
                        slides_data.append(slide)
                        slide_index += 1

                if slide_index == 9:
                    break

    printer.pprint(slides_data)

    return slides_data

def load(address, google_links, slides_data):
    address_format = address.lower().replace(' ', '-').replace(',', '')

    prs = Presentation()
    i = 0
    for slide_data in slides_data:

        # Set layout
        if slide_data["slide_index"] == 0:
            layout = prs.slide_layouts[0]
        elif slide_data["slide_index"] == 1:
            layout = prs.slide_layouts[1]
        else:
            layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(layout)

        # Title for each slide
        title = slide.shapes.title
        title.text = slide_data['title']

        # Slide #1
        if slide_data["slide_index"] == 0:
            subtitle = slide.placeholders[1]
            subtitle.text = "MLS Matrix Bot"
        # Google Links Slide
        elif slide_data["slide_index"] == 1:
            body_shape = slide.shapes.placeholders[1]

            tf = body_shape.text_frame
            tf.text = 'Top 10:'
            # Google Links
            for link in google_links:
                p = tf.add_paragraph()
                p.text = link
                p.size = Pt(6)
                p.level = 1
                p.font.size = Pt(15)

        # Insert image if have a image path
        if slide_data["image_path"]:
            left = Inches(0.1)
            top = Inches(2) if len(slide_data['title']) > 40 else Inches(1.8)

            if slide_data["slide_index"] in [9, 10, 11]:
                height = Inches(5.5)
                width = None
                if slide_data['slide_index'] == 9:
                    top = Inches(1.5)
                    left = Inches(3)
                    width = Inches(3)
                    height = None
            else:
                width = Inches(9.9)
                height = None
            image_path = slide_data["image_path"]

            if os.path.isfile(image_path):
                pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)

        i += 1

    # Save
    prs.save(f'{address_format}.pptx')


if __name__ == "__main__":

    address = "416 SW 24th"

    # Create if not exists this folders
    verify_folder_exists('results')
    verify_folder_exists('criteria')
    verify_folder_exists('county')
    # Extract
    google_links, mls_data, county_data = extract(address, 2, 2, 500)
    # Transform
    slides_data = transform(address=address, mls_data=mls_data, county_data=county_data)
    # Load
    load(address=address, google_links=google_links, slides_data=slides_data)
