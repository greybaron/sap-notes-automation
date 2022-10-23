import os
import time
from pathlib import Path

import chromedriver_autoinstaller
import keyring
from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, ElementNotInteractableException

# from selenium.common.exceptions import StaleElementReferenceException

def scrape(formatted_date, system_readable):
    if system_readable == 'DE & CH & AT':
        system = "'PY-DE*'~'PY-CH*'~'PY-AT*'~'PA-PA-DE*'~'PA-PA-XX*'"
    elif system_readable == 'ESS':
        system = "'PT-RC-UI-XS*'"
    elif system_readable == 'Reisemanagement':
        system = "'FI-TV*'"
    elif system_readable == 'Successfactors':
        system = "'BC-ESI-WS-ABA*'"

    # ensuring chromedriver path exists
    try:
        os.makedirs(Path.home().joinpath("saphinweise-chromedriver"))
    except FileExistsError:
        pass

    
    # installing driver
    chromedriver_path = chromedriver_autoinstaller.install(path=Path.home().joinpath("saphinweise-chromedriver"))
    if chromedriver_path is None:
        raise EnvironmentError("Failed to download chromedriver - is Chrome (x86_64) installed?")

    # driver options
    opshuns = Options()
    # headless breaks downloads and selenium is annoying
    # opshuns.add_argument("--headless")
    
    # instantiating driver
    driver = webdriver.Chrome(executable_path=chromedriver_path, options=opshuns)

    # stays valid for the whole process (driver will wait >up to< 30 secs for any element)
    driver.implicitly_wait(30)

    # login page
    driver.get('https://launchpad.support.sap.com/')

    # username
    search_box = driver.find_element("name", "j_username")
    search_box.send_keys(keyring.get_password("system", "launchpad_username"))
    search_box.submit()

    # password
    search_box = driver.find_element("name", "j_password")
    search_box.send_keys(keyring.get_password("system", "launchpad_password"))
    search_box.submit()

    # go to Notes url (components pre-filled)
    driver.get(f"https://launchpad.support.sap.com/#/mynotes?tab=Search&sortBy=Relevance&filters=themk%25253Aeq~{system}%25252BreleaseStatus%25253Aeq~'NotRestricted'%25252BsecurityPatchDay%25253Aeq~'NotRestricted'%25252BfuzzyThreshold%25253Aeq~'0.9'")



    # enter date since SAP is too incompetent to correctly parse it from url
    # the sites' JS is cool and clears it while init'ing the page,
    # therefore this sweet while True loop






    while True:
        try:
            time.sleep(0.1)
            driver.find_element('id', '__xmlview1--idReleasedOnFree-inner').send_keys(formatted_date)
            break
        except StaleElementReferenceException:
            continue
        except ElementNotInteractableException:
            continue



    submit_button = driver.find_element('id', '__xmlview1--filterBar-btnGo-content')
    submit_button.click()


    csv_button = driver.find_element('id', '__button12-inner')
    csv_button.click()



    while not os.path.exists(Path.home().joinpath("Downloads/data.csv")):
        time.sleep(0.1)

    driver.quit()
