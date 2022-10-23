from importlib.resources import path
import os
from pathlib import Path
import shutil
import sys
import keyring
from playwright.sync_api import sync_playwright
import requests


def check_browser_install():
    if sys.platform == "win32":
        if os.path.exists(Path.home().joinpath("sap-automation-chromium/chrome.exe")):
            return True
        else:
            download_windows()
    elif sys.platform == "darwin":
        if os.path.exists(Path.home().joinpath("sap-automation-chromium/Chromium.app/Contents/MacOS/Chromium")):
            return True
        else:
            download_macos()
            return True
    else:
        raise NotImplementedError(f"not running windows or macos, sys {sys.platform} not supported")


def download_macos():
    print("downloading Chromium for macOS")

    path = Path.home().joinpath("sap-automation-chromium/")


    zip_path = path.joinpath("chrome-mac.zip")
 
    response = requests.get("https://github.com/greybaron/sap-notes-automation/raw/main/chromium/chrome-mac.zip")

    zip_path.write_bytes(response.content)



    # shutil.make_archive(Path.home().joinpath("sap-automation-chromium/chrome_mac"), "zip", Path.home().joinpath("sap-automation-chromium/Chromium.app"))
    shutil.unpack_archive(zip_path, path)

download_macos()
# check_browser_install()


def scrape(formatted_date, system_readable):

    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ""

    if system_readable == 'DE & CH & AT':
        system = "'PY-DE*'~'PY-CH*'~'PY-AT*'~'PA-PA-DE*'~'PA-PA-XX*'"
    elif system_readable == 'ESS':
        system = "'PT-RC-UI-XS*'"
    elif system_readable == 'Reisemanagement':
        system = "'FI-TV*'"
    elif system_readable == 'Successfactors':
        system = "'BC-ESI-WS-ABA*'"




    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False, executable_path=Path.home().joinpath("sap-automation-chromium/Chromium.app/Contents/MacOS/Chromium"))
        context = browser.new_context()
        page = context.new_page()

        page.goto('http://launchpad.support.sap.com')

        # username input
        page.locator("#j_username").fill(keyring.get_password("system", "launchpad_username"))
        page.locator("#j_username").press("Enter")

        # password input
        page.locator("#j_password").fill(keyring.get_password("system", "launchpad_password"))
        page.locator("#j_password").press("Enter")


        # go to Notes url (components pre-filled)
        req_url = f"https://launchpad.support.sap.com/#/mynotes?tab=Search&sortBy=Relevance&filters=themk%25253Aeq~{system}%25252BreleaseStatus%25253Aeq~'NotRestricted'%25252BsecurityPatchDay%25253Aeq~'NotRestricted'%25252BfuzzyThreshold%25253Aeq~'0.9'"
        page.goto(req_url)

        # enter date since SAP is too incompetent to correctly parse it from url

        # filling the "date box" with a dummy string and wait until it is empty
        # since the sites' JS is cool and clears it while init'ing the page
        datebox_dummy = page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]")
        datebox_dummy.fill("dummy")
        

        while datebox_dummy.input_value() == "dummy":
            # waiting 50ms, then checking again if datebox has been reset
            page.wait_for_timeout(50)
            continue

        page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]").fill(formatted_date)


        # page.locator("[placeholder=\"MMM d\\, y - MM  y\"]").fill("Okt. 10, 2022 - Okt. 16, 2022")

    
        # Click button:has-text("Start")
        page.locator("button:has-text(\"Start\")").click()


        # request csv
        with page.expect_download() as download_info:
            page.locator("text=Liste als CSV-Datei exportierenListe als CSV-Datei exportieren").click()

        shutil.copy(download_info.value.path(), Path.home().joinpath("Downloads/data.csv"))


        page.close()
        context.close()
        browser.close()