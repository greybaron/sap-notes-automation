import os
import shutil
import stat
import sys
import traceback
from pathlib import Path

import keyring
import requests
from playwright._impl._api_types import TimeoutError
from playwright.sync_api import sync_playwright
from PyQt6.QtCore import QThread, pyqtSignal


def check_browser_install():
    if sys.platform == "win32":
        chromepath = Path.home().joinpath("sap-automation-chromium/chrome.exe")
    elif sys.platform == "darwin":
        chromepath = Path.home().joinpath("sap-automation-chromium/Chromium.app/Contents/MacOS/Chromium")
    else:
        raise NotImplementedError(f"not running windows or macos, sys {sys.platform} not supported")
    
    if not chromepath.exists:
        download_chromium(('win' if sys.platform == 'win32' else 'mac'))
    
    return chromepath


        


def download_chromium(platform):
    print(f"downloading Chromium for {'Windows' if platform == 'win' else 'macOS'}")

    path = Path.home().joinpath("sap-automation-chromium/")

    try:
        shutil.rmtree(path)
    except FileNotFoundError:
        pass

    os.makedirs(path)


    zip_path = path.joinpath("chromium.zip")

    zip_response = requests.get(f"https://github.com/greybaron/sap-notes-automation/raw/main/chromium/chromium-{platform}.zip")

    # this is dumb as the zip will be written to ram first and not streamed in chunks to storage
    zip_path.write_bytes(zip_response.content)


    # shutil.make_archive(Path.home().joinpath("sap-automation-chromium/chrome_mac"), "zip", Path.home().joinpath("sap-automation-chromium/Chromium.app"))
    shutil.unpack_archive(zip_path, path)

    print("Download finished")

    os.remove(zip_path)

    if platform == "mac":
        os.chmod(path.joinpath("Chromium.app/Contents/MacOS/chromium"), stat.S_IEXEC)



class ScrapeThread(QThread):

    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()
    error_signal = pyqtSignal(str)

    def __init__(self, formatted_date, system_readable):
        super().__init__()
        self.formatted_date = formatted_date
        self.system_readable = system_readable

    def run(self):
        try:
            chromepath = check_browser_install()
            self.progress_signal.emit(10)

            if self.system_readable == 'DE & CH & AT':
                system = "'PY-DE*'~'PY-CH*'~'PY-AT*'~'PA-PA-DE*'~'PA-PA-XX*'"
            elif self.system_readable == 'ESS':
                system = "'PT-RC-UI-XS*'"
            elif self.system_readable == 'Reisemanagement':
                system = "'FI-TV*'"
            elif self.system_readable == 'Successfactors':
                system = "'BC-ESI-WS-ABA*'"




            with sync_playwright() as p:
        
                print("Starting browser backend")

                browser = p.chromium.launch(headless=True, executable_path=chromepath)
                context = browser.new_context()
                page = context.new_page()
                self.progress_signal.emit(20)

                print("Logging in")
                page.goto('http://launchpad.support.sap.com')

                # username input
                page.locator("#j_username").fill(keyring.get_password("system", "launchpad_username"))
                page.locator("#j_username").press("Enter")

                # password input
                page.locator("#j_password").fill(keyring.get_password("system", "launchpad_password"))
                page.locator("#j_password").press("Enter")

                try:
                    page.wait_for_url("https://launchpad.support.sap.com/*", timeout=10000)
                except TimeoutError as e:
                    raise PermissionError("Redirect from login page failed. Username/Password is likely incorrect.") from e
                self.progress_signal.emit(25)

                print("Accessing Notes page")
                # go to Notes url (components pre-filled)
                req_url = f"https://launchpad.support.sap.com/#/mynotes?tab=Search&sortBy=Relevance&filters=themk%25253Aeq~{system}%25252BreleaseStatus%25253Aeq~'NotRestricted'%25252BsecurityPatchDay%25253Aeq~'NotRestricted'%25252BfuzzyThreshold%25253Aeq~'0.9'"
                page.goto(req_url)
                self.progress_signal.emit(40)

                print("Waiting for date input box")
                # enter date since SAP is too incompetent to correctly parse it from url
                # filling the "date box" with a dummy string and wait until it is empty
                # since the sites' JS is cool and clears it while init'ing the page
                datebox_dummy = page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]")
                datebox_dummy.fill("dummy")
                self.progress_signal.emit(60)
                
                print("Working around SAP js pain")
                tries = 20
                while datebox_dummy.input_value() == "dummy" and tries != 0:
                    tries -= 1
                    # waiting 50ms, then checking again if datebox has been reset
                    page.wait_for_timeout(50)
                    continue

                page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]").fill(self.formatted_date)
                self.progress_signal.emit(70)


                print("Sending form request")
                page.locator("button:has-text(\"Start\")").click()
                self.progress_signal.emit(80)


                print("Sent CSV request")
                # request csv
                with page.expect_download() as download_info:
                    page.locator("text=Liste als CSV-Datei exportierenListe als CSV-Datei exportieren").click()
                    self.progress_signal.emit(90)
                
                self.progress_signal.emit(100)
                shutil.copy(download_info.value.path(), Path.home().joinpath("Downloads/data.csv"))
                self.finished_signal.emit()
                print("Downloaded successfully")
                

                page.close()
                context.close()
                browser.close()

        except Exception:
            self.error_signal.emit(traceback.format_exc())
