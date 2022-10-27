import csv
from multiprocessing import AuthenticationError
import time
import traceback


import keyring
from playwright.sync_api import sync_playwright
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSlot
import chromium_utils



class NewScrapeThread(QThread):

    browser_ready_signal = pyqtSignal(bool)
    progress_signal = pyqtSignal(int)
    result_signal = pyqtSignal(set)
    error_signal = pyqtSignal(str)



    def __init__(self):
        self.browser_ready = False
        self.tasks = list()
        super().__init__()

    def run(self):
        try:
            chromepath = chromium_utils.check_browser_install()

            with sync_playwright() as p:
        
                print("Starting browser backend")

                browser = p.chromium.launch(headless=True, executable_path=chromepath)
                context = browser.new_context()
                # this way, other functions of scrape.py have an easy ref to page
                self.page = context.new_page()
                self.progress_signal.emit(20)

                self.page.goto('http://launchpad.support.sap.com')

                # self.progress_signal.emit(25)

                # username input
                uname_box = self.page.locator("#j_username")
                uname_box.click()
                uname_box.fill(keyring.get_password("system", "launchpad_username"))
                uname_box.press("Enter")

                self.progress_signal.emit(40)

                # password input
                password_box = self.page.locator("#j_password")
                password_box.click()
                password_box.fill(keyring.get_password("system", "launchpad_password"))
                password_box.press("Enter")

                # on auth success, this would be the next url
                if self.page.url[-14:] != '?redirect=true':
                    raise AuthenticationError(f"\n\nAccount username/Password is probably wrong. Expected URL ending with '?redirect=true'\nGot URL='{self.page.url}'")

                self.progress_signal.emit(43)

                print("Accessing Notes page")
                # go to Notes url (components pre-filled)
                req_url = f"https://launchpad.support.sap.com/#/mynotes?tab=Search"
                self.page.goto(req_url)


                self.progress_signal.emit(45)

                print("Waiting for site load")
                self.page.wait_for_load_state("networkidle")

                self.progress_signal.emit(85)

                self.browser_ready = True
                self.browser_ready_signal.emit(self.browser_ready)
                print("Browser is ready.\n")
                
                # i guess this is an event loop now but no way is this how you're supposed to do it
                self.check_tasks()

        except Exception as e:
            self.error_signal. emit(traceback.format_exc())
    
    @pyqtSlot(str, str)
    def add_task(self, formatted_date, system):
        self.tasks.append((formatted_date, system))
    
    def check_tasks(self):
        while True:
            if len(self.tasks) != 0:
                print(f"there are {len(self.tasks)} tasks")
                formatted_date, system = self.tasks.pop(0)
                self.run_task(formatted_date, system)
            time.sleep(1)

    def run_task(self, formatted_date, system_readable):

        self.progress_signal.emit(20)

        if   system_readable == 'DE & CH & AT':
            system = ['PY-DE*', 'PY-CH*', 'PY-AT*', 'PA-PA-DE*', 'PA-PA-XX*']
        elif system_readable == 'ESS':
            system = ['PT-RC-UI-XS*']
        elif system_readable == 'Reisemanagement':
            system = ['FI-TV*']
        elif system_readable == 'Successfactors':
            system = ['BC-ESI-WS-ABA*']

        datebox = self.page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]")
        datebox.click()
        # according to doc, this is actually how you're supposed to do it
        datebox.fill("")
        datebox.fill(formatted_date)

        self.progress_signal.emit(40)

        components_box = self.page.get_by_role("textbox", name="Komponenten (exakt)")
        components_box.click()

        components_box.press("Control+a")

        for comp in system:
            components_box.fill(comp)
            components_box.press("Enter")

        self.page.locator("button:has-text(\"Start\")").click()
        self.progress_signal.emit(60)

        with self.page.expect_download() as download_info:
            self.page.locator("text=Liste als CSV-Datei exportierenListe als CSV-Datei exportieren").click()
        
        self.progress_signal.emit(90)
        notes_from_sap = read_csv(download_info.value.path())



        self.progress_signal.emit(100)
        print(f"emitting for sys {system_readable}")
        self.result_signal.emit(notes_from_sap)





# class ScrapeThread(QThread):

#     progress_signal = pyqtSignal(int)
#     result_signal = pyqtSignal(set)
#     error_signal = pyqtSignal(str)

#     def __init__(self, formatted_date, system_readable):
#         super().__init__()
#         self.formatted_date = formatted_date
#         self.system_readable = system_readable

#     def run(self):
#         try:
#             chromepath = chromium_utils.check_browser_install()
#             self.progress_signal.emit(10)

#             if self.system_readable == 'DE & CH & AT':
#                 system = "'PY-DE*'~'PY-CH*'~'PY-AT*'~'PA-PA-DE*'~'PA-PA-XX*'"
#             elif self.system_readable == 'ESS':
#                 system = "'PT-RC-UI-XS*'"
#             elif self.system_readable == 'Reisemanagement':
#                 system = "'FI-TV*'"
#             elif self.system_readable == 'Successfactors':
#                 system = "'BC-ESI-WS-ABA*'"




#             with sync_playwright() as p:
        
#                 print("Starting browser backend")

#                 browser = p.chromium.launch(headless=True, executable_path=chromepath)
#                 context = browser.new_context()
#                 page = context.new_page()
#                 self.progress_signal.emit(20)

#                 print("Logging in")
#                 page.goto('http://launchpad.support.sap.com')

#                 self.progress_signal.emit(25)

#                 # username input
#                 uname_box = page.locator("#j_username")
#                 uname_box.click()
#                 uname_box.fill(keyring.get_password("system", "launchpad_username"))
#                 uname_box.press("Enter")

#                 self.progress_signal.emit(40)

#                 # password input
#                 password_box = page.locator("#j_password")
#                 password_box.click()
#                 password_box.fill(keyring.get_password("system", "launchpad_password"))
#                 password_box.press("Enter")

#                 # on auth success, this would be the next url
#                 if page.url[-14:] != '?redirect=true':
#                     raise AuthenticationError(f"\n\nAccount username/Password is probably wrong. Expected URL ending with '?redirect=true'\nGot URL='{page.url}'")

#                 self.progress_signal.emit(43)

#                 print("Accessing Notes page")
#                 # go to Notes url (components pre-filled)
#                 req_url = f"https://launchpad.support.sap.com/#/mynotes?tab=Search&sortBy=Relevance&filters=themk%25253Aeq~{system}%25252BreleaseStatus%25253Aeq~'NotRestricted'%25252BsecurityPatchDay%25253Aeq~'NotRestricted'%25252BfuzzyThreshold%25253Aeq~'0.9'"
#                 page.goto(req_url)


#                 self.progress_signal.emit(45)

#                 print("Waiting for site load")
#                 page.wait_for_load_state("networkidle")

#                 self.progress_signal.emit(85)
#                 # enter date since SAP is too incompetent to correctly parse it from url
#                 # filling the "date box" with a dummy string and wait until it is empty
#                 # since the sites' JS is cool and clears it while init'ing the page
#                 datebox = page.locator("[placeholder=\"MMM d\\, y - MMM d\\, y\"]")
#                 datebox.click()
#                 datebox.fill(self.formatted_date)

#                 self.progress_signal.emit(90)
                
#                 # print("Working around SAP js pain")
#                 # tries = 20
#                 # while datebox_dummy.input_value() == "dummy" and tries != 0:
#                 #     tries -= 1
#                 #     # waiting 50ms, then checking again if datebox has been reset
#                 #     page.wait_for_timeout(50)
#                 #     continue

#                 # clicking components to make sure js processing is done
#                 page.get_by_role("textbox", name="Komponenten (exakt)").click()
                

#                 print("Sending form request")
#                 page.locator("button:has-text(\"Start\")").click()

#                 self.progress_signal.emit(90)


#                 print("Sending CSV request")
#                 # request csv
#                 with page.expect_download() as download_info:
#                     page.locator("text=Liste als CSV-Datei exportierenListe als CSV-Datei exportieren").click()
                
#                 self.progress_signal.emit(95)
#                 notes_from_sap = read_csv(download_info.value.path())

#                 self.progress_signal.emit(100)
#                 self.result_signal.emit(notes_from_sap)

#                 print("Done.\n")
#                 page.close()
#                 context.close()
#                 browser.close()

#         except Exception:
#             self.error_signal.emit(traceback.format_exc())



def read_csv(csv_path):
    # get note numbers from SAP CSV, save to set
    # csv_path = Path.home().joinpath('Downloads/data.csv')
    with open(csv_path, newline='') as csvfile:
        note_numbers = set()
        csv_reader = csv.reader(csvfile, delimiter=';', quotechar='|')
        
        # skipping the first line
        next(csv_reader)

        for row in csv_reader:
            note_numbers.add(int(row[1]))
    
    return note_numbers
