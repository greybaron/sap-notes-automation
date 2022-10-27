import csv
from multiprocessing import AuthenticationError
import time
import traceback


import keyring
from playwright.sync_api import sync_playwright
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSlot
import chromium_utils



class NewScrapeThread(QThread):

    progress_signal = pyqtSignal(int)
    result_signal = pyqtSignal(set)
    error_signal = pyqtSignal(str)



    def __init__(self):
        self.tasks = list()

        super().__init__()

    def run(self):
        try:
            chromepath = chromium_utils.check_browser_install()

            with sync_playwright() as p:
                print("You don't need to wait for this to finish\n")
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
                
                # i guess this is an event loop now but no way is this how you're supposed to do it
                self.check_tasks()

        except Exception as e:
            self.error_signal.emit(traceback.format_exc())
    
    @pyqtSlot(str, str)
    def add_task(self, formatted_date, system):
        self.tasks.append((formatted_date, system))
    
    def check_tasks(self):
        print("Browser is ready.\n")
        while True:
            if len(self.tasks) != 0:
                formatted_date, system = self.tasks.pop(0)
                self.run_task(formatted_date, system)
            time.sleep(0.1)

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
        self.result_signal.emit(notes_from_sap)


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
