import asyncio
import os
import shutil
import traceback
import webbrowser
from datetime import date, timedelta
from pathlib import Path
from random import randint

import keyring
from playwright.async_api import async_playwright
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (QApplication, QComboBox, QDialog, QHBoxLayout,
                             QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
                             QProgressDialog, QPushButton, QScrollArea,
                             QVBoxLayout, QWidget)
from qasync import QEventLoop, asyncSlot

import chromium_utils
from analysis import analysis
from scrape import ScrapeThread


def main():
    app = QApplication([])
    loop = QEventLoop(app)
    asyncio.set_event_loop(loop)

    mw = mainWindow()
    mw.show()

    with loop:
        loop.run_forever()


# simple popup window, only needs to be instantiated with title+text
class alert(QMessageBox):
    
    def __init__(self, title, text):
        super().__init__()

        self.setWindowTitle(title) 
        self.setText(text)

        self.show()


class exception_viewer(QPlainTextEdit):
    def __init__(self, title, text):
        super().__init__()

        self.setWindowTitle(title)
        self.setMinimumSize(800, 600)
        self.setReadOnly(True)
        self.setPlainText(text)

        self.show()



class accountSetupWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Launchpad-Account")
        self.setFixedSize(300, 100)
        self.main_layout = QVBoxLayout()

        self.unameInput = QLineEdit()
        self.unameInput.setPlaceholderText("E-Mail")

        self.pwInput = QLineEdit()
        self.pwInput.setPlaceholderText("Passwort")
        self.pwInput.setEchoMode(QLineEdit.EchoMode.Password)

        self.saveButton = QPushButton("Speichern")
        self.saveButton.clicked.connect(self.saveChanges)


        self.main_layout.addWidget(self.unameInput)
        self.main_layout.addWidget(self.pwInput)
        self.main_layout.addWidget(self.saveButton)
        self.setLayout(self.main_layout)

    def saveChanges(self):
        uname = self.unameInput.text()
        pw = self.pwInput.text()

        if uname != "" and pw != "":
            keyring.set_password("system", "launchpad_username", uname)
            keyring.set_password("system", "launchpad_password", pw)

            self.close()


        





class mainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # window bastelei
        self.setWindowTitle("SAP Hinweise")
        self.setFixedSize(360, 200)

        self.main_layout = QVBoxLayout()
        self.buttonLine0 = QHBoxLayout()
        self.buttonLine1 = QHBoxLayout()
        self.buttonLine2 = QHBoxLayout()


        self.weekSelectorLabel = QLabel("KW auswählen:")
        self.main_layout.addWidget(self.weekSelectorLabel)

        self.weekSelector = QComboBox()
        self.weekSelector.setFont(QFont('Arial', 11))
        self.main_layout.addWidget(self.weekSelector)


        self.main_layout.addStretch()


        # self.prepareOutputButton = QPushButton("PDF-Verzeichnis vorbereiten...")
        self.prepareOutputButton = QPushButton("2% chance")
        self.prepareOutputButton.clicked.connect(self.confirm_prepare_output_dir)

        self.AccSetupButton = QPushButton("Launchpad-Account...")
        self.AccSetupButton.clicked.connect(self.startAccountSetup)

        self.buttonLine0.addWidget(self.prepareOutputButton)
        self.buttonLine0.addWidget(self.AccSetupButton)

        self.main_layout.addLayout(self.buttonLine0)

        self.main_layout.addStretch()

        self.DECHATButton = QPushButton("DE && CH && AT")
        self.DECHATButton.clicked.connect(lambda: self.start_processing(self.possibleWeekChoices, self.weekSelector, "DE & CH & AT"))

        self.ReisemanagementButton = QPushButton("Reisemanagement")
        self.ReisemanagementButton.clicked.connect(lambda: self.start_processing(self.possibleWeekChoices, self.weekSelector, "Reisemanagement"))

        self.buttonLine1.addWidget(self.DECHATButton)
        self.buttonLine1.addWidget(self.ReisemanagementButton)

        self.main_layout.addLayout(self.buttonLine1)

        self.ESSButton = QPushButton("ESS")
        self.ESSButton.clicked.connect(lambda: self.start_processing(self.possibleWeekChoices, self.weekSelector, "ESS"))

        self.SuccButton = QPushButton("SuccessFactors")
        self.SuccButton.clicked.connect(lambda: self.start_processing(self.possibleWeekChoices, self.weekSelector, "Successfactors"))

        self.buttonLine2.addWidget(self.ESSButton)
        self.buttonLine2.addWidget(self.SuccButton)



        self.main_layout.addLayout(self.buttonLine2)
    
        self.setLayout(self.main_layout)

        self.loadWeekSelectorContent()


    def loadWeekSelectorContent(self):

        # filling the KW selector with 20 weeks (beginning at last week)

        monate = ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez."]

        self.possibleWeekChoices = []

        for i in range(100):
            wasLetzteWoche = date.today() - timedelta(weeks = 1+i)
            kw = wasLetzteWoche.isocalendar().week
            year = wasLetzteWoche.isocalendar().year

            # those -1s because arrays start at a number between 1 and stromboli
            start_formatted_date = monate[(date.fromisocalendar(year, kw, 1).month)-1] + " " + str(date.fromisocalendar(year, kw, 1).day) + ", " + str(date.fromisocalendar(year, kw, 1).year)
            end_formatted_date = monate[(date.fromisocalendar(year, kw, 7).month)-1] + " " + str(date.fromisocalendar(year, kw, 7).day) + ", " + str(date.fromisocalendar(year, kw, 7).year)

            formatted_date = f"{start_formatted_date} - {end_formatted_date}"
            

            self.possibleWeekChoices.append({
                "formatted_date": formatted_date,
                "kw": kw,
                "year": year
                })
            self.weekSelector.addItem(f"KW {kw}: {formatted_date}")

    def prepare_output_dir(self, output_dir):
        selected_date = self.possibleWeekChoices[self.weekSelector.currentIndex()]
        year = selected_date['year']
        kw = selected_date['kw']

        try:
            shutil.rmtree(output_dir)
        except FileNotFoundError:
            pass

        os.mkdir(output_dir)


        for name in ['DE_CH_AT', 'ESS', 'Reisemanagement', 'Successfactors']:
            open(Path.joinpath(output_dir, f"SAP Hinweise {name} KW {kw}_{year}.pdf"), 'w').close()


    def confirm_prepare_output_dir(self):
        if randint(0, 49) == 0:
            webbrowser.open("https://pbs.twimg.com/media/FYfy7mGUIAEBMVZ?format=jpg&name=small")
    

    def startAccountSetup(self):
        self.accountSetupWindow = accountSetupWindow()
        self.accountSetupWindow.show()


    # has to be declared before use
    def startAnalysis(self, system, xlsx_india, notes_from_sap):
        try:
            results = analysis(system, xlsx_india, notes_from_sap)

            if len(results) == 0:
                self.a = alert(system, "Keine fehlenden Hinweise")
            else:
                self.results_window = results_window(system, results)
                self.results_window.show()
            
        except Exception:
            self.excv = exception_viewer("Analysis failed", traceback.format_exc())
            return



    def start_processing(self, possibleWeekChoices, weekSelector, system):
        
        selection_data = possibleWeekChoices[weekSelector.currentIndex()]

        output_dir = Path.home().joinpath(f"Downloads/Ergebnis SAP Hinweise")

        xlsx_india_path = Path.home().joinpath(f"Downloads/SAP-Notes Week {selection_data['kw']} -{selection_data['year']}.xlsx")
        xlsx_target_path = output_dir.joinpath(f"SAP Hinweise KW {selection_data['kw']}_{selection_data['year']}.xlsx")


        if keyring.get_password("system", "launchpad_username") is None or keyring.get_password("system", "launchpad_password") is None:
            self.a = alert("Keine Logindaten vorhanden", "Zuerst Launchpad-Account festlegen")
            return


        if not xlsx_target_path.is_file():
            if not xlsx_india_path.is_file():
                # xlsx wasn't found at either location
                self.a = alert("XLSX fehlt", f'"{xlsx_india_path}" existiert nicht')
                return
            else:
                # xlsx exists @ india_path → move to target_path (and rename)
                self.prepare_output_dir(output_dir)
                shutil.move(xlsx_india_path, xlsx_target_path)
                
                webbrowser.open(f"file://{output_dir}")
                # self.a = alert("XLSX wurde verschoben", f'XLSX ist jetzt hier: "{xlsx_target_path}"')

        try:
            delete_data_csv()
        except Exception:
            self.excv = exception_viewer("Deleting temp data failed", traceback.format_exc())
            return
            
        self.progressView = QProgressDialog()
        self.progressView.setFixedWidth(250)
        self.progressView.setWindowTitle("Hinweise werden abgerufen")
        self.progressView.show()

        self.p_thread = ScrapeThread(selection_data['formatted_date'], system)
        self.p_thread.progress_signal.connect(self.progressView.setValue)
        self.p_thread.result_signal.connect(lambda notes_from_sap: self.startAnalysis(system, xlsx_target_path, notes_from_sap))
        # self.p_thread.result_signal.connect(self.startAnalysis)
        self.p_thread.error_signal.connect(self.show_scraping_error)
        self.p_thread.start()

    def show_scraping_error(self, error):
        self.progressView.close()
        self.excv = exception_viewer("Scraping failed", error)
        
    



class results_window(QDialog):
    def __init__(self, system, results):
        # importing to instance so that playwright note opener has access
        self.results = results
        self.context = None

        super().__init__()
        self.setWindowTitle(f"{system}  —  Ergebnisse")
        # self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        self.mainLayout = QVBoxLayout()
        self.mainLayout.addWidget(QLabel("Fehlende Hinweise:"))
        

        self.scrollView = QScrollArea()
        self.scrollContent = QWidget()
        self.scrollContentLayout = QVBoxLayout()

        for notenumber in results:
            individualNoteButton = QPushButton(str(notenumber))
            # i'm not sure exactly what lambda state does, but otherwise all buttons call
            # opennote with the same argument/number, so i guess it keeps the state of notenumber inside x
            individualNoteButton.clicked.connect(lambda state, x=notenumber: self.playwright_noteviewer({x}))

            self.scrollContentLayout.addWidget(individualNoteButton)    

        self.scrollContent.setLayout(self.scrollContentLayout)

        self.scrollView.setWidgetResizable(True)
        self.scrollView.setWidget(self.scrollContent)

        self.mainLayout.addWidget(self.scrollView)
        
        self.openAllNotesButton = QPushButton("Alle Notes anzeigen")
        self.openAllNotesButton.clicked.connect(lambda: self.playwright_noteviewer(self.results))

        self.buttonRow = QHBoxLayout()
        self.buttonRow.addWidget(self.openAllNotesButton)
        self.mainLayout.addLayout(self.buttonRow)
        self.setLayout(self.mainLayout)

        self.openAllNotesButton.setFocus()
        self.openAllNotesButton.setDefault(True)

    @asyncSlot()
    async def playwright_noteviewer(self, notes):
        async with async_playwright() as p:
            # browser context instance is in window instance,
            # so that it can be reused for multiple calls of playwright_noteviewer from ui
            
            # first checking if context has been set up; if not, run login once
            if self.context is None:
                chromepath = chromium_utils.check_browser_install()
                browser = await p.chromium.launch(headless=False, executable_path=chromepath)
                self.context = await browser.new_context()
                login_page = await self.context.new_page()

                await login_page.goto('http://launchpad.support.sap.com')

                # username input
                await login_page.locator("#j_username").fill(keyring.get_password("system", "launchpad_username"))
                await login_page.locator("#j_username").press("Enter")

                # password input
                await login_page.locator("#j_password").fill(keyring.get_password("system", "launchpad_password"))
                await login_page.locator("#j_password").press("Enter")

                await login_page.close()


            # opening all others in new tabs
            await asyncio.gather(*[self.open_note(self.context, note) for note in notes])


    async def open_note(self, context, note):
        page = await context.new_page()
        await page.goto(f"https://launchpad.support.sap.com/#/notes/{note}")
        # wait up to 16.667 minutes :)
        await page.wait_for_timeout(1000000)



def delete_data_csv():
    # Exception Handling is now in GUI
    csv_path = Path.home().joinpath("Downloads/data.csv")

    if csv_path.exists():
        os.remove(csv_path)




if __name__ == "__main__":
    main()
