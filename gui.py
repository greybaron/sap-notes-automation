import asyncio
import sys
import os
import shutil
import time
import traceback
import webbrowser
from datetime import date, timedelta
from pathlib import Path
from random import randint

import keyring
from playwright.async_api import async_playwright
from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtGui import QFont, QTextCursor
from PyQt6.QtWidgets import (QApplication, QComboBox, QDialog, QHBoxLayout,
                             QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QTextEdit,
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

    main_window = MainWindow()
    main_window.show()

    with loop:
        loop.run_forever()


# simple popup window, only needs to be instantiated with title+text
class Alert(QMessageBox):
    
    def __init__(self, title, text):
        super().__init__()

        self.setWindowTitle(title) 
        self.setText(text)

        self.show()


class ExceptionViewer(QPlainTextEdit):
    def __init__(self, title, text):
        super().__init__()

        self.setWindowTitle(title)
        self.setMinimumSize(800, 600)
        self.setReadOnly(True)
        self.setPlainText(text)

        self.show()



class AccountSetupWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Launchpad-Account")
        self.setFixedSize(300, 100)
        self.main_layout = QVBoxLayout()

        self.uname_input = QLineEdit()
        self.uname_input.setPlaceholderText("E-Mail")

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Passwort")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        self.save_button = QPushButton("Speichern")
        self.save_button.clicked.connect(self.saveChanges)


        self.main_layout.addWidget(self.uname_input)
        self.main_layout.addWidget(self.password_input)
        self.main_layout.addWidget(self.save_button)
        self.setLayout(self.main_layout)

    def saveChanges(self):
        uname = self.uname_input.text()
        pw = self.password_input.text()

        if uname != "" and pw != "":
            keyring.set_password("system", "launchpad_username", uname)
            keyring.set_password("system", "launchpad_password", pw)

            self.close()


        
class Stream(QObject):
    newText = pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # window bastelei
        self.setWindowTitle("SAP Hinweise")


        self.main_layout = QVBoxLayout()
        self.buttonLine0 = QHBoxLayout()
        self.buttonLine1 = QHBoxLayout()
        self.buttonLine2 = QHBoxLayout()


        self.weekSelectorLabel = QLabel("KW auswählen:")
        self.main_layout.addWidget(self.weekSelectorLabel)

        self.weekSelector = QComboBox()
        self.weekSelector.setFont(QFont('Arial', 11))
        self.main_layout.addWidget(self.weekSelector)


        # self.main_layout.addStretch()
        # self.main_layout.addSpacing(10)


        # self.uselessButton = QPushButton("PDF-Verzeichnis vorbereiten...")
        self.uselessButton = QPushButton("2% chance")
        self.uselessButton.clicked.connect(self.confirm_prepare_output_dir)

        self.AccSetupButton = QPushButton("Launchpad-Account...")
        self.AccSetupButton.clicked.connect(self.startAccountSetup)

        self.buttonLine0.addWidget(self.uselessButton)
        self.buttonLine0.addWidget(self.AccSetupButton)

        self.main_layout.addLayout(self.buttonLine0)

        # self.main_layout.addStretch()
        # self.main_layout.addSpacing(15)

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

        # sys.stdout = Stream(newText=self.onUpdateText)
        # streaming starts only when stdout_viewer is visible
        self.stdout_viewer = QTextEdit()
        self.stdout_viewer.setDisabled(True)
        self.stdout_viewer.moveCursor(QTextCursor.MoveOperation.Start)
        self.stdout_viewer.ensureCursorVisible()
        self.stdout_viewer.setLineWrapColumnOrWidth(500)
        self.stdout_viewer.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.main_layout.addWidget(self.stdout_viewer)
        self.stdout_viewer.hide()


        self.setLayout(self.main_layout)
        self.loadWeekSelectorContent()

        self.base_size = self.sizeHint()
        print(self.base_size)
        self.base_size.setWidth(self.base_size.width()+70)
        print(self.base_size)

        self.setFixedSize(self.base_size)

        # self.setFixedSize(self.sizeHint().grownBy(QSize)
    
    # 'text' (~ a single line) is received from stdout and then streamed here
    def onUpdateText(self, text):
        cursor = self.stdout_viewer.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        # text is actually written here
        cursor.insertText(text)
        self.stdout_viewer.setTextCursor(cursor)
        self.stdout_viewer.ensureCursorVisible()


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
        self.AccountSetupWindow = AccountSetupWindow()
        self.AccountSetupWindow.show()

    def set_stdoutviewer_enabled(self, enabled):
        if enabled:
            self.stdout_viewer.show()
            sys.stdout = Stream(newText=self.onUpdateText)
            sys.stderr = Stream(newText=self.onUpdateText)
        else:
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
            self.stdout_viewer.hide()
        

    def start_processing(self, possibleWeekChoices, weekSelector, system):
        # setting up stdout stream
        self.set_stdoutviewer_enabled(True)

        selection_data = possibleWeekChoices[weekSelector.currentIndex()]

        output_dir = Path.home().joinpath(f"Downloads/Ergebnis SAP Hinweise")

        xlsx_india_path = Path.home().joinpath(f"Downloads/SAP-Notes Week {selection_data['kw']} -{selection_data['year']}.xlsx")
        xlsx_target_path = output_dir.joinpath(f"SAP Hinweise KW {selection_data['kw']}_{selection_data['year']}.xlsx")


        if keyring.get_password("system", "launchpad_username") is None or keyring.get_password("system", "launchpad_password") is None:
            self.a = Alert("Keine Logindaten vorhanden", "Zuerst Launchpad-Account festlegen")
            return


        if not xlsx_target_path.is_file():
            if not xlsx_india_path.is_file():
                # xlsx wasn't found at either location
                self.a = Alert("XLSX fehlt", f'"{xlsx_india_path}" existiert nicht')
                return
            else:
                # xlsx exists @ india_path → move to target_path (and rename)
                self.prepare_output_dir(output_dir)
                shutil.move(xlsx_india_path, xlsx_target_path)
                
                webbrowser.open(f"file://{output_dir}")
                # self.a = Alert("XLSX wurde verschoben", f'XLSX ist jetzt hier: "{xlsx_target_path}"')

        try:
            delete_data_csv()
        except Exception:
            self.excv = ExceptionViewer("Deleting temp data failed", traceback.format_exc())
            return
            
        self.progressView = QProgressDialog()
        self.progressView.setFixedWidth(250)
        self.progressView.setWindowTitle("Hinweise werden abgerufen")
        self.progressView.canceled.connect(self.cancel_scraping)
        self.progressView.show()

        self.p_thread = ScrapeThread(selection_data['formatted_date'], system)
        self.p_thread.progress_signal.connect(self.progressView.setValue)
        self.p_thread.result_signal.connect(lambda notes_from_sap: self.start_analysis(system, xlsx_target_path, notes_from_sap))
        # self.p_thread.result_signal.connect(self.startAnalysis)
        self.p_thread.error_signal.connect(self.show_scraping_error)
        self.p_thread.start()

        # has to be declared before use
    def start_analysis(self, system, xlsx_india, notes_from_sap):
        self.set_stdoutviewer_enabled(False)
        self.resize(self.sizeHint())

        try:
            results = analysis(system, xlsx_india, notes_from_sap)

            if len(results) == 0:
                self.a = Alert(system, "Keine fehlenden Hinweise")
            else:
                self.results_window = ResultsWindow(system, results)
                self.results_window.show()
            
        except Exception:
            self.excv = ExceptionViewer("Analysis failed", traceback.format_exc())
            return

    def show_scraping_error(self, error):
        self.progressView.close()
        self.excv = ExceptionViewer("Scraping failed", error)
    
    def cancel_scraping(self):
        self.set_stdoutviewer_enabled(False)
        print(self.sizeHint())
        self.resize(self.sizeHint())
        self.p_thread.terminate()
        time.sleep(1)
        print(self.sizeHint())
        
        
    



class ResultsWindow(QDialog):
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
