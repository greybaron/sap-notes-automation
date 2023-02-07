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
# from playwright.async_api import async_playwright
from PyQt6.QtCore import QObject, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QTextCursor
from PyQt6.QtWidgets import (QApplication, QComboBox, QDialog, QHBoxLayout,
                             QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QTextEdit,
                             QProgressDialog, QPushButton, QScrollArea,
                             QVBoxLayout, QWidget)
# from qasync import QEventLoop, asyncSlot

# import chromium_utils
from analysis import analysis
from scrape import NewScrapeThread


def main():
    app = QApplication([])
    # loop = QEventLoop(app)
    # asyncio.set_event_loop(loop)

    main_window = MainWindow()
    # main_window = ResultsWindow('DE & CH & AT', {2647329, 2647300, 3267658, 3267087, 3267669, 2647350, 3267030, 2647416, 3258201}, {3251933}, [{'domain': 'launchpad.support.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'mdsoucrsvirwxsbijufrtzruvjuh', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '1%2FZDQwTsWYUsYMnYzbEQzbfzCjQSJE%2BOZkqcpipYpexBnUR762Ua3%2FJguzj3HZCDh49hBLSZVY%2BtxOId%2B%2FXhndf8eQ3ogXue%2BXdp3QZgr7K1y94DcII%2BIDQ0LCmf8pAnvpjjrtYHl5wJ6jIKFM3h6rE6u%2BVTRPmHGog07HQ6VD%2FBK1ewO0s%2BO%2B5%2F0m2PnUfL%2BUXblOXBwVygRZPy2iUr8Dq%2B4Nfe%2FrE5tiKcAUG6UMv2QUC4rKfbZiuyBsqi1QYMzRW9guoubDGYsg90e6oW7%2FZ5L%2Fr3q0HlvVfeNDrm0ocEHalg%2F0%2FllAv6qHv6er8CM2I%2BO6npHlQ5nNLXenBnKg%3D%3D'}, {'domain': 'launchpad.support.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'BIGipServerdispatcher.factory.customdomain', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '!RfX+0s1izdG7eesrNj2u8iYv3vui4JNVspHm2wvYmL0O90pNX+h/ArGFg9AgIIjpBAHa90gFHC3CxQ=='}, {'domain': 'authn.hana.ondemand.com', 'expires': -1, 'httpOnly': True, 'name': 'oucrsvirwxsbijufrtzruvjuh', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'ge7D3S0MGOadEs5WV8PiOwIP%2F5PiR6X%2FCTePlyB0zCSdC6V5NUlUtSUD14EI812GwPRWTPq69I3O%2F1i1pWTrXQJJFwOp5KAiR6n5V5yTs%2BBu3wBgYQ%2FBIpsi7JLUMK83DZvitliq%2FESJ6TNCJfbow3n0Qz73w7mjU7N09BSnKOeTs5Wfz6rouuTP7hLXRctveEHTvRFpd3euI2GTkQsNkDdmarjphDBLpm6Ymbz%2BhFZpcPmOGNICMgsPhuSvarLqsswaB2zAJZtVyNL4S1%2B00cFdLX2wlk2UVJSmKFzChhbsFdHPj1IHSRW06I8gvgONykgQg4XRp11uxcMtRZekBQ%3D%3D'}, {'domain': 'authn.hana.ondemand.com', 'expires': -1, 'httpOnly': True, 'name': 'BIGipServerssoendpointssecurity.hana.ondemand.com', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '!ijy5bdHeU0UHf4hl3wqwDvMjD5TPACL74QdxaLRjz9wnPXfrypp05uToXRNEGj1gEtXnYxX+EuSbZk4='}, {'domain': 'accounts.sap.com', 'expires': 1668518572.957474, 'httpOnly': True, 'name': 'authIdentifierDataTemporary', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'AAAADOJceAq7Fa3dWwVJjznPXknmsROc9zB6oPPRZWwt1bUyxazPIdnIUY43ZqGpB1ymIgCoxg%3D%3D'}, {'domain': '.accounts.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'IDP_USER', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'NzY0OTYwMzY0NjEyNjczNzM2MTEyNGJhOWEyM2NjMDQ%3D'}, {'domain': '.accounts.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'IDP_J_COOKIE', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'S-IDP-188c7416-0aeb-4a85-909a-301689ef47a0'}, {'domain': '.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'IDP_SESSION_MARKER_accounts', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'eyJhbGciOiJFQ0RILUVTK0EyNTZLVyIsImVuYyI6IkExMjhDQkMtSFMyNTYiLCJraWQiOm51bGwsImN0eSI6IkpXVCIsImVwayI6eyJrdHkiOiJFQyIsIngiOiIwSDNEQ1BvLW12UzIzY0kwLU1jSHg3RkMweVJ4d0F2VlFMVmR4TXJxcHZjIiwieSI6Ijk4cFFMNlVnTVUxdnBqYVgtUnVvTDRxWkJNUDdMSlYwam05Ml9GaDV2Nm8iLCJjcnYiOiJQLTI1NiJ9fQ.aU6SzjohbrrPQkLMpH07RlWCud0NjCKcR_RSGzPqcAdjN-vTaQWMJQ.DwPp8LK796aj8b-6jqGtwA.D2z7D2mRSRPu0gf0dzbHPf6yYi_PiJ8lDLWX16_oErK0ZO5J88CCb2bZ0CwYiSiuEr4ve7YhMwFYeEXDZglUA91ryYDtI-2gvbHzUpNK67NaXIJ0kuAAO6xWwPirCtrsNzEBmJzQGlpiQEzScCItJ9JZE9ZNYaJgA352kCbtsOECVjapQiGjiIxOnMm0dJiq5rzoSuma_3x2BiFO9i6uBa-TGyTj0oFu_IkbrzZpnn9DPgkLu_fSFwcrJJB9nQKlF03CG1bQQgJus5rcDyUmWM0pO4r8e_QkTAWMMjQsTL82PRzF-xMh6GRDtwTG5oVDI2r8kMrgmvZsBt2TpNKKFriOVVWy3xWPmDMF1Y2cXiBV-v7PaHIlzV9y00CR2nhzFq9aEskCxTYcAH1cgzjysvufVf0aj9OpH7robHm_9n3mS4ltJY07U4BIlznjhvuLiw5YB4EMDvBdcJvmT2b-IVnM3Qmwtzg5Q7jzUtZt3RJ2bnctErCS8E2iIojkGC9wWXf1lu4MUvGncSJKeTeH2NdL6d-DldzvcZgfnrxN0HHz2FjLMyBLPbK2KTUEjwob5vK0ImRpu-fHo1NpT87XRpvGv2T7WjS31Fqul_G-9_U._nhCPk5eCx8Xruy9-MrKhQ'}, {'domain': 'accounts.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'XSRF_COOKIE', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '"QglyEyPvv70ELXEx77+9UlhvCu+/ve+/vTnvv73vv73vv70MB3F+77+977+9BAXvv706MTY2ODUxODQ5NjcyMw=="'}, {'domain': 'accounts.sap.com', 'expires': 1703078512.957828, 'httpOnly': True, 'name': 'authIdentifierData', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': 'AAAADEV1R76yN%2FZVOiBk24wnv5W2hllthJzC%2FrEjkotmxl8ysRcq77tGnnRp83c5Dq%2FV2zjNAA%3D%3D'}, {'domain': 'accounts.sap.com', 'expires': -1, 'httpOnly': True, 'name': 'JSESSIONID', 'path': '/', 'sameSite': 'None', 'secure': True, 'value': '35D7C030F7BE319D45DC2828668FA49D'}])
    main_window.show()

    app.exec()
    
    # with loop:
    #     loop.run_forever()


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
    add_scrape_task_signal = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()

        sys.stdout = Stream(newText=self.on_update_text)
        sys.stderr = Stream(newText=self.on_update_text)

        # setting up browser async
        self.NewScrapeThread = NewScrapeThread()
        self.add_scrape_task_signal.connect(self.NewScrapeThread.add_task)
        self.NewScrapeThread.start()

        # setting up browser signals
        self.NewScrapeThread.result_signal.connect(self.start_analysis)
        self.NewScrapeThread.error_signal.connect(self.show_scraping_error)
        # also getting cookies from browser
        self.NewScrapeThread.send_cookies_signal.connect(self.import_cookies)

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


        self.uselessButton = QPushButton("3% chance")
        self.uselessButton.clicked.connect(self.komplett_useless)

        self.AccSetupButton = QPushButton("Launchpad-Account...")
        self.AccSetupButton.clicked.connect(self.startAccountSetup)

        self.buttonLine0.addWidget(self.uselessButton)
        self.buttonLine0.addWidget(self.AccSetupButton)
        self.main_layout.addLayout(self.buttonLine0)

        self.main_layout.addSpacing(10)

        self.DECHATButton = QPushButton("DE && CH && AT")
        self.DECHATButton.clicked.connect(lambda: self.start_processing(self.possible_week_choices, self.weekSelector, "DE & CH & AT"))

        self.ReisemanagementButton = QPushButton("Reisemanagement")
        self.ReisemanagementButton.clicked.connect(lambda: self.start_processing(self.possible_week_choices, self.weekSelector, "Reisemanagement"))

        self.buttonLine1.addWidget(self.DECHATButton)
        self.buttonLine1.addWidget(self.ReisemanagementButton)
        self.main_layout.addLayout(self.buttonLine1)

        self.ESSButton = QPushButton("ESS")
        self.ESSButton.clicked.connect(lambda: self.start_processing(self.possible_week_choices, self.weekSelector, "ESS"))

        self.SuccButton = QPushButton("SuccessFactors")
        self.SuccButton.clicked.connect(lambda: self.start_processing(self.possible_week_choices, self.weekSelector, "Successfactors"))

        self.buttonLine2.addWidget(self.ESSButton)
        self.buttonLine2.addWidget(self.SuccButton)
        self.main_layout.addLayout(self.buttonLine2)

        # setting up stdout/stderr viewer
        self.stdout_viewer = QTextEdit()
        self.stdout_viewer.setDisabled(True)
        self.stdout_viewer.moveCursor(QTextCursor.MoveOperation.Start)
        self.stdout_viewer.ensureCursorVisible()
        self.stdout_viewer.setLineWrapColumnOrWidth(500)
        self.stdout_viewer.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.main_layout.addWidget(self.stdout_viewer)


        self.setLayout(self.main_layout)
        self.load_week_selector_content()


        # sizehint calculates the minimum size so that all contents fit
        self.size_initial = self.sizeHint()
        
        # sizehint works on macos, but is a bit cramped (buttons are sized to text and therefore vary in width)
        if sys.platform == "darwin":
            self.size_initial += QSize(70, -50)
        else:
            self.size_initial += QSize(0, -80)

        self.resize(self.size_initial)

    
    # 'text' (~ a single line) is received from stdout and then streamed here
    def on_update_text(self, text):
        cursor = self.stdout_viewer.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        # text is actually written here
        cursor.insertText(text)
        self.stdout_viewer.setTextCursor(cursor)
        self.stdout_viewer.ensureCursorVisible()

    def import_cookies(self, cookies):
        # print("GOT COOKIES")
        self.cookies = cookies


    def load_week_selector_content(self):
        # filling the KW selector with 20 weeks (beginning at last week)

        monate = ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sept.", "Okt.", "Nov.", "Dez."]

        self.possible_week_choices = []

        for i in range(100):
            was_letzte_woche = date.today() - timedelta(weeks = 1+i)
            kw = was_letzte_woche.isocalendar().week
            year = was_letzte_woche.isocalendar().year

            # those -1s because arrays start at a number between 1 and stromboli
            start_formatted_date = monate[(date.fromisocalendar(year, kw, 1).month)-1] + " " + str(date.fromisocalendar(year, kw, 1).day) + ", " + str(date.fromisocalendar(year, kw, 1).year)
            end_formatted_date = monate[(date.fromisocalendar(year, kw, 7).month)-1] + " " + str(date.fromisocalendar(year, kw, 7).day) + ", " + str(date.fromisocalendar(year, kw, 7).year)

            formatted_date = f"{start_formatted_date} - {end_formatted_date}"
            

            self.possible_week_choices.append({
                "formatted_date": formatted_date,
                "kw": kw,
                "year": year
                })
            self.weekSelector.addItem(f"KW {kw}: {formatted_date}")

    # this is only called if the xlsx is not in the directory
    # therefore, deleting the output dir (using rmtree) should fail
    def prepare_target_dir(self, output_dir):
        selected_date = self.possible_week_choices[self.weekSelector.currentIndex()]
        year = selected_date['year']

        try:
            shutil.rmtree(output_dir)
        except FileNotFoundError:
            pass

        os.mkdir(output_dir)

        for name in ['DE_CH_AT', 'ESS', 'Reisemanagement', 'Successfactors']:
            open(Path.joinpath(output_dir, f"SAP Hinweise {name} KW {kw}_{year}.pdf"), 'w').close()


    def komplett_useless(self):
        rand = randint(0, 66)
        # 2 cases out of 67 ≅ 3% chance
        if rand == 0:
            webbrowser.open("https://pbs.twimg.com/media/FYfy7mGUIAEBMVZ?format=jpg&name=small")
        elif rand == 1:
            webbrowser.open("https://i.imgflip.com/6urpor.png")
    

    def startAccountSetup(self):
        self.AccountSetupWindow = AccountSetupWindow()
        self.AccountSetupWindow.show()


    def set_UI_disabled(self, state):
        self.weekSelectorLabel.setDisabled(state)
        self.weekSelector.setDisabled(state)
        self.uselessButton.setDisabled(state)
        self.AccSetupButton.setDisabled(state)
        self.DECHATButton.setDisabled(state)
        self.ReisemanagementButton.setDisabled(state)
        self.ESSButton.setDisabled(state)
        self.SuccButton.setDisabled(state)


    def start_processing(self, possible_week_choices, weekSelector, system):

        selection_date = possible_week_choices[weekSelector.currentIndex()]

        # .xlsx is sometimes sent with single digig
        if len(str(selection_date['kw'])) == 1:
            selection_date = 0 + selection_date

        # ensuring that top level target path exists
        container_dir = Path.home().joinpath("Downloads/Ergebnis SAP Hinweise")
        try:
            os.mkdir(container_dir)
        except FileExistsError:
            pass

        target_dir = container_dir.joinpath(f"{selection_date['year']} KW{selection_date['kw']}")

        # the file naming might change in the future
        xlsx_india_path = Path.home().joinpath(f"Downloads/SAP Notes Week {selection_date['kw']} {selection_date['year']}.xlsx")

        xlsx_new_target_path = target_dir.joinpath(f"SAP Hinweise KW {selection_date['kw']}_{selection_date['year']}.xlsx")

        # recently India has sent .xlsx like this (why must they change it???)
        xlsx_new_new_target_path = target_dir.joinpath(f"SAP Notes Week {selection_date['kw']} {selection_date['year']}.xlsx")

        if keyring.get_password("system", "launchpad_username") is None or keyring.get_password("system", "launchpad_password") is None:
            self.a = Alert("Keine Logindaten vorhanden", "Zuerst Launchpad-Account festlegen")
            return



        if xlsx_new_target_path.is_file():
            # xlsx exists @ india_path → move to target_path (and rename)
            self.prepare_target_dir(target_dir)
            shutil.move(xlsx_india_path, xlsx_new_target_path)    
            webbrowser.open(f"file://{target_dir}")
        elif xlsx_new_new_target_path.is_file():
            self.prepare_target_dir(target_dir)
            shutil.move(xlsx_india_path, xlsx_new_new_target_path)    
            webbrowser.open(f"file://{target_dir}")   
        elif not xlsx_india_path.is_file():
            # xlsx wasn't found at either of the 3 locations
            self.a = Alert("XLSX fehlt", f'"{xlsx_india_path}" existiert nicht')
            return      

        
        # setting up references for result signal callback
        self.system = system
        self.xlsx_target_path = xlsx_new_target_path

        self.set_UI_disabled(True)

        self.progressView = QProgressDialog()
        self.progressView.setFixedWidth(250)
        self.progressView.setWindowTitle("Hinweise werden abgerufen")
        self.progressView.setCancelButton(None)
        self.progressView.show()

        self.NewScrapeThread.progress_signal.connect(self.progressView.setValue)

        # sending the task to NewScrapeThread queue
        self.add_scrape_task_signal.emit(selection_date['formatted_date'], system)



        # has to be declared before use
    def start_analysis(self, notes_from_sap):# system, xlsx_india, notes_from_sap):

        try:
            only_in_sap, only_in_xlsx = analysis(self.system, self.xlsx_target_path, notes_from_sap)

            if len(only_in_sap) == 0 and len(only_in_xlsx) == 0:
                self.a = Alert(self.system, "Keine fehlenden Hinweise")
            else:
                self.results_window = ResultsWindow(self.system, only_in_sap, only_in_xlsx, cookies=self.cookies)
                self.results_window.show()
          
        except Exception:
            self.excv = ExceptionViewer("Analysis failed", traceback.format_exc())
            
        self.set_UI_disabled(False)

    def show_scraping_error(self, error):
        try:
            self.progressView.close()
        except:
            pass
        
        self.excv = ExceptionViewer("Scraping failed", error)



class ResultsWindow(QDialog):
    def __init__(self, system, only_in_sap, only_in_xlsx, cookies):

        # print(f"mw = ResultsWindow({repr(system)}, {repr(only_in_sap)}, {repr(only_in_xlsx)}, {repr(cookies)})")

        # this var will store the playwright cookies, used by the notes viewer browser
        # self.cookies = cookies

        # importing to instance so that playwright note opener has access
        self.only_in_sap = only_in_sap
        self.only_in_xlsx = only_in_xlsx

        ##########c
        self.browser = None

        super().__init__()
        self.setWindowTitle(system)
        # self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)

        self.mainLayout = QVBoxLayout()
        
        
        # this isn't very nice, this should be a function or something and not basically just repeated
        if len(only_in_sap) > 0:
            self.mainLayout.addWidget(QLabel("Fehlende Hinweise:"))
            
            self.OnlySapScrollView = QScrollArea()
            self.OnlySapScrollContent = QWidget()
            self.OnlySapScrollContentLayout = QVBoxLayout()
            
            for notenumber in only_in_sap:
                individualNoteButton = QPushButton(str(notenumber))
                # i'm not sure exactly what lambda state does, but otherwise all buttons call
                # opennote with the same argument/number, so i guess it keeps the state of notenumber inside x
                # individualNoteButton.clicked.connect(lambda state, x=notenumber: self.playwright_noteviewer(cookies=cookies, notes={x}))
                individualNoteButton.clicked.connect(lambda state, x=notenumber: self.open_notes(notes={x}))

                self.OnlySapScrollContentLayout.addWidget(individualNoteButton)    

            self.OnlySapScrollContent.setLayout(self.OnlySapScrollContentLayout)

            self.OnlySapScrollView.setWidgetResizable(True)
            self.OnlySapScrollView.setWidget(self.OnlySapScrollContent)

            self.mainLayout.addWidget(self.OnlySapScrollView)
        
            self.openMissingNotesButton = QPushButton("Fehlende Notes anzeigen")
            # self.openMissingNotesButton.clicked.connect(lambda: self.playwright_noteviewer(cookies=cookies, notes=self.only_in_sap))
            self.openMissingNotesButton.clicked.connect(lambda: self.open_notes(notes=self.only_in_sap))
            self.mainLayout.addWidget(self.openMissingNotesButton)

        if len(only_in_xlsx) > 0:
            self.mainLayout.addWidget(QLabel("Fragwürdige Hinweise:\n(Nur in XLSX vorhanden)"))
            
            self.OnlyXLSXScrollView = QScrollArea()
            self.OnlyXLSXScrollContent = QWidget()
            self.OnlyXLSXScrollContentLayout = QVBoxLayout()
            
            for notenumber in only_in_xlsx:
                individualNoteButton = QPushButton(str(notenumber))
                # i'm not sure exactly what lambda state does, but otherwise all buttons call
                # opennote with the same argument/number, so i guess it keeps the state of notenumber inside x
                # individualNoteButton.clicked.connect(lambda state, x=notenumber: self.playwright_noteviewer(cookies=cookies, notes={x}))
                individualNoteButton.clicked.connect(lambda state, x=notenumber: self.open_notes(notes={x}))

                self.OnlyXLSXScrollContentLayout.addWidget(individualNoteButton)    

            self.OnlyXLSXScrollContent.setLayout(self.OnlyXLSXScrollContentLayout)

            self.OnlyXLSXScrollView.setWidgetResizable(True)
            self.OnlyXLSXScrollView.setWidget(self.OnlyXLSXScrollContent)

            self.mainLayout.addWidget(self.OnlyXLSXScrollView)
        
            self.openExcessNotesButton = QPushButton("Unbekannte Notes anzeigen")
            # self.openExcessNotesButton.clicked.connect(lambda: self.playwright_noteviewer(cookies=cookies, notes=self.only_in_xlsx))
            self.openExcessNotesButton.clicked.connect(lambda: self.open_notes(notes=self.only_in_xlsx))
            self.mainLayout.addWidget(self.openExcessNotesButton)


        if len(only_in_sap) != 0:
            self.openMissingNotesButton.setFocus()
            self.openMissingNotesButton.setDefault(True)

        self.setLayout(self.mainLayout)

    def open_notes(self, notes):
        for note in notes:
            webbrowser.open(f"https://launchpad.support.sap.com/#/notes/{note}")
#     @asyncSlot()
#     async def playwright_noteviewer(self, cookies, notes):
#         async with async_playwright() as p:
#             # browser context instance is in window instance,
#             # so that it can be reused for multiple calls of playwright_noteviewer from ui

#                 # first checking if context has been set up; if not, run login once
#                 # if self.browser is None:
#                 if True:
#                     print("neuer browser")
                    
#                     chromepath = chromium_utils.check_browser_install()
#                     self.browser = await p.chromium.launch(headless=False, executable_path=chromepath)
#                     self.context = await self.browser.new_context()
                    
#                     await self.context.add_cookies(cookies)
#                     print("cookies added")

#                 else:
#                     print("browser existiert")


#                 # opening all others in new tabs
#                 # page = await self.context.new_page()
#                 # await page.goto("https://launchpad.support.sap.com/#/notes/4545454")
#                 await asyncio.gather(*[self.open_note(self.context, note) for note in notes])

#                 # for note in notes:
#                 #     print(note)

#                 # for note in notes:
#                 #     print(f"for {note} in len={len(notes)}")
#                 #     loop = asyncio.get_event_loop()

#                 #     loop.create_task(self.open_note(self.context, note))
#                 #     # await asyncio.sleep(1)
                    




#     async def open_note(self, context, note):
#         print("open_note")

#         # print("vor con und page")
#         # context = await self.browser.new_context()
#         # page = await context.new_page()
#         # print("danach")
#         # await context.add_cookies(cookies)
#         page = await context.new_page()

#         note_page_reached = False
#         while not note_page_reached:
#             print("was it reached?")
#             try:
#                 await page.goto(f"https://launchpad.support.sap.com/#/notes/{note}")
#                 await page.wait_for_load_state("networkidle", timeout=0)
#                 if page.url == f"https://launchpad.support.sap.com/#/notes/{note}":
#                     note_page_reached = True
#             except:
#                 # exception occurs if browser is closed. If another exc occurs, good luck 👍
#                 pass
#             # else:
#             #     note_page_reached = False
#             #     print("page NOT reached", page.url)

#             #     # sometimes SAP says no
#             #     # just trying again is really elegant i think
#             #     await page.close()
#             #     await self.open_note(context, note)
                
#         # wait up to 16.667 minutes :)
#         await page.wait_for_timeout(1000000)


if __name__ == "__main__":
    main()
