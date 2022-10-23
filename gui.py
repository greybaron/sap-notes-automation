import os
import shutil
import traceback
import webbrowser
from datetime import date, timedelta
from pathlib import Path

import keyring
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (QApplication, QComboBox, QDialog, QHBoxLayout,
                             QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
                             QPushButton, QVBoxLayout, QWidget, QScrollArea)

from analysis import analysis
from scrape import scrape


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

        



app = QApplication([])
window = QWidget()

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


        self.prepareOutputButton = QPushButton("PDF-Verzeichnis vorbereiten...")
        self.prepareOutputButton.clicked.connect(self.prepareOutputFiles)

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


    def prepareOutputFiles(self):
        dialog = QMessageBox(parent=self, text="This will delete everything inside '~/Downloads/Hinweise PDFs'!")
        dialog.setWindowTitle("Warning")
        dialog.setIcon(QMessageBox.Icon.Warning)
        dialog.setStandardButtons(QMessageBox.StandardButton.Ok|
                        QMessageBox.StandardButton.Cancel)
        

        # 1024 corresponds to 'ok' button press.
        if dialog.exec() == 1024:
            output_dir= Path.home().joinpath(f"Downloads/Hinweise PDFs")

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
        

    def startAccountSetup(self):
        self.accountSetupWindow = accountSetupWindow()
        self.accountSetupWindow.show()


    def start_processing(self, possibleWeekChoices, weekSelector, system):
        selection_data = possibleWeekChoices[weekSelector.currentIndex()]

        xlsx_india = Path.home().joinpath(f"Downloads/SAP Hinweise KW {selection_data['kw']}_{selection_data['year']}.xlsx")

        if not xlsx_india.is_file():
            self.a = alert("XLSX fehlt", f"~/Downloads/SAP Hinweise KW {selection_data['kw']}_{selection_data['year']}.xlsx existiert nicht")
            return
        elif keyring.get_password("system", "launchpad_username") is None or keyring.get_password("system", "launchpad_password") is None:
            self.a = alert("Keine Logindaten vorhanden", "Zuerst Launchpad-Account festlegen")
            return
        else:
            try:
                delete_data_csv()
            except Exception:
                self.excv = exception_viewer("Deleting temp data failed", traceback.format_exc())
                return
            
            try:
                scrape(selection_data['formatted_date'], system)
            except Exception:
                self.excv = exception_viewer("Scraping failed", traceback.format_exc())
                return

            try:
                results = analysis(system, selection_data['kw'], selection_data['year'])

                if len(results) == 0:
                    self.a = alert(system, "Keine fehlenden Hinweise")
                else:
                    self.results_window = results_window(system, results)
                    self.results_window.show()
                
            except Exception:
                self.excv = exception_viewer("Analysis failed", traceback.format_exc())
                return
    

        





class results_window(QDialog):
    def __init__(self, system, results):
        super().__init__()
        self.setWindowTitle(f"{system}  —  Ergebnisse")
        self.mainLayout = QVBoxLayout()
        self.mainLayout.addWidget(QLabel("Fehlende Hinweise:"))

        self.scrollView = QScrollArea()
        self.scrollContent = QWidget()
        self.scrollContentLayout = QVBoxLayout()

        
        

        

        for notenumber in results:
            individualNoteButton = QPushButton(str(notenumber))
            # i'm not sure exactly what lambda state does, but otherwise all buttons call
            # opennote with the same argument/number, so i guess it keeps the state of notenumber inside x
            individualNoteButton.clicked.connect(lambda state, x=notenumber: open_note(x))
            self.scrollContentLayout.addWidget(individualNoteButton)    
        self.scrollContent.setLayout(self.scrollContentLayout)

        # self.scrollView.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        # self.scrollView.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scrollView.setWidgetResizable(True)
        self.scrollView.setWidget(self.scrollContent)

        self.mainLayout.addWidget(self.scrollView)
        # self.mainLayout.addStretch()

        self.ensureLoginButton = QPushButton("Anmelden (falls nötig)")
        self.ensureLoginButton.clicked.connect(lambda: webbrowser.open("https://launchpad.support.sap.com"))
        
        self.openAllNotesButton = QPushButton("Alle Notes anzeigen")
        self.openAllNotesButton.clicked.connect(lambda: open_all_notes(results))

        self.buttonRow = QHBoxLayout()
        self.buttonRow.addWidget(self.ensureLoginButton)
        self.buttonRow.addWidget(self.openAllNotesButton)
        self.mainLayout.addLayout(self.buttonRow)
        self.setLayout(self.mainLayout)

        self.openAllNotesButton.setFocus()
        self.openAllNotesButton.setDefault(True)


def open_note(notenumber):
    webbrowser.open("https://launchpad.support.sap.com/#/notes/"+str(notenumber))

def open_all_notes(notes):
    for note in notes:
        open_note(note)



def delete_data_csv():
    # Exception Handling is now in GUI
    csv_path = Path.home().joinpath("Downloads/data.csv")

    if os.path.exists(csv_path):
        os.remove(csv_path)

mw = mainWindow()
mw.show()
app.exec()
