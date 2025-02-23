import sys
import time
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, 
    QTableWidget, QTableWidgetItem, QFileDialog
)
from PyQt5.QtCore import Qt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

class DragDropWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Impostazioni finestra
        self.setWindowTitle("Drag and Drop - Carica Excel")
        self.setGeometry(100, 100, 800, 600)

        # Layout
        self.layout = QVBoxLayout()

        # Etichetta per il drag-and-drop
        self.label = QLabel("Trascina un file Excel qui", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label)

        # Tabella per visualizzare i dati del file Excel
        self.table = QTableWidget(self)
        self.layout.addWidget(self.table)

        # Pulsante per caricare il file tramite finestra di dialogo
        self.load_button = QPushButton("Carica File Excel", self)
        self.load_button.clicked.connect(self.load_file_dialog)
        self.layout.addWidget(self.load_button)
        
        # Pulsanti per esportare in JSON e PDF
        self.export_json_button = QPushButton("Esporta in JSON", self)
        self.export_json_button.clicked.connect(self.export_to_json)
        self.export_json_button.setEnabled(False)  # Disabilitato finché non si carica un file
        self.layout.addWidget(self.export_json_button)

        self.export_pdf_button = QPushButton("Esporta in PDF", self)
        self.export_pdf_button.clicked.connect(self.export_to_pdf)
        self.export_pdf_button.setEnabled(False)  # Disabilitato finché non si carica un file
        self.layout.addWidget(self.export_pdf_button)

        # Pulsante per compilare il form web
        self.fill_form_button = QPushButton("Compila Form Web", self)
        self.fill_form_button.clicked.connect(self.fill_web_form)
        self.fill_form_button.setEnabled(False)
        self.layout.addWidget(self.fill_form_button)

        # Imposta la finestra per supportare il drag-and-drop
        self.setAcceptDrops(True)
        self.setLayout(self.layout)

        # Percorso del file Excel caricato
        self.file_path = None
        self.df = None
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        # Ottieni il file path dal drag-and-drop
        file_path = event.mimeData().urls()[0].toLocalFile()
        self.label.setText(f"File caricato: {file_path}")
        self.read_excel(file_path)
    
    def load_file_dialog(self):
        # Apri la finestra di dialogo per selezionare un file Excel
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleziona File Excel", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.label.setText(f"File caricato: {file_path}")
            self.read_excel(file_path)

    def read_excel(self, file_path):
        try:
            # Leggi il file Excel usando pandas
            if file_path.endswith(".xlsx"):
                self.df = pd.read_excel(file_path, engine="openpyxl")
            elif file_path.endswith(".xls"):
                self.df = pd.read_excel(file_path, engine="xlrd")
            else:
                raise ValueError("Formato non supportato. Usa .xls o .xlsx")

            self.display_table(self.df)
            # Abilita i pulsanti di esportazione
            self.export_json_button.setEnabled(True)
            self.export_pdf_button.setEnabled(True)
            self.fill_form_button.setEnabled(True)
            self.file_path = file_path

        except Exception as e:
            self.label.setText(f"Errore nel caricare il file: {str(e)}")

    def display_table(self, df):
        # Imposta il numero di righe e colonne della tabella
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)
        
        # Aggiungi i dati al widget della tabella
        for row in range(len(df)):
            for col in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[row, col]))
                self.table.setItem(row, col, item)
    
    def export_to_json(self):
            if self.df is not None:
                file_path, _ = QFileDialog.getSaveFileName(self, "Salva come JSON", "", "JSON Files (*.json)")
                if file_path:
                    try:
                        self.df.to_json(file_path, orient="records", indent=4, force_ascii=False), 
                        self.label.setText(f"File salvato in JSON: {file_path}")
                    except Exception as e:
                        self.label.setText(f"Errore nell'esportazione JSON: {str(e)}")

    def export_to_pdf(self):
        if self.df is not None:
            file_path, _ = QFileDialog.getSaveFileName(self, "Salva come PDF", "", "PDF Files (*.pdf)")
            if file_path:
                try:
                    self.create_pdf(file_path)
                    self.label.setText(f"File salvato in PDF: {file_path}")
                except Exception as e:
                    self.label.setText(f"Errore nell'esportazione PDF: {str(e)}")

    def create_pdf(self, file_path):
        c = canvas.Canvas(file_path, pagesize=A4)
        width, height = A4
        x_offset = 50
        y_offset = height - 50
        line_height = 20

        c.setFont("Helvetica", 10)
        c.drawString(x_offset, y_offset, "Dati esportati da Excel")
        y_offset -= 30

        # Disegna l'intestazione della tabella
        if not self.df.empty:
            columns = list(self.df.columns)
            for i, col in enumerate(columns):
                c.drawString(x_offset + i * 100, y_offset, col)
            y_offset -= line_height

            # Disegna i dati della tabella
            for index, row in self.df.iterrows():
                for i, col in enumerate(columns):
                    c.drawString(x_offset + i * 100, y_offset, str(row[col]))
                y_offset -= line_height

                # Se si supera la pagina, creare una nuova pagina
                if y_offset < 50:
                    c.showPage()
                    c.setFont("Helvetica", 10)
                    y_offset = height - 50

        c.save()
    
    def fill_web_form(self):
        if self.df is None:
            self.label.setText("Errore: Carica prima un file Excel!")
            return

        # Avvia il browser
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get("https://my.drimify.com/en/register?lang=it")
        #time.sleep(3)  # Attendi il caricamento della pagina

        for _, row in self.df.iterrows():
            nome = row["Nome"]
            cognome = row["Cognome"]
            email = row["Email"]
            password = row["Password"]

            try:
                
                # Compila altri campi (esempio per "Nome", "Email", etc.)
                driver.find_element(By.ID, "registration_form_firstname").send_keys(row["Nome"])
                driver.find_element(By.ID, "registration_form_surname").send_keys(row["Cognome"])
                driver.find_element(By.ID, "registration_form_email").send_keys(row["Email"])
                driver.find_element(By.ID, "registration_form_plainPassword").send_keys(row["Password"])
                print(f"Compilato modulo per {email}")
                #time.sleep(2)  # Aspetta prima di passare all'utente successivo

            except Exception as e:
                print(f"Errore nella compilazione: {str(e)}")

        #time.sleep(5)  # Aspetta prima di chiudere il browser
        #driver.quit()
        self.label.setText("Compilazione completata!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DragDropWindow()
    window.show()
    sys.exit(app.exec_())
