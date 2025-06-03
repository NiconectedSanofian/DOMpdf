import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QProgressBar, QLabel, QMessageBox)
from PyQt5.QtGui import QIcon 
from PyQt5.QtCore import Qt
import fitz 
import re
import pandas as pd

def get_text(pdf_path):
    """
    Extracts and returns the concatenated text content from all pages of a PDF file.

    Args:
        pdf_path (str): The file path to the PDF document.

    Returns:
        str: The combined text extracted from all pages of the PDF.

    Raises:
        FileNotFoundError: If the specified PDF file does not exist.
        fitz.FileDataError: If the file is not a valid PDF or cannot be opened.
    """
    with fitz.open(pdf_path) as doc:
        text = ""
        for page in doc:
            text += page.get_text()
    return text

def extract_fields(texto):
    """
    Extracts specific fields from a given text string, typically obtained from a PDF document related to Mexican tax identification (SAT).
    The function preprocesses the input text by removing irrelevant sections and then uses regular expressions to extract key-value pairs for fields such as CURP, RFC, address details, and other personal or fiscal information.
    Args:
        texto (str): The input text from which to extract fields.
    Returns:
        dict: A dictionary where keys are field names (e.g., 'CURP', 'RFC', 'Código Postal') and values are the corresponding extracted values from the text. If a field is not found, its value is 'N/A'.
    """
    
    texto = texto.split('Actividades Económicas:')[0]  
    texto = texto.split('Susdatospersonales sonincorporados yprotegidos enlossistemas delSAT,deconformidad conlosLineamientos deProtección deDatos')[0]  
    texto = texto.replace('Página  [2] de [2]', '')  
    texto = texto.replace('Página  [2] de [3]', '')  
    texto = texto.replace('Página  [2] de [4]', '')  
    texto = texto.replace('Página  [2] de [5]', '')
    texto = texto.replace('Página  [2] de [6]', '')

    patrones = {
        'Cédula de Identificación fiscal': r'CÉDULA DE IDENTIFICACIÓN FISCAL \n([\w]*)?',
        'CURP': r'CURP:\n([\w]*)?',
        'RFC': r'RFC:\n([\w]*)?',
        'Nombre o Razón Social': r'Registro\s*Federal\s*de\s*Contribuyentes\n([^\n]*)?\nNombre,\s*denominación\s*o\s*razón',
        'Código Postal': r'Código\s*Postal:\s*(.*?)?\nTipo\s*de\s*Vialidad:',
        'Tipo Vialidad': r'Tipo\s*de\s*Vialidad:\s*(.*?)?\nNombre\s*de\s*Vialidad:',
        'Nombre Vialidad': r'Nombre\s*de\s*Vialidad:\s*(.*?)?\s*Número\s*Exterior:',
        'Número Exterior': r'Número\s*Exterior:\s*([^\n]*)?',
        'Número Interior': r'Número\s*Interior:\s*(.*?)?\n*Nombre',
        'Nombre Colonia': r'Nombre\s*de\s*la\s*Colonia:\s*([^\n]*)?',
        'Nombre Localidad': r'Nombre\s*de\s*la\s*Localidad:\s*(.*?)?\s*Nombre',
        'Nombre Municipio o Demarcación Territorial': r'\s*Territorial:\s*(.*?)?\n(.*?)?\s*Nombre\s*de\s*la\s*Entidad\s*Federativa:',	
        'Nombre Entidad Federativa': r'Nombre\s*de\s*la\s*Entidad\s*Federativa:\s*(.*?)?\s*\n',
        'Entre Calle': r'Entre\s*Calle:\s*(.*?)?\n(.*?)?\s*Y\s*Calle:',
        'Y Calle': r'Y\s*Calle:\s*([^\n]*)?',
        'Fecha y lugar de emisión': r'Lugar y Fecha de Emisión\n*([^\n]*)\n*([^\n]*)?',
    }

    resultado = {}
    for campo, patron in patrones.items():
        coincidencia = re.search(patron, texto)
        if coincidencia:
            resultado[campo] = coincidencia.group(1).strip() + ' ' + coincidencia.group(2).strip() if len(coincidencia.groups()) > 1 else coincidencia.group(1).strip()
            if campo == "Fecha y lugar de emisión":
                resultado[campo] = resultado[campo].replace(resultado["RFC"],"")
        else:
            resultado[campo] = 'N/A'
    
    return resultado
def is_fiscal(texto):
    # Verifies if the PDF is indeed a fiscal document by checking for specific keywords.
    return "CONSTANCIA DE SITUACIÓN FISCAL" in texto

class PDFExtractorApp(QWidget):
    """
    A PyQt5-based GUI application for extracting text from PDF files in a selected folder,
    processing them to extract specific fields, and exporting the results to an Excel file.
    Features:
        - Allows the user to select a folder containing PDF files.
        - Processes each PDF to extract text and validate if it is a fiscal document.
        - Extracts relevant fields from fiscal documents using a generic extraction function.
        - Displays a progress bar during processing.
        - Exports the extracted data to an Excel (.xlsx) file.
        - Provides user feedback via message boxes for errors and successful operations.
    Methods:
        __init__():
            Initializes the GUI components and layout.
        select_folder():
            Opens a dialog for the user to select a folder and initiates PDF processing.
        process_pdfs(folder_path):
            Processes all PDF files in the given folder, extracts data, updates progress,
            and saves the results to an Excel file.
    """
    def __init__(self):
        """
        Initializes the main window for the PDF text extractor application.
        Sets up the window icon, title, and geometry. Creates and arranges the main UI components,
        including a label with instructions, a button to select a folder containing PDF files,
        and a progress bar to display processing progress.
        """
        super().__init__()
        self.setWindowIcon(QIcon("DomPDF_icon.ico"))
        self.setWindowTitle("Extractor de Texto desde PDFs")
        self.setGeometry(100, 100, 400, 200)
        
        self.layout = QVBoxLayout()

        self.label = QLabel("Selecciona una carpeta con archivos PDF")
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Seleccionar carpeta")
        self.select_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)

        self.progress = QProgressBar()
        self.progress.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.progress)

        self.setLayout(self.layout)

    def select_folder(self):
        """
        Opens a dialog for the user to select a directory containing PDF files.
        If a directory is selected, initiates processing of the PDFs in the chosen folder.

        Uses:
            QFileDialog.getExistingDirectory: To prompt the user for a folder selection.

        Side Effects:
            Calls self.process_pdfs with the selected folder path if a folder is chosen.
        """
        folder_path = QFileDialog.getExistingDirectory(self, "Selecciona carpeta con PDFs")
        if folder_path:
            self.process_pdfs(folder_path)

    def process_pdfs(self, folder_path):
        """
        Processes all PDF files in the specified folder, extracts relevant fields from each PDF,
        and saves the results to an Excel file.
        Args:
            folder_path (str): The path to the folder containing PDF files to process.
        Workflow:
            - Lists all PDF files in the given folder.
            - For each PDF:
                - Extracts text from the PDF.
                - Checks if the PDF is a fiscal document.
                - If fiscal, extracts generic fields and adds the filename as the first column.
                - Updates a progress bar in the UI.
            - Prompts the user to select a location to save the results as an Excel file.
            - Saves the extracted data to the specified Excel file.
            - Notifies the user upon successful completion.
        UI:
            - Displays warnings if no PDFs are found.
            - Updates progress bar during processing.
            - Shows information dialog when results are saved.
        Note:
            This method relies on external functions: get_text, is_fiscal, extraer_campos_generico,
            and uses PyQt5 widgets for UI interactions.
        """
        archivos_pdf = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        total = len(archivos_pdf)
        if total == 0:
            QMessageBox.warning(self, "Sin PDFs", "No se encontraron archivos PDF en la carpeta.")
            return
        
        resultados = []
        for i, archivo in enumerate(archivos_pdf, 1):
            ruta = os.path.join(folder_path, archivo)
            texto = get_text(ruta)
            if is_fiscal(texto):
                resultado = extract_fields(texto)
                resultado['Archivo'] = archivo
                resultado = {k: resultado[k] for k in ['Archivo'] + [key for key in resultado if key != 'Archivo']}
                resultados.append(resultado)
            self.progress.setValue(int((i / total) * 100))
            QApplication.processEvents()

        save_path, _ = QFileDialog.getSaveFileName(self, "Guardar Excel", "", "Excel Files (*.xlsx)")
        if save_path:
            if not save_path.endswith(".xlsx"):
                save_path += ".xlsx"
            df = pd.DataFrame(resultados)
            df.to_excel(save_path, index=False)
            QMessageBox.information(self, "Éxito", f"Resultados guardados en:\n{save_path}")
        self.progress.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = PDFExtractorApp()
    ventana.show()
    sys.exit(app.exec_())