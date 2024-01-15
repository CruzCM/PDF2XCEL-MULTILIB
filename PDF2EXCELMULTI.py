import PySimpleGUI as sg
from tkinter import filedialog
import os
import fitz  # PyMuPDF
import pandas as pd
from tabula import read_pdf



# Função para converter PDF em Excel
def convert_pdf_to_excel(pdf_file, excel_file,type_combo_lib):
    try:
        if type_combo_lib == 0: #TABULA
            df_list = read_pdf(pdf_file, pages='all')
            df = pd.concat(df_list, ignore_index=True)
            df.to_excel(excel_file, index=False)

        elif type_combo_lib_index == 1:
            # Abrir o arquivo PDF
            doc = fitz.open(pdf_file)
            text_cells = []
            # Extrair o texto das páginas
            for page in doc:
                for block in page.get_text("blocks"):
                    text_cells.append(block[4])
            # Criar um DataFrame a partir do texto
            data = {'Texto do PDF': text_cells}
            df = pd.DataFrame(data)
            # Salvar o DataFrame em um arquivo Excel
            df.to_excel(excel_file, index=False)
            sg.popup("Conversão concluída com sucesso!", title="Conversão de PDF para Excel")

    except Exception as e:
        sg.popup_error(f"Ocorreu um erro: {str(e)}", title="Erro")

# Função para escolher PDF
def browse_button_click():
    file_path = filedialog.askopenfilename(
        title="Escolha um arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    if file_path:
        window["-PDF-"].update(file_path)


# Função para escolher onde salvar EXCEL
def save_as_button_click(excel_filename):
    file_path = filedialog.asksaveasfilename(
        title="Escolha onde salvar o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        defaultextension=".xlsx",
        initialfile=excel_filename
    )
    if file_path:
        window["-Excel-"].update(file_path)


# Layout
sg.theme('LightGrey3')

elementos = [
    "TABULA : Extração de Tabelas",
    "CAMELOT : Extração de Tabelas (Detecção Automática)",
    "PYMUPDF : Manipulação Geral de PDF",
    "PDFMINER : Extração de Texto, Imagens e Metadados",
    "PDFPLUMBER : Extração de Textos e Tabelas Específicos"
]

layout = [
    [sg.Text("Escolha um arquivo PDF para converter em Excel:")],
    [sg.Input(key="-PDF-", disabled=True), sg.Button("Browse", key="-BROWSE-BUTTON-")],
    [sg.Text("Escolha onde salvar o arquivo Excel:")],
    [sg.Input(key="-Excel-", disabled=True), sg.Button("Save As", key="-SAVEAS-BUTTON-")],
    [sg.Text("Escolha a forma de conversão:")],
    [sg.Combo(elementos, readonly=True, default_value=elementos[0], key="-type-")],
    [sg.Text("", size=(1, 1))],
    [sg.Button("Converter")]
]

window = sg.Window("Conversor PDF para Excel", layout)

# EVENTOS
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break
    elif event == "-BROWSE-BUTTON-":
        browse_button_click()
    elif event == "-SAVEAS-BUTTON-":
        pdf_filename = os.path.splitext(os.path.basename(values["-PDF-"]))[0]
        save_as_button_click(pdf_filename)
    elif event == "Converter":
        pdf_file_path = values["-PDF-"]
        excel_file_path = values["-Excel-"]
        type_combo_lib = elementos.index(values["-type-"])
        if pdf_file_path and excel_file_path:
            convert_pdf_to_excel(pdf_file_path, excel_file_path, type_combo_lib)
            sg.popup("Conversão concluída com sucesso!", title="Conversão de PDF para Excel")
            window.close()
