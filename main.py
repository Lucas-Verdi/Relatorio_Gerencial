import tkinter as tk
from tkinter import *
from tkinter import filedialog
import xlwings
import pyautogui
import win32com.client as win32
import pythoncom
import pywintypes
import ctypes
from pyautogui import sleep


def main():
    # Criando janela para selecionar o arquivo base
    root = tk.Tk()
    root.withdraw()
    global arquivo
    arquivo = filedialog.askopenfilename()

    # Abrindo a planilha selecionada
    pastadetrabalho = xlwings.Book(arquivo)

    # Abre o Excel em tela cheia
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.Visible = True
    xl.Workbooks.Open(arquivo)
    xl.ActiveWindow.WindowState = win32.constants.xlMaximized

    #Formatando a planilha
    planilha = pastadetrabalho.sheets[0]
    planilha.range('1:8').delete()
    planilha.range('A:AD').column_width = 6
    planilha.range('A:AD').unmerge()
    planilha.range('A:B').delete()
    sleep(0.3)
    planilha.range('B:B').delete()
    sleep(0.3)
    planilha.range('C:P').delete()
    sleep(0.3)
    planilha.range('D:K').delete()
    sleep(0.3)
    planilha.range('C:C').column_width = 45
    planilha.range('F:J').column_width = 6

#Interface
janela = Tk()
janela.title('Relatório Gerencial')
Label1 = Label(janela, text='Insira a pasta de trabalho:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='Inserir')
Botao1.bind("<Button>",  lambda e: main())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()