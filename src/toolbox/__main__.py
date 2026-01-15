import pyautogui
import openpyxl
import time
import pyperclip
import webbrowser
import os
import tkinter as tk
import schedule
import argparse

from toolbox.ToolboxApp import ToolboxApp

from toolbox import tabelaChefe
from toolbox import exportaBases
from toolbox import totvsTools
from toolbox import core

def main(args=None):
    parser =  argparse.ArgumentParser("Toolbox")
    parser.add_argument("-a", "--atualizar", action="store_true", help="Atualiza todas as bases do PCP")
    parser.add_argument("-e", "--executar", action="store_true", help="Executa os comandos cadastrados na planilha de input.xlsx")

    args = parser.parse_args()
    
    if args.atualizar:
        core.resizeTotvs()
        exportaBases.exportaDiaria()
        exportaBases.exportaHorarios()
        exportaBases.exportaCadastros()
        exportaBases.exportaTabelaSalarial()

    if args.executar:
        ## pegar o diretorio com o argumento talvez?
        # Open input sheet
        workbook = openpyxl.load_workbook("input.xlsx")
        sheet = workbook.active

        contador = 0
        for row in sheet.iter_rows():
            if row[0].value == None:
                break

            # tabelaChefe.insereSupervisor(row[0].value, "0018628-7", row[1].value, "01/01/2026")
            
            contador += 1
            print(contador)
        pyautogui.alert("Execucao concluida")


    schedule.every(5).minutes.do(core.mantemLigado)
    schedule.every(30).minutes.do(exportaBases.exportaDiaria)
    schedule.every().day.at("09:00").do(exportaBases.exportaCadastros)

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    main()
