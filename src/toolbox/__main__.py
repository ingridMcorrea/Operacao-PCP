import pyautogui
import openpyxl
import time
import pyperclip
import webbrowser
import os
import tkinter as tk
import schedule

from toolbox.ToolboxApp import ToolboxApp

from toolbox import exportaBases
from toolbox import totvsTools
from toolbox import core

def main(args=None):
    core.resizeTotvs()
    
    

    exportaBases.exportaDiaria()

    # schedule.every(25).minutes.do(exportaBases.exportaDiaria)
    # schedule.every().day.at("07:00")

    # schedule.every(5).minutes.do(core.mantemLigado)

    # while 1:
    #     schedule.run_pending()
    #     time.sleep(1)

    # Open input sheet
    workbook = openpyxl.load_workbook("input.xlsx")
    sheet = workbook.active

    contador = 0
    for row in sheet.iter_rows():

        if row[0].value == None:
            break
         

        contador += 1
        print(contador)

    pyautogui.alert("Execucao concluida")

if __name__ == "__main__":
    main()
