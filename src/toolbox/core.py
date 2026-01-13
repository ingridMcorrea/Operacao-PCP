# Core Functions

import time
import pyautogui
import os
import pygetwindow

SCREEN_WIDTH = 1280
SCREEN_HEIGHT = 768

## Get TOTVS Windows, Resize and set it to 0,0 on primary monitor
def resizeTotvs():
    ## Primary monitor size 
    screenWidth, screenHeight = pyautogui.size()
    print(screenWidth, screenHeight)

    totvsWindow = pygetwindow.getWindowsWithTitle("TOTVS Linha RM - Varejo  Alias: CORPORERM_RH | 1-DROGARIA ARAUJO S/A")[0]
    print(f"Dando resize na tela para {SCREEN_WIDTH}x{SCREEN_HEIGHT}")
    totvsWindow.minimize()
    totvsWindow.maximize()
    totvsWindow.move(0, 0)
    totvsWindow.size = (SCREEN_WIDTH, SCREEN_HEIGHT)
    time.sleep(1.0)

def mantemLigado():
    pyautogui.press("pause")

def getImage(inputImagePath):
    script_dir = os.path.dirname(__file__)
    image_path = os.path.join(script_dir, "imagens", inputImagePath)

    return image_path

def buscaClicaBotao(imagemBotaoExemploNome, sleepTime = 0.2, doubleClick = False):
    botao = pyautogui.locateCenterOnScreen(getImage(imagemBotaoExemploNome), confidence=0.88, grayscale=True)
    if botao:
        # pyautogui.click(botao)
        if(doubleClick):
            pyautogui.doubleClick(botao)
        else:
            pyautogui.click(botao)
        time.sleep(sleepTime)
        pyautogui.move(-150, -150)
    else:
        print(f"Botao não encontrado: {imagemBotaoExemploNome}")
    
def ficaBuscandoClicaBotaoAteAcharEClica(imagemBotaoExemploNome, numTentativas = 0):
    contadorTentativas = 0
    exitLoop = False
    while(not exitLoop):
        try:
            botao = pyautogui.locateCenterOnScreen(getImage(imagemBotaoExemploNome), confidence=0.88, grayscale=True)
            if botao:
                time.sleep(0.1)
                pyautogui.click(botao)
                print("Botao de fechar clicado")
                exitLoop = True
                # pyautogui.moveTo(0, 0)
                pyautogui.move(-150, -150)
        except pyautogui.ImageNotFoundException:
            # else:
                print(f"Botão não encontrado: {imagemBotaoExemploNome} | TENTANDO NOVAMENTE")
                exitLoop = False
                time.sleep(1)
                contadorTentativas += 1
                if contadorTentativas == numTentativas:
                    break
                elif numTentativas == 0:
                    continue

def tentaAcharImagem(imagemExemploNome, numTentativas = 0):
    contadorTentativas = 0
    exitLoop = False
    while(not exitLoop):
        try:
            botao = pyautogui.locateCenterOnScreen(getImage(imagemExemploNome), confidence=0.88, grayscale=True)
            if botao:
                time.sleep(0.1)
                print("Imagem Encontrada")
                exitLoop = True
                pyautogui.move(25, 25)
                return True
        except pyautogui.ImageNotFoundException:
                print(f"Imagem não encontrada: {imagemExemploNome} | TENTANDO NOVAMENTE")
                exitLoop = False
                time.sleep(1)
                contadorTentativas += 1
                if contadorTentativas == numTentativas:
                    return False
                elif numTentativas == 0:
                    continue