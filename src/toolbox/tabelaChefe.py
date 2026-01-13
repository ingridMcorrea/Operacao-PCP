import pyautogui
import time
import pyperclip
import webbrowser

from toolbox.core import buscaClicaBotao
from toolbox.core import ficaBuscandoClicaBotaoAteAcharEClica
from toolbox.core import tentaAcharImagem

def excluiChefeAcessoInativo(codSecao, nivelSecao, secaoAtiva):

    ## Ignora secoes ativas e cabecalho
    if secaoAtiva == "ATIVA" or  secaoAtiva == "SEÇÃO ATIVA?":
        return

    print(f"Removendo acessos {codSecao}")

    if nivelSecao == 32:
        excluiVinculo(codSecao, 12)  

    if nivelSecao == 27:
        excluiVinculo(codSecao, 11)

    if nivelSecao == 24:
        excluiVinculo(codSecao, 10)

    if nivelSecao == 21:
        excluiVinculo(codSecao, 9)
    
    if nivelSecao == 18:
        excluiVinculo(codSecao, 8)
    
    if nivelSecao == 15:
        excluiVinculo(codSecao, 7)
    
    if nivelSecao == 12:
        excluiVinculo(codSecao, 6)

    if nivelSecao == 9:
        excluiVinculo(codSecao, 5)

    if nivelSecao == 6:
        excluiVinculo(codSecao, 4)

    if nivelSecao == 4:
        excluiVinculo(codSecao, 3)

def insereChefe(codSecao, matricula): 
    print(f"Inserindo chefe {matricula} na {codSecao}")

    nivelSecao = len(codSecao)

    if nivelSecao == 32:
        insereChefeWrapper(codSecao, 11, matricula)
    if nivelSecao == 27:
        insereChefeWrapper(codSecao, 10, matricula)
    if nivelSecao == 24:
        insereChefeWrapper(codSecao, 9, matricula)
    if nivelSecao == 21:
        insereChefeWrapper(codSecao, 8, matricula)
    if nivelSecao == 18:
        insereChefeWrapper(codSecao, 7, matricula)
    if nivelSecao == 15:
        insereChefeWrapper(codSecao, 6, matricula)
    if nivelSecao == 12:
        insereChefeWrapper(codSecao, 5, matricula)
    if nivelSecao == 9:
        insereChefeWrapper(codSecao, 4, matricula)
    if nivelSecao == 6:
        insereChefeWrapper(codSecao, 3, matricula)
    if nivelSecao == 4:
        insereChefeWrapper(codSecao, 2, matricula)
        
def insereChefeWrapper(codSecao, profundidade, matricula):    
    buscaClicaBotao("filtroCodSecaoLike.png")
    pyperclip.copy(codSecao)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(0.5)

    pyautogui.PAUSE = 0.010
    for i in range(profundidade):
        pyautogui.press("right")
    pyautogui.PAUSE = 0.1

    time.sleep(0.3)
    buscaClicaBotao("novoChefe.png")
    pyperclip.copy(matricula)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(0.8)

    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write("01/01/2026", interval=0.008)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")

    # time.sleep(0.250)
    pyautogui.press("enter")
    
def excluiVinculo(codSecao, profundidade):
    buscaClicaBotao("filtroCodSecaoLike.png")
    pyperclip.copy(codSecao)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(0.5)

    buscaClicaBotao("drogariaAraujoNivel.png")

    pyautogui.PAUSE = 0.010
    for i in range(profundidade):
        pyautogui.press("right")
    pyautogui.PAUSE = 0.1

    buscaClicaBotao("excluiTabelaChefe.png")
    time.sleep(0.250)
    pyautogui.press("enter")

def insereSupervisor(codSecao, matriculaSupervisor, codigoEquipeNome, dataInicio):
    if matriculaSupervisor == "MATRÍCULA":
        return

    print(f"Tentando inserir supervisor {matriculaSupervisor}")

    buscaClicaBotao("filtroCodSecaoLike.png")
    pyperclip.copy(codSecao)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    time.sleep(0.5)

    for i in range(11):
        pyautogui.press("right")

    
    # pyautogui.PAUSE = 1

    pyautogui.PAUSE = 0.1
    buscaClicaBotao("novoSupervisor.png")
    pyperclip.copy(matriculaSupervisor)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.7)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(codigoEquipeNome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.7)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write(dataInicio, interval=0.016)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")

    erro = tentaAcharImagem("erroSupervisor.png", 2)
    if erro:
        pyautogui.press("enter")
        pyautogui.hotkey("alt", "f4")
        pyautogui.press("enter")
        print(f"Falha ao inserir supervisor {matriculaSupervisor} em {codigoEquipeNome}")

    buscaClicaBotao("drogariaAraujoNivel.png")
