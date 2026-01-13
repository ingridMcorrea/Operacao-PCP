import pyautogui
import time
import pyperclip

from toolbox.core import buscaClicaBotao
from toolbox.core import ficaBuscandoClicaBotaoAteAcharEClica
from toolbox.core import tentaAcharImagem

## Abre aumento de quadro
def abreAumentoDeQuadro(matriculaRequisitante, justificativa, filialCod, secaoCod, funcaoCod, salario, horarioTrabalhoCod, escalaTrabalhoCod, sexoCod):
    buscaClicaBotao("inserir.png")
    pyperclip.copy(matriculaRequisitante)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(justificativa)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyperclip.copy(filialCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(secaoCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(funcaoCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write(str(salario), interval=0.016)
    # pyperclip.copy(str(salario))
    # pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    ## MUDA DE ABA PARA A DE CAMPOS COMPLEMENTARES
    pyautogui.PAUSE = 0.150
    buscaClicaBotao("camposComplementares.png")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(horarioTrabalhoCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(escalaTrabalhoCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(sexoCod)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy("1")
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    print()

## Abre Transferencias Simples - Necessita de validação
def abreTransferencia(matriculaRequisitante, matriculaColaborador, justificativa, numeroFilial, codSecao, geraSub, motivoMudancaSecao, motivoTransferenciaCamposComplementares):

    # Ignora Cabecalho
    if matriculaRequisitante == "MATRICULA REQUISITANTE":
        return

    pyautogui.hotkey("ctrl", "insert")

    pyperclip.copy(matriculaRequisitante)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    
    pyperclip.copy(matriculaColaborador)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")

    pyperclip.copy(justificativa)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(numeroFilial)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(codSecao)
    pyautogui.hotkey("ctrl", "v")
    
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    # NAO GERA SUBSTITUICAO
    if geraSub == False:
        pyautogui.press("space")
        pyautogui.press("enter")

    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.write(motivoMudancaSecao, interval=0.016)

    buscaClicaBotao("camposComplementares.png", 1)
    pyautogui.press("tab")
    pyperclip.copy(motivoTransferenciaCamposComplementares)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")

    time.sleep(3)

## Abre Alteração de dados - Cheio de remendos para casos especificos, precisa de melhor estruturação
def abrirAlteracaoDadosFuncionais(matriculaRequisitante, matriculaColaborador, justificativa, codFuncao, 
                                  salario, dataPrevisao, sendoMovimentado, salarioAtual, alteraFuncao = False):

    if sendoMovimentado == False:
        print(f"Colaborador {matriculaColaborador} já possui Req")
        return 

    print(f"Abrindo requisição de alteração de dados funcionais chapa num {matriculaColaborador} para {codFuncao}")

    pyautogui.hotkey("ctrl", "insert")
    time.sleep(0.7)
    # buscaClicaBotao("inserir.png")
    pyperclip.copy(matriculaRequisitante)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.8)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(matriculaColaborador)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.8)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(dataPrevisao, interval=0.016)
    pyautogui.press("tab")
    pyperclip.copy(justificativa)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(codFuncao)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.6)
    # pyautogui.PAUSE = 1.0
    # if nomeFuncao == "VENDEDOR":
    #     print("É vendedor")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyperclip.copy("0,0")
    #     pyautogui.hotkey("ctrl", "v")
    #     pyautogui.press("tab")
    #     pyautogui.press("right")
    #     pyautogui.press("enter")
    #     pyautogui.press("tab")
    #     pyperclip.copy("04")
    #     pyautogui.hotkey("ctrl", "v")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyperclip.copy("04")
    #     pyautogui.hotkey("ctrl", "v")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("enter")
    # else:
    #     print("Possui salario normal")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyperclip.copy("04")    
    #     pyautogui.hotkey("ctrl", "v")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("tab")
    #     pyautogui.press("enter")
    
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyperclip.copy(salario)
    # pyautogui.hotkey("ctrl", "a") 
    pyautogui.write(str(salario), interval=0.016)
    pyautogui.press("tab")
    

    # MERITO E PROMOCAO
    necessistaEscalonamento = tentaAcharImagem("salarioExcede.png", 2)
    if necessistaEscalonamento == True:
        pyautogui.press("enter")


    if alteraFuncao == "SIM":
        justificativa = "PROMOÇÃO"

    if justificativa == "ENQUADRAMENTO":
        pyautogui.press("tab")
        pyautogui.write("08")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")

    if justificativa == "MÉRITO":
        pyautogui.press("tab")
        pyautogui.write("10")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")

    if justificativa == "PROMOÇÃO":
        pyautogui.press("tab")
        pyautogui.write("04")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.write("12")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")

    if justificativa == "UNIFICAÇÃO":
        pyautogui.press("tab")
        pyautogui.write("04")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.write("08")
        pyautogui.press("tab")
        pyautogui.press("tab")
        pyautogui.press("tab")

    ## EXCECAO TATI
    if justificativa == "Movimentações de Loja Vigência 01/01/2026":
        if salarioAtual == salario:
            pyautogui.press("tab")
            pyautogui.write("03")
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.press("tab")
        else:
            pyautogui.press("tab")
            pyautogui.write("03")
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.write("04")
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.press("tab")

    pyautogui.press("enter")

    alertaEscalonamento = tentaAcharImagem("simboloDuvida.png", 2)
    if alertaEscalonamento == True:
        pyautogui.press("right")
        pyautogui.press("enter")
    if necessistaEscalonamento == True:
        pyautogui.press("right")
        pyautogui.press("enter")

    time.sleep(2)