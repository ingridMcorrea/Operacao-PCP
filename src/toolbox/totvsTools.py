import pyautogui
import time
import pyperclip
import webbrowser

from toolbox.core import buscaClicaBotao
from toolbox.core import ficaBuscandoClicaBotaoAteAcharEClica
from toolbox.core import tentaAcharImagem


def alteraAprovador(requisicaoNum, motivo, matriculaNovoAprovador, matriculaResponsavel):
    #clica no filtro
    pyautogui.moveTo(653, 205)
    pyautogui.click()
    time.sleep(0.2)
    # escreve num requisicao
    pyautogui.write(str(requisicaoNum).removesuffix(".0"), interval=0.008)
    # confirma
    pyautogui.press('enter')
    time.sleep(0.5)
    # seleciona req
    pyautogui.moveTo(44, 311)
    pyautogui.click()
    time.sleep(0.2)
    # abre processos
    buscaClicaBotao("processosFiltros.png")
    # alterar aprovador
    buscaClicaBotao("alterarAprovadorAtualDaRequisicao.png")
    # novo aprovador
    pyautogui.moveTo(717, 284)
    pyautogui.click()
    time.sleep(0.8)
    pyautogui.write(str(matriculaNovoAprovador).removesuffix(".0"), interval=0.008)
    time.sleep(0.2)
    pyautogui.press('tab')
    # responsavel
    pyautogui.moveTo(341, 345)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.write(str(matriculaResponsavel).removesuffix(".0"), interval=0.008)
    time.sleep(0.2)
    pyautogui.press('tab')
    # seleciona caixa de texto
    pyautogui.moveTo(341, 402)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.write(motivo, interval=0.016)
    # executa
    pyautogui.moveTo(865, 539)
    pyautogui.click()
    time.sleep(5.5)
    # fecha caixa
    # pyautogui.moveTo(965, 601)
    # pyautogui.click()
    # time.sleep(0.2)
    ficaBuscandoClicaBotaoAteAcharEClica("botaoFechar.png")

def suspende(value, motivo):
    #clica no filtro
    pyautogui.moveTo(600, 199)
    pyautogui.click()
    time.sleep(0.7)
    # escreve num requisicao
    pyautogui.write(str(value), interval=0.016)
    # confirma
    pyautogui.press('enter')
    # seleciona
    pyautogui.moveTo(45, 279)
    pyautogui.click()
    time.sleep(0.5)
    # expande processos
    pyautogui.moveTo(498, 205)
    pyautogui.click()
    time.sleep(0.5)
    # clica suspender
    pyautogui.moveTo(500, 252)
    pyautogui.click()
    time.sleep(0.5)
    # clica caixa de texto
    pyautogui.moveTo(451, 373)
    pyautogui.click()
    time.sleep(0.5)
    # escreve motivo
    pyautogui.write(motivo, interval=0.016)
    # executa suspensao
    pyautogui.moveTo(715, 597)
    pyautogui.click()
    time.sleep(5)
    # fecha janela
    pyautogui.moveTo(791, 597)
    pyautogui.click()
    time.sleep(0.5)

def excluiVinculoSalarial(value):
    # abri vinculo
    pyautogui.moveTo(485, 121)
    pyautogui.click()
    time.sleep(1)
    # avanca processo
    pyautogui.moveTo(853, 673)
    pyautogui.click()
    time.sleep(0.2)
    # seleciona tabela
    pyautogui.moveTo(343, 469)
    pyautogui.doubleClick()
    time.sleep(0.2)
    # avanca processo
    pyautogui.moveTo(853, 673)
    pyautogui.click()
    time.sleep(0.2)
    # abre filtros
    pyautogui.moveTo(394, 250)
    pyautogui.click()
    time.sleep(0.5)
    # gerenciar filtros
    pyautogui.moveTo(394, 284)
    pyautogui.click()
    time.sleep(0.2)
    # novo filtro
    pyautogui.moveTo(895, 219)
    pyautogui.click()
    time.sleep(0.2)
    # seleciona codigo da secao
    pyautogui.moveTo(524, 268)
    pyautogui.click()
    time.sleep(0.2)
    # seleciona caixa de texto
    pyautogui.moveTo(820, 263)
    pyautogui.click()
    time.sleep(0.2)
    # escreve secao
    # value = "0346.1.02.06.00.02.31.01.00.0001"
    pyautogui.write(value, interval=0.016)
    # adiciona
    pyautogui.moveTo(944, 335)
    pyautogui.click()
    time.sleep(0.2)
    # filtrar por
    pyautogui.moveTo(716, 340)
    pyautogui.click()
    time.sleep(0.2)
    # codigo
    pyautogui.moveTo(631, 355)
    pyautogui.click()
    time.sleep(0.2)
    # seleciona caixa de texto
    pyautogui.moveTo(774, 336)
    pyautogui.click()
    time.sleep(0.2)
    # escreve secao
    pyautogui.write(value, interval=0.016)
    # filtra
    pyautogui.moveTo(899, 339)
    pyautogui.click()
    time.sleep(0.2)
    # avanca ok
    pyautogui.moveTo(832, 567)
    pyautogui.click()
    time.sleep(0.2)
    # adiciona
    pyautogui.moveTo(944, 335)
    pyautogui.click()
    time.sleep(0.2)
    # executa pesquisa
    pyautogui.moveTo(801, 701)
    pyautogui.click()
    time.sleep(1)
    # seleciona quadro de lotacao
    pyautogui.moveTo(546, 310)
    pyautogui.click()
    time.sleep(0.2)
    # exclui vinculos
    pyautogui.moveTo(487, 624)
    pyautogui.click()
    time.sleep(0.2)
    # aperta enter para sim
    pyautogui.press('enter')
    # avanca processo
    pyautogui.moveTo(838, 675)
    pyautogui.click()
    time.sleep(0.2)
    # avanca processo
    pyautogui.moveTo(838, 675)
    pyautogui.click()
    time.sleep(0.2)
    # EXECUTA
    pyautogui.moveTo(838, 675)
    pyautogui.click()
    time.sleep(6)
    # finaliza 
    pyautogui.moveTo(926, 676)
    pyautogui.click()
    time.sleep(0.2)

def addMotivoMudancaSalarial():
    pyautogui.PAUSE = 0.3
    pyautogui.moveTo(973, 315)
    pyautogui.doubleClick()
    # seleciona campo de motivo salarial
    pyautogui.moveTo(676, 548)
    pyautogui.click()
    # escreve campo de motivo salarial
    pyautogui.write("04", interval=0.016)
    pyautogui.press("tab")
    # clica em ok
    pyautogui.moveTo(796, 598)
    pyautogui.click()
    # confirma alteracoes
    pyautogui.moveTo(747, 523)
    pyautogui.click()
    # escreve motivo da alteracao
    pyautogui.write("Ok", interval=0.016)
    # clica em ok
    pyautogui.press("tab")
    pyautogui.press("enter")
    # pyautogui.moveTo(253, 416)
    # pyautogui.click()
    # pyautogui.moveTo(193, 355)
    # pyautogui.click()
    ficaBuscandoClicaBotaoAteAcharEClica("okDiferente.png")
    pyautogui.PAUSE = 0.1
    print()

def cancelaRequisicao(value, motivo):
    print(f"Cancelando requisicao: {value}")

    pyautogui.PAUSE = 0.3
    pyautogui.moveTo(653, 205)
    pyautogui.click()
    time.sleep(0.2)
    pyautogui.write(str(value).removesuffix(".0"), interval=0.008)
    pyautogui.press('enter')
    time.sleep(0.5)

    pyautogui.moveTo(45, 282)
    pyautogui.click()
    time.sleep(0.2)
    
    pyautogui.hotkey("ctrl", "p")
    pyautogui.press("down")
    pyautogui.press("enter")
    for i in range(8):
        pyautogui.press("tab")

    pyperclip.copy(motivo)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)

    pyautogui.press("tab")
    pyautogui.press("enter")
    ficaBuscandoClicaBotaoAteAcharEClica("botaoFechar.png")
    print()