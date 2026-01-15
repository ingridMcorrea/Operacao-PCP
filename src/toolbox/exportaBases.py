import pyautogui
import time
import pyperclip
import webbrowser
from datetime import date
from dateutil.relativedelta import relativedelta

from toolbox.core import buscaClicaBotao
from toolbox.core import ficaBuscandoClicaBotaoAteAcharEClica
from toolbox.core import tentaAcharImagem


# Definir todas as bases por aqui, talvez num arquivo XLSX que ele le
class DetalhadamentoBase:
    def __init__(self, nomeBase, codigo, nomeArquivo):
        self.nomeBase = nomeBase
        self.codigo = codigo
        self.nomeArquivo = nomeArquivo

def exportaCadastros():
    setupConsultas()
    procuraSentenca("PCP.01.099", True)
    exportaSentencaV2("PCP.01.099 - Centro de Custo SAP x Centro de CustoTOTVS.XLSX")

    procuraSentenca("PCP.01.101")
    exportaSentencaV2("PCP.01.101 - Todos os colaboradores ativos com o código de equipe.XLSX", True)

    procuraSentenca("PCP.01.106")
    exportaSentencaV2("PCP.01.106 - Histórico de Exclusão de Requisições.xlsx")

    ## Precisa de um pequeno Sleep
    procuraSentenca("PCP.02.002")
    time.sleep(3)
    exportaSentencaV2("PCP.02.002 - Demitidos.XLSX", False, False, False)

    procuraSentenca("PCP.01.049")
    exportaSentencaV2("PCP.01.049 - Detalhamento Seções.XLSX", True)

    procuraSentenca("PCP.02.007")
    pyautogui.press("right")
    hoje = date.today()
    pyautogui.write(f"01/01/{hoje.year}", interval=0.016)
    pyautogui.press("down")
    pyautogui.write(f"31/12/{hoje.year}", interval=0.016)
    pyautogui.press("tab")
    exportaSentencaV2("PCP.02.007 - Admitidos no Mês com Requisição.XLSX", False, True)

    procuraSentenca("PCP.01.015")
    exportaSentencaV2("PCP.01.015 - Atendente e Atendente Substituto.XLSX")

    procuraSentenca("PCP.01.104")
    exportaSentencaV2("PCP.01.104 - Todos os Chefes e Supervisores das seções ativas.XLSX")

    procuraSentenca("PCP.01.046")
    exportaSentencaV2("PCP.01.046 - Descrição Seções e Seções Superiores.XLSX")

    procuraSentenca("PCP.01.085")
    exportaSentencaV2("PCP.01.085 - Função x Nível.XLSX")

    procuraSentenca("PCP.01.013.1")
    exportaSentencaV2("PCP.01.013.1 - Somente Tabela de Chefes (PSUBSTCHEFE).XLSX")

    procuraSentenca("PCP.01.013.3")
    exportaSentencaV2("PCP.01.013.3 - Chefes e Supervisores Seção e Função.XLSX")

    procuraSentenca("PCP.03.007")
    exportaSentencaV2("PCP.03.007 - Sindicatos.XLSX")

    procuraSentenca("PCP 01.079")
    exportaSentencaV2("PCP.01.079 - Cadastro de Função.XLSX")

    procuraSentenca("PCP.01.093")
    exportaSentencaV2("PCP.01.093 - Cargo e agrupamento PPR.XLSX")

    atualizaGoogleSheets("https://docs.google.com/spreadsheets/d/1oqeE36_FSK7VYgAGSx45ZuwDGU6O1PFmeGpNB9B1NII/edit?gid=664006522#gid=664006522")

def exportaDiaria():
    setupConsultas()

    procuraSentenca("PCP.01.006", True)
    ## Insere hoje menos 8 meses como parametro da sentença
    pyautogui.press("right")
    hoje = date.today()
    hojeMenosOitoMeses = hoje + relativedelta(months=-8)
    formatted_date = hojeMenosOitoMeses.strftime("%m/%d/%Y")
    pyautogui.write(formatted_date)
    pyautogui.press("tab")
    time.sleep(0.3)

    exportaSentencaV2("PCP.01.006 - Datas de Retorno de Licença.XLSX", False, True)           

    procuraSentenca("PCP.01.001")
    exportaSentencaV2("PCP.01.001 - Requisições de Substituição e Aumento de Quadro Abertas no Sistema.XLSX")

    procuraSentenca("PCP.01.002")
    exportaSentencaV2("PCP.01.002 - Funcionários não Desligados.XLSX", True)

    ## NAO FILTRA SOZINHO, demora "muito" para abrir time sleep para isso
    procuraSentenca("PCP.01.011")
    time.sleep(3)
    exportaSentencaV2("PCP.01.011 - Férias Abertas.XLSX", False, False, False)

    ## Filtra entre os hoje menos tres meses e hoje
    procuraSentenca("PCP.02.001")
    pyautogui.press("right")
    hojeMenosTresMeses = hoje + relativedelta(months=-3)
    formatted_date = hojeMenosTresMeses.strftime("%m/%d/%Y")
    pyautogui.write(formatted_date)
    pyautogui.press("down")
    formatted_date = hoje.strftime("%m/%d/%Y")
    pyautogui.write(formatted_date)
    pyautogui.press("tab")
    exportaSentencaV2("PCP.02.001 - Atestados", True, True, True)

    procuraSentenca("PCP.01.031")
    exportaSentencaV2("PCP.01.031 - Checklist (Movimentações).XLSX")

    procuraSentenca("PCP.01.032")
    exportaSentencaV2("PCP.01.032 - Requisições de Desligamento Abertas.XLSX")

    atualizaGoogleSheets("https://docs.google.com/spreadsheets/d/1crwocHnyA8yrE66Jlv7Q9YSxT_uaDahUrUfar7kn_X4/edit?gid=1926191882")

def exportaHorarios():
    setupConsultas()
    procuraSentenca("PCP.01.120", True)
    exportaSentencaV2("PCP.01.120 - Log de Importação WFM Ultimos 3 dias - Período Ativo.XLSX")

    atualizaGoogleSheets("https://docs.google.com/spreadsheets/d/1gmsPYOM9kdf7PZgLxtn_Hb_E01I_rw0N-P7ph-4Q2hE/edit?gid=0#gid=0")

def exportaTabelaSalarial():
    setupConsultas()
    procuraSentenca("PCP.02.038", True)
    exportaSentencaV2("PCP.02.038 - Tabelas Salariais com suas vinculações de seção e função.XLSX", True)

    atualizaGoogleSheets("https://docs.google.com/spreadsheets/d/1fashYFlb0fa9pwwX5VPfxdG_gsMiK7Sn6R1qFPHYQVw/edit?gid=0#gid=0")

def exportaSentenca(nomeSentenca, fotoExemploNome, maisDe10KLinhas = False):
    # pyautogui.PAUSE = 0.3
    buscaClicaBotao(fotoExemploNome, 4, True)
    time.sleep(0.2)
    ficaBuscandoClicaBotaoAteAcharEClica("removerFiltrosSentenca.png", 2)
    time.sleep(0.2)
    buscaClicaBotao("exportar.png", 1)
    buscaClicaBotao("exportarXLSX.png", 2)
    buscaClicaBotao("avancar.png")
    pyperclip.copy(nomeSentenca)
    time.sleep(0.8)
    pyautogui.hotkey("ctrl", "v")
    # buscaClicaBotao("salvar.png")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")

    time.sleep(0.5)
    buscaClicaBotao("sim.png")
    if maisDe10KLinhas:
        ficaBuscandoClicaBotaoAteAcharEClica("simComOutline.png")

    # maior10KLinhas = tentaAcharImagem("simboloDuvida.png", 2)
    # if maior10KLinhas == True:
    #     pyautogui.press("enter")
    

    time.sleep(0.8)
    ficaBuscandoClicaBotaoAteAcharEClica("nao.png")
    pyautogui.hotkey("ctrl", "w")
    time.sleep(0.2)
    buscaClicaBotao("consultasSQL.png")
    # pyautogui.PAUSE = 0.2

def exportaSentencaV2(nomeSentenca, maisDe10KLinhas = False, executarPrimeiro = False, removeFiltro = True):

    if executarPrimeiro:
        buscaClicaBotao("executar.png", 1)

    time.sleep(0.2)
    if removeFiltro:
        ficaBuscandoClicaBotaoAteAcharEClica("removerFiltrosSentenca.png", 2)
        time.sleep(0.2)
    
    buscaClicaBotao("exportar.png", 2)
    buscaClicaBotao("exportarXLSX.png", 2)
    buscaClicaBotao("avancar.png")
    pyperclip.copy(nomeSentenca)
    time.sleep(0.8)
    pyautogui.hotkey("ctrl", "v")
    # buscaClicaBotao("salvar.png")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")

    time.sleep(0.5)
    buscaClicaBotao("sim.png")
    if maisDe10KLinhas:
        ficaBuscandoClicaBotaoAteAcharEClica("simComOutline.png")

    # maior10KLinhas = tentaAcharImagem("simboloDuvida.png", 2)
    # if maior10KLinhas == True:
    #     pyautogui.press("enter")
    

    time.sleep(0.8)
    ficaBuscandoClicaBotaoAteAcharEClica("nao.png")
    pyautogui.hotkey("ctrl", "w")
    time.sleep(0.2)
    buscaClicaBotao("consultasSQL.png")
    # pyautogui.PAUSE = 0.2

def procuraSentenca(codigoSentenca, ocultarRegistros = False):
    if ocultarRegistros == False:
        pyautogui.press("tab")
        pyautogui.press("tab")
    pyperclip.copy(codigoSentenca)
    pyautogui.hotkey("ctrl", "v")
    if ocultarRegistros == True:
        buscaClicaBotao("ocultarRegistros.png", 1)
    time.sleep(0.3)
    buscaClicaBotao("simboloSQL.png", 1, True)
    time.sleep(0.4)

def setupConsultas():
    time.sleep(0.5)
    pyautogui.PAUSE = 0.1
    pyautogui.hotkey("alt", "g")
    pyautogui.press("s")
    time.sleep(0.5)
    pyautogui.hotkey("v", "i")
    time.sleep(2)

    for i in range(12):
        pyautogui.press("down")
    time.sleep(0.3)
    pyautogui.press("enter")
    pyautogui.press("enter")
    time.sleep(3)

    buscaClicaBotao("lupaSentencas.png", 2)
    time.sleep(0.2)

def atualizaGoogleSheets(googleSheetsUrl):
    pyautogui.hotkey("ctrl", "w")
    time.sleep(1)
    webbrowser.open(googleSheetsUrl)
    time.sleep(15)
    pyautogui.hotkey("ctrl", "alt", "shift", "1")
    pyautogui.hotkey("ctrl", "w")