import tkinter as tk
from tkinter import font

from toolbox import totvsTools

class ToolboxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Toolbox")
        self.root.geometry("600x400")

        self.setup()

    def setup(self):
        main_frame = tk.Frame(self.root, pady=10)
        
        # main_frame = tk.Label(self.root, text="Toolbox")
        main_frame.pack(fill=tk.BOTH, expand=True)
        title = tk.Label(main_frame, text="Toolbox", font=(font.nametofont("TkDefaultFont").actual(), 24))
        title.pack(fill=tk.X)

        optionSelected = tk.StringVar()
        options = (("Cancelar requisições", "cancelaReq"),
                   ("Alterar aprovador", "alteraAprovador"),
                   ("Suspender requisições", "suspendeReq"), 
                   ("Excluir Vinculo Salarial", "excluiVinculo"),
                   ("Abrir alteração de dados funcionais", "abreAlteracaoDados"),
                   ("Adicionar motivo de mudança salarial", "addMotivoMudancaSalarial"),
                   ("Abrir aumento de quadro", "abreAumentoDeQuadro"),
                   ("Inserir Supervisor na Tabela-Chefe", "insereSupervisor"))


        function_title = tk.Label(main_frame, text="Selecione uma função para executar")
        function_title.pack(padx=10, pady=10, anchor="nw")
        functions_frame = tk.Frame(main_frame)
        functions_frame.pack(fill=tk.X)
        # cancelaRequisicao_rBtn = tk.Radiobutton(functions_frame, text="Cancelar Requisições", value="cancelaReq", variable=optionSelected)
        # alterarAprovador_rBtn = tk.Radiobutton(functions_frame, text="Alterar aprovador", value="alteraAprovador", variable=optionSelected)
        # cancelaRequisicao_rBtn.pack()
        # alterarAprovador_rBtn.pack()
        for option in options:
            r = tk.Radiobutton(functions_frame, text=option[0], value=option[1], variable=optionSelected)
            r.pack(side=tk.TOP, anchor="w", padx=15)

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=20, anchor="e")

        run_local_btn = tk.Button(btn_frame, text="Executar localmente")
        run_local_btn.pack(side=tk.RIGHT, padx=5, anchor="e")

        run_cloud_btn = tk.Button(btn_frame, text="Executar na máquina")
        run_cloud_btn.pack(side=tk.RIGHT, padx=5, anchor="e")

        run_btn = tk.Button(btn_frame, text="Atualizar Diaria", bg="#ffae00", command=self.exportaDiaria)
        run_btn.pack(side=tk.RIGHT, padx=5, anchor="e")

    def exportaDiaria(self):
        # toolboxWin = pygetwindow.getWindowsWithTitle("TOTVS Linha RM - Varejo  Alias: CORPORERM_RH | 1-DROGARIA ARAUJO S/A")[0]
        self.root.iconify()
        totvsTools.exportaDiaria()
        # print("exporta Diaria")