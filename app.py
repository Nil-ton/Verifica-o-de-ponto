import pandas as pd
import re
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

path_xlsx = 'dados_caminhoes.xlsx'

class GUI:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.ordenacao = {
            "Número do Registro": True,
            "Data": True,
            "Placa do Caminhão": True,
            "Q. GB": True
        }
        self.configurar_tamanho_janela()
        self.configurar_grid()
        self.criar_widgets()

    def configurar_tamanho_janela(self):
        largura = 800
        altura = 600
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        pos_x = (screen_width - largura) // 2
        pos_y = (screen_height - altura) // 2
        self.root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
        self.root.title("Grupo Compel")
        root.wm_iconbitmap("icon.ico")

    def configurar_grid(self):
        self.root.grid_rowconfigure(1, weight=0)  # Entrada de Placa
        self.root.grid_rowconfigure(2, weight=0)  # Entrada de GB
        self.root.grid_rowconfigure(3, weight=0)  # Botão de Adicionar Dados
        self.root.grid_rowconfigure(4, weight=1)  # Tabela de Dados (expandível)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=2)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=1)
        self.root.grid_columnconfigure(4, weight=1)  # Coluna intermediária para o espaçamento

    def criar_widgets(self):
        self.criar_entradas_placa_gb()
        self.criar_botoes()
        self.criar_tabela()
        self.carregar_dados_planilha()

    def criar_entradas_placa_gb(self):
        # Entrada da placa
        self.label_placa = ttk.Label(self.root, text="Placa do Caminhão:")
        self.label_placa.grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_placa = ttk.Entry(self.root)
        self.entry_placa.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.entry_placa.bind("<KeyPress>", self.converter_maiúsculas)

        # Configurar o evento de pressionamento da tecla Enter no campo de placa
        self.entry_placa.bind("<Return>", lambda event: self.adicionar_dados())

        # Entrada da quantidade de GB
        self.label_gb = ttk.Label(self.root, text="Quantidade de GB:")
        self.label_gb.grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_gb = ttk.Entry(self.root)
        self.entry_gb.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Configurar o evento de pressionamento da tecla Enter no campo de GB
        self.entry_gb.bind("<Return>", lambda event: self.adicionar_dados())

    def criar_botoes(self):
        # Chama os métodos individuais para criar os botões
        self.criar_botao_adicionar()
        self.criar_botao_pesquisar()
        self.criar_botao_excluir()

    def criar_botao_adicionar(self):
        # Botão para adicionar dados
        self.botao_adicionar = ttk.Button(self.root, text="Adicionar Dados", command=self.adicionar_dados)
        self.botao_adicionar.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

    def criar_botao_pesquisar(self):
        # Campo de pesquisa de placa
        self.label_pesquisa = ttk.Label(self.root, text="Pesquisar por Placa:")
        self.label_pesquisa.grid(row=1, column=2, padx=10, pady=5, sticky="e")
        self.entry_pesquisa = ttk.Entry(self.root)
        self.entry_pesquisa.grid(row=1, column=3, padx=10, pady=5, sticky="ew")

        # Configurar o evento de pressionamento da tecla Enter no campo de pesquisa
        self.entry_pesquisa.bind("<Return>", lambda event: self.pesquisar_placa())

        # Botão para pesquisar
        self.botao_pesquisar = ttk.Button(self.root, text="Pesquisar", command=self.pesquisar_placa)
        self.botao_pesquisar.grid(row=2, column=3, padx=10, pady=5, sticky="ew")

    def criar_botao_excluir(self):
        # Botão para excluir dados
        self.botao_excluir = ttk.Button(self.root, text="Excluir Dados", command=self.excluir_dados)
        self.botao_excluir.grid(row=3, column=3, padx=10, pady=10, sticky="ew")

    def criar_tabela(self):
        self.tree_frame = ttk.Frame(self.root)
        self.tree_frame.grid(row=4, column=0, columnspan=4, pady=10, sticky="nsew")

        self.tree = ttk.Treeview(self.tree_frame, columns=("Número do Registro", "Data", "Placa do Caminhão", "Q. GB"), show="headings")
        self.tree.heading("Número do Registro", text="Número do Registro", command=lambda: self.ordenar_coluna("Número do Registro"))
        self.tree.heading("Data", text="Data", command=lambda: self.ordenar_coluna("Data"))
        self.tree.heading("Placa do Caminhão", text="Placa do Caminhão", command=lambda: self.ordenar_coluna("Placa do Caminhão"))
        self.tree.heading("Q. GB", text="Q. GB", command=lambda: self.ordenar_coluna("Q. GB"))
        self.tree.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")

        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Delete>", self.excluir_dados_evento)

    def excluir_dados_evento(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        confirm = messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir este registro?")
        if confirm:
            numero_registro = self.tree.item(selected_item)["values"][0]
            self.app.excluir_registro(numero_registro)
            self.carregar_dados_planilha()
            messagebox.showinfo("Sucesso", "Registro excluído com sucesso.")

    def converter_maiúsculas(self, event):
        char = event.char.upper()
        if char != event.char:
            self.entry_placa.insert(self.entry_placa.index(tk.INSERT), char)
            return "break"

    def adicionar_dados(self):
        placa = self.entry_placa.get().strip()
        gb = self.entry_gb.get().strip()

        if not placa or not gb:
            messagebox.showerror("Erro", "Preencha todos os campos.")
            return

        try:
            self.app.validar_placa(placa)
            self.app.validar_gb(gb)
            dados = [{'Número do Registro': self.app.gerar_numero_registro(),
                      'Data': self.app.data_atual(),
                      'Placa do Caminhão': placa,
                      'Q. GB': int(gb)}]
            self.app.adicionar_dados(dados)
            self.entry_placa.delete(0, tk.END)
            self.entry_gb.delete(0, tk.END)
            messagebox.showinfo("Sucesso", "Dados adicionados com sucesso.")
            self.carregar_dados_planilha()
        except ValueError as e:
            messagebox.showerror("Erro", str(e))

    def excluir_dados(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um registro para excluir.")
            return
        confirm = messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir este registro?")
        if confirm:
            numero_registro = self.tree.item(selected_item)["values"][0]
            self.app.excluir_registro(numero_registro)
            self.carregar_dados_planilha()
            messagebox.showinfo("Sucesso", "Registro excluído com sucesso.")

    def carregar_dados_planilha(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        df = self.app.carregar_data_frame()

        if df is not None and not df.empty:
            for idx, (index, row) in enumerate(df.iterrows()):
                data_formatada = datetime.strptime(row["Data"], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
                tag = "linha_clara" if idx % 2 == 0 else "linha_escura"
                self.tree.insert("", "end", values=(row.name, data_formatada, row["Placa do Caminhão"], row["Q. GB"]), tags=(tag,))

            self.tree.tag_configure("linha_clara", background="#f2f2f2")
            self.tree.tag_configure("linha_escura", background="#e0e0e0")

    def pesquisar_placa(self):
        placa_pesquisada = self.entry_pesquisa.get().strip().upper()
        if placa_pesquisada:
            df = self.app.carregar_data_frame()
            if df is not None and not df.empty:
                df_filtrado = df[df["Placa do Caminhão"].str.contains(placa_pesquisada, case=False, na=False)]
                for row in self.tree.get_children():
                    self.tree.delete(row)
                for _, row in df_filtrado.iterrows():
                    self.tree.insert("", "end", values=(row.name, row["Data"], row["Placa do Caminhão"], row["Q. GB"]))
        else:
            self.carregar_dados_planilha()

    def ordenar_coluna(self, coluna):
        ordem_crescente = self.ordenacao[coluna]
        df = self.app.carregar_data_frame()

        if df is not None and not df.empty:
            df_sorted = df.sort_values(by=coluna, ascending=ordem_crescente)
            for row in self.tree.get_children():
                self.tree.delete(row)
            for _, row in df_sorted.iterrows():
                self.tree.insert("", "end", values=(row.name, row["Data"], row["Placa do Caminhão"], row["Q. GB"]))

        self.ordenacao[coluna] = not ordem_crescente

class App:
    def criar_dados(self):
        placa = self.pedir_placa()
        gb = self.pedir_gb()
        dados = [{'Número do Registro': self.gerar_numero_registro(), 'Data': self.data_atual(), 'Placa do Caminhão': placa, 'Q. GB': gb}]
        print('Dados Criados com Sucesso.')
        return dados

    def gerar_numero_registro(self):
        df = self.carregar_data_frame()
        if df is not None and not df.empty:
            # Pegando o próximo número de registro
            ultimo_registro = df.index.max()
            return ultimo_registro + 1  # Incrementando o número de registro
        return 1  # Se não houver dados, começa do número 1

    def criar_data_frame(self, dados):
        df = pd.DataFrame(dados)
        df.set_index('Número do Registro', inplace=True)
        return df

    def salvar_data_frame(self, df):
        df.to_excel(path_xlsx, index=True)
        print('Data frame criado com sucesso em documentos/registros/dados_caminhoes.xlsx')

    def data_atual(self):
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def validar_placa(self, placa):
        pattern = r'^[A-Z]{3}\d[A-Z]\d{2}$'
        if re.match(pattern, placa):
            print(f'Placa {placa} válida.')
            return True
        else:
            raise ValueError(f'Placa {placa} inválida. O formato correto é ABC1D23.')

    def pedir_placa(self):
        while True:
            try:
                placa = input('Digite a placa do caminhão (formato ABC1D23): ')
                self.validar_placa(placa)
                return placa
            except ValueError as e:
                print(e)

    def validar_gb(self, gb):
        if gb.isdigit() and int(gb) > 0:
            return True
        else:
            raise ValueError(f'"Q. GB" ({gb}) deve ser um número inteiro positivo.')

    def pedir_gb(self):
        while True:
            try:
                gb = input('Digite a quantidade de GB (número inteiro positivo): ')
                self.validar_gb(gb)
                return int(gb)
            except ValueError as e:
                print(e)

    def carregar_data_frame(self):
        try:
            df = pd.read_excel(path_xlsx, index_col=0)
            print("Arquivo Excel carregado com sucesso.")
            return df
        except FileNotFoundError:
            print("Arquivo não encontrado. Criando novo DataFrame.")
            return None

    def adicionar_dados(self, dados):
        df = self.carregar_data_frame()

        if df is None:
            df = self.criar_data_frame(dados)
        else:
            new_data_frame = self.criar_data_frame(dados)
            df = pd.concat([df.dropna(), new_data_frame.dropna()], ignore_index=True)

        self.salvar_data_frame(df)

    def excluir_registro(self, numero_registro):
        df = self.carregar_data_frame()

        if df is not None and not df.empty:
            # Excluir o registro pelo número
            if numero_registro in df.index:
                df = df.drop(numero_registro)
                self.salvar_data_frame(df)  # Salvar a planilha após a exclusão
            else:
                messagebox.showwarning("Erro", "Registro não encontrado.")
        else:
            messagebox.showwarning("Erro", "Não há dados para excluir.")



# Iniciar o aplicativo
app = App()
root = tk.Tk()
gui = GUI(root, app)
root.mainloop()
