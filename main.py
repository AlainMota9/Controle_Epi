import numpy as np
import pandas as pd
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import StringProperty
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.list import OneLineListItem
from kivymd.uix.menu import MDDropdownMenu


def ler_excel(item, nome):
    """ Realiza a leitura do arquivo Excel"""
    df = pd.read_excel('ENTREGAS DE EPIS (FICHA).xlsm', sheet_name=nome, usecols=[0, 1, 3, 5], skiprows=5)
    """Verifica quais os valores na coluna Qtd do dataframe são
    maiores que zero e retorna um valor booleano True ou False"""
    filtro = df['Qtd'] > 0
    """Exibe as linhas do dataframe nos valores os valores true"""
    dado = df[filtro]
    """Exibe a quantidade de linhas do dataframe"""
    linhas = dado.shape[0]
    """Substituem os valores NaN por traço e outros valores específicos"""
    dado3 = dado.replace(regex={np.NaN: '-', 'LUVA PARA PROTEÇÃO CONTRA AGENTES MECÂNICOS': 'LUVA AG MEC'})
    return dado3, linhas


class IconListItem(OneLineListItem):
    icon = StringProperty()


class AppSSO(MDApp):
    abas_excel = ['Colaborador1', 'Colaborador2', 'Colaborador3', 'Colaborador4']

    """Lista criada após realizada a pesquisa pelo que foi digitado"""
    encontrados = []

    data_tables = None
    title = 'Controle de Epi'
    icon = 'icon.png'
    colaborador = ""
    button = None
    numero_linhas = 0

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.menu = None
        self.menu_items = None
        self.screen = Builder.load_file('appsso.kv')

    def on_enter(self):

        self.menu_items = [
            {
                "viewclass": "IconListItem",
                "height": dp(56),
                "text": self.encontrados[i],
                "on_release":
                    lambda x=self.encontrados[i]: self.set_item(text__item=x),
            } for i in range(len(self.encontrados))]

        self.menu = MDDropdownMenu(
            caller=self.screen.ids.text_field_error,
            items=self.menu_items,
            position="center",
            width_mult=10,
            ver_growth="up",
            elevation=4,
            radius=[24, 24, 24, 24],
            hor_growth="left",
            border_margin=dp(48),
            max_height=dp(0),
        )
        self.menu.open()

    def build(self):

        """Envia True quando IconButton for clicado"""
        self.screen.ids.icon_button.bind(on_press=lambda x: on_text(None, None, True))

        def insert_text(objeto, texto):
            """converte o dado do text_fiel para string"""
            nome_string = str(texto)
            """converte para maiúscula a string"""
            nome_upper = nome_string.upper()
            """Enviar a string em maiúsculas para o variável colaborador"""
            self.colaborador = nome_upper
            """Pesquisa na lista o nome do colaborador conforme digitado text field"""
            self.encontrados = [s for s in self.abas_excel if nome_upper in s]

        """Busca o nome do colaborador na planilha após IconButton ser clicado"""
        def on_text(valor, valor2, search=False):
            """Se Search for True (o iconButton foi clicado) e executa o bloco if"""
            if search:
                try:
                    self.add_row(self.colaborador)
                except ValueError:
                    self.screen.ids.text_field_error.error = True

        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Orange"
        """Captura o texto digitado no text_field"""
        self.screen.ids.text_field_error.bind(
            text=insert_text,
        )

        """Cria o root layout"""
        layout = MDFloatLayout()

        """Criará Tabela"""
        self.data_tables = MDDataTable(
            pos_hint={"center_y": 0.6, "center_x": 0.5},
            size_hint=(0.9, 0.8),
            use_pagination=True,
            rows_num=15,
            column_data=[
                ("Qtd", dp(8)),
                ("EPI", dp(25)),
                ("C.A", dp(15)),
                ("Entrega", dp(15)),
            ],

        )

        """Adiciona o data_tables ao layout"""
        layout.add_widget(self.data_tables)
        """Adiciona o Screen ao layout"""
        layout.add_widget(self.screen)
        """Adiciona as linhas com nome do colaborador específico"""

        return layout

    def set_item(self, text__item):
        """Envia o text__item para variável colaborador"""
        self.colaborador = text__item
        """Insere o nome selecionado em text_field"""
        self.screen.ids.text_field_error.text = text__item
        """Fecha o menu_itens"""
        self.menu.dismiss()

    def add_row(self, nome) -> None:
        """Remove todas das linhas do data tables"""
        if len(self.data_tables.row_data) > 1:
            for i in range(self.numero_linhas):
                self.data_tables.remove_row(self.data_tables.row_data[-1])
        """Chama a função ler_excel a primeira vez pra capturar a quantidade
         de linha do dataframe"""
        dado_df, self.numero_linhas = ler_excel(0, nome)
        """Utiliza o for para ler linha por linha do dataframe e adicionar
        linha no MDdataTable"""
        for i in range(self.numero_linhas):
            """Exibe a linha específica do DataFrame"""
            dado5 = dado_df.iloc[i]
            self.data_tables.add_row((str(dado5[0]).replace('.0', ''), dado5[1].capitalize(), str(dado5[2]).replace('.0', ''), dado5[3].strftime("%d/%m/%y"),))


AppSSO().run()
