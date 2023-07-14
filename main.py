import numpy as np
import pandas as pd
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import StringProperty
from kivy.utils import platform
from kivymd.app import MDApp
from kivymd.font_definitions import theme_font_styles
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.floatlayout import MDFloatLayout
from kivymd.uix.label import MDLabel
from kivymd.uix.list import OneLineListItem
from kivymd.uix.menu import MDDropdownMenu


def ler_excel(item, nome):
    if platform == 'android':
        from android.permissions import request_permissions, Permission
        """Solicita a permisão no Android"""
        try:
            request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
            try:
                df = pd.read_excel('ficha.xlsm', sheet_name=nome, engine='openpyxl',
                                   usecols=[0, 1, 3, 5], skiprows=5)
                # 'ENTREGAS DE EPIS (FICHA).xlsm'
                # '/storage/emulated/0/Download/ficha.xlsm'
            except FileNotFoundError:
                AppSSO.erro_arquivo = True
                print('arquivo excel não encontrado')
            except ImportError:
                AppSSO.erro_important = True
                print('Erro de importação de openpyxl')
            else:
                print(df)
                """Verifica quais os valores na coluna Qtd do dataframe são
                maiores que zero e retorna um valor booleano True ou False"""
                filtro = df['Qtd'] > 0
                print(filtro)
                """Exibe as linhas do dataframe nos valores os valores true"""
                dado = df[filtro]
                print(dado)
                """Exibe a quantidade de linhas do dataframe"""
                linhas = dado.shape[0]
                print(linhas)
                """Substituem os valores NaN por traço e outros valores específicos"""
                dado3 = dado.replace(regex={np.NaN: '-', 'LUVA PARA PROTEÇÃO CONTRA AGENTES MECÂNICOS': 'LUVA AG MEC'})
                print(dado3)
                return dado3, linhas
        except PermissionError:
            print('Permissão Negada')
    """ Realiza a leitura do arquivo Excel"""


class IconListItem(OneLineListItem):
    icon = StringProperty()


class AppSSO(MDApp):
    abas_excel = ['COLABORADORES']

    """Lista criada após realizada a pesquisa pelo que foi digitado"""
    encontrados = []
    erro_arquivo = False  # Variável pra erro na leitura de arquivo
    erro_important = False
    data_tables = None
    title = 'Controle de Epi'
    icon = 'icon.png'
    colaborador = ""
    button = None
    numero_linhas = 0

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.label = None
        self.label2 = None
        self.menu = None
        self.menu_items = None
        self.screen = Builder.load_file('appsso.kv')

    def on_enter(self):
        print('A função On enter foi chamado')

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
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            """Solicita a permisão no Android"""
            try:
                request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
            except PermissionError:
                print('Permissão Negada')

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
            print(self.encontrados)
            print(texto)

        """Busca o nome do colaborador na planilha após IconButton ser clicado"""

        def on_text(valor, valor2, search=False):
            """Se Search for True (o iconButton foi clicado) e executa o bloco if"""
            if search:
                try:
                    self.add_row(self.colaborador)
                except ValueError:
                    self.screen.ids.text_field_error.error = True
                    print("erro ao adicionar linha!!!!")
            else:
                print("Else de on text foi acionado")
            print(self, valor, self.colaborador, search)

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
            pos_hint={"center_y": 0.6, "center_x": 0.5},  # Posicionamento da data table
            size_hint=(0.9, 0.8),  # Tamanho da área da tabela
            use_pagination=True,  # Cria a varias páginas de dados
            rows_num=15,  # Números de linhas a exibir em cada páginas
            column_data=[
                ("Qtd", dp(8)),
                ("EPI", dp(25)),
                ("C.A", dp(15)),
                ("Entrega", dp(15)),
            ],

        )
        """Cria um Label para inserir o texto de erro de importação"""
        self.label2 = MDLabel(
            text="OPENPYXL NÃO IMPORTADO!",
            pos_hint={"center_x": .6, "center_y": .1},
            theme_text_color="Error",
            font_style=theme_font_styles[6],
        )

        """Adiciona o label no layout se a variável erro_arquivo2 for True"""
        if self.erro_important:
            layout.add_widget(self.label2)

        """Cria um Label para inserir o texto de erro"""
        self.label = MDLabel(
            text="ARQUIVO EXCEL NÃO LOCALIZADO!",
            pos_hint={"center_x": .6, "center_y": .1},
            theme_text_color="Error",
            font_style=theme_font_styles[6],
        )
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            """Solicita a permisão no Android"""
            try:
                request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
                try:
                    """Chama a função ler_excel a primeira vez pra capturar a quantidade
                     de linha do dataframe"""
                    dado_df, self.numero_linhas = ler_excel(0, "AMANDA")
                except TypeError:
                    print('erro ao ler excel para obter numero de linhas')
                """Adiciona o label no layout se a variável erro_arquivo for True"""
            except PermissionError:
                print('Permissão Negada')

        if self.erro_arquivo:
            layout.add_widget(self.label)
        """Adiciona o data_tables ao layout"""
        layout.add_widget(self.data_tables)
        """Adiciona o Screen ao layout"""
        layout.add_widget(self.screen)

        return layout

    def set_item(self, text__item):
        print("set item foi chamado {0}".format(text__item))
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
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            """Solicita a permisão no Android"""
            try:
                request_permissions([Permission.READ_EXTERNAL_STORAGE, Permission.WRITE_EXTERNAL_STORAGE])
                try:
                    """Chama a função ler_excel a primeira vez pra capturar a quantidade
                     de linha do dataframe"""
                    dado_df, self.numero_linhas = ler_excel(0, nome)
                    print(self.numero_linhas)
                except TypeError:
                    print('Erro ao ler excel pra obter o número de linhas')
                else:
                    """Utiliza o for para ler linha por linha do dataframe e adicionar
                    linha no MD dataTable"""
                    for i in range(self.numero_linhas):
                        """Exibe a linha específica do DataFrame"""
                        dado5 = dado_df.iloc[i]
                        print(dado5)
                        self.data_tables.add_row((str(dado5[0]).replace('.0', ''), dado5[1].capitalize(),
                                                  str(dado5[2]).replace('.0', ''), dado5[3].strftime("%d/%m/%y"),))
                        print(type(dado5[0]))
            except PermissionError:
                print('Permissão Negada')


AppSSO().run()
