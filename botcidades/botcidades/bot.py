from botcity.core import DesktopBot
import pandas as pd
import openpyxl
import numpy as np

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *
print("-----------------------------------------------------")
print("         Buscador de dados dos Estados do Brasil     ")
print("-----------------------------------------------------")
class Bot(DesktopBot):
    def action(self, execution=None):
        estado = input("Escreva, por extenso, o nome do estado que deseja procurar: ")

        estados = ['Acre', 'Alagoas', 'Amapá', 'Amazonas', 'Bahia', 'Ceará', 'Distrito Federal', 'Espírito Santo',
                   'Goiás', 'Maranhão', 'Mato Grosso', 'Mato Grosso do Sul', 'Minas Gerais', 'Pará', 'Paraíba',
                   'Paraná', 'Pernambuco', 'Piauí', 'Rio de Janeiro', 'Rio Grande do Norte', 'Rio Grande do Sul',
                   'Rondônia', 'Roraima', 'Santa Catarina', 'São Paulo', 'Sergipe', 'Tocantins']

        populacao_estimada = 0
        populacao_estimada2 = 0
        populacao_estimada3 = 0
        idh = 0
        idh2 = 0
        idh3 = 0

        if any(estado.lower() == estado.lower() for estado in estados):
            # Abre o portal Cidades do IBGE
            self.browse("https://cidades.ibge.gov.br/")

            if not self.find( "lupa", matching=0.97, waiting_time=10000):
                self.not_found("lupa")
            self.click()

            self.paste(estado)

            if not self.find( "selecionarEstado", matching=0.97, waiting_time=10000):
                self.not_found("selecionarEstado")
            self.click_relative(950, 210)

            if not self.find( "gentilico", matching=0.97, waiting_time=10000):
                self.not_found("gentilico")
            self.double_click_relative(26, 25)
            self.control_c()
            gentilico = self.get_clipboard()
            print(gentilico)

            if not self.find( "capital", matching=0.97, waiting_time=10000):
                self.not_found("capital")
            self.triple_click_relative(12, 21)
            self.control_c()
            capital = self.get_clipboard()
            print(capital)

            if not self.find( "governador", matching=0.97, waiting_time=10000):
                self.not_found("governador")
            self.triple_click_relative(27, 24)
            self.control_c()
            governador = self.get_clipboard()
            print(governador)

            if not self.find( "popEstimada", matching=0.97, waiting_time=100):
                if not self.find( "popEstimada2", matching=0.97, waiting_time=100):
                    if not self.find( "popEstimada3", matching=0.97, waiting_time=10000):
                        self.not_found("popEstimada3")
                    self.double_click_relative(289, 5)
                    self.control_c()
                    populacao_estimada3 = self.get_clipboard()
                    print(populacao_estimada3)
                else:
                    self.double_click_relative(287, 2)
                    self.control_c()
                    populacao_estimada2 = self.get_clipboard()
                    print(populacao_estimada2)
            else:
                self.double_click_relative(290, -5)
                self.control_c()
                populacao_estimada = self.get_clipboard()
                print(populacao_estimada)

            if not self.find( "economia", matching=0.97, waiting_time=10000):
                if not self.find( "economia2", matching=0.97, waiting_time=10000):
                    if not self.find( "economia3", matching=0.97, waiting_time=10000):
                        self.not_found("economia3")
                    self.click()
                else:
                    self.click()
            else:
                self.click()

            if not self.find( "idh", matching=0.97, waiting_time=10000):
                if not self.find( "idh2", matching=0.97, waiting_time=10000):
                    if not self.find("idh3", matching=0.97, waiting_time=10000):
                        self.not_found("idh3")
                    self.double_click_relative(278, 16)
                    self.control_c()
                    idh3 = self.get_clipboard()
                    print(idh3)
                else:
                    self.double_click_relative(280, 16)
                    self.control_c()
                    idh2 = self.get_clipboard()
                    print(idh2)
            else:
                self.double_click_relative(273, 16)
                self.control_c()
                idh = self.get_clipboard()
                print(idh)

            a = {'Estado': [estado],
                 'Gentílico': [gentilico],
                 'Capital': [capital],
                 'Governador': [governador],
                 'PopulaçãoEstimada': [populacao_estimada],
                 'IDH': [idh]}

            a2 = {'Estado': [estado],
                 'Gentílico': [gentilico],
                 'Capital': [capital],
                 'Governador': [governador],
                 'PopulaçãoEstimada': [populacao_estimada],
                 'IDH': [idh2]}

            a3 = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada],
                  'IDH': [idh3]}

            b = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada2],
                  'IDH': [idh]}

            b2 = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada2],
                  'IDH': [idh2]}

            b3 = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada2],
                  'IDH': [idh3]}

            c = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada3],
                  'IDH': [idh]}

            c2 = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada3],
                  'IDH': [idh2]}

            c3 = {'Estado': [estado],
                  'Gentílico': [gentilico],
                  'Capital': [capital],
                  'Governador': [governador],
                  'PopulaçãoEstimada': [populacao_estimada3],
                  'IDH': [idh3]}


            tabela1 = pd.DataFrame(data=a)
            tabela2 = pd.DataFrame(data=a2)
            tabela3 = pd.DataFrame(data=a3)
            tabela4 = pd.DataFrame(data=b)
            tabela5 = pd.DataFrame(data=b2)
            tabela6 = pd.DataFrame(data=b3)
            tabela7 = pd.DataFrame(data=c)
            tabela8 = pd.DataFrame(data=c2)
            tabela9 = pd.DataFrame(data=c3)

            tabela1.to_excel('tabela_principal.xlsx', index=False)
            tabela1.to_excel('tabela1.xlsx', index=False)
            tabela2.to_excel('tabela2.xlsx', index=False)
            tabela3.to_excel('tabela3.xlsx', index=False)
            tabela4.to_excel('tabela4.xlsx', index=False)
            tabela5.to_excel('tabela5.xlsx', index=False)
            tabela6.to_excel('tabela6.xlsx', index=False)
            tabela7.to_excel('tabela7.xlsx', index=False)
            tabela8.to_excel('tabela8.xlsx', index=False)
            tabela9.to_excel('tabela9.xlsx', index=False)

            # Carregando a tabela principal
            df_principal = pd.read_excel('tabela_principal.xlsx')

            # Carregando as outras tabelas
            df_tabela1 = pd.read_excel('tabela1.xlsx')
            df_tabela2 = pd.read_excel('tabela2.xlsx')
            df_tabela3 = pd.read_excel('tabela3.xlsx')
            df_tabela4 = pd.read_excel('tabela4.xlsx')
            df_tabela5 = pd.read_excel('tabela5.xlsx')
            df_tabela6 = pd.read_excel('tabela6.xlsx')
            df_tabela7 = pd.read_excel('tabela7.xlsx')
            df_tabela8 = pd.read_excel('tabela8.xlsx')
            df_tabela9 = pd.read_excel('tabela9.xlsx')

            # Verificando e substituindo os dados
            for i in range(len(df_principal)):
                for j in range(len(df_principal.columns)):
                    if df_principal.iloc[i, j] == 0 or pd.isnull(df_principal.iloc[i, j]):
                        for df in [df_tabela1, df_tabela2, df_tabela3, df_tabela4, df_tabela5, df_tabela6, df_tabela7,
                                   df_tabela8, df_tabela9]:
                            if df.iloc[i, j] != 0 and not pd.isnull(df.iloc[i, j]):
                                df_principal.iloc[i, j] = df.iloc[i, j]

            # Salvando a tabela principal com as substituições
            df_principal.to_excel('tabela_pfinal.xlsx', index=False)

            def replay_main():
                print("1 - Escolher outro Estado")
                print("2 - Sair")
                opcao = int(input("Informe a sua opção: "))

                if opcao == 1:
                    Bot.main()
                elif opcao == 2:
                    exit()
                else:
                    print("Informe uma opção válida, informe 1 ou 2")
                    replay_main()
            replay_main()
        else:
            print("Informe corretamente o nome do Estado brasileiro que você deseja acessar.")

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    Bot.main()



