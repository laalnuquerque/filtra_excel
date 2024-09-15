from datetime import datetime
from venv import logger
from webbrowser import Error

import pandas as pd
from pathlib import Path
import openpyxl

# Defina o nome dos arquivos de entrada e saída
diretorio_atual = Path(__file__).parent
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
input_file = diretorio_atual.parent / 'arquivo' / 'Estado.xlsx'
output_file = diretorio_atual.parent / 'arquivo' / f"filtro_do_dia_{timestamp}.xlsx"

class FiltraExcel:
    def loadExcel(self, input_file):
        try:
            print("Convertendo Excel em DataFrame")
            #Carregar uma planilha específica de um arquivo Excel
            df_excel = pd.read_excel(input_file, sheet_name='CidadeEstado')
            # df_excel = pd.read_excel(input_file)
            print(df_excel.columns)  # Exibe as colunas disponíveis
            return df_excel
        except Error :
            print(Error)

    def filtraExcel(self, df_excel):
    # Filtrar os dados com base em duas colunas
        try:
            print("Filtrando dados da planilha")
            df_filtrado = df_excel[(df_excel['ESTADO '] == 'São Paulo') & (df_excel['Letra'] == 'H')]
            return df_filtrado
        except Error:
            print(Error)

    def selectColumns(self, df_filtrado):
        try:
            print("Selecionando Colunas para Salvar")
            # Selecionar apenas as colunas desejadas para o novo arquivo
            # Por exemplo: apenas 'Nome' e 'Valor'
            df_selecionado = df_filtrado[['ESTADO ', 'CIDADE']]
            return df_selecionado
        except Error:
            print(Error)

    def salvarExcel(self, df_selecionado):
        try:
            print("Salvando dados filtrados em uma nova planilha")
            # Salvar o DataFrame filtrado e selecionado em um novo arquivo Excel
            df_selecionado.to_excel(output_file, index=False)
        except Error:
            print(Error)

    print(f'Dados filtrados e salvos em {output_file}')

def main():
    # Criar uma instância da classe Greeter
    filtra_excel = FiltraExcel()

    df_excel = filtra_excel.loadExcel(input_file)
    df_filtrado = filtra_excel.filtraExcel(df_excel)
    select_columns = filtra_excel.selectColumns(df_filtrado)
    filtra_excel.salvarExcel(select_columns)

if __name__ == "__main__":
    main()


