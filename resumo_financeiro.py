import os
from Interface import openFile
import pandas as pd
from datetime import datetime

date = datetime.now().strftime("%d_%m_%Y")

def tratamento_csv():
    # Lê e seleciona as colunas de interesse na variável tabela_tratada
    df = pd.read_csv(openFile())
    tabela_tratada = df.loc[:,['ID da venda','Código','SKU', 'Data', 'Frete', 'Valor Unit.', 'Qtd.', 'Frete Vendedor (-)']]
    tabela_tratada['Frete Vendedor (-)'] = tabela_tratada['Frete Vendedor (-)'].str.replace(',','.').astype(float)
    tabela_tratada['Frete/Quant.'] = tabela_tratada['Frete Vendedor (-)'] / tabela_tratada['Qtd.']

    # Criação do arquivo .XLSX
    sku = tabela_tratada.loc[0,'SKU']
    file = f"{sku}_{datetime.now().day}_{datetime.now().minute}_{datetime.now().second}.xlsx"
    writer = pd.ExcelWriter(file, engine="xlsxwriter")
    # Convert the dataframe to an XlsxWriter Excel object. Turn off the default
    # header and index and skip one row to allow us to insert a user defined
    # header.
    tabela_tratada.to_excel(writer, sheet_name=date, startrow=1, header=False, index=False)
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[date]
    # Get the dimensions of the dataframe.
    (max_row, max_col) = tabela_tratada.shape
    # Create a list of column headers, to use in add_table().
    column_settings = []
    for header in tabela_tratada.columns:
        column_settings.append({'header': header})
    # Add the table.
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
    # Make the columns wider for clarity.
    worksheet.set_column(0, max_col - 1, 12)
    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    os.startfile(file)

tratamento_csv()