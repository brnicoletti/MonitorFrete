import os
import pandas as pd
import subprocess
from datetime import datetime
from Interface import openFile

conta = "Lufiara"
excel_path = "C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"
onlyoffice_path = "C:/Program Files/ONLYOFFICE/DesktopEditors/DesktopEditors.exe"
# Read the exported file
df = pd.read_excel(openFile())

# Converts Column "Data" string in datetime
dates_list = []
hours_list = []
for date in df.Data:
    date_string = date[:19] + " " + date[23:]
    format_string = "%a %b %d %H:%M:%S %Y"
    dt = datetime.strptime(date_string, format_string)
    date = dt.date()
    hour = dt.time()
    dates_list.append(date)
    hours_list.append(hour)
df.Data = dates_list
df["Hour"] = hours_list

new_df = df[["Data", "Hour", "ID Venda Atual", "Título", "ID Anúncio", "SKU", "Valor Frete agora","Diferença" ]].copy()

#Concatena o caractere "#" no início dos dados
##TODO 2: Reescrever a código sem utilizar o for. Simplificar com pandas
new_list = []
for item in new_df["ID Venda Atual"]:
    new_list.append(f'#{item}')
new_df["ID Venda Atual"] = new_list

# Formata os dados das colunas "Valor Frete agora e Diferença"
new_df[['Valor Frete agora', 'Diferença']] = new_df[['Valor Frete agora', 'Diferença']].div(100)
if os.path.isfile('teste.xlsx'):
    old_df = pd.read_excel("teste.xlsx")
    new_df_2 = pd.concat([old_df, new_df])
else:
    new_df_2 = new_df

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(f"{conta}@.xlsx", engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
new_df_2.to_excel(writer, sheet_name='MonitorFrete', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets["MonitorFrete"]

# Add some cell formats.
currencyForm = workbook.add_format({'num_format': 'R$# ##0.00'})
timeForm = workbook.add_format({'num_format': "*hh:mm:ss"})
textForm = workbook.add_format({'num_format': 49})

worksheet.set_column(0,0,10,textForm)
worksheet.set_column(1,4,18,textForm)
worksheet.set_column(5,6,9,currencyForm)

# Get the dimensions of the dataframe.
(max_row, max_col) = new_df_2.shape
# Create a list of column headers, to use in add_table().
column_settings = []
for header in new_df_2.columns:
    column_settings.append({'header': header})
# Add the table.
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

# Close the Pandas Excel writer and output the Excel file.

writer.close()
# subprocess.Popen([onlyoffice_path, f"{conta}@.xlsx"])
os.startfile(f"{conta}@.xlsx")