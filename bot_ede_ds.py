# Bot EDE - ERP Data Extract - Daily Sales

# Imports
from datetime import date, datetime, timedelta
from selenium import webdriver
import pyautogui, pydirectinput, time, os, shutil
import numpy as np
import pandas as pd

############################################################################################################################
# Variables Definition
############################################################################################################################
# Period Definition
hour_now = int(str(datetime.today())[11:13])

if hour_now > 10:
    dti = datetime.today()
    dtf = datetime.today()
else:
    dti = datetime.today() - timedelta(days=1)
    dtf = datetime.today()
    
date_ini = str(dti)[8:10] + str(dti)[5:7] + str(dti)[0:4]
date_fin = str(dtf)[8:10] + str(dtf)[5:7] + str(dtf)[0:4]

# Files Name Definition
file_name = str(dti)[0:4] + str(dti)[5:7] + str(dti)[8:10] + "-" + str(dti)[11:13] + str(dti)[14:16]
file_name_pd = file_name + "-01PD.XLS" # Vendas por Produto
file_name_fp = file_name + "-02FP.XLS" # Vendas por Forma de Pagamento
file_name_cr = file_name + "-03CR.XLS" # Vendas por Tipo de Carnê
file_name_cl = file_name + "-04CL.XLS" # Cadastro de Clientes

# Softwares Paths Definition
sys_adm = 'D:/RPA/ERP/CEVidal-Admin.rdp' # Modulo Admin ERP
sys_fin = 'D:/RPA/ERP/CEVidal-Financeiro.rdp' # Modulo Financeiro ERP
sys_res = 'D:/RPA/ERP/Reiniciar.rdp' # Botão para Reiniciar o ERP

# Directories Definition
sourc_file = 'C:/Dataweb/' # Fonte de Arquivos (Onde o ERP exporta os Arquivos)
input_pend = 'D:/RPA/ETL/INPUT/01_PENDING/'
input_done = 'D:/RPA/ETL/INPUT/02_DONE/'
output_pend = 'D:/RPA/ETL/OUTPUT/01_PENDING/'
output_done = 'D:/RPA/ETL/OUTPUT/02_DONE/'
output_app = '//192.168.9.2/htdocs/pro/src/datasets_csv/venda_diaria/'

# Login Definition
user_name = 'username'
pass_word = 'password'

# Definition of number of stores
stores_number = 13

############################################################################################################################
# Functions Definition
############################################################################################################################
def loginSystem(username, password, system_path, system_module):
    
    # Function to System Login    
    pyautogui.hotkey("win", "m")
    os.system('start ' + system_path)
    time.sleep(60)
    pydirectinput.write(username)
    pydirectinput.press("tab")
    pydirectinput.write(password)
    pydirectinput.press("tab")
    pydirectinput.write("m")
    pydirectinput.press("enter", presses=2)
    time.sleep(30)
    
    # Navigating the system
    if(system_module == "Admin"):
        
        pydirectinput.press("alt")
        pydirectinput.press("f")
        pydirectinput.press("i")
        pydirectinput.press("e")
        time.sleep(30)
        
    elif(system_module == "Financeiro"):
        
        pydirectinput.press("alt")
        pydirectinput.press("f")
        pydirectinput.press("p")
        time.sleep(30)

############################################################################################################################
def extractDataPD(date_ini, date_fin, stores_number):
    
    # 01PD | Function to Extract Data from Sales by Product
    
    # Click para Análise de Vendas
    pyautogui.click(x=63, y=104)
    time.sleep(30)

    # Click em Personalizar
    pyautogui.click(x=526, y=61)
    time.sleep(1)
    pydirectinput.press("down")
    time.sleep(1)
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Personalizar
    pyautogui.click(x=526, y=61)
    time.sleep(1)
    pydirectinput.press("down", presses=2)
    time.sleep(1)
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Modelos
    pyautogui.click(x=778, y=66)
    time.sleep(1)
    pydirectinput.write("01")
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Aplicar Modelo
    pyautogui.click(x=818, y=61)
    time.sleep(5)

    # Click na lista de lojas
    pyautogui.click(x=257, y=113)
    time.sleep(2)    
    for tag_stores_one in range(stores_number):        
        pydirectinput.press("down")
        pydirectinput.press("space")
    for tag_stores_two in range(9):
        pydirectinput.press("down")        
    pydirectinput.press("space")
    pydirectinput.press("enter")

    # Click na lista de períodos
    pyautogui.click(x=697, y=117)
    time.sleep(2)
    pydirectinput.write("p")
    pydirectinput.press("enter")

    # Click na data inicial
    pyautogui.click(x=732, y=120)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_ini)

    # Click na data final
    pyautogui.click(x=833, y=120)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_fin)

    # Click na Lupa Pesquisar
    pyautogui.click(x=238, y=52)
    time.sleep(15)

    # Click no incone excel
    pyautogui.click(x=466, y=63)
    time.sleep(2)
    pydirectinput.press("enter")
    time.sleep(10)

############################################################################################################################
def extractDataFP(date_ini, date_fin, stores_number):
    
    # 02FP | Function to Extract Data from Sales by Payment Method
    
    # Click na Consulta Personalizada
    pyautogui.click(x=88, y=343)
    time.sleep(30)

    # Click na lista de consultas
    pyautogui.click(x=302, y=64)
    time.sleep(3)
    pydirectinput.write("ma")
    pydirectinput.press("enter")

    # Click na lista de lojas
    pyautogui.click(x=282, y=118)
    time.sleep(2)
    for tag_stores_one in range(stores_number):        
        pydirectinput.press("down")
        pydirectinput.press("space")
    for tag_stores_two in range(9):
        pydirectinput.press("down")        
    pydirectinput.press("space")
    pydirectinput.press("enter")

    # Click na lista de períodos
    pyautogui.click(x=515, y=118)
    time.sleep(2)
    pydirectinput.write("p")
    pydirectinput.press("enter")

    # Click na data inicial
    pyautogui.click(x=550, y=118)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_ini)

    # Click na data final
    pyautogui.click(x=651, y=118)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_fin)

    # Click para consultar F9
    pydirectinput.press('f9')
    time.sleep(60)

    # Click para exportar os dados
    pyautogui.click(x=648, y=63)
    pydirectinput.write("3fp")
    pydirectinput.press("enter", presses=3)
    time.sleep(10)

############################################################################################################################
def extractDataCL(date_ini, date_fin, stores_number):

    # 04CL | Function to Extract Data from Clients

    # Click para Análise de Vendas
    pyautogui.click(x=63, y=104)
    time.sleep(30)

    # Click em Personalizar
    pyautogui.click(x=526, y=61)
    time.sleep(1)
    pydirectinput.press("down")
    time.sleep(1)
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Personalizar
    pyautogui.click(x=526, y=61)
    time.sleep(1)
    pydirectinput.press("down", presses=2)
    time.sleep(1)
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Modelos
    pyautogui.click(x=778, y=66)
    time.sleep(1)
    pydirectinput.write("04")
    pydirectinput.press("enter")
    time.sleep(1)

    # Click em Aplicar Modelo
    pyautogui.click(x=818, y=61)
    time.sleep(5)

    # Click na lista de lojas
    pyautogui.click(x=257, y=113)
    time.sleep(2)    
    for tag_stores_one in range(stores_number):        
        pydirectinput.press("down")
        pydirectinput.press("space")
    for tag_stores_two in range(9):
        pydirectinput.press("down")        
    pydirectinput.press("space")
    pydirectinput.press("enter")

    # Click na lista de períodos
    pyautogui.click(x=697, y=117)
    time.sleep(2)
    pydirectinput.write("p")
    pydirectinput.press("enter")

    # Click na data inicial
    pyautogui.click(x=732, y=120)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_ini)

    # Click na data final
    pyautogui.click(x=833, y=120)
    time.sleep(2)
    pydirectinput.press('right', presses=10)
    pydirectinput.press('backspace', presses=10)
    pydirectinput.write(date_fin)

    # Click na Lupa Pesquisar
    pyautogui.click(x=238, y=52)
    time.sleep(15)

    # Click no incone excel
    pyautogui.click(x=466, y=63)
    time.sleep(2)
    pydirectinput.press("enter")
    time.sleep(10)

############################################################################################################################
def extractDataCR(date_ini, stores_number):
    
    # 03CR | Function to Extract Data from Sales by Paym Slip (Carnê)
    # d - Date Period
    # n - Stores Number

    # Navegar até os filtros
    pyautogui.click(x=50, y=55)
    time.sleep(1)
    pyautogui.click(x=155, y=60)
    time.sleep(1)
    
    # Click na lista de filiais
    pyautogui.click(x=34, y=191)
    time.sleep(2)
    for tag_stores_one in range(stores_number):        
        pydirectinput.press("down")
        pydirectinput.press("space")
    for tag_stores_two in range(9):
        pydirectinput.press("down")        
    pydirectinput.press("space")
    pydirectinput.press("enter")

    # Click na lista de situações
    pyautogui.click(x=34, y=254)
    pydirectinput.press("up")
    pydirectinput.press("enter")
    time.sleep(2)

    # Click na opção Receber
    pyautogui.click(x=225, y=309)
    time.sleep(2)

    # Click Minimizar Selecao Principal
    pyautogui.click(x=23, y=145)

    # Click Maximizar Selecao Outras Datas
    pyautogui.click(x=24, y=187)

    # Click em Emissao
    pyautogui.click(x=63, y=233)
    time.sleep(2)
    pydirectinput.write(date_ini)

    # Click Minimizar Selecao Outras Datas
    pyautogui.click(x=24, y=187)

    # Click Maximizar Selecao Outros
    pyautogui.click(x=24, y=266)
    time.sleep(2)

    # Click na opção Forma pgto
    pyautogui.click(x=91, y=316)
    time.sleep(2)

    pydirectinput.press("down", presses=5)
    pydirectinput.press("enter")
    time.sleep(2)

    # Click Minimizar Selecao Outros
    pyautogui.click(x=24, y=266)
    time.sleep(2)

    pydirectinput.press("f3")
    time.sleep(5)

    # Click na com direito na aba Plano de Contas
    pyautogui.click(button='right', x=383, y=150)
    time.sleep(2)

    pydirectinput.press("down")
    pydirectinput.press("enter")
    time.sleep(5)

    pydirectinput.press("enter")
    time.sleep(10)

############################################################################################################################
def renameFiles(source_directory, file_name):
    
    # Function to Rename Files
    for file in os.listdir(source_directory):

        if "01pd_" in file or "01PD_" in file:
            os.rename(source_directory + file, source_directory + file_name)
            
        if "3fp.XLS" in file or "3FP.XLS" in file:
            os.rename(source_directory + file, source_directory + file_name)

        if "04cl_" in file or "04CL_" in file:
            os.rename(source_directory + file, source_directory + file_name)

        if "PlanoContas" in file:
            os.rename(source_directory + file, source_directory + file_name)

############################################################################################################################
def orgFiles(source_directory, destiny_directory, move_or_copy):
    
    # Function to Move or Copy Files    
    for file in os.listdir(source_directory):
        if '.XLS' in file or '.xls' in file or '.CSV' in file or '.csv' in file:
            if move_or_copy == 'move':
                shutil.move(source_directory + file, destiny_directory + file)
            elif move_or_copy == 'copy':
                shutil.copy(source_directory + file, destiny_directory + file)

############################################################################################################################
def clearDirectories(source_directory):

    # Function to Clear Directories
    for file in os.listdir(source_directory):
        if ".XLS" in file or ".xls" in file or ".CSV" in file or ".csv" in file or ".XML" in file or ".xml" in file:            
            os.unlink(source_directory + file)

############################################################################################################################
def dataETL(source_directory, destiny_directory):
    
    for file in os.listdir(source_directory):

        if '.XLS' in file or '.xls' in file:

            if '-01PD' in file or '-01pd' in file:

                df = pd.read_excel(source_directory + file)

                # Delete unnecessary columns
                df.drop(['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 4', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 16', 'Unnamed: 18', 'Unnamed: 20', 'Unnamed: 22', 'Unnamed: 27'], inplace=True, axis=1)

                # Rename the Columns
                df.columns = ['Filial', 'Vendedor', 'Venda', 'CFE', 'Cliente', 'CPF_Cliente', 'Id_Item', 'Desc_Item', 'Familia_Item', 'Emissao', 'Campanha', 'Quant', 'Vlr_Compra_Unit', 'Valor_Bruto', 'Valor_Liquido']

                # Delete unnecessary lines
                df = df.drop([0, 1, 2])
                indexNames = df[ df['Filial'] == 'Total geral' ].index
                df.drop(indexNames, inplace=True)

                # Delete the blank spaces before and after the seller's name
                df['Vendedor'] = df['Vendedor'].str.strip()

                # Fix merged cells
                df['Familia_Item'] = df['Familia_Item'].replace(np.nan, 'NONE')
                df = df.fillna(method='ffill')
                df['Familia_Item'] = df['Familia_Item'].replace('NONE', np.nan)

                # Fix spaces between words
                df['Desc_Item'] = df['Desc_Item'].replace(['  ', '   '], ' ', regex=True)
                df['Familia_Item'] = df['Familia_Item'].replace(['  ', '   '], ' ', regex=True)

                # Convert the Data to the database model
                df['Emissao'] = df['Emissao'].str[6:] + df['Emissao'].str[3:5] + df['Emissao'].str[:2]

                # Converting values from portuguese to float
                df['Vlr_Compra_Unit'] = df['Vlr_Compra_Unit'].str.slice(start=3)
                df['Vlr_Compra_Unit'] = df['Vlr_Compra_Unit'].str.replace('.', '')
                df['Vlr_Compra_Unit'] = df['Vlr_Compra_Unit'].str.replace(',', '.')
                df['Vlr_Compra_Unit'] = df['Vlr_Compra_Unit'].astype(float)

                df['Valor_Bruto'] = df['Valor_Bruto'].str.slice(start=3)
                df['Valor_Bruto'] = df['Valor_Bruto'].str.replace('.', '')
                df['Valor_Bruto'] = df['Valor_Bruto'].str.replace(',', '.')
                df['Valor_Bruto'] = df['Valor_Bruto'].astype(float)

                df['Valor_Liquido'] = df['Valor_Liquido'].str.slice(start=3)
                df['Valor_Liquido'] = df['Valor_Liquido'].str.replace('.', '')
                df['Valor_Liquido'] = df['Valor_Liquido'].str.replace(',', '.')
                df['Valor_Liquido'] = df['Valor_Liquido'].astype(float)

                # Exporting File
                basename = os.path.basename(source_directory + file)
                file_name = os.path.splitext(basename)[0] + '.csv'
                df.to_csv(destiny_directory + file_name, encoding='utf-8', index=False)

            elif '-02FP' in file or '-02fp' in file:

                df = pd.read_excel(source_directory + file)

                # Rename the Columns
                df.columns = ['Filial', 'Venda', 'Vendedor', 'Emissao', 'Vlr_Bruto', 'Vlr_Desc', 'Vlr_Liquido', 'Forma_Pagamento', 'Vlr_Parcela', 'CFE', 'Chave_CFE']

                # Delete unnecessary lines
                df = df.dropna(subset=['Filial'])

                # Fix '.0' from 'Venda'
                df['Venda'] = df['Venda'].astype(str)
                df['Venda'] = df['Venda'].str.slice(start=0, stop=-2)

                # Delete the blank spaces before and after the seller's name
                df['Vendedor'] = df['Vendedor'].str.strip()

                # Fix '-' from 'Emissao'
                df['Emissao'] = df['Emissao'].astype(str)
                df['Emissao'] = df['Emissao'].replace('-', '', regex=True)

                # Convert to Float
                df['Vlr_Bruto'] = df['Vlr_Bruto'].astype(float)
                df['Vlr_Desc'] = df['Vlr_Desc'].astype(float)
                df['Vlr_Liquido'] = df['Vlr_Liquido'].astype(float)
                df['Vlr_Parcela'] = df['Vlr_Parcela'].astype(float)

                # Fix '.0' from 'CFE'
                df['CFE'] = df['CFE'].astype(str)
                df['CFE'] = df['CFE'].str.slice(start=0, stop=-2)

                # Sort by 'Filial' and 'CFE'
                df = df.sort_values(by=['Filial', 'CFE'])

                # Exporting File
                basename = os.path.basename(source_directory + file)
                file_name = os.path.splitext(basename)[0] + '.csv'
                df.to_csv(destiny_directory + file_name, encoding='utf-8', index=False)

            elif '-03CR' in file or '-03cr' in file:

                df = pd.read_excel(source_directory + file)

                # Rename the Columns
                df.columns = ['Filial', 'N_Doc', 'Emissao', 'Pagamento', 'Recebimento', 'Vencimento', 'Valor', 'ID_Cliente', 'Cad_Pessoa', 'Tipo_Crediario']

                # Delete unnecessary lines
                df = df.dropna(subset=['Filial'])

                # Fix ' 00:00:00' and '-' from 'Emissao'
                df['Emissao'] = df['Emissao'].astype(str)
                df['Emissao'] = df['Emissao'].str.replace(' 00:00:00', '')
                df['Emissao'] = df['Emissao'].str.replace('-', '')

                # Fix '-' from 'Vencimento'
                df['Vencimento'] = df['Vencimento'].astype(str)
                df['Vencimento'] = df['Vencimento'].str.replace('-', '')

                # Fix '.0' from 'ID_Cliente'
                df['ID_Cliente'] = df['ID_Cliente'].astype(str)
                df['ID_Cliente'] = df['ID_Cliente'].str.slice(start=0, stop=-2)

                # Fix '.0' from 'Cad_Pessoa'
                df['Cad_Pessoa'] = df['Cad_Pessoa'].astype(str)
                df['Cad_Pessoa'] = df['Cad_Pessoa'].str.slice(start=0, stop=-2)

                # Exporting File
                basename = os.path.basename(source_directory + file)
                file_name = os.path.splitext(basename)[0] + '.csv'
                df.to_csv(destiny_directory + file_name, encoding='utf-8', index=False)

            elif '-04CL' in file or '-04cl' in file:

                df = pd.read_excel(source_directory + file)

                indexNames = df[ df['Unnamed: 0'] == 'Total geral' ].index
                df.drop(indexNames, inplace=True)

                df = df.fillna(method='ffill')
                df.drop(['Empresa', 'NF: número', 'Unnamed: 2', 'Unnamed: 4', 'Unnamed: 6', 'Unnamed: 8', 'Unnamed: 10', 'Unnamed: 12', 'Unnamed: 14', 'Unnamed: 16', 'Unnamed: 18', 'Unnamed: 20', 'Unnamed: 22', 'Unnamed: 24', 'Unnamed: 26', 'Unnamed: 28', 'Unnamed: 30', 'Unnamed: 32', 'Unnamed: 34', 'Unnamed: 36', 'Unnamed: 37'], inplace=True, axis=1)
                df = df.drop([0, 1, 2])
                df.columns = ['Filial', 'Emissao', 'Status', 'Id_Cli', 'Tip_Pes', 'CPF', 'Nome_Cli', 'Data_Nasc', 'Fone', 'E-mail', 'Cep', 'Endereco', 'Bairro', 'Cidade', 'Estado', 'Data_Cad', 'Sexo']

                df['Emissao'] = df['Emissao'].str[6:] + df['Emissao'].str[3:5] + df['Emissao'].str[:2]

                df['Id_Cli'] = df['Id_Cli'].astype(str)
                df['Id_Cli'] = df['Id_Cli'].str.slice(start=0, stop=-2)
                df['Id_Cli'] = df['Id_Cli'].str.zfill(7)

                df['Tip_Pes'] = df['Tip_Pes'].str.strip()

                df['CPF'] = df['CPF'].astype(str)
                df['CPF'] = df['CPF'].str.slice(start=0, stop=-2)
                df['CPF'] = df['CPF'].str.zfill(11)

                df['Nome_Cli'] = df['Nome_Cli'].str.strip()
                df['Nome_Cli'] = df['Nome_Cli'].replace(['  ', '   '],' ', regex=True)

                df['Data_Nasc'] = df['Data_Nasc'].str[6:] + df['Data_Nasc'].str[3:5] + df['Data_Nasc'].str[:2]

                df['E-mail'] = df['E-mail'].str.lower()
                df['E-mail'] = df['E-mail'].str.strip()

                df['Cep'] = df['Cep'].replace('        ', np.nan)

                df['Endereco'] = df['Endereco'].astype(str)
                df['Endereco'] = df['Endereco'].replace(['  ', '   '],' ', regex=True)
                df['Endereco'] = df['Endereco'].replace(['  ', '   '],' ', regex=True)
                df['Endereco'] = df['Endereco'].replace(' ,',',', regex=True)
                df['Endereco'] = df['Endereco'].replace(',',', ', regex=True)
                df['Endereco'] = df['Endereco'].replace(['  ', '   '],' ', regex=True)
                df['Endereco'] = df['Endereco'].replace('nan', np.nan)

                df['Bairro'] = df['Bairro'].str.strip()
                df['Bairro'] = df['Bairro'].replace(['  ', '   '],' ', regex=True)
                df['Bairro'] = df['Bairro'].replace(['  ', '   '],' ', regex=True)
                df['Bairro'] = df['Bairro'].replace('nan', np.nan)

                df['Cidade'] = df['Cidade'].str.strip()
                df['Cidade'] = df['Cidade'].replace(['  ', '   '],' ', regex=True)
                df['Cidade'] = df['Cidade'].replace(['  ', '   '],' ', regex=True)
                df['Cidade'] = df['Cidade'].replace('nan', np.nan)

                df['Data_Cad'] = df['Data_Cad'].str[6:] + df['Data_Cad'].str[3:5] + df['Data_Cad'].str[:2]

                df['Sexo'] = df['Sexo'].str.strip()
                df['Sexo'] = df['Sexo'].replace('(Não definido)', np.nan)

                df['Dia_Nasc'] = df['Data_Nasc'].str[6:]
                df['Mes_Nasc'] = df['Data_Nasc'].str[4:6]

                df['CNPJ'] = df['CPF']
                df['Celular'] = np.nan
                df['Rua'] = np.nan
                df['Numero'] = np.nan
                df['Complemento'] = np.nan
                df['Pais'] = np.nan
                df['Primeira_Compra'] = np.nan
                df['Ultima_Compra'] = np.nan
                df['Tipo_Cliente'] = np.nan
                df['Filial_Origem'] = np.nan

                df = df[['Status', 'Id_Cli', 'Tip_Pes', 'CPF', 'CNPJ', 'Nome_Cli', 'Data_Nasc', 'Dia_Nasc', 'Mes_Nasc', 'Fone', 'Celular', 'E-mail', 'Cep', 'Rua' ,'Numero', 'Complemento', 'Endereco', 'Bairro', 'Cidade', 'Estado', 'Pais', 'Sexo', 'Data_Cad', 'Primeira_Compra', 'Ultima_Compra', 'Tipo_Cliente', 'Filial_Origem', 'Filial', 'Emissao']]

                df = df.sort_values(by='Id_Cli')

                indexNames = df[ df['Nome_Cli'] == 'CONSUMIDOR' ].index
                df.drop(indexNames, inplace=True)

                # Exporting File
                basename = os.path.basename(source_directory + file)
                file_name = os.path.splitext(basename)[0] + '.csv'
                df.to_csv(destiny_directory + file_name, encoding='utf-8', index=False)

############################################################################################################################
def processA():

    clearDirectories(sourc_file)
    clearDirectories(input_pend)
    clearDirectories(output_pend)
    clearDirectories(output_app)

    counter = 0
    files_number = 0

    while files_number < 1:

        loginSystem(user_name, pass_word, sys_adm, "Admin")
        extractDataFP(date_ini, date_fin, stores_number)

        for file in os.listdir(sourc_file):
            if "3fp.XLS" in file or "3FP.XLS" in file:
                renameFiles(sourc_file, file_name_fp)
                files_number = files_number + 1

        counter = counter + 1
    
        if files_number < 1:
            os.system('start ' + sys_res)
            pyautogui.hotkey("win", "m")
            clearDirectories(sourc_file)

        if counter == 5:
            break

    return files_number

############################################################################################################################
def processB():

    extractDataPD(date_ini, date_fin, stores_number)
    renameFiles(sourc_file, file_name_pd)
    os.system('start ' + sys_res)
    time.sleep(10)

    loginSystem(user_name, pass_word, sys_adm, "Admin")
    extractDataCL(date_ini, date_fin, stores_number)
    renameFiles(sourc_file, file_name_cl)

    loginSystem(user_name, pass_word, sys_fin, "Financeiro")
    extractDataCR(date_ini, stores_number)
    renameFiles(sourc_file, file_name_cr)
    os.system('start ' + sys_res)

    counter = 0

    for file in os.listdir(sourc_file):
        if ".XLS" in file or ".xls" in file:
            counter = counter + 1
    
    if counter >= 3:
        return 1
    else:
        clearDirectories(sourc_file)
        return 0

############################################################################################################################
def processC():

    orgFiles(sourc_file, input_pend, 'move')
    dataETL(input_pend, output_pend)
    orgFiles(input_pend, input_done, 'move')
    orgFiles(output_pend, output_app, 'copy')
    orgFiles(output_pend, output_done, 'move')

    counter = 0

    for file in os.listdir(output_app):
        if ".CSV" in file or ".csv" in file:
            counter = counter + 1
    
    if counter >= 3:
        return 1
    else:
        clearDirectories(output_app)
        return 0

############################################################################################################################
def execBot():

    hour = int(str(datetime.today())[11:13])
    weekday = date.today().weekday()

    if weekday != 6:
        if processA() == 1:
            if processB() == 1:
                processC()    
    else:
        if hour > 14:
            if processA() == 1:
                if processB() == 1:
                    processC()

############################################################################################################################
execBot()