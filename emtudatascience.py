from cv2 import distanceTransformWithLabels
import pyodbc
from tabula import read_pdf
import pandas as pd
from tabulate import tabulate
from matplotlib import pyplot as plt
import fitz

ft = pd.read_excel('populacao_rmsp.xlsx')

list_demografia = []

print(ft.loc[ft['REGIAO METROPOLITANA'] == 'Região Metropolitana de São Paulo'])


dados_conexao = (
        
        "Driver={SQL Server};"
        "Server=LAPTOP-SG53A25P\SQLEXPRESS;"
        "DataBase=populacao_rmsp;"
        
        )

conexao = pyodbc.connect(dados_conexao)

print('Conexão bem sucedida')


populacao2010 = ft.loc[55, 'Pop_2010']

list_demografia.append(f'{(populacao2010/10**6):.1f}')

populacao2011 = ft.loc[55, 'Pop_2011']

list_demografia.append(f'{(populacao2011/10**6):.1f}')

populacao2012 = ft.loc[55, 'Pop_2012']

list_demografia.append(f'{(populacao2012/10**6):.1f}')

populacao2013 = ft.loc[55, 'Pop_2013']

list_demografia.append(f'{(populacao2013/10**6):.1f}')

populacao2014 = ft.loc[55, 'Pop_2014']

list_demografia.append(f'{(populacao2014/10**6):.1f}')

populacao2015 = ft.loc[55, 'Pop_2015']

list_demografia.append(f'{(populacao2015/10**6):.1f}')

populacao2016 = ft.loc[55, 'Pop_2016']

list_demografia.append(f'{(populacao2016/10**6):.1f}')

populacao2017 = ft.loc[55, 'Pop_2017']

list_demografia.append(f'{(populacao2017/10**6):.1f}')

populacao2018 = ft.loc[55, 'Pop_2018']

list_demografia.append(f'{(populacao2018/10**6):.1f}')

populacao2019 = ft.loc[55, 'Pop_2019']

list_demografia.append(f'{(populacao2019/10**6):.1f}')

populacao2020 = ft.loc[55, 'Pop_2020']

list_demografia.append(f'{(populacao2020/10**6):.1f}')


cursor = conexao.cursor()

comando = f"""INSERT INTO dados_populacao(populacao2010, populacao2011, populacao2012, populacao2013, populacao2014, populacao2015, populacao2016, populacao2017, populacao2018, populacao2019, populacao2020)
    
    VALUES
        ('{list_demografia[0]}', '{list_demografia[1]}', '{list_demografia[2]}', '{list_demografia[3]}', '{list_demografia[4]}', '{list_demografia[5]}', '{list_demografia[6]}', '{list_demografia[7]}', '{list_demografia[8]}', '{list_demografia[9]}', '{list_demografia[10]}' )"""

cursor.execute(comando)
    
cursor.commit()

def retornar_conexao_sql():

    server = "LAPTOP-SG53A25P\SQLEXPRESS"
    database = "populacao_rmsp"
    string_conexao = 'Driver={SQL Server Native Client 11.0};Server='+server+';Database='+database+';Trusted_Connection=yes;'
    return conexao.cursor()

retornar_conexao_sql()

select_sql = "SELECT * FROM dados_populacao"

cursor.execute(select_sql)

linha = cursor.fetchone()

if linha:
    
    print(linha)

print(select_sql)

#def abrir_pdf():

    #conteudo = ""

    #with fitz.open('ftpd.pdf') as pdf:

        #for pagina in pdf:
            
            #conteudo += pagina.getText()