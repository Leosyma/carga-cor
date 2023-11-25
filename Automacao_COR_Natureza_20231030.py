# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 13:53:36 2022

@author: 2018459
"""

#%% Bibliotecas
import pandas as pd
import keyring
import cx_Oracle
import os
from openpyxl import load_workbook



#%% Dados de entrada
#Origem dos dados

#Caminho de rede onde está o arquivo
pasta = r"W:\Inteligência Regulatória Analítica - IRA\2. Projetos\2022\COR\Dados"


#Arquivo
arquivo_entrada = r'COR_NATUREZA_referencia.xlsx'
arquivo_saida = r'COR_NATUREZA_20231030.xlsx'



#%%Parâmetros Adicionais
#Tabelas Banco de Dados Oracle
tabela_oracle_agrupada = 'VW_COR_NATUREZA_AGRUPADA'
tabela_oracle_classifica = 'VW_COR_VERIFICA_CLASSIFICACAO_COR'
tabela_oracle_bmp = 'COR_DEPARA_NATUREZA_ACAO_BMP'
tabela_oracle_juridico = 'COR_DEPARA_NATUREZA_ACAO_JURIDICO'



#%% Acessar o banco de dados para importar os dados 


#Definir as variáveis para conexão no banco de dados
aplicacao_usuario = "USER_IRA"
aplicacao_senha = "BD_IRA"
aplicacao_dsn = "DSN"
usuario = "IRA"


#Definir conexão com o banco de dados     
try:
    connection = cx_Oracle.connect(user = keyring.get_password(aplicacao_usuario, usuario),
                                   password = keyring.get_password(aplicacao_senha,usuario),
                                   dsn= keyring.get_password(aplicacao_dsn, usuario),
                                   encoding="UTF-8")

#Se der erro na conexão com o banco, irá aparecer a mensagem abaixo
except Exception as err:
    print('Erro na Conexao:', err)    

#Se estiver tudo certo na conexão, irá aparecer a mensagem abaixo
else:
    print('Conexao com o Banco de Dados efetuada com sucesso. Versao da conexao: ' + connection.version)
    
    #O cursor abaixo irá importar a tabela do Banco de Dados Oracle
    try:
        #Criar o cursor para cada tabela
        cursor_agrupada = connection.cursor()
        cursor_classifica = connection.cursor()
        cursor_bmp = connection.cursor()
        cursor_juridico = connection.cursor()
        
        #Importação das tabelas
        print('Carregando tabela NATUREZA AGRUPADA...')
        cursor_agrupada.execute('''SELECT * FROM ''' + tabela_oracle_agrupada) #Importa a tabela 'NATUREZA AGRUPADA'
        print('Carregando tabela CLASSIFICACAO COR...')
        cursor_classifica.execute('''SELECT * FROM ''' + tabela_oracle_classifica) #Importa a tabela 'CLASSIFICACAO COR' 
        print('Carregando tabela NATUREZA BMP...')
        cursor_bmp.execute('''SELECT * FROM ''' + tabela_oracle_bmp) #Importa a tabela 'NATUREZA BMP' 
        print('Carregando tabela NATUREZA JURIDICA...')
        cursor_juridico.execute('''SELECT * FROM ''' + tabela_oracle_juridico) #Importa a tabela 'NATUREZA JURIDICA'


        
        #Transforma as tabelas em DataFrame
        df_agrupada_oracle=pd.DataFrame(cursor_agrupada)
        df_classifica_oracle = pd.DataFrame(cursor_classifica)
        df_bmp_oracle = pd.DataFrame(cursor_bmp)
        df_juridico_oracle = pd.DataFrame(cursor_juridico)
        
        
    except Exception as err:
        print('Erro no Importação:', err)
    else:
        print('Tabelas importadas com sucesso!')
    finally:    
        cursor_agrupada.close()
        cursor_classifica.close()
        cursor_bmp.close()
        cursor_juridico.close()
        connection.close()


#%%Tratamento da tabela importada do Banco de Dados Oracle
#TABELA AGRUPADA
#Renomear o nome das colunas
df_agrupada_oracle = df_agrupada_oracle.rename(columns={0:'FONTE'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={1:'CODIGO_EMPRESA'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={2:'NATUREZA'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={3:'ANO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={4:'JANEIRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={5:'FEVEREIRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={6:'MARCO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={7:'ABRIL'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={8:'MAIO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={9:'JUNHO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={10:'JULHO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={11:'AGOSTO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={12:'SETEMBRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={13:'OUTUBRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={14:'NOVEMBRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={15:'DEZEMBRO'})
df_agrupada_oracle = df_agrupada_oracle.rename(columns={16:'TOTAL'})



#TABELA VERIFICA CLASSIFICACAO COR
#Renomear o nome das colunas
df_classifica_oracle = df_classifica_oracle.rename(columns={0:'CLASSE_DE_CUSTO'})
df_classifica_oracle = df_classifica_oracle.rename(columns={1:'DESCRICAO_DA_CLASSE_DE_CUSTO$'})
df_classifica_oracle = df_classifica_oracle.rename(columns={2:'CONTROLADORIA_COR_NAO_COR%'})
df_classifica_oracle = df_classifica_oracle.rename(columns={3:'CONTROLADORIA_NATUREZA$'})
df_classifica_oracle = df_classifica_oracle.rename(columns={4:'REGULATORIO_COR_NAO_COR%'})
df_classifica_oracle = df_classifica_oracle.rename(columns={5:'REGULATORIO_NATUREZA$'})



#TABELA DE PARA NATUREZA BMP
#Renomear o nome das colunas
df_bmp_oracle = df_bmp_oracle.rename(columns={0:'NUMERO'})
df_bmp_oracle = df_bmp_oracle.rename(columns={1:'DESCRICAO'})
df_bmp_oracle = df_bmp_oracle.rename(columns={2:'CLASSIFICACAO'})
df_bmp_oracle = df_bmp_oracle.rename(columns={3:'CLASSIFICACAO_PADRONIZADA'})




#TABELA DE PARA NATUREZA JURIDICO
#Renomear o nome das colunas
df_juridico_oracle = df_juridico_oracle.rename(columns={0:'NATUREZA_DA_ACAO'})
df_juridico_oracle = df_juridico_oracle.rename(columns={1:'NATUREZA_DA_ACAO_PADRONIZADA'})





#%%Inserção dos novos dados nas respectivas abas da planilha de referência

#Abrir a planilha para inserir o DataFrame editado
book = load_workbook(os.path.join(pasta, arquivo_entrada))
writer = pd.ExcelWriter(os.path.join(pasta, arquivo_saida), engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

'''
#Limpar os dados da aba
ws_bmp = book['BMP']
ws_cobertura = book['Cobertura']
ws_despesa = book['Orçamento Despesa']
ws_receita = book['Orçamento Receita']

#Limpa os dados da aba 'BMP'
for row in ws_bmp['A2:Z100000']:
    for cell in row:
        cell.value = None
        
#Limpa os dados da aba 'Cobertura'
for row in ws_cobertura['A2:Z100000']:
    for cell in row:
        cell.value = None
        
#Limpa os dados da aba 'Orçamento Despesa'
for row in ws_despesa['A3:Z100000']:
    for cell in row:
        cell.value = None

#Limpa os dados da aba 'Orçamento Receita'
for row in ws_receita['A2:Z100000']:
    for cell in row:
        cell.value = None
'''

df_agrupada_oracle.to_excel(writer, sheet_name='NATUREZA AGRUPADA',index=False) #Insere o dataframe 'AGRUPADA' na planilha existente
df_classifica_oracle.to_excel(writer, sheet_name = 'VERIFICA CLASSIFICACAO COR', index = False) #Insere o dataframe 'CLASSIFICA COR' na planilha
df_bmp_oracle.to_excel(writer, sheet_name = 'DEPARA NATUREZA BMP', index = False) #Insere o dataframe 'NATUREZA BMP' na planilha
df_juridico_oracle.to_excel(writer, sheet_name = 'DEPARA NATUREZA JURIDICO', index = False) #Insere o dataframe 'NATUREZA JURIDICO' na planilha
print('Tabelas Exportadas')


#Salva o arquivo e fecha
writer.save()
writer.close()
writer.handles = None



