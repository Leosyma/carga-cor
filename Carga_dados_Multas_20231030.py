# -*- coding: utf-8 -*-
"""
Created on Wed Nov  9 14:47:17 2022

@author: 2018459
"""

#%% Bibliotecas
import pandas as pd
import numpy as np
import keyring
import cx_Oracle
import os



#%% Dados de entrada
#Origem dos dados

#Caminho referencia
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2023 - 3TRI - 15ªReunião\1. Arquivos\Controladoria"

#Arquivos
#Como os dados das empresas estão em arquivos separados, listar todos que serão carregados conforme layout abaixo.
arquivo = 'Multas_2023_Fernando Cesar.xlsx'
aba = 'Resumo'


#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_MULTAS'


#Ano dos dados
ano = '2023'
ano_oracle = "'%23'"



#%% Abertura dos arquivos


#Carregar arquivos 
df = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba
                   ,header = 1
                   ,nrows = 8
                   ,usecols='B:R')   



#Formatar as colunas
#Dropa as linhas e colunas nulas
df = df.dropna(axis=0,how='all')
df = df.dropna(axis=1,how='all')

df.insert(3,'ANO',ano) #Inserir uma coluna com o nome 'ANO'


#Mudar o nome das colunas
df.columns.values[0] = 'EMPRESA'
df.columns.values[1] = 'CONTA'
df.columns.values[2] = 'TIPO'
df.columns.values[4] = 'TOTAL'
df.columns.values[5] = 'JAN'
df.columns.values[6] = 'FEV'
df.columns.values[7] = 'MAR'
df.columns.values[8] = 'ABR'
df.columns.values[9] = 'MAI'
df.columns.values[10] = 'JUN'
df.columns.values[11] = 'JUL'
df.columns.values[12] = 'AGO'
df.columns.values[13] = 'SET'
df.columns.values[14] = 'OUT'
df.columns.values[15] = 'NOV'
df.columns.values[16] = 'DEZ'


#Limpeza e Tratamento dos dados
df = df.astype(str)
df['EMPRESA'] = df['EMPRESA'].replace('nan','Total')
df['TIPO'] = df['TIPO'].replace('nan',df.iat[0,2])
df['JAN'] = df['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['FEV'] = df['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['MAR'] = df['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['ABR'] = df['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['MAI'] = df['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['JUN'] = df['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['JUL'] = df['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['AGO'] = df['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['SET'] = df['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['OUT'] = df['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['NOV'] = df['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['DEZ'] = df['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)
df['TOTAL'] = df['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0).round(2)





#%% Inserir dados no banco de dados

#Criar a lista para inserção no banco SQL com os dados da base editada
dados_list = df.values.tolist()


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
    
    #O cursor abaixo irá executar o insert de cada uma das linhas da base editada no Banco de Dados Oracle
    try:
        cursor = connection.cursor()
        cursor.execute('''DELETE FROM ''' + tabela_oracle + ''' WHERE ANO LIKE ''' + ano_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:    
        cursor.close()
        connection.close()


