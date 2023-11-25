# -*- coding: utf-8 -*-
"""
Insert de dados para estudos COR - Fonte de dados Contabilidade

@author: IRA
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
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2023 - 3TRI - 15ªReunião\1. Arquivos\Contabilidade"

#Arquivos
#Como os dados das empresas estão em arquivos separados, listar todos que serão carregados conforme layout abaixo.
arquivo = ['BMP__10_2023_D009.xlsx'
           ,'BMP_10.2023_D002.xlsx'
           ,'BMP_10.2023_D006.xlsx'
           ,'BMP_10_2023_D001.xlsx']

#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_BMP'



#%% Abertura dos arquivos


#Carregar arquivos das 3 abas (uma por mês, somando as 3 abas temos o trimestre)
df = pd.DataFrame(columns=['codigo_empresa','ano','mes','numero','debito','credito','saldo'])

for j in arquivo:
    for i in range(1):
        
        df_temp = pd.read_excel(os.path.join(pasta, j),sheet_name = i,decimal=',')
        df = pd.concat([df, df_temp]) 
        print('Carregou a aba: ',i)


#Formatar as colunas
df = df.astype("str")
df['debito'] = df['debito'].astype('float').replace('.', ',').replace(np.nan,0)
df['credito'] = df['credito'].astype('float').replace('.', ',').replace(np.nan,0)
df['saldo'] = df['saldo'].astype('float').replace('.', ',').replace(np.nan,0)




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
        #cursor.execute('''TRUNCATE TABLE ''' + tabela_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:
        cursor.close()
        connection.close()

