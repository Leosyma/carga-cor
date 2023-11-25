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
import pyxlsb




#%% Dados de entrada
#Origem dos dados

#Caminho referencia
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2. Demais Informações"


#Arquivos
#Como os dados das empresas estão em arquivos separados, listar todos que serão carregados conforme layout abaixo.
arquivo = '1. DE_PARA_PLUZ-BMP_2023.xlsb'
aba = 'DE-PARA Controladoria'


#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_DEPARA_PLUZ_CONTROLADORIA'



#%% Abertura dos arquivos


#Carregar arquivos 
df = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba
                   ,engine='pyxlsb'
                   ,na_filter = False)


#Formatar as colunas
df = df.drop(df.columns[[0,8,11]],axis=1) #Exclui as colunas em branco
df = df.drop(df.index[[0,1,2,3]]) #Remover as linhas em branco e que não interessam
df = df.reset_index(drop=True)  #Reseta o índice


#Mudar o nome das colunas
df.columns.values[0] = 'CLASSE DE CUSTO CONTROLADORIA'
df.columns.values[1] = 'DESCRICAO DA CLASSE DE CUSTO CONTROLADORIA'
df.columns.values[2] = 'NATUREZA CONTROLADORIA'
df.columns.values[3] = 'COR/NÃO COR CONTROLADORIA'
df.columns.values[4] = 'CONTA ANEEL REGULATORIO'
df.columns.values[5] = 'NATUREZA REGULATORIO'
df.columns.values[6] = 'COR/NÃO COR REGULATORIO'
df.columns.values[7] = 'CONTAS OFICIO'
df.columns.values[8] = 'AJUSTE'
df.columns.values[9] = 'CONTA NOVA 2021'


#Excluir as linhas em branco
df = df.dropna(subset=df.columns[[1]], axis=0)


#Limpeza e Tratamento dos dados
df = df.astype(str)
df['NATUREZA CONTROLADORIA'] = df['NATUREZA CONTROLADORIA'].replace('nan','-')
df['COR/NÃO COR CONTROLADORIA'] = df['COR/NÃO COR CONTROLADORIA'].replace('nan','-' )
df['CONTAS OFICIO'] = df['CONTAS OFICIO'].replace('nan','-')
df['AJUSTE'] = df['AJUSTE'].replace('nan','-')
df['CONTA NOVA 2021'] = df['CONTA NOVA 2021'].replace('nan','-')




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
        cursor.execute('''TRUNCATE TABLE ''' + tabela_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:    
        cursor.close()
        connection.close()
