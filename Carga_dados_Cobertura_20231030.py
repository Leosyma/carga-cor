# -*- coding: utf-8 -*-
"""
Insert de dados para estudos COR - Fonte de dados RRE

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
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2023 - 3TRI - 15ªReunião\1. Arquivos\RRE"

#Arquivos
arquivo = 'Cobertura COR+XT_Valores Colados_Giovanna Esteves.xlsx'

#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_COBERTURA'

ano = "2023"
ano_filtro = '23'


#%% Abertura dos arquivos


#Carregar arquivos
#Definir quais colunas e linhas serão utilizadas
df = pd.read_excel(os.path.join(pasta, arquivo)
                   ,usecols=(1,2,3,4,5,6,7,8,9) #Definir quais colunas serão utilizadas
                   ,skiprows=[1,2,3] #Definir as linhas do cabeçalho que serão IGNORADAS 
                   ).reset_index(drop=True) #Excluir a coluna Index


#Renomear as colunas
df.rename(columns = {
    'Unnamed: 1':'Mes'
    ,'Unnamed: 2':'Piratininga_R$'
    ,'Unnamed: 3':'Piratininga_%'
    ,'Unnamed: 4':'Paulista_R$'
    ,'Unnamed: 5':'Paulista_%'                      
    ,'Unnamed: 6':'RGE_R$'
    ,'Unnamed: 7':'RGE_%'                        
    ,'Unnamed: 8':'Santa_Cruz_R$'
    ,'Unnamed: 9':'Santa_Cruz_%'                        
    }
    ,inplace = True)


#Formatar as colunas
df = df.astype("str")
df['Mes'] = pd.to_datetime(df['Mes'], format="%Y-%m-%d")
df['Mes'] = df['Mes'].dt.strftime('%m/%y')
df['Mes'] = df['Mes'].replace('-','').replace('NaT','')
df['Mes'] = df['Mes'].astype("str").replace('nan','').replace('-','').replace('NaT','')

df['Piratininga_R$'] = df['Piratininga_R$'].astype('float').replace('.', ',').replace(np.nan,0)
df['Piratininga_%'] = df['Piratininga_%'].astype('float').replace('.', ',').replace(np.nan,0)

df['Paulista_R$'] = df['Paulista_R$'].astype('float').replace('.', ',').replace(np.nan,0)
df['Paulista_%'] = df['Paulista_%'].astype('float').replace('.', ',').replace(np.nan,0)

df['RGE_R$'] = df['RGE_R$'].astype('float').replace('.', ',').replace(np.nan,0)
df['RGE_%'] = df['RGE_%'].astype('float').replace('.', ',').replace(np.nan,0)

df['Santa_Cruz_R$'] = df['Santa_Cruz_R$'].astype('float').replace('.', ',').replace(np.nan,0)
df['Santa_Cruz_%'] = df['Santa_Cruz_%'].astype('float').replace('.', ',').replace(np.nan,0)


#Preparação do dataframe da carga do ano em questão
ano_oracle = "'%" + ano[2:4]+"'"
df = df[df.loc[:,'Mes'].str.endswith(ano_filtro)] #Filtra somente o ano de interesse para insert




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
        cursor.execute('''DELETE FROM ''' + tabela_oracle + ''' WHERE MES LIKE ''' + ano_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:
        cursor.close()
        connection.close()

