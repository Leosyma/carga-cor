# -*- coding: utf-8 -*-
"""
Insert de dados para estudos COR - Fonte de dados Juridico

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
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2023 - 3TRI - 15ªReunião\1. Arquivos\Jurídico"

#Arquivos
arquivo = '09_PGTOS CONSOLIDADOS_2023_Anderson Ferrari_Ajustado.xlsx'

#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_PAGAMENTOS_CONSOLIDADOS'

ano_oracle = "'%/23'"




#%% Abertura dos arquivos


#Carregar arquivos
df = pd.read_excel(os.path.join(pasta, arquivo)
                   ).reset_index(drop=True) #Excluir a coluna Index


#Validações
colunas_chaves = ['MÊS/ANO','EMPRESA DO GRUPO CPFL','NATUREZA DA AÇÃO','DEPARTAMENTO','TOTAL PAGAMENTOS']
for coluna in colunas_chaves:
    if df[coluna].isnull().any() == True:
        print('Existe Campo Nula: ',coluna)

#Formatar as colunas
df = df.astype("str")
df['MÊS/ANO'] = pd.to_datetime(df['MÊS/ANO'], format="%Y-%m-%d")
df['MÊS/ANO'] = df['MÊS/ANO'].dt.strftime('%m/%y')
df['MÊS/ANO'] = df['MÊS/ANO'].replace('-','').replace('NaT','')
df['MÊS/ANO'] = df['MÊS/ANO'].astype("str").replace('nan','').replace('-','').replace('NaT','')

df['DATA DO CADASTRO'] = pd.to_datetime(df['DATA DO CADASTRO'], format="%Y-%m-%d")
df['DATA DO CADASTRO'] = df['DATA DO CADASTRO'].dt.strftime('%d/%m/%y')
df['DATA DO CADASTRO'] = df['DATA DO CADASTRO'].replace('-','').replace('NaT','')
df['DATA DO CADASTRO'] = df['DATA DO CADASTRO'].astype("str").replace('nan','').replace('-','').replace('NaT','')

df['PAGAMENTO'] = df['PAGAMENTO'].astype('float').replace('.', ',').replace(np.nan,0)
df['BAIXA DEPÓSITO'] = df['BAIXA DEPÓSITO'].astype('float').replace('.', ',').replace(np.nan,0)
df['TOTAL PAGAMENTOS'] = df['TOTAL PAGAMENTOS'].astype('float').replace('.', ',').replace(np.nan,0)

df['OUTROS'] = df['OUTROS'].replace('nan', '').replace(' ','')
df['OBSERVAÇÃO'] = df['OBSERVAÇÃO'].replace('nan', '').replace(' ','')
df = df.replace('nan', '')



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
        cursor.execute('''DELETE FROM ''' + tabela_oracle + ''' WHERE MES_ANO LIKE ''' + ano_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17,:18,:19)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:
        cursor.close()
        connection.close()

