# -*- coding: utf-8 -*-
"""
Insert de dados para estudos COR - Fonte de dados Controladoria

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
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\2023 - 3TRI - 15ªReunião\1. Arquivos\Controladoria"

#Arquivos
#Como os dados das empresas estão em arquivos separados, listar todos que serão carregados conforme layout abaixo.
arquivo = 'Relatorio_PMSO_2023_Cor_NCor_Set_23_Fernando Cesar.xlsx'
aba = 'Base'

#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_ORCAMENTO_DESPESA'

#Ano dos dados
ano = '2023'
ano_oracle = "'2023'"




#%% Abertura dos arquivos


df = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba) 



#Formatar as colunas
df = df.astype(str)
df['JAN'] = df['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['FEV'] = df['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['MAR'] = df['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ABR'] = df['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['MAI'] = df['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['JUN'] = df['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['JUL'] = df['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['AGO'] = df['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['SET'] = df['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['OUT'] = df['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['NOV'] = df['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['DEZ'] = df['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ACUMULADO'] = df['ACUMULADO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ ANO'] = df['ORÇ ANO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['REAL MÊS'] = df['REAL MÊS'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ MÊS'] = df['ORÇ MÊS'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ/REAL MÊS'] = df['ORÇ/REAL MÊS'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['REAL ACUM'] = df['REAL ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ ACUM'] = df['ORÇ ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ/REAL ACUM'] = df['ORÇ/REAL ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ANTECIP ACUM'] = df['ANTECIP ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['INCORP ACUM'] = df['INCORP ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['TRANSF'] = df['TRANSF'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['TRANSF ACUM'] = df['TRANSF ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['SALDO YTG'] = df['SALDO YTG'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['2021 MÊS NOMINAL'] = df['2021 MÊS NOMINAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['2021 MÊS ACUM NOMINAL'] = df['2021 MÊS ACUM NOMINAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['REAL 21/REAL 22 MÊS'] = df['REAL 21/REAL 22 MÊS'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['REAL 21/REAL 22 MÊS ACUM'] = df['REAL 21/REAL 22 MÊS ACUM'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ. MOV.'] = df['ORÇ. MOV.'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['OPORT. CAPTURA'] = df['OPORT. CAPTURA'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ FUTURO - MÊS'] = df['ORÇ FUTURO - MÊS'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['ORÇ FUTURO - %'] = df['ORÇ FUTURO - %'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)
df['COR/ NÃO COR'] = df['COR/ NÃO COR'].replace('nan','N/A')
df['Natureza'] = df['Natureza'].replace('nan','N/A')
df['COR/ NÃO COR_ant'] = df['COR/ NÃO COR_ant'].replace('nan','N/A')
df['Natureza_ant'] = df['Natureza_ant'].replace('nan','N/A')
df['ENT_OBZ (ant)'] = df['ENT_OBZ (ant)'].replace('nan','-')
df['Condição Exceção'] = df['Condição Exceção'].replace('nan','-')
df['Condição Exceção (ant)'] = df['Condição Exceção (ant)'].replace('nan','-')
df['Ajuste UNIO'] = df['Ajuste UNIO'].replace('nan','-')
df['Class. Rotina 1'] = df['Class. Rotina 1'].replace('nan','-').replace(np.nan,'-')
df['Class. Rotina 2'] = df['Class. Rotina 2'].replace('nan','-')
df['Rotina Despesa FM'] = df['Rotina Despesa FM'].replace('nan','-')





#Acrescentar a coluna ano no arquivo que será inserido no banco de dados
df.insert(0,'ANO',ano) #Inserir uma coluna com o nome 'TIPO'



#Montar um dataframe apenas com as colunas que serão inseridas no banco de dados
df_carga = df[['ANO','VP','DIRETORIA','DEPARTAMENTO','ENTIDADE','GRUPO BMP','PACOTE','PACKAGE','SUBPACKAGE','SUBPACOTE O&M RENOVÁVEIS','SUBPACKAGE O&M RENEWABLE','DESCRIÇÃO CC','Target Execão','TARGET CORPORATIVO','TARGET VP','TARGET CNPJ','TARGET DIRETORIA','CLASSIFICAÇÃO REPROT','NEGÓCIO','CLASSES FM','CONTROLE ORÇAMENTÁRIO','GRUPO BIU','SUB GRUPO BIU','CUSTO/DESPESA','CATEGORIA ORIGINAL','EMPRESA','DESCRIÇÃO DA EMPRESA','SUBPACOTE','CENTRO DE CUSTO','CLASSE DE CUSTO','DESCRIÇÃO DA CLASSE DE CUSTO','PMSO','JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ','ACUMULADO','ORÇ ANO','REAL MÊS','ORÇ MÊS','ORÇ/REAL MÊS','REAL ACUM','ORÇ ACUM','ORÇ/REAL ACUM','ANTECIP ACUM','INCORP ACUM','TRANSF','TRANSF ACUM','SALDO YTG','2021 MÊS NOMINAL','2021 MÊS ACUM NOMINAL','REAL 21/REAL 22 MÊS','REAL 21/REAL 22 MÊS ACUM','ORÇ. MOV.','OPORT. CAPTURA','ORÇ FUTURO - MÊS','ORÇ FUTURO - %','COR/ NÃO COR','Natureza','COR/ NÃO COR_ant','Natureza_ant','ENT_OBZ (ant)','spç','Condição Exceção','Condição Exceção (ant)','Ajuste UNIO','Rotina','Class. Rotina 1','Class. Rotina 2','Custos e Despesas','Rotina Despesa FM']]



#%% Inserir dados no banco de dados

#Criar a lista para inserção no banco SQL com os dados da base editada
dados_list = df_carga.values.tolist()


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
        cursor.execute('''DELETE FROM ''' + tabela_oracle + ''' WHERE ANO = ''' + ano_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17,:18,:19,:20,:21,:22,:23,:24,:25,:26,:27,:28,:29,:30,:31,:32,:33,:34,:35,:36,:37,:38,:39,:40,:41,:42,:43,:44,:45,:46,:47,:48,:49,:50,:51,:52,:53,:54,:55,:56,:57,:58,:59,:60,:61,:62,:63,:64,:65,:66,:67,:68,:69,:70,:71,:72,:73,:74,:75,:76,:77,:78,:79)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:    
        cursor.close()
        connection.close()

