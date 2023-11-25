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
arquivo = 'Modelo Dados_CB05.2024_Gustavo Nicolau.xlsx'
aba = 'Mercado'
aba1 = 'Compensação'
aba2 = 'CPFL Paulista'
aba3 = 'CPFL Piratininga'
aba4 = 'CPFL Santa Cruz'
aba5 = 'RGE'


#Nome da tabela Oracle onde será dada a carga
tabela_oracle = 'COR_ORCAMENTO_RECEITA'

#Ano dos dados
ano = '2023'
ano_oracle = "'%23'"




#%% Abertura dos arquivos


#Carregar arquivos 
df = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba)   #aba 'Mercado'

df1 = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba1)   #aba 'Compensação'

df2 = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba2)   #aba 'CPFL Paulista'

df3 = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba3)   #aba 'CPFL Piratininga'

df4 = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba4)   #aba 'CPFL Santa Cruz'

df5 = pd.read_excel(os.path.join(pasta, arquivo)
                   ,sheet_name = aba5)   #aba 'CPFL RGE'




#Formatar as colunas
##Formatação da aba 'Mercado'
#Copia o código para as outras linhas
df.iat[1,0] = df.iat[0,0]
df.iat[2,0] = df.iat[0,0]
df.iat[3,0] = df.iat[0,0]
df.iat[4,0] = df.iat[0,0]
df.iat[5,0] = df.iat[0,0]
df.iat[11,0] = df.iat[10,0]
df.iat[12,0] = df.iat[10,0]
df.iat[13,0] = df.iat[10,0]
df.iat[14,0] = df.iat[10,0]
df.iat[15,0] = df.iat[10,0]


df.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'
df = df.drop(df.index[[0,6,7,8,9,10]]) #Remover as linhas em branco e que não interessam
df = df.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df.columns.values[0] = 'CODIGO'
df.columns.values[3] = 'DISTRIBUIDORA'
df.columns.values[4] = 'JAN'
df.columns.values[5] = 'FEV'
df.columns.values[6] = 'MAR'
df.columns.values[7] = 'ABR'
df.columns.values[8] = 'MAI'
df.columns.values[9] = 'JUN'
df.columns.values[10] = 'JUL'
df.columns.values[11] = 'AGO'
df.columns.values[12] = 'SET'
df.columns.values[13] = 'OUT'
df.columns.values[14] = 'NOV'
df.columns.values[15] = 'DEZ'
df.columns.values[16] = 'TOTAL'

#Mudar o nome das linhas
df.iat[5,1] = 'Melhor Estimativa' #Muda o nome do indice 5 e coluna 1 para 'Melhor Estimativa'
df.iat[6,1] = 'Melhor Estimativa' #Muda o nome do indice 6 e coluna 1 para 'Melhor Estimativa'
df.iat[7,1] = 'Melhor Estimativa' #Muda o nome do indice 7 e coluna 1 para 'Melhor Estimativa'
df.iat[8,1] = 'Melhor Estimativa' #Muda o nome do indice 8 e coluna 1 para 'Melhor Estimativa'
df.iat[9,1] = 'Melhor Estimativa' #Muda o nome do indice 9 e coluna 1 para 'Melhor Estimativa'


#Limpeza e Tratamento dos dados
df = df.astype(str)
df['DISTRIBUIDORA'] = df['DISTRIBUIDORA'].replace('nan','TOTAL/MES').replace(np.nan,0)
df['JAN'] = df['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['FEV'] = df['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['MAR'] = df['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['ABR'] = df['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['MAI'] = df['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['JUN'] = df['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['JUL'] = df['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['AGO'] = df['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['SET'] = df['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['OUT'] = df['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['NOV'] = df['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['DEZ'] = df['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df['TOTAL'] = df['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000






##Formatação da aba 'Compensação'
#Copia o código para as outras linhas
df1.iat[1,0] = df1.iat[0,0]
df1.iat[2,0] = df1.iat[0,0]
df1.iat[3,0] = df1.iat[0,0]
df1.iat[4,0] = df1.iat[0,0]
df1.iat[5,0] = df1.iat[0,0]
df1.iat[11,0] = df1.iat[10,0]
df1.iat[12,0] = df1.iat[10,0]
df1.iat[13,0] = df1.iat[10,0]
df1.iat[14,0] = df1.iat[10,0]
df1.iat[15,0] = df1.iat[10,0]

df1.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df1.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'

#Muda o nome da coluna 'CLASSE' para 'DIC/FIC/DMIC/DICRI/VUP/Nível de Tensão'
df1.iat[0,2] = df1.iat[0,3]
df1.iat[1,2] = df1.iat[0,3]
df1.iat[2,2] = df1.iat[0,3]
df1.iat[3,2] = df1.iat[0,3]
df1.iat[4,2] = df1.iat[0,3]
df1.iat[5,2] = df1.iat[0,3]
df1.iat[11,2] = df1.iat[0,3]
df1.iat[12,2] = df1.iat[0,3]
df1.iat[13,2] = df1.iat[0,3]
df1.iat[14,2] = df1.iat[0,3]
df1.iat[15,2] = df1.iat[0,3]


df1 = df1.drop(df1.index[[0,6,7,8,9,10]]) #Remover as linhas em branco e que não interessam
df1 = df1.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df1.columns.values[0] = 'CODIGO'
df1.columns.values[3] = 'DISTRIBUIDORA'
df1.columns.values[4] = 'JAN'
df1.columns.values[5] = 'FEV'
df1.columns.values[6] = 'MAR'
df1.columns.values[7] = 'ABR'
df1.columns.values[8] = 'MAI'
df1.columns.values[9] = 'JUN'
df1.columns.values[10] = 'JUL'
df1.columns.values[11] = 'AGO'
df1.columns.values[12] = 'SET'
df1.columns.values[13] = 'OUT'
df1.columns.values[14] = 'NOV'
df1.columns.values[15] = 'DEZ'
df1.columns.values[16] = 'TOTAL'

#Mudar o nome das linhas
df1.iat[5,1] = 'Melhor Estimativa' #Muda o nome do indice 5 e coluna 1 para 'Melhor Estimativa'
df1.iat[6,1] = 'Melhor Estimativa' #Muda o nome do indice 6 e coluna 1 para 'Melhor Estimativa'
df1.iat[7,1] = 'Melhor Estimativa' #Muda o nome do indice 7 e coluna 1 para 'Melhor Estimativa'
df1.iat[8,1] = 'Melhor Estimativa' #Muda o nome do indice 8 e coluna 1 para 'Melhor Estimativa'
df1.iat[9,1] = 'Melhor Estimativa' #Muda o nome do indice 9 e coluna 1 para 'Melhor Estimativa'


#Limpeza e Tratamento dos dados
df1 = df1.astype(str)
df1['DISTRIBUIDORA'] = df1['DISTRIBUIDORA'].replace('nan','TOTAL/MES').replace(np.nan,0)
df1['JAN'] = df1['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['FEV'] = df1['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['MAR'] = df1['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['ABR'] = df1['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['MAI'] = df1['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['JUN'] = df1['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['JUL'] = df1['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['AGO'] = df1['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['SET'] = df1['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['OUT'] = df1['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['NOV'] = df1['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['DEZ'] = df1['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df1['TOTAL'] = df1['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)




##Formatação da aba 'CPFL Paulista'
df2.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df2.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'
df2.insert(16,'TOTAL', 'TOTAL') #Inserir uma coluna com o nome 'TOTAL'
df2 = df2.drop(df2.index[[0,2,3,5,6,7,8,10,11]]) #Remover as linhas em branco e que não interessam
df2 = df2.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df2.columns.values[0] = 'CODIGO'
df2.columns.values[3] = 'DISTRIBUIDORA'
df2.columns.values[4] = 'JAN'
df2.columns.values[5] = 'FEV'
df2.columns.values[6] = 'MAR'
df2.columns.values[7] = 'ABR'
df2.columns.values[8] = 'MAI'
df2.columns.values[9] = 'JUN'
df2.columns.values[10] = 'JUL'
df2.columns.values[11] = 'AGO'
df2.columns.values[12] = 'SET'
df2.columns.values[13] = 'OUT'
df2.columns.values[14] = 'NOV'
df2.columns.values[15] = 'DEZ'

#Mudar o nome das linhas
df2.iat[1,1] = 'Melhor Estimativa' #Muda o nome do indice 1 e coluna 1 para 'Melhor Estimativa'
df2.iat[3,1] = 'Melhor Estimativa' #Muda o nome do indice 3 e coluna 1 para 'Melhor Estimativa'
df2.iat[0,2] = 'Cobertura COR' #Muda o nome do indice 0 e coluna 2 para 'Cobertura COR'
df2.iat[1,2] = 'Cobertura COR' #Muda o nome do indice 1 e coluna 2 para 'Cobertura COR'
df2.iat[2,2] = 'Receita de Fio B Total' #Muda o nome do indice 2 e coluna 2 para 'Receita de Fio B Total'
df2.iat[3,2] = 'Receita de Fio B Total' #Muda o nome do indice 3 e coluna 2 para 'Receita de Fio B Total'
df2.iat[0,3] = 'CPFL Paulista' #Muda o nome do indice 0 e coluna 3 para 'CPFL Paulista'
df2.iat[1,3] = 'CPFL Paulista' #Muda o nome do indice 0 e coluna 3 para 'CPFL Paulista'
df2.iat[2,3] = 'CPFL Paulista' #Muda o nome do indice 0 e coluna 3 para 'CPFL Paulista'
df2.iat[3,3] = 'CPFL Paulista' #Muda o nome do indice 0 e coluna 3 para 'CPFL Paulista'

#Limpeza e Tratamento dos dados
df2 = df2.astype(str)
df2['DISTRIBUIDORA'] = df2['DISTRIBUIDORA'].replace(np.nan,0)
df2['JAN'] = df2['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['FEV'] = df2['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['MAR'] = df2['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['ABR'] = df2['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['MAI'] = df2['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['JUN'] = df2['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['JUL'] = df2['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['AGO'] = df2['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['SET'] = df2['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['OUT'] = df2['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['NOV'] = df2['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df2['DEZ'] = df2['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000



#Soma Total dos Meses
df2['TOTAL'] = df2['JAN'] + df2['FEV'] + df2['MAR'] + df2['ABR'] + df2['MAI'] + df2['JUN'] + df2['JUL'] + df2['AGO'] + df2['SET'] + df2['OUT'] + df2['NOV'] + df2['DEZ']
df2['TOTAL'] = df2['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)





##Formatação da aba 'CPFL Piratininga'
df3.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df3.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'
df3 = df3.drop(df3.index[[0,2,3,5,6,7,8,10,11]]) #Remover as linhas em branco e que não interessam
df3 = df3.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df3.columns.values[0] = 'CODIGO'
df3.columns.values[3] = 'DISTRIBUIDORA'
df3.columns.values[4] = 'JAN'
df3.columns.values[5] = 'FEV'
df3.columns.values[6] = 'MAR'
df3.columns.values[7] = 'ABR'
df3.columns.values[8] = 'MAI'
df3.columns.values[9] = 'JUN'
df3.columns.values[10] = 'JUL'
df3.columns.values[11] = 'AGO'
df3.columns.values[12] = 'SET'
df3.columns.values[13] = 'OUT'
df3.columns.values[14] = 'NOV'
df3.columns.values[15] = 'DEZ'

#Mudar o nome das linhas
df3.iat[1,1] = 'Melhor Estimativa' #Muda o nome do indice 1 e coluna 1 para 'Melhor Estimativa'
df3.iat[3,1] = 'Melhor Estimativa' #Muda o nome do indice 3 e coluna 1 para 'Melhor Estimativa'
df3.iat[0,2] = 'Cobertura COR' #Muda o nome do indice 0 e coluna 2 para 'Cobertura COR'
df3.iat[1,2] = 'Cobertura COR' #Muda o nome do indice 1 e coluna 2 para 'Cobertura COR'
df3.iat[2,2] = 'Receita de Fio B Total' #Muda o nome do indice 2 e coluna 2 para 'Receita de Fio B Total'
df3.iat[3,2] = 'Receita de Fio B Total' #Muda o nome do indice 3 e coluna 2 para 'Receita de Fio B Total'
df3.iat[0,3] = 'CPFL Piratininga' #Muda o nome do indice 0 e coluna 3 para 'CPFL Piratininga'
df3.iat[1,3] = 'CPFL Piratininga' #Muda o nome do indice 0 e coluna 3 para 'CPFL Piratininga'
df3.iat[2,3] = 'CPFL Piratininga' #Muda o nome do indice 0 e coluna 3 para 'CPFL Piratininga'
df3.iat[3,3] = 'CPFL Piratininga' #Muda o nome do indice 0 e coluna 3 para 'CPFL Piratininga'

#Limpeza e Tratamento dos dados
df3 = df3.astype(str)
df3['DISTRIBUIDORA'] = df3['DISTRIBUIDORA'].replace(np.nan,0)
df3['JAN'] = df3['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['FEV'] = df3['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['MAR'] = df3['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['ABR'] = df3['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['MAI'] = df3['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['JUN'] = df3['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['JUL'] = df3['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['AGO'] = df3['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['SET'] = df3['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['OUT'] = df3['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['NOV'] = df3['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df3['DEZ'] = df3['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000

#Soma Total dos Meses
df3['TOTAL'] = df3['JAN'] + df3['FEV'] + df3['MAR'] + df3['ABR'] + df3['MAI'] + df3['JUN'] + df3['JUL'] + df3['AGO'] + df3['SET'] + df3['OUT'] + df3['NOV'] + df3['DEZ']
df3['TOTAL'] = df3['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)






##Formatação da aba 'CPFL Santa Cruz'
df4.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df4.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'
df4 = df4.drop(df4.index[[0,2,3,5,6,7,8,10,11]]) #Remover as linhas em branco e que não interessam
df4 = df4.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df4.columns.values[0] = 'CODIGO'
df4.columns.values[3] = 'DISTRIBUIDORA'
df4.columns.values[4] = 'JAN'
df4.columns.values[5] = 'FEV'
df4.columns.values[6] = 'MAR'
df4.columns.values[7] = 'ABR'
df4.columns.values[8] = 'MAI'
df4.columns.values[9] = 'JUN'
df4.columns.values[10] = 'JUL'
df4.columns.values[11] = 'AGO'
df4.columns.values[12] = 'SET'
df4.columns.values[13] = 'OUT'
df4.columns.values[14] = 'NOV'
df4.columns.values[15] = 'DEZ'

#Mudar o nome das linhas
df4.iat[1,1] = 'Melhor Estimativa' #Muda o nome do indice 1 e coluna 1 para 'Melhor Estimativa'
df4.iat[3,1] = 'Melhor Estimativa' #Muda o nome do indice 3 e coluna 1 para 'Melhor Estimativa'
df4.iat[0,2] = 'Cobertura COR' #Muda o nome do indice 0 e coluna 2 para 'Cobertura COR'
df4.iat[1,2] = 'Cobertura COR' #Muda o nome do indice 1 e coluna 2 para 'Cobertura COR'
df4.iat[2,2] = 'Receita de Fio B Total' #Muda o nome do indice 2 e coluna 2 para 'Receita de Fio B Total'
df4.iat[3,2] = 'Receita de Fio B Total' #Muda o nome do indice 3 e coluna 2 para 'Receita de Fio B Total'
df4.iat[0,3] = 'CPFL Santa Cruz' #Muda o nome do indice 0 e coluna 3 para 'CPFL Santa Cruz'
df4.iat[1,3] = 'CPFL Santa Cruz' #Muda o nome do indice 0 e coluna 3 para 'CPFL Santa Cruz'
df4.iat[2,3] = 'CPFL Santa Cruz' #Muda o nome do indice 0 e coluna 3 para 'CPFL Santa Cruz'
df4.iat[3,3] = 'CPFL Santa Cruz' #Muda o nome do indice 0 e coluna 3 para 'CPFL Santa Cruz'

#Limpeza e Tratamento dos dados
df4 = df4.astype(str)
df4['DISTRIBUIDORA'] = df4['DISTRIBUIDORA'].replace(np.nan,0)
df4['JAN'] = df4['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['FEV'] = df4['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['MAR'] = df4['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['ABR'] = df4['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['MAI'] = df4['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['JUN'] = df4['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['JUL'] = df4['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['AGO'] = df4['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['SET'] = df4['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['OUT'] = df4['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['NOV'] = df4['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df4['DEZ'] = df4['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000

#Soma Total dos Meses
df4['TOTAL'] = df4['JAN'] + df4['FEV'] + df4['MAR'] + df4['ABR'] + df4['MAI'] + df4['JUN'] + df4['JUL'] + df4['AGO'] + df4['SET'] + df4['OUT'] + df4['NOV'] + df4['DEZ']
df4['TOTAL'] = df4['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)






##Formatação da aba 'RGE'
df5.insert(1,'TIPO','Orçamento') #Inserir uma coluna com o nome 'TIPO'
df5.insert(2,'CLASSE', 'Efeito Crescimento de Mercado') #Inserir uma coluna com o nome 'CLASSE'
df5 = df5.drop(df5.index[[0,2,3,5,6,7,8,10,11]]) #Remover as linhas em branco e que não interessam
df5 = df5.reset_index(drop=True)  #Reseta o índice

#Mudar o nome das colunas
df5.columns.values[0] = 'CODIGO'
df5.columns.values[3] = 'DISTRIBUIDORA'
df5.columns.values[4] = 'JAN'
df5.columns.values[5] = 'FEV'
df5.columns.values[6] = 'MAR'
df5.columns.values[7] = 'ABR'
df5.columns.values[8] = 'MAI'
df5.columns.values[9] = 'JUN'
df5.columns.values[10] = 'JUL'
df5.columns.values[11] = 'AGO'
df5.columns.values[12] = 'SET'
df5.columns.values[13] = 'OUT'
df5.columns.values[14] = 'NOV'
df5.columns.values[15] = 'DEZ'

#Mudar o nome das linhas
df5.iat[1,1] = 'Melhor Estimativa' #Muda o nome do indice 1 e coluna 1 para 'Melhor Estimativa'
df5.iat[3,1] = 'Melhor Estimativa' #Muda o nome do indice 3 e coluna 1 para 'Melhor Estimativa'
df5.iat[0,2] = 'Cobertura COR' #Muda o nome do indice 0 e coluna 2 para 'Cobertura COR'
df5.iat[1,2] = 'Cobertura COR' #Muda o nome do indice 1 e coluna 2 para 'Cobertura COR'
df5.iat[2,2] = 'Receita de Fio B Total' #Muda o nome do indice 2 e coluna 2 para 'Receita de Fio B Total'
df5.iat[3,2] = 'Receita de Fio B Total' #Muda o nome do indice 3 e coluna 2 para 'Receita de Fio B Total'
df5.iat[0,3] = 'RGE' #Muda o nome do indice 0 e coluna 3 para 'RGE'
df5.iat[1,3] = 'RGE' #Muda o nome do indice 0 e coluna 3 para 'RGE'
df5.iat[2,3] = 'RGE' #Muda o nome do indice 0 e coluna 3 para 'RGE'
df5.iat[3,3] = 'RGE' #Muda o nome do indice 0 e coluna 3 para 'RGE'

#Limpeza e Tratamento dos dados
df5 = df5.astype(str)
df5['DISTRIBUIDORA'] = df5['DISTRIBUIDORA'].replace(np.nan,0)
df5['JAN'] = df5['JAN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['FEV'] = df5['FEV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['MAR'] = df5['MAR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['ABR'] = df5['ABR'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['MAI'] = df5['MAI'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['JUN'] = df5['JUN'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['JUL'] = df5['JUL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['AGO'] = df5['AGO'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['SET'] = df5['SET'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['OUT'] = df5['OUT'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['NOV'] = df5['NOV'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000
df5['DEZ'] = df5['DEZ'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)*1000

#Soma Total dos Meses
df5['TOTAL'] = df5['JAN'] + df5['FEV'] + df5['MAR'] + df5['ABR'] + df5['MAI'] + df5['JUN'] + df5['JUL'] + df5['AGO'] + df5['SET'] + df5['OUT'] + df5['NOV'] + df5['DEZ']
df5['TOTAL'] = df5['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)





#Juntar os dataframes
df6 = pd.concat([df,df1,df2,df3,df4,df5])

#Limpeza e Tratamento dos dados
df6['TOTAL'] = df6['TOTAL'].astype('float').replace('.', ',').replace('nan', '0').replace(np.nan,0)



#Montar um dataframe apenas com as colunas que serão inseridas no banco de dados
df_carga = df6[['CODIGO','TIPO','CLASSE','DISTRIBUIDORA','JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ','TOTAL' ]]

#Acrescentar a coluna ano no arquivo que será inserido no banco de dados
df_carga.insert(0,'ANO',ano) #Inserir uma coluna com o ANO




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
        cursor.execute('''DELETE FROM ''' + tabela_oracle + ''' WHERE ANO LIKE ''' + ano_oracle) #Limpar a tabela antes de executar o insert
        sql = '''INSERT INTO ''' + tabela_oracle +''' VALUES (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12,:13,:14,:15,:16,:17,:18)''' #Deve ser igual ao número de colunas da tabela do banco de dados
        cursor.executemany(sql, dados_list)
    except Exception as err:
        print('Erro no Insert:', err)
    else:
        print('Carga executada com sucesso!')
        connection.commit() #Caso seja executado com sucesso, esse comando salva a base de dados
    finally:    
        cursor.close()
        connection.close()
