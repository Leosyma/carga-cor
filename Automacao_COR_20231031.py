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
pasta = r"X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\1. PMSO Grupo CPFL\1. Arquivos IRA\Layout"
pasta_saida = r'X:\Estudos Regulatórios\Alocação de Gastos\12. PMSO CPFL e VPR OPEX\1. PMSO Grupo CPFL\1. Arquivos IRA\2023.10 - 3T2023'

#Empresas
empresas = ['Paulista','Piratininga','RGE','Santa Cruz']


#%%Parâmetros Adicionais
#Tabelas Banco de Dados Oracle
tabela_oracle_bmp = 'VW_COR_BMP'
tabela_oracle_cobertura = 'VW_COR_COBERTURA'
tabela_oracle_despesa = 'VW_COR_ORCAMENTO_DESPESA'
tabela_oracle_receita = 'VW_COR_ORCAMENTO_RECEITA'
tabela_oracle_multas = 'VW_COR_MULTAS'


#%% Função para consolidar os dados
def consolidacao(empresa):
    #Acessar o banco de dados para importar os dados 
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
            cursor_bmp = connection.cursor()
            cursor_despesa = connection.cursor()
            cursor_receita = connection.cursor()
            cursor_cobertura = connection.cursor()
            cursor_multas = connection.cursor()
            
            #Importação das tabelas
            print('Carregando tabela BMP...')
            cursor_bmp.execute('''SELECT * FROM ''' + tabela_oracle_bmp) #Importa a tabela 'BMP'
            print('Carregando tabela ORCAMENTO_DESPESA...')
            cursor_despesa.execute('''SELECT * FROM ''' + tabela_oracle_despesa) #Importa a tabela 'Orçamento Despesa' 
            print('Carregando tabela ORCAMENTO_RECEITA...')
            cursor_receita.execute('''SELECT * FROM ''' + tabela_oracle_receita) #Importa a tabela 'Orçamento Receita' 
            print('Carregando tabela COBERTURA...')
            cursor_cobertura.execute('''SELECT * FROM ''' + tabela_oracle_cobertura) #Importa a tabela 'Cobertura'
            print('Carregando tabela MULTAS...')
            cursor_multas.execute('''SELECT * FROM ''' + tabela_oracle_multas) #Importa a tabela 'Multas'
    
            
            #Transforma as tabelas em DataFrame
            df_bmp_oracle=pd.DataFrame(cursor_bmp)
            df_despesa_oracle = pd.DataFrame(cursor_despesa)
            df_receita_oracle = pd.DataFrame(cursor_receita)
            df_cobertura_oracle = pd.DataFrame(cursor_cobertura)
            df_multas_oracle = pd.DataFrame(cursor_multas)
            
            
        except Exception as err:
            print('Erro no Importação:', err)
        else:
            print('Tabelas importadas com sucesso!')
        finally:    
            cursor_bmp.close()
            cursor_cobertura.close()
            cursor_despesa.close()
            cursor_receita.close()
            cursor_multas.close()
            connection.close()
    
    
    #Tratamento da tabela importada do Banco de Dados Oracle
    #TABELA BMP
    #Renomear o nome das colunas
    df_bmp_oracle = df_bmp_oracle.rename(columns={0:'número&ano&mês'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={1:'codigo_empresa'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={2:'ano'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={3:'mes'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={4:'numero'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={5:'debito'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={6:'credito'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={7:'saldo'})
    df_bmp_oracle = df_bmp_oracle.rename(columns={8:'variação(debito-credito)'})
    
    #Limpeza e tratamento dos dados
    df_bmp_oracle = df_bmp_oracle.astype('str')
    df_bmp_oracle['debito'] = df_bmp_oracle['debito'].astype('float')
    df_bmp_oracle['credito'] = df_bmp_oracle['credito'].astype('float')
    df_bmp_oracle['saldo'] = df_bmp_oracle['saldo'].astype('float')
    df_bmp_oracle['variação(debito-credito)'] = df_bmp_oracle['variação(debito-credito)'].astype('float')
    
    
    
    #TABELA COBERTURA
    #Renomear o nome das colunas
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={0:'MES'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={1:'PIRATININGA_R$'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={2:'PIRATININGA_%'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={3:'PAULISTA_R$'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={4:'PAULISTA_%'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={5:'RGE_R$'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={6:'RGE_%'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={7:'SANTA_CRUZ_R$'})
    df_cobertura_oracle = df_cobertura_oracle.rename(columns={8:'SANTA_CRUZ_%'})
    
    
    #TABELA ORÇAMENTO DESPESA
    #Renomear o nome das colunas
    df_despesa_oracle = df_despesa_oracle.rename(columns={0:'CHAVE'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={1:'EMPRESA'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={2:'TIPO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={3:'COR_NAO_COR'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={4:'NATUREZA'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={5:'ANO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={6:'JANEIRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={7:'FEVEREIRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={8:'MARCO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={9:'ABRIL'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={10:'MAIO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={11:'JUNHO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={12:'JULHO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={13:'AGOSTO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={14:'SETEMBRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={15:'OUTUBRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={16:'NOVEMBRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={17:'DEZEMBRO'})
    df_despesa_oracle = df_despesa_oracle.rename(columns={18:'ACUMULADO'})
    
    
    
    #TABELA ORÇAMENTO RECEITA
    #Renomear o nome das colunas
    df_receita_oracle = df_receita_oracle.rename(columns={0:'CHAVE'})
    df_receita_oracle = df_receita_oracle.rename(columns={1:'ANO'})
    df_receita_oracle = df_receita_oracle.rename(columns={2:'CODIGO'})
    df_receita_oracle = df_receita_oracle.rename(columns={3:'TIPO'})
    df_receita_oracle = df_receita_oracle.rename(columns={4:'CLASSE'})
    df_receita_oracle = df_receita_oracle.rename(columns={5:'DISTRIBUIDORA'})
    df_receita_oracle = df_receita_oracle.rename(columns={6:'JAN'})
    df_receita_oracle = df_receita_oracle.rename(columns={7:'FEV'})
    df_receita_oracle = df_receita_oracle.rename(columns={8:'MAR'})
    df_receita_oracle = df_receita_oracle.rename(columns={9:'ABR'})
    df_receita_oracle = df_receita_oracle.rename(columns={10:'MAI'})
    df_receita_oracle = df_receita_oracle.rename(columns={11:'JUN'})
    df_receita_oracle = df_receita_oracle.rename(columns={12:'JUL'})
    df_receita_oracle = df_receita_oracle.rename(columns={13:'AGO'})
    df_receita_oracle = df_receita_oracle.rename(columns={14:'SET'})
    df_receita_oracle = df_receita_oracle.rename(columns={15:'OUT'})
    df_receita_oracle = df_receita_oracle.rename(columns={16:'NOV'})
    df_receita_oracle = df_receita_oracle.rename(columns={17:'DEZ'})
    df_receita_oracle = df_receita_oracle.rename(columns={18:'TOTAL'})
    
    
    #TABELA MULTAS
    #Renomear o nome das colunas
    df_multas_oracle = df_multas_oracle.rename(columns={0:'CHAVE'})
    df_multas_oracle = df_multas_oracle.rename(columns={1:'EMPRESA'})
    df_multas_oracle = df_multas_oracle.rename(columns={2:'CONTA'})
    df_multas_oracle = df_multas_oracle.rename(columns={3:'TIPO'})
    df_multas_oracle = df_multas_oracle.rename(columns={4:'ANO'})
    df_multas_oracle = df_multas_oracle.rename(columns={5:'TOTAL'})
    df_multas_oracle = df_multas_oracle.rename(columns={6:'JAN'})
    df_multas_oracle = df_multas_oracle.rename(columns={7:'FEV'})
    df_multas_oracle = df_multas_oracle.rename(columns={8:'MAR'})
    df_multas_oracle = df_multas_oracle.rename(columns={9:'ABR'})
    df_multas_oracle = df_multas_oracle.rename(columns={10:'MAI'})
    df_multas_oracle = df_multas_oracle.rename(columns={11:'JUN'})
    df_multas_oracle = df_multas_oracle.rename(columns={12:'JUL'})
    df_multas_oracle = df_multas_oracle.rename(columns={13:'AGO'})
    df_multas_oracle = df_multas_oracle.rename(columns={14:'SET'})
    df_multas_oracle = df_multas_oracle.rename(columns={15:'OUT'})
    df_multas_oracle = df_multas_oracle.rename(columns={16:'NOV'})
    df_multas_oracle = df_multas_oracle.rename(columns={17:'DEZ'})
    
    
    
    
    #Filtrar a empresa desejada
    #df_bmp_oracle = df_bmp_oracle[df_bmp_oracle['codigo_empresa'] == codigo_empresa]
    #df_despesa_oracle = df_despesa_oracle[df_despesa_oracle['EMPRESA'] == codigo_distribuidora]
    #df_receita_oracle = df_receita_oracle[df_receita_oracle['DISTRIBUIDORA'] == nome_distribuidora]
    
    df_bmp_oracle = df_bmp_oracle.loc[df_bmp_oracle['codigo_empresa'].isin(codigo_empresa)] 
    df_despesa_oracle = df_despesa_oracle.loc[df_despesa_oracle['EMPRESA'].isin(codigo_distribuidora)] 
    df_receita_oracle = df_receita_oracle.loc[df_receita_oracle['DISTRIBUIDORA'].isin(nome_distribuidora)] 
    df_multas_oracle = df_multas_oracle.loc[df_multas_oracle['EMPRESA'].isin(codigo_distribuidora)] 
    
    
    
    
    #Filtrar colunas da aba 'Cobertura'
    if empresa == 'Paulista':
        df_cobertura_oracle = df_cobertura_oracle.drop(df_cobertura_oracle.columns[[1,2,5,6,7,8]], axis = 1)
    elif empresa == 'Piratininga':
        df_cobertura_oracle = df_cobertura_oracle.drop(df_cobertura_oracle.columns[[3,4,5,6,7,8]], axis = 1)
    
    elif empresa == 'RGE':
        df_cobertura_oracle = df_cobertura_oracle.drop(df_cobertura_oracle.columns[[1,2,3,4,7,8]], axis = 1)
    
    elif empresa == 'Santa Cruz':
        df_cobertura_oracle = df_cobertura_oracle.drop(df_cobertura_oracle.columns[[1,2,3,4,5,6]], axis = 1)
    
    
        
    #Inserir uma coluna na aba 'Cobertura' com o código da empresa
    df_cobertura_oracle.insert(1,'DISTRIBUIDORA', empresa)
    
    
    
    
    #Inserção dos novos dados nas respectivas abas da planilha de referência
    
    #Abrir a planilha para inserir o DataFrame editado
    book = load_workbook(os.path.join(pasta, arquivo))
    writer = pd.ExcelWriter(os.path.join(pasta_saida, nome_arquivo + '.xlsx'), engine='openpyxl')
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
    
    df_bmp_oracle.to_excel(writer, sheet_name='BMP',index=False) #Insere o dataframe 'BMP' na planilha existente
    df_cobertura_oracle.to_excel(writer, sheet_name = 'Cobertura', index = False) #Insere o dataframe 'Cobertura' na planilha
    df_despesa_oracle.to_excel(writer, sheet_name = 'Orçamento Despesa', index = False, startrow = 1) #Insere o dataframe 'Orçamento Despesa' na planilha
    df_receita_oracle.to_excel(writer, sheet_name = 'Orçamento Receita', index = False, startrow = 1) #Insere o dataframe 'Orçamento Receita' na planilha
    df_multas_oracle.to_excel(writer, sheet_name = 'Multas Real', index = False, startrow = 1) #Insere o dataframe 'Multas Real' na planilha
    
    
    #Salva o arquivo e fecha
    writer.save()
    writer.close()
    writer.handles = None


#%%Consolida os dados
for empresa in empresas:
    if empresa == 'Paulista':
        codigo_empresa = ['63']
        codigo_distribuidora = ['D001']
        nome_distribuidora = ['CPFL Paulista']
        nome_arquivo = 'CPFL Paulista'
        arquivo = r'CPFL Paulista_layout.xlsx'
        consolidacao(empresa)
        print('Arquivo CPFL Paulista Exportado\n')
        
    elif empresa == 'Piratininga':
        codigo_empresa = ['2937']
        codigo_distribuidora = ['D002']
        nome_distribuidora = ['CPFL Piratininga']
        nome_arquivo = 'CPFL Piratininga'
        arquivo = r'CPFL Piratininga_layout.xlsx'
        consolidacao(empresa)
        print('Arquivo CPFL Piratininga Exportado\n')
        
    elif empresa == 'RGE':
        codigo_empresa = ['397', '396']
        codigo_distribuidora = ['D009']
        nome_distribuidora = ['RGE']
        nome_arquivo = 'RGE'
        arquivo = r'RGE_layout.xlsx'
        consolidacao(empresa)
        print('Arquivo RGE Exportado\n')
        
    elif empresa == 'Santa Cruz':
        codigo_empresa = ['69']
        codigo_distribuidora = ['D006']
        nome_distribuidora = ['CPFL Santa Cruz']
        nome_arquivo = 'CPFL Santa Cruz'
        arquivo = r'CPFL Santa Cruz_layout.xlsx'
        consolidacao(empresa)
        print('Arquivo CPFL Santa Cruz Exportado\n')
        
    



