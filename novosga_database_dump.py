# encoding=utf8
import sys
import xlsxwriter
import psycopg2
from datetime import date

try:
   conn = psycopg2.connect(
      host='localhost',
      database='novosga',
      user='postgres',
      password='123456')
except psycopg2.Error as err:
      print(err)
else:
   cursor = conn.cursor()
   arquivo_excel = xlsxwriter.Workbook('NovoSGADatabaseDump.xlsx')
   planilha = arquivo_excel.add_worksheet()

   titulo_colunas = [
         'ATENDIMENTO CODIFICADO : ATENDIMENTO_ID',
         'ATENDIMENTO CODIFICADO : SERVICO_ID',
         'ATENDIMENTO CODIFICADO : VALOR_PESO',
         'ATENDIMENTOS : ID',
         'ATENDIMENTOS : UNIDADE_ID',
         'ATENDIMENTOS : USUARIO_ID',
         'ATENDIMENTOS : USUARIO_TRI_ID',
         'ATENDIMENTOS : SERVICO_ID',
         'ATENDIMENTOS : PRIORIDADE_ID',
         'ATENDIMENTOS : ATENDIMENTO_ID',
         'ATENDIMENTOS : STATUS',
         'ATENDIMENTOS : SIGLA_SENHA',
         'ATENDIMENTOS : NUM_SENHA',
         'ATENDIMENTOS : NUM_SENHA_SERV',
         'ATENDIMENTOS : NM_CLI',
         'ATENDIMENTOS : NUM_LOCAL',
         'ATENDIMENTOS : DT_CHEGADA',
         'ATENDIMENTOS : DT_CHAMADA',
         'ATENDIMENTOS : DT_INICIO',
         'ATENDIMENTOS : DT_FIM',
         'ATENDIMENTOS : IDENT_CLI',
         'HIST. ATEND. CODIF. : ATENDIMENTO_ID',
         'HIST. ATEND. CODIF. : SERVICO_ID',
         'HIST. ATEND. CODIF. : VALOR_PESO',
         'HIST. ATEND. : ID',
         'HIST. ATEND. : UNIDADE_ID',
         'HIST. ATEND. : USUARIO_ID',
         'HIST. ATEND. : USUARIO_TRI_ID',
         'HIST. ATEND. : SERVICO_ID',
         'HIST. ATEND. : PRIORIDADE_ID',
         'HIST. ATEND. : STATUS',
         'HIST. ATEND. : SIGLA_SENHA',
         'HIST. ATEND. : NUM_SENHA',
         'HIST. ATEND. : NUM_SENHA_SERV',
         'HIST. ATEND. : NM_CLI',
         'HIST. ATEND. : NUM_LOCAL',
         'HIST. ATEND. : DT_CHEGADA',
         'HIST. ATEND. : DT_CHAMADA',
         'HIST. ATEND. : DT_INICIO',
         'HIST. ATEND. : DT_FIM',
         'HIST. ATEND. : IDENT_CLI',
         'SERVICOS : ID',
         'SERVICOS : MACRO_ID',
         'SERVICOS : DESCRICAO',
         'SERVICOS : NOME',
         'SERVICOS : STATUS',
         'SERVICOS : PESO',
         'USUARIOS : ID',
         'USUARIOS : LOGIN',
         'USUARIOS : NOME',
         'USUARIOS : SOBRENOME',
         'USUARIOS : ULT_ACESSO',
         'USUARIOS : STATUS'
         ]

   coluna = 0

   for titulo in titulo_colunas:
         planilha.write(0, coluna, titulo)
         coluna += 1

   linha = 1
   print('Exportando tabela atend_codif...\n')
   cursor.execute('SELECT * FROM atend_codif;')
   
   for row in cursor.fetchall():
         planilha.write(linha, 0, row[0])
         planilha.write(linha, 1, row[1])
         planilha.write(linha, 2, row[2])
         
         linha += 1

   linha = 1
   print('Exportando tabela atendimentos...\n')
   cursor.execute(' SELECT * FROM atendimentos;')

   for row in cursor.fetchall():
         planilha.write(linha, 3, row[0])
         planilha.write(linha, 4, row[1])
         planilha.write(linha, 5, row[2])
         planilha.write(linha, 6, row[3])
         planilha.write(linha, 7, row[4])
         planilha.write(linha, 8, row[5])
         planilha.write(linha, 9, row[6])
         planilha.write(linha, 10, row[7])
         planilha.write(linha, 11, row[8])
         planilha.write(linha, 12, row[9])
         planilha.write(linha, 13, row[10])
         planilha.write(linha, 14, row[11])
         planilha.write(linha, 15, row[12])
         if row[13] is not None : planilha.write(linha, 16, row[13].strftime('%d-%m-%Y %H:%M:%S'))
         if row[14] is not None : planilha.write(linha, 17, row[14].strftime('%d-%m-%Y %H:%M:%S'))
         if row[15] is not None : planilha.write(linha, 18, row[15].strftime('%d-%m-%Y %H:%M:%S'))
         if row[16] is not None : planilha.write(linha, 19, row[16].strftime('%d-%m-%Y %H:%M:%S'))
         planilha.write(linha, 20, row[17])
         
         linha += 1

   linha = 1
   print('Exportando tabela historico_atend_codif...\n')
   cursor.execute(' SELECT * FROM historico_atend_codif;')

   for row in cursor.fetchall():
         planilha.write(linha, 21, row[0])
         planilha.write(linha, 22, row[1])
         planilha.write(linha, 23, row[2])
         
         linha += 1

   linha = 1
   print('Exportando tabela histico_atendimentos...\n')
   cursor.execute(' SELECT * FROM historico_atendimentos;')

   for row in cursor.fetchall():
         planilha.write(linha, 24, row[0])
         planilha.write(linha, 25, row[1])
         planilha.write(linha, 26, row[2])
         planilha.write(linha, 27, row[3])
         planilha.write(linha, 28, row[4])
         planilha.write(linha, 29, row[5])
         planilha.write(linha, 30, row[6])
         planilha.write(linha, 31, row[7])
         planilha.write(linha, 32, row[8])
         planilha.write(linha, 33, row[9])
         planilha.write(linha, 34, row[10])
         planilha.write(linha, 35, row[11])
         if row[12] is not None : planilha.write(linha, 36, row[12].strftime('%d-%m-%Y %H:%M:%S'))
         if row[13] is not None : planilha.write(linha, 37, row[13].strftime('%d-%m-%Y %H:%M:%S'))
         if row[14] is not None : planilha.write(linha, 38, row[14].strftime('%d-%m-%Y %H:%M:%S'))
         if row[15] is not None : planilha.write(linha, 39, row[15].strftime('%d-%m-%Y %H:%M:%S'))
         planilha.write(linha, 40, row[16])
         
         linha += 1

   linha = 1
   print('Exportando tabela servicos...\n')
   cursor.execute(' SELECT * FROM servicos;')

   for row in cursor.fetchall():
         planilha.write(linha, 41, row[0])
         planilha.write(linha, 42, row[1])
         planilha.write(linha, 43, row[2])
         planilha.write(linha, 44, row[3])
         planilha.write(linha, 45, row[4])
         planilha.write(linha, 46, row[5])
         
         linha += 1

   linha = 1
   print('Exportando tabela usuarios...\n')
   cursor.execute(' SELECT * FROM usuarios;')

   for row in cursor.fetchall():
         planilha.write(linha, 47, row[0])
         planilha.write(linha, 48, row[1])
         planilha.write(linha, 49, row[2])
         planilha.write(linha, 50, row[3])
         if row[5] is not None : planilha.write(linha, 51, row[5].strftime('%d-%m-%Y %H:%M:%S'))
         planilha.write(linha, 52, row[6])
         
         linha += 1
   
   print('Finalizando o arquivo...')
   arquivo_excel.close()

   print('Arquivo gerado com sucesso! Fim da execução...')