#Enviando email com dados de vendas para seus respectivos vendedores de forma dinâmica
#pip install pywin32 <- biblioteca necessária
#https://www.udemy.com/course/python-rpa-e-excel-aprenda-automatizar-processos-e-planilhas/learn/lecture/27930634#overview

from openpyxl import load_workbook
import os
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
emailOutlook = outlook.CreateItem(0)

#Abrir o arquivo em segundo plano
nome_arquivo = "C:\\Users\\Windows\\Desktop\\Python Projetos\\openpyxl\\ExcelEmail\ListaEmail.xlsx"
planilha_aberta = load_workbook(filename=nome_arquivo)

#Indicando qual Sheet será lida.
sheet_selecionada = planilha_aberta['Dados']

for linha in range(2, len(sheet_selecionada['A']) + 1): #len(sheet_selecionada['A']) + 1): Deixa DINÂMICO o script.

    nome = sheet_selecionada['A%s' % linha].value #Vai passar linha a linha e pegar os valores
    nomeCompleto = sheet_selecionada['B%s' % linha].value  # Vai passar linha a linha e pegar os valores
    email = sheet_selecionada['C%s' % linha].value  # Vai passar linha a linha e pegar os valores

#Trabalhando a parte do envio do email:
emailOutlook.To = email
emailOutlook.Subject = "Lista de Vendas " + nomeCompleto
emailOutlook.HTMLBody = f"""
<p>Boa noite <b> {nome}</b>.</p>
<p>Segue o relatório com as suas vendas</p>

<p>Atenciosamente; Carlos Henrique de Camargo</p>
"""

#Indicando o caminho do email e os arquivos que serão enviados para cada vendedor.
anexoEmail = "C:\\Users\\Windows\\Desktop\\Python Projetos\\openpyxl\\ExcelEmail\\" + nomeCompleto + ".xlsx"
emailOutlook.Attachments.Add(anexoEmail)

emailOutlook.save() #save = Cria e salva o email, Send() - Enviar o email