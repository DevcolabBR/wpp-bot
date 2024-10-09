#!usr/bin/env python3

__version__ = "0.1.4"
__author__ = "r0bert,   ,   "
__license__ = "Unlicense"

# %%
# import bibliotecas
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from datetime import datetime

#%%
webbrowser.open("https://web.whatsapp.com")
sleep(18)

# importando tabelas
workbook = openpyxl.load_workbook("tabela.xlsx")
pagina_clientes = workbook["Sheet1"]  # nome da página no libreoficce

# %%
def corrigir_numero(numero):
    numero = "".join(filter(str.isdigit, numero))

    if len(numero) < 8:
        return None

    if len(numero) == 8:
        return f"55919{numero}"

    if len(numero) == 10 and not numero.startswith("55"):
        return f"55{numero[:2]}9{numero[2:]}"

    if len(numero) == 11 and numero.startswith("55"):
        return f"{numero[:4]}9{numero[4:]}"

    if len(numero) == 12 and numero.startswith("55"):
        if numero[4] == "9" and len(numero[5:]) == 8:
            return numero
        else:
            return f"{numero[:4]}9{numero[4:]}"

    return numero

# Definindo um valor padrão de vencimento
vencimento_padrao = datetime.strptime('01/01/2024', '%d/%m/%Y')

# Lendo os números de telefone da tabela e aplicando a função corrigir_numero
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = str(linha[1].value)  # Convertendo telefone para string
    telefone_corrigido = corrigir_numero(telefone)

    # Usando o valor padrão de vencimento
    vencimento = vencimento_padrao

    mensagem = f"""Olá {nome}, seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Por favor, pagar no link
    https://www.link_do_pagamento.com """
    
    try:
        link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone_corrigido}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_wpp)
        sleep(15)
        pyautogui.press('enter')
        #seta = pyautogui.locateCenterOnScreen('seta.jpeg')
        sleep(3)
        #pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w') # pra fechar a aba

    except Exception as e:
        print(f"Não foi possível enviar mensagem para {nome}: {e}")
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo_erros:
            arquivo_erros.write(f"{nome};{telefone}\n")
