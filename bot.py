#!usr/bin/env python3

__version__ = "0.1.2"
__author__ = "r0bert,   ,   "
__license__ = "Unlicense"

# %%
# import bibliotecas
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import pillow
#%%
webbrowser.open("https://web.whatsapp.com")
sleep(5)

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


# Lendo os números de telefone da tabela e aplicando a função corrigir_numero
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = str(linha[1].value)  # Convertendo telefone para string
    vencimento = linha[2].value

    telefone_corrigido = corrigir_numero(telefone)
    mensagem = f"""Olá {nome}, seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Por favor, pagar no link
    https://www.link_do_pagamento.com """
    # print(nome)
    # print(telefone_corrigido)
    # print(vencimento)


try:
        link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone_corrigido}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_wpp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.jpeg')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w') # pra fechar a aba

except: 
    print("Não foi possível enviar mensagem para {nome}")
    with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo_erros:
        arquivo_erros.write(f"{nome};{telefone}\n")