#!usr/bin/env python3

__version__ = "0.1.1"
__author__ = "r0bert"
__license__ = "Unlicense"

# %%
# import bibliotecas
import openpyxl

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
    data = linha[2].value

    telefone_corrigido = corrigir_numero(telefone)

    print(nome)
    print(telefone_corrigido)
    print(data)

# %%
