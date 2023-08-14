import openpyxl as op
from metodos.get_row import coord_row
from metodos.get_values import obter_valores_convertidos

def input_values(nome_planilha, key, taxas_conversao, nova_aba):
    workbook = op.load_workbook('teste.xlsx')
    planilha = workbook[nome_planilha]
    row_teste = coord_row(nome_planilha, key)
    min_row = 9
    max_row = row_teste
    min_col = 2
    max_col = 4

    valores_obtidos = obter_valores_convertidos(planilha, min_row, max_row, min_col, max_col, taxas_conversao)

    soma_valores = [0, 0, 0, 0, 0, 0]
    num_valores = len(valores_obtidos)

    for valor in valores_obtidos:
        soma_valores = [soma + val for soma, val in zip(soma_valores, valor[1:])]

    media_valores = [soma / num_valores for soma in soma_valores]

    nova_aba.append([nome_planilha, '', 'USD (BUYER)', 'R$ (BUYER)', 'USD (SELLER)', 'R$ (SELLER)'])

    for valor in valores_obtidos:
        data, val2_dolar, val2_real, val3_dolar, val3_real = valor
        nova_aba.append([data, f"USD {val2_dolar:,.2f}", f"R$ {val2_real:,.2f}",
                         f"USD {val3_dolar:,.2f}", f"R$ {val3_real:,.2f}"])

    nova_aba.append(['Média', f"USD {media_valores[0]:,.2f}", f"R$ {media_valores[1]:,.2f}",
                     f"USD {media_valores[2]:,.2f}", f"R$ {media_valores[3]:,.2f}"])

    return nova_aba