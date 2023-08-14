from bot import new_caminho
import openpyxl as op

def coord_row(a, b):
    workbook = op.load_workbook(new_caminho)
    nome_planilha = a
    planilha = workbook[nome_planilha]
    palavra_chave = b

    row_teste = None

    for row in planilha.iter_rows():
        for cell in row:
            if palavra_chave in str(cell.value):
                row_teste = cell.row - 1
                break

    workbook.close()

    return row_teste