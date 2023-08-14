from datetime import datetime
from metodos.get_convert_bcb import obter_taxas_de_conversao


def obter_valores_convertidos(planilha, min_row, max_row, min_col, max_col, taxas_conversao):
    valores_convertidos = []

    for row in planilha.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        col1_value = row[0].value
        col2_value = row[1].value
        col3_value = row[2].value

        if isinstance(col1_value, datetime):
            data_str = col1_value.date().strftime('%d/%m/%Y')
            data = data_str
        else:
            data_str = col1_value
            data = col1_value

        taxa_conversao = taxas_conversao.get(data_str)

        if taxa_conversao is not None:
            col2_value_reais = col2_value * taxa_conversao
            col3_value_reais = col3_value * taxa_conversao
            valores_convertidos.append((data, col2_value, col2_value_reais, col3_value, col3_value_reais))
        else:
            valores_convertidos.append((data, col2_value, col2_value, col3_value, col3_value))

    return valores_convertidos