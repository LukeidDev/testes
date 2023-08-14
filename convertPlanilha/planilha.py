

workbook = op.load_workbook('teste.xlsx')
nome_planilha = 'Copper'
planilha = workbook[nome_planilha]

row_teste = (coord_row(nome_planilha, "Average"))

min_row = 9
max_row = row_teste
min_col = 2
max_col = 4



valores_obtidos = obter_valores_da_planilha(planilha, min_row, max_row, min_col, max_col)

soma_valores = [0, 0, 0, 0]  # Inicializar soma dos valores
for valor in valores_obtidos:
    data, val2_dolar, val2_real, val3_dolar, val3_real = valor
    soma_valores[0] += val2_dolar
    soma_valores[1] += val2_real
    soma_valores[2] += val3_dolar
    soma_valores[3] += val3_real
    
num_valores = len(valores_obtidos)
media_valores = [soma / num_valores for soma in soma_valores]

nova_aba = workbook.create_sheet(title='Valores Convertidos')
nova_aba.append(['Sheet', 'Data', 'Valor 2 em Dólar', 'Valor 2 em Reais', 'Valor 3 em Dólar', 'Valor 3 em Reais'])  # Cabeçalho

for valor in valores_obtidos:
    data, val2_dolar, val2_real, val3_dolar, val3_real = valor
    nova_aba.append([nome_planilha, data, f"USD {val2_dolar:,.2f}", f"R$ {val2_real:,.2f}", f"USD {val3_dolar:,.2f}", f"R$ {val3_real:,.2f}"])

nova_aba.append(['Média', '', f"USD {media_valores[0]:,.2f}", f"R$ {media_valores[1]:,.2f}", f"USD {media_valores[2]:,.2f}", f"R$ {media_valores[3]:,.2f}"])

for col in nova_aba.columns:
    max_length = 0
    column = col[0].column_letter 
    for cell in col:
        try: 
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2) 
    nova_aba.column_dimensions[column].width = adjusted_width
    
workbook.save('teste.xlsx')
