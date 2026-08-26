import openpyxl

# 1. Carregar as planilhas
wb_origem = openpyxl.load_workbook("Cartão Ponto fechamento 26-07 a 25-08_3.xlsx", data_only=True)
wb_destino = openpyxl.load_workbook("PONTO_S.DIOGO_PREENCHIDO (3).xlsx")

ws_origem = wb_origem.active
ws_destino = wb_destino.active

# 2. Mapear as colunas de dias na linha 2 da planilha de DESTINO
# Transforma os cabeçalhos em inteiros para evitar falha de comparação (ex: "30" vs 30)
colunas_dias_destino = {}
for col in range(1, ws_destino.max_column + 1):
    val = ws_destino.cell(row=2, column=col).value
    if val is not None:
        try:
            # Extrai apenas o número do dia
            dia_num = int(str(val).strip())
            colunas_dias_destino[dia_num] = col
        except ValueError:
            pass

# 3. Mapear as linhas dos colaboradores por Matrícula (Coluna B) no DESTINO
linhas_matriculas_destino = {}
for row in range(3, ws_destino.max_row + 1):
    mat = ws_destino.cell(row=row, column=2).value
    if mat is not None:
        mat_limpa = str(mat).strip()
        linhas_matriculas_destino[mat_limpa] = row

# 4. Processar a planilha de ORIGEM
linha_atual_origem = 1
max_rows = ws_origem.max_row

while linha_atual_origem <= max_rows:
    col_a = ws_origem.cell(row=linha_atual_origem, column=1).value
    
    # Identificar quando encontra uma nova matrícula na Coluna A ou B
    # Ajuste o texto de busca conforme a estrutura da sua origem:
    if col_a and "Matrícula" in str(col_a):
        # Lê a matrícula da origem (ex: pegando o valor ao lado ou na mesma célula)
        matricula_origem = str(ws_origem.cell(row=linha_atual_origem, column=2).value).strip()
        
        # Verifica se essa matrícula existe no destino
        if matricula_origem in linhas_matriculas_destino:
            linha_dest = linhas_matriculas_destino[matricula_origem]
            
            # Percorre os dias do colaborador na origem até encontrar o final do cartão dele
            for r in range(linha_atual_origem, linha_atual_origem + 35):
                data_val = ws_origem.cell(row=r, column=1).value
                
                # Se for uma data válida (ex: 30/07/2026)
                if data_val:
                    try:
                        # Trata a data seja ela texto ou objeto datetime
                        if hasattr(data_val, 'day'):
                            dia = data_val.day
                        else:
                            dia = int(str(data_val).split('/')[0])
                        
                        # Se o dia existe no cabeçalho de destino
                        if dia in colunas_dias_destino:
                            col_dest = colunas_dias_destino[dia]
                            
                            # Pega o valor da coluna X (coluna 24 no Excel)
                            valor_x = ws_origem.cell(row=r, column=24).value
                            
                            if valor_x is not None:
                                ws_destino.cell(row=linha_dest, column=col_dest).value = valor_x
                    except (ValueError, IndexError):
                        pass

                # Identificar a linha de TOTAIS da origem
                # Exemplo: se a palavra "Total" estiver na Coluna A/B dessa seção
                if ws_origem.cell(row=r, column=1).value == "Total":
                    # Mapeie aqui as colunas exatas dos totais da origem para o destino
                    # Ex: copiar da coluna X (24) da origem para a coluna Z (26) do destino
                    ws_destino.cell(row=linha_dest, column=35).value = ws_origem.cell(row=r, column=24).value
                    break

    linha_atual_origem += 1

# 5. Salvar o arquivo processado
wb_destino.save("PONTO_S.DIOGO_PREENCHIDO_CORRIGIDO.xlsx")
