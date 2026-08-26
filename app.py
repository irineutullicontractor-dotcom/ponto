import streamlit as st
import pandas as pd
import openpyxl
import io
import re

st.set_page_config(page_title="Processador de Ponto - S. DIOGO", layout="wide")

st.title("⏱️ Processador de Espelho de Ponto")
st.markdown("""
### Instruções de uso:
1. Faça o upload do **Cartão Ponto em Aberto/Bruto** (Relatório vindo do sistema automatizado).
2. Faça o upload do **Modelo Manual em Branco (PONTO S.DIOGO)**.
3. Clique em **Processar e Preencher Ponto** para gerar e baixar a planilha tratada.
""")

col1, col2 = st.columns(2)
with col1:
    file_origem = st.file_uploader("1. Cartão Ponto Fechamento (Sistema Automatizado)", type=['xlsx', 'xls'])
with col2:
    file_destino = st.file_uploader("2. PONTO S.DIOGO (Modelo Manual em Branco)", type=['xlsx', 'xls'])

def extrair_apenas_numeros(v):
    if pd.isna(v): return ""
    return "".join(filter(str.isdigit, str(v)))

def converter_horas(val):
    if pd.isna(val) or val is None or str(val).strip() in ["", "-", "00:00", "0"]:
        return 0.0
    s_val = str(val).strip()
    # Se já for número
    try:
        return float(s_val)
    except ValueError:
        pass
    
    # Formato HH:MM
    if ":" in s_val:
        partes = s_val.split(":")
        try:
            h = int(partes[0])
            m = int(partes[1])
            return h + (m / 60.0)
        except:
            return 0.0
    return 0.0

def processar_espelho_ponto(file_orig, file_dest):
    # 1. Leitura bruta dos dados do sistema
    df_raw = pd.read_excel(file_orig, header=None)
    
    col_a = df_raw.iloc[:, 0].astype(str)
    
    # Identifica onde começam os registros de cada colaborador
    indices_inicio = col_a[col_a.str.contains("CONTRACTOR ENGENHARIA LTDA", case=False, na=False)].index.tolist()
    
    dados_colaboradores = {}

    for i in range(len(indices_inicio)):
        idx_st = indices_inicio[i]
        idx_end = indices_inicio[i+1] if i + 1 < len(indices_inicio) else len(df_raw)
        
        block = df_raw.iloc[idx_st:idx_end]
        
        # Leitura da Matrícula (Coluna M = Índice 12) e Nome (Coluna A = Índice 0)
        matricula_raw = ""
        nome_raw = ""
        
        for idx_row, row in block.iterrows():
            txt_a = str(row[0]).strip()
            
            # Identifica o Nome (Linha que não é cabeçalho padrão nem data)
            if txt_a and not any(k in txt_a for k in ["CONTRACTOR", "Data", "TOTAIS", "Emissão", "Página"]):
                if not re.match(r'^\d{2}/\d{2}/\d{4}', txt_a) and len(txt_a) > 3:
                    nome_raw = txt_a
            
            # Identifica a Matrícula
            val_m = str(row[12]).strip() if len(row) > 12 else ""
            if val_m and val_m.lower() != "nan" and val_m.isdigit():
                matricula_raw = val_m

        matricula_clean = extrair_apenas_numeros(matricula_raw)
        
        if not matricula_clean:
            continue
            
        # Extração de Horas por Dia e Totais
        dias_dict = {}
        totais = {"50%": 0.0, "70%": 0.0, "120%": 0.0, "atrasos": 0.0, "faltas": 0.0}
        
        for idx_row, row in block.iterrows():
            txt_a = str(row[0]).strip()
            
            # Registros Diários (ex: 26/07/2026 - Dom)
            if re.match(r'^\d{2}/\d{2}/\d{4}', txt_a):
                dia_num = int(txt_a.split('/')[0]) # Extrai o dia (26, 27, ...)
                
                he_50 = converter_horas(row[23]) if len(row) > 23 else 0.0   # Coluna X
                he_70 = converter_horas(row[26]) if len(row) > 26 else 0.0   # Coluna AA
                he_120 = converter_horas(row[30]) if len(row) > 30 else 0.0  # Coluna AE
                faltas = converter_horas(row[35]) if len(row) > 35 else 0.0  # Coluna AJ
                
                # Guarda as HE do dia (se houver mais de um tipo no dia, pega a maior/vigente)
                val_dia = max(he_50, he_70, he_120)
                if val_dia > 0:
                    dias_dict[dia_num] = val_dia
            
            # Linha de Totais no Fim do Bloco
            elif "TOTAIS" in txt_a.upper():
                totais["50%"] = converter_horas(row[23]) if len(row) > 23 else 0.0     # Coluna X
                totais["70%"] = converter_horas(row[26]) if len(row) > 26 else 0.0     # Coluna AA
                totais["120%"] = converter_horas(row[30]) if len(row) > 30 else 0.0    # Coluna AE
                totais["atrasos"] = converter_horas(row[33]) if len(row) > 33 else 0.0 # Coluna AH
                totais["faltas"] = converter_horas(row[35]) if len(row) > 35 else 0.0  # Coluna AJ

        dados_colaboradores[matricula_clean] = {
            "nome": nome_raw,
            "totais": totais,
            "dias": dias_dict
        }

    # 2. Preenchimento do Modelo Manual com OpenPyXL
    wb = openpyxl.load_workbook(file_dest)
    ws = wb.active
    
    # Mapear os dias das colunas a partir da Linha 2 (Coluna H = 8 em diante)
    mapa_colunas_dias = {}
    for col in range(8, ws.max_column + 1):
        val_cabecalho = ws.cell(row=2, column=col).value
        if val_cabecalho is not None:
            dia_limpo = extrair_apenas_numeros(val_cabecalho)
            if dia_limpo:
                mapa_colunas_dias[int(dia_limpo)] = col

    # Mapear a Coluna AM (Total de Atrasos)
    col_atraso = 39 # Coluna AM
    
    # Preencher a partir da Linha 3 (onde começam os colaboradores)
    for row in range(3, ws.max_row + 1):
        val_mat = ws.cell(row=row, column=2).value # Coluna B = Matrícula
        mat_dest = extrair_apenas_numeros(val_mat)
        
        if mat_dest in dados_colaboradores:
            info = dados_colaboradores[mat_dest]
            
            # Atualiza o Nome se a coluna B bater a Matrícula
            if info["nome"]:
                ws.cell(row=row, column=1).value = info["nome"] # Coluna A
            
            # Lança os Totais nas Colunas D, E, F, G
            ws.cell(row=row, column=4).value = info["totais"]["faltas"]  # Coluna D: Faltas
            ws.cell(row=row, column=5).value = info["totais"]["50%"]     # Coluna E: HE 50%
            ws.cell(row=row, column=6).value = info["totais"]["70%"]     # Coluna F: HE 70%
            ws.cell(row=row, column=7).value = info["totais"]["120%"]    # Coluna G: HE 120%
            
            # Lança o Total de Atrasos na Coluna AM
            ws.cell(row=row, column=col_atraso).value = info["totais"]["atrasos"]
            
            # Lança os Valores Diários nos dias correspondentes
            for dia_num, col_idx in mapa_colunas_dias.items():
                if dia_num in info["dias"]:
                    ws.cell(row=row, column=col_idx).value = info["dias"][dia_num]

    # Salva o arquivo preenchido em buffer de memória
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Botão para processar
if st.button("🚀 Processar e Preencher Ponto"):
    if file_origem and file_destino:
        try:
            excel_processado = processar_espelho_ponto(file_origem, file_destino)
            st.success("✅ Planilha preenchida com sucesso!")
            st.download_button(
                label="📥 Baixar PONTO S.DIOGO Preenchido",
                data=excel_processado,
                file_name="PONTO_S.DIOGO_PREENCHIDO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
    else:
        st.warning("⚠️ Por favor, faça o upload de ambos os arquivos para prosseguir.")
