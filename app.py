import streamlit as st
import pandas as pd
import openpyxl
import re
import io

# Configuração da página
st.set_page_config(page_title="Automação de Ponto - RH", layout="wide")

st.title("⏱️ Processador de Cartão Ponto")
st.markdown("""
### Instruções de uso:
1. Carregue o arquivo de **Origem** (`Cartão Ponto fechamento...`).
2. Carregue o arquivo modelo de **Destino em Branco** (`PONTO S.DIOGO...`).
3. Clique em **Processar e Preencher Ponto** para gerar a planilha preenchida mantendo a formatação original.
""")

col1, col2 = st.columns(2)
with col1:
    file_origem = st.file_uploader("1. Relatório de Origem (Cartão Ponto)", type=['xlsx'])
with col2:
    file_destino = st.file_uploader("2. Modelo Destino em Branco (Ponto RH)", type=['xlsx'])

def extrair_dados_origem(file):
    """Extrai os dados da planilha de origem estruturada por colaboradores."""
    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb.active

    colaboradores = {}
    colab_atual = None

    for row in range(1, sheet.max_row + 1):
        col_a = str(sheet.cell(row=row, column=1).value or "").strip()
        
        # 1. Identifica início de novo colaborador
        if "CONTRACTOR ENGENHARIA LTDA" in col_a:
            col_a_nome = str(sheet.cell(row=row+1, column=1).value or "").strip()
            col_m_mat = str(sheet.cell(row=row+1, column=13).value or "").strip()
            
            # Limpa matrícula de caracteres desnecessários
            matricula = "".join(filter(str.isdigit, col_m_mat))
            
            colab_atual = {
                'nome': col_a_nome,
                'matricula': matricula,
                'dias': {},
                'totais': {'faltas': 0, 'he50': 0, 'he70': 0, 'he120': 0, 'atrasos': 0}
            }
            chave = (col_a_nome, matricula)
            colaboradores[chave] = colab_atual
            continue

        if not colab_atual:
            continue

        # 2. Captura linha de TOTAIS
        if col_a == "TOTAIS":
            colab_atual['totais']['he50'] = sheet.cell(row=row, column=24).value or 0   # Coluna X
            colab_atual['totais']['he70'] = sheet.cell(row=row, column=27).value or 0   # Coluna AA
            colab_atual['totais']['he120'] = sheet.cell(row=row, column=31).value or 0  # Coluna AE
            colab_atual['totais']['atrasos'] = sheet.cell(row=row, column=34).value or 0# Coluna AH
            colab_atual['totais']['faltas'] = sheet.cell(row=row, column=36).value or 0 # Coluna AJ
            continue

        # 3. Captura dias individuais (Formato: DD/MM/YYYY - Dia)
        match_data = re.search(r'^(\d{2})/\d{2}/\d{4}', col_a)
        if match_data:
            dia = int(match_data.group(1))
            he50 = sheet.cell(row=row, column=24).value or 0   # Coluna X
            he70 = sheet.cell(row=row, column=27).value or 0   # Coluna AA
            he120 = sheet.cell(row=row, column=31).value or 0  # Coluna AE
            faltas = sheet.cell(row=row, column=36).value or 0 # Coluna AJ
            
            colab_atual['dias'][dia] = {
                'he50': he50,
                'he70': he70,
                'he120': he120,
                'faltas': faltas
            }

    return colaboradores

def preencher_destino(file_dest, dados_colab):
    """Preenche a planilha modelo em branco mantendo todos os estilos."""
    wb = openpyxl.load_workbook(file_dest)
    sheet = wb.active

    # 1. Mapeia cabeçalho de dias (Linha 2, da coluna H em diante)
    mapa_dias_col = {}
    for col in range(8, sheet.max_column + 1):
        val = sheet.cell(row=2, column=col).value
        if val is not None and str(val).isdigit():
            mapa_dias_col[int(val)] = col

    # 2. Preenche dados a partir da linha 3
    for row in range(3, sheet.max_row + 1):
        nome = str(sheet.cell(row=row, column=1).value or "").strip()
        mat_raw = str(sheet.cell(row=row, column=2).value or "").strip()
        matricula = "".join(filter(str.isdigit, mat_raw))

        if not nome and not matricula:
            continue

        # Busca dados por Nome + Matrícula ou apenas Matrícula (segunda camada)
        colab_data = None
        for (c_nome, c_mat), data in dados_colab.items():
            if (nome and c_nome == nome) or (matricula and c_mat == matricula):
                colab_data = data
                break

        if colab_data:
            # Lança Totais
            sheet.cell(row=row, column=4).value = colab_data['totais']['faltas']  # Coluna D
            sheet.cell(row=row, column=5).value = colab_data['totais']['he50']    # Coluna E
            sheet.cell(row=row, column=6).value = colab_data['totais']['he70']    # Coluna F
            sheet.cell(row=row, column=7).value = colab_data['totais']['he120']   # Coluna G
            sheet.cell(row=row, column=39).value = colab_data['totais']['atrasos']# Coluna AM (Índice 39)

            # Lança Dias
            for dia, info in colab_data['dias'].items():
                if dia in mapa_dias_col:
                    col_idx = mapa_dias_col[dia]
                    
                    # Preenche com valor de hora extra se existir
                    val_he = info['he50'] or info['he70'] or info['he120']
                    if val_he:
                        sheet.cell(row=row, column=col_idx).value = val_he

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- PROCESSAMENTO ---

if st.button("🚀 Processar e Preencher Ponto"):
    if not file_origem or not file_destino:
        st.error("⚠️ Por favor, envie ambos os arquivos antes de continuar.")
        st.stop()

    try:
        dados_extraidos = extrair_dados_origem(file_origem)
        excel_processado = preencher_destino(file_destino, dados_extraidos)

        st.success("✅ Ponto processado com sucesso!")
        st.download_button(
            label="📥 Baixar Ponto Preenchido",
            data=excel_processado,
            file_name="PONTO_S.DIOGO_PREENCHIDO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Ocorreu um erro ao processar: {e}")
