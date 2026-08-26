import io
import re
import unicodedata
from datetime import datetime, timedelta

import openpyxl
import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(
    page_title="Preenchimento de Ponto - S. Diogo",
    page_icon="🕘",
    layout="wide",
)

st.title("🕘 Preenchimento automático do Ponto - S. Diogo")

st.markdown("""
### Como funciona
1. Carregue o **Cartão Ponto** exportado do sistema.
2. Carregue a **planilha PONTO S.DIOGO em branco**.
3. O sistema identifica os colaboradores por **nome e matrícula**.
4. Transfere totais, faltas (F), atestados (AT) e horas extras para os respectivos dias.
5. Mantém a planilha modelo, incluindo fórmulas, formatação e anotações já existentes.
6. Antes do download, apresenta qualquer divergência encontrada.
""")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def normalizar_nome(valor):
    """Normaliza nome para comparação sem acentos, espaços ou pontuação."""
    if valor is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    texto = texto.upper()
    return re.sub(r"[^A-Z0-9]", "", texto)


def limpar_matricula(valor):
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return re.sub(r"\D", "", texto)


def converter_hora(valor):
    """
    Converte valores do tipo HH:MM / HH:MM:SS ou timedelta
    para timedelta, permitindo que as fórmulas SUM do Excel funcionem.
    """
    if valor is None:
        return None

    if isinstance(valor, timedelta):
        return valor

    if isinstance(valor, datetime):
        return timedelta(
            hours=valor.hour,
            minutes=valor.minute,
            seconds=valor.second,
        )

    texto = str(valor).strip()

    if not texto or ":" not in texto:
        return None

    partes = texto.split(":")

    try:
        horas = int(partes[0])
        minutos = int(partes[1])
        segundos = int(partes[2]) if len(partes) > 2 else 0
        return timedelta(hours=horas, minutes=minutos, seconds=segundos)
    except (ValueError, TypeError):
        return None


def extrair_dia(valor):
    """
    Lê datas no padrão:
    26/07/2026 - Dom
    """
    if isinstance(valor, datetime):
        return valor.day

    texto = str(valor).strip() if valor is not None else ""

    resultado = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})", texto)

    if not resultado:
        return None

    return int(resultado.group(1))


def encontrar_linha_totais(ws, linha_inicio):
    for linha in range(linha_inicio + 1, ws.max_row + 1):
        valor = ws.cell(linha, 1).value

        if str(valor).strip().upper() == "TOTAIS":
            return linha

    return None


def verificar_status_dia(val_e, val_p):
    """
    Verifica se ambas as colunas (E e P) possuem 'FALTA' ou 'ATESTAD'.
    Retorna 'F' para falta, 'AT' para atestado ou None caso não atenda.
    """
    txt_e = str(val_e).strip().upper() if val_e is not None else ""
    txt_p = str(val_p).strip().upper() if val_p is not None else ""

    # Verifica se ambas são FALTA
    if "FALTA" in txt_e and "FALTA" in txt_p:
        return "F"

    # Verifica se ambas são ATESTAD (ou ATESTADO)
    if "ATESTAD" in txt_e and "ATESTAD" in txt_p:
        return "AT"

    return None


# ============================================================
# LEITURA DO CARTÃO PONTO
# ============================================================
def ler_cartao_ponto(arquivo):
    wb = openpyxl.load_workbook(arquivo, data_only=False)
    ws = wb[wb.sheetnames[0]]

    colaboradores = []
    linhas_inicio = []

    for linha in range(1, ws.max_row + 1):
        valor = ws.cell(linha, 1).value

        if str(valor).strip().upper() == "CONTRACTOR ENGENHARIA LTDA":
            linhas_inicio.append(linha)

    for linha_inicio in linhas_inicio:
        linha_nome = None
        linha_matricula = None
        linha_funcao = None

        # Localiza os campos do cabeçalho do colaborador.
        limite = min(linha_inicio + 35, ws.max_row)

        for linha in range(linha_inicio, limite + 1):
            valor_a = str(ws.cell(linha, 1).value).strip().upper()
            valor_m = str(ws.cell(linha, 13).value).strip().upper()

            if valor_a == "NOME:":
                linha_nome = linha + 2

            if valor_m == "Nº FOLHA:":
                linha_matricula = linha + 2

            if valor_a == "FUNÇÃO:":
                linha_funcao = linha + 2

        linha_totais = encontrar_linha_totais(ws, linha_inicio)

        if not linha_totais:
            continue

        nome = ws.cell(linha_nome, 1).value if linha_nome else None
        matricula = ws.cell(linha_matricula, 13).value if linha_matricula else None
        funcao = ws.cell(linha_funcao, 1).value if linha_funcao else None

        nome = str(nome).strip() if nome is not None else ""
        matricula = limpar_matricula(matricula)
        funcao = str(funcao).strip() if funcao is not None else ""

        if not nome or not matricula:
            continue

        # Dados diários.
        diarios = {}

        for linha in range(linha_inicio + 1, linha_totais):
            dia = extrair_dia(ws.cell(linha, 1).value)

            if dia is None:
                continue

            # Leitura das colunas E (coluna 5) e P (coluna 16) para FALTA / ATESTAD
            val_e = ws.cell(linha, 5).value
            val_p = ws.cell(linha, 16).value
            status_dia = verificar_status_dia(val_e, val_p)

            diarios[dia] = {
                "status_dia": status_dia,                        # 'F', 'AT' ou None
                "50": converter_hora(ws.cell(linha, 24).value),   # X
                "70": converter_hora(ws.cell(linha, 27).value),   # AA
                "120": converter_hora(ws.cell(linha, 31).value),  # AE
            }

        # Totais mensais.
        totais = {
            "50": converter_hora(ws.cell(linha_totais, 24).value),   # X
            "70": converter_hora(ws.cell(linha_totais, 27).value),   # AA
            "120": converter_hora(ws.cell(linha_totais, 31).value),  # AE
            "atrasos": converter_hora(ws.cell(linha_totais, 34).value), # AH
            "faltas": converter_hora(ws.cell(linha_totais, 36).value),  # AJ
        }

        colaboradores.append({
            "nome": nome,
            "matricula": matricula,
            "funcao": funcao,
            "nome_normalizado": normalizar_nome(nome),
            "diarios": diarios,
            "totais": totais,
        })

    return colaboradores


# ============================================================
# LEITURA DA PLANILHA DESTINO
# ============================================================
def localizar_colaboradores_destino(ws):
    """
    Retorna somente linhas de colaboradores.
    Ignora cabeçalhos, linhas em branco e a linha NOME/Nº existente
    no meio da planilha.
    """
    resultado = []

    for linha in range(3, ws.max_row + 1):
        nome = ws.cell(linha, 1).value
        matricula = ws.cell(linha, 2).value

        if not nome or not limpar_matricula(matricula):
            continue

        resultado.append({
            "linha": linha,
            "nome": str(nome).strip(),
            "matricula": limpar_matricula(matricula),
            "nome_normalizado": normalizar_nome(nome),
        })

    return resultado


def construir_mapa_origem(colaboradores):
    por_matricula = {}
    por_nome = {}

    for colaborador in colaboradores:
        por_matricula[colaborador["matricula"]] = colaborador
        por_nome.setdefault(colaborador["nome_normalizado"], []).append(colaborador)

    return por_matricula, por_nome


def localizar_colaborador(destino, por_matricula, por_nome):
    """
    Regra de cruzamento:
    1. Matrícula exata + nome compatível.
    2. Matrícula exata, mesmo que o nome esteja abreviado/divergente.
    3. Nome exato quando não houver matrícula correspondente.
    4. Se nome e matrícula existirem mas forem conflitantes, sinaliza.
    """
    matricula = destino["matricula"]
    nome_normalizado = destino["nome_normalizado"]

    candidato_matricula = por_matricula.get(matricula)
    candidatos_nome = por_nome.get(nome_normalizado, [])

    if candidato_matricula:
        if candidato_matricula["nome_normalizado"] == nome_normalizado:
            return candidato_matricula, "OK - matrícula e nome conferem"

        return candidato_matricula, "ATENÇÃO - matrícula confere, nome diverge"

    if len(candidatos_nome) == 1:
        candidato_nome = candidatos_nome[0]

        return (
            candidato_nome,
            f"ATENÇÃO - nome confere, matrícula diverge "
            f"(modelo: {matricula} / origem: {candidato_nome['matricula']})",
        )

    if len(candidatos_nome) > 1:
        return None, "ERRO - nome duplicado na origem"

    return None, "NÃO ENCONTRADO na origem"


# ============================================================
# PREENCHIMENTO
# ============================================================
def preencher_planilha(arquivo_destino, colaboradores):
    wb = openpyxl.load_workbook(arquivo_destino)
    ws = wb[wb.sheetnames[0]]

    por_matricula, por_nome = construir_mapa_origem(colaboradores)

    linhas_destino = localizar_colaboradores_destino(ws)

    # Datas do modelo: H:AL.
    colunas_por_dia = {}

    for coluna in range(8, 39):  # H até AL
        valor = ws.cell(2, coluna).value

        if isinstance(valor, datetime):
            colunas_por_dia[valor.day] = coluna
        elif isinstance(valor, int):
            # Modelo também possui o número do dia na linha 2.
            colunas_por_dia[valor] = coluna

    # Estilo para AM, já que a coluna está reservada no modelo
    # mas não possui estilo/cabeçalho.
    if ws["AM3"].style_id == 0:
        ws["AM3"]._style = ws["D3"]._style
        ws["AM3"].number_format = "[h]:mm:ss;@"

    resultados = []
    preenchidos = set()

    for destino in linhas_destino:
        colaborador, status = localizar_colaborador(
            destino,
            por_matricula,
            por_nome,
        )

        linha = destino["linha"]

        if colaborador is None:
            resultados.append({
                "linha": linha,
                "destino": destino["nome"],
                "matricula_modelo": destino["matricula"],
                "origem": "",
                "matricula_origem": "",
                "status": status,
            })
            continue

        # Evita que a mesma pessoa da origem seja lançada duas vezes.
        if colaborador["matricula"] in preenchidos:
            resultados.append({
                "linha": linha,
                "destino": destino["nome"],
                "matricula_modelo": destino["matricula"],
                "origem": colaborador["nome"],
                "matricula_origem": colaborador["matricula"],
                "status": "ERRO - colaborador da origem já utilizado em outra linha",
            })
            continue

        preenchidos.add(colaborador["matricula"])

        # A/B permanecem como estão no modelo.
        # D = total de faltas
        ws.cell(linha, 4).value = colaborador["totais"]["faltas"] or timedelta(0)

        # E/F/G já possuem fórmulas no modelo.
        # Não substituímos as fórmulas; elas serão recalculadas pelo Excel.

        # Dias H:AL = faltas, atestados ou horas extras dos respectivos dias.
        for dia, valores in colaborador["diarios"].items():
            coluna = colunas_por_dia.get(dia)

            if not coluna:
                continue

            status_dia = valores.get("status_dia")

            # 1. Se identificou Falta (F) ou Atestado (AT) em E e P:
            if status_dia in ("F", "AT"):
                ws.cell(linha, coluna).value = status_dia
                ws.cell(linha, coluna).number_format = "@"  # Formato Texto

            # 2. Caso contrário, verifica se existem horas extras na origem:
            else:
                for chave in ("50", "70", "120"):
                    valor = valores.get(chave)

                    if valor is not None:
                        ws.cell(linha, coluna).value = valor
                        ws.cell(linha, coluna).number_format = "h:mm;@"

        # AM = total de atrasos.
        ws.cell(linha, 39).value = colaborador["totais"]["atrasos"] or timedelta(0)
        ws.cell(linha, 39).number_format = "[h]:mm:ss;@"

        resultados.append({
            "linha": linha,
            "destino": destino["nome"],
            "matricula_modelo": destino["matricula"],
            "origem": colaborador["nome"],
            "matricula_origem": colaborador["matricula"],
            "status": status,
        })

    # Força o Excel a recalcular as fórmulas E/F/G ao abrir.
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), resultados


# ============================================================
# INTERFACE
# ============================================================
col1, col2 = st.columns(2)

with col1:
    arquivo_origem = st.file_uploader(
        "1. 📥 Cartão Ponto exportado do sistema",
        type=["xlsx"],
        key="origem",
    )

with col2:
    arquivo_destino = st.file_uploader(
        "2. 📄 PONTO S.DIOGO em branco",
        type=["xlsx"],
        key="destino",
    )

st.divider()

if st.button("🚀 Processar Ponto", type="primary", use_container_width=True):

    if not arquivo_origem or not arquivo_destino:
        st.error("⚠️ Carregue os dois arquivos antes de processar.")
        st.stop()

    try:
        colaboradores = ler_cartao_ponto(arquivo_origem)

        if not colaboradores:
            st.error("❌ Nenhum colaborador foi identificado no Cartão Ponto.")
            st.stop()

        arquivo_final, resultados = preencher_planilha(
            arquivo_destino,
            colaboradores,
        )

        ok = sum("OK" in r["status"] for r in resultados)
        alertas = sum("ATENÇÃO" in r["status"] for r in resultados)
        erros = sum(
            ("ERRO" in r["status"] or "NÃO ENCONTRADO" in r["status"])
            for r in resultados
        )

        st.success(
            f"✅ Processamento concluído: {len(colaboradores)} colaboradores "
            f"identificados na origem."
        )

        m1, m2, m3 = st.columns(3)
        m1.metric("Conferidos", ok)
        m2.metric("Com divergência", alertas)
        m3.metric("Não encontrados / erros", erros)

        # Relatório de conferência.
        st.subheader("🔎 Conferência dos cruzamentos")

        df_resultado = pd.DataFrame(resultados)

        st.dataframe(
            df_resultado,
            use_container_width=True,
            hide_index=True,
        )

        if erros:
            st.error(
                "Existem colaboradores não encontrados ou com erro de "
                "cruzamento. Revise a tabela antes de enviar ao RH."
            )
        elif alertas:
            st.warning(
                "Existem divergências de nome/matrícula. "
                "O sistema realizou o lançamento, mas recomendamos conferir "
                "os registros destacados antes do envio ao RH."
            )

        nome_saida = arquivo_destino.name

        st.download_button(
            "📥 Baixar PONTO preenchido",
            data=arquivo_final,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as erro:
        st.exception(erro)
