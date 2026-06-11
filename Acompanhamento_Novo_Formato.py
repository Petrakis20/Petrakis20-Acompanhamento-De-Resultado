import io
import json
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

from processador_novo_formato import (
    aplicar_adicoes,
    aplicar_bi_resultados,
    aplicar_resultados,
    detectar_mes_por_nome,
    extrair_balancete_xls_bytes,
)

# Carrega funções do pacote bi_fiscal_pkg
_BI_PKG = Path(__file__).parent / "bi_fiscal_pkg"
if str(_BI_PKG) not in sys.path:
    sys.path.insert(0, str(_BI_PKG))

try:
    from bi_fiscal_extractor import (  # type: ignore[import]
        _process_es_sheet,
        extract_bi_data,
        extract_bi_data_from_env,
        test_connection,
        test_connection_from_env,
    )
    BI_DISPONIVEL = True
except Exception:
    BI_DISPONIVEL = False


CONFIG_PATH = Path("mapeamentos_novo_formato.json")

MESES_LABELS = {
    "janeiro": "Janeiro",
    "fevereiro": "Fevereiro",
    "marco": "Março",
    "abril": "Abril",
    "maio": "Maio",
    "junho": "Junho",
    "julho": "Julho",
    "agosto": "Agosto",
    "setembro": "Setembro",
    "outubro": "Outubro",
    "novembro": "Novembro",
    "dezembro": "Dezembro",
}

MODELOS_LABELS = {
    "lucro_presumido": "Lucro Presumido",
    "lucro_real": "Lucro Real",
    "mensal": "Mensal",
}


@st.cache_data
def carregar_config(_mtime: float = 0.0):
    """Carrega o JSON de mapeamentos. O parâmetro mtime força recarregamento
    automático sempre que o arquivo for modificado no disco."""
    with CONFIG_PATH.open("r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def extrair_bi_para_mes(
    company_code: str,
    ano: int,
    mes_nome: str,
    config: dict,
    usar_env: bool = True,
    host: str = "",
    database: str = "",
    user: str = "",
    password: str = "",
):
    """Extrai dados do BI Fiscal diretamente do banco para um mês específico.
    Retorna (result_entrada, result_saida) no formato normalizado de load_bi_multisheet.
    """
    mes_num = config["months"][mes_nome]["number"]
    start_date = datetime(ano, mes_num, 1)
    next_month = mes_num % 12 + 1
    next_year = ano if mes_num < 12 else ano + 1
    end_date = datetime(next_year, next_month, 1)

    if usar_env:
        df_e, df_s, msg = extract_bi_data_from_env(company_code, start_date, end_date)
    else:
        df_e, df_s, msg = extract_bi_data(host, database, user, password, company_code, start_date, end_date)

    if df_e is None and df_s is None:
        raise ValueError(msg)

    result_e = _process_es_sheet(df_e) if df_e is not None and not df_e.empty else None
    result_s = _process_es_sheet(df_s) if df_s is not None and not df_s.empty else None
    return result_e, result_s


def calcular_bi_valores(config, modelo_key, result_entrada, result_saida):
    """Calcula o somatório por CFOP do BI para cada regra bi_cfop_mappings do modelo."""
    modelo_config = config["models"][modelo_key]
    bi_mappings = modelo_config.get("bi_cfop_mappings", [])
    resultado = {}

    for regra in bi_mappings:
        cfop_pos = {str(c) for c in regra["cfop_positivo"]}
        cfop_neg = {str(c) for c in regra["cfop_negativo"]}

        total_pos = 0.0
        if result_entrada is not None:
            df_e, cfop_e = result_entrada
            if not df_e.empty:
                mask = cfop_e.isin(cfop_pos)
                total_pos = float(df_e.loc[mask, "v_cont"].sum())

        total_neg = 0.0
        if result_saida is not None:
            df_s, cfop_s = result_saida
            if not df_s.empty:
                mask = cfop_s.isin(cfop_neg)
                total_neg = float(df_s.loc[mask, "v_cont"].sum())

        resultado[regra["id"]] = round(total_pos - total_neg, 2)

    return resultado


def _somar_cfops_bi(result, cfops):
    if result is None:
        return {str(cfop): 0.0 for cfop in cfops}

    df, cfop_series = result
    if df.empty:
        return {str(cfop): 0.0 for cfop in cfops}

    base = pd.DataFrame(
        {
            "cfop": cfop_series.astype(str),
            "valor": df["v_cont"].astype(float),
        }
    )
    somas = base.groupby("cfop", dropna=False)["valor"].sum().to_dict()
    return {str(cfop): round(float(somas.get(str(cfop), 0.0)), 2) for cfop in cfops}


def _celula_destino_bi(config, modelo_config, mes, regra):
    trimestre = config["months"][mes]["quarter"]
    if "month_columns_by_quarter" in modelo_config:
        coluna = modelo_config["month_columns_by_quarter"][trimestre][mes]
        return f"{coluna}{regra['rows_by_quarter'][trimestre]}"

    coluna = modelo_config["month_columns"][mes]
    return f"{coluna}{regra['row']}"


def montar_preview_bi(config, modelo_key, resultados_por_mes):
    modelo_config = config["models"][modelo_key]
    resumo = []
    detalhes = []

    for mes, resultados in resultados_por_mes.items():
        result_entrada, result_saida = resultados
        for regra in modelo_config.get("bi_cfop_mappings", []):
            valores_entrada = _somar_cfops_bi(result_entrada, regra["cfop_positivo"])
            valores_saida = _somar_cfops_bi(result_saida, regra["cfop_negativo"])
            total_entrada = round(sum(valores_entrada.values()), 2)
            total_saida = round(sum(valores_saida.values()), 2)
            valor_final = round(total_entrada - total_saida, 2)

            resumo.append(
                {
                    "mes": MESES_LABELS[mes],
                    "regra": regra["id"],
                    "entradas_cfop": total_entrada,
                    "saidas_cfop": total_saida,
                    "valor_aplicado": valor_final,
                    "celula_destino": _celula_destino_bi(config, modelo_config, mes, regra),
                }
            )

            for cfop, valor in valores_entrada.items():
                detalhes.append(
                    {
                        "mes": MESES_LABELS[mes],
                        "regra": regra["id"],
                        "tipo": "Entrada (+)",
                        "cfop": cfop,
                        "valor_cfop": valor,
                        "valor_no_calculo": valor,
                    }
                )

            for cfop, valor in valores_saida.items():
                detalhes.append(
                    {
                        "mes": MESES_LABELS[mes],
                        "regra": regra["id"],
                        "tipo": "Saída (-)",
                        "cfop": cfop,
                        "valor_cfop": valor,
                        "valor_no_calculo": round(-valor, 2),
                    }
                )

    return pd.DataFrame(resumo), pd.DataFrame(detalhes)


def extrair_preview_bi(config, modelo_key, bi_config, progress=None):
    if bi_config["usar_env"]:
        conexao_ok, mensagem_conexao = test_connection_from_env()
    else:
        conexao_ok, mensagem_conexao = test_connection(
            bi_config["host"],
            bi_config["database"],
            bi_config["user"],
            bi_config["password"],
        )

    if not conexao_ok:
        raise ConnectionError(mensagem_conexao)

    resultados_por_mes = {}
    bi_valores_por_mes = {}
    meses_bi = bi_config["meses"]
    total = len(meses_bi)

    for i, mes in enumerate(meses_bi):
        if progress is not None:
            progress.progress(
                (i + 1) / total,
                text=f"BI Fiscal: extraindo {MESES_LABELS[mes]}...",
            )

        result_e, result_s = extrair_bi_para_mes(
            company_code=bi_config["company_code"],
            ano=bi_config["ano"],
            mes_nome=mes,
            config=config,
            usar_env=bi_config["usar_env"],
            host=bi_config["host"],
            database=bi_config["database"],
            user=bi_config["user"],
            password=bi_config["password"],
        )
        resultados_por_mes[mes] = (result_e, result_s)
        bi_valores_por_mes[mes] = calcular_bi_valores(config, modelo_key, result_e, result_s)

    resumo_df, detalhes_df = montar_preview_bi(config, modelo_key, resultados_por_mes)
    return resultados_por_mes, bi_valores_por_mes, resumo_df, detalhes_df


def montar_resumo_linhas(modelo_config):
    linhas = []
    for regra in modelo_config.get("result_mappings", []):
        destino = regra.get("row") or regra.get("rows_by_quarter")
        linhas.append(
            {
                "grupo": "Resultado",
                "id": regra["id"],
                "classificacoes": ", ".join(regra["classifications"]),
                "match": regra["match"],
                "destino": str(destino),
            }
        )

    for regra in modelo_config.get("additions_mappings", []):
        destino = regra.get("row") or regra.get("rows_by_quarter")
        linhas.append(
            {
                "grupo": "Adições e Exclusões",
                "id": regra["id"],
                "classificacoes": ", ".join(regra["classifications"]),
                "match": regra["match"],
                "destino": str(destino),
            }
        )

    for regra in modelo_config.get("bi_cfop_mappings", []):
        destino = regra.get("row") or regra.get("rows_by_quarter")
        cfops_pos = ", ".join(regra["cfop_positivo"])
        cfops_neg = ", ".join(f"-{c}" for c in regra["cfop_negativo"])
        linhas.append(
            {
                "grupo": "BI Fiscal",
                "id": regra["id"],
                "classificacoes": f"(+) {cfops_pos} | (-) {cfops_neg}",
                "match": "cfop",
                "destino": str(destino),
            }
        )

    return linhas


def preencher_workbook(
    config,
    modelo_key,
    arquivos,
    meses_por_arquivo,
    nome_empresa,
    cnpj_empresa,
    bi_valores_por_mes=None,
):
    modelo_config = config["models"][modelo_key]
    workbook = openpyxl.load_workbook(modelo_config["template_file"])
    sheet_resultado = workbook[modelo_config["result_sheet"]]

    if nome_empresa:
        sheet_resultado[modelo_config["company_cells"]["nome"]] = nome_empresa
    if cnpj_empresa:
        sheet_resultado[modelo_config["company_cells"]["cnpj"]] = cnpj_empresa

    resumo = []
    meses_processados = set()

    for arquivo in arquivos:
        mes = meses_por_arquivo[arquivo.name]
        trimestre = config["months"][mes]["quarter"]
        valores_balancete = extrair_balancete_xls_bytes(
            arquivo.getvalue(),
            config,
            filename=arquivo.name,
        )

        aplicar_resultados(sheet_resultado, modelo_config, mes, trimestre, valores_balancete)
        aplicar_adicoes(workbook, modelo_config, mes, trimestre, valores_balancete)

        if bi_valores_por_mes and mes in bi_valores_por_mes:
            aplicar_bi_resultados(
                sheet_resultado, modelo_config, mes, trimestre, bi_valores_por_mes[mes]
            )

        meses_processados.add(mes)
        resumo.append(
            {
                "arquivo": arquivo.name,
                "mes": MESES_LABELS[mes],
                "trimestre": trimestre,
                "classificacoes_lidas": len(valores_balancete),
            }
        )

    # Meses com BI mas sem balancete correspondente
    if bi_valores_por_mes:
        for mes, bi_valores in bi_valores_por_mes.items():
            if mes not in meses_processados:
                trimestre = config["months"][mes]["quarter"]
                aplicar_bi_resultados(sheet_resultado, modelo_config, mes, trimestre, bi_valores)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, resumo


# =============================================================================
# UI
# =============================================================================
st.set_page_config(page_title="Acompanhamento Novo Formato", layout="wide")

if Path("logo.png").exists():
    st.image("logo.png", width=100)

st.title("Acompanhamento de Resultado - Novo Formato")
st.caption("Leitura do balancete XLS novo por classificação contábil e preenchimento via JSON.")

config = carregar_config(_mtime=CONFIG_PATH.stat().st_mtime)

with st.sidebar:
    st.header("Configuração")
    modelo_label = st.selectbox("Modelo", list(MODELOS_LABELS.values()), key="modelo")
    modelo_key = next(key for key, label in MODELOS_LABELS.items() if label == modelo_label)

    modelo_config = config["models"][modelo_key]
    st.write("Modelo Excel:")
    st.code(modelo_config["template_file"])
    st.write("Aba de resultado:")
    st.code(modelo_config["result_sheet"])

import streamlit.components.v1 as components

nome_empresa = st.text_input("Nome da Empresa")
cnpj_empresa = st.text_input("CNPJ da Empresa", max_chars=18)

components.html(
    """
    <script>
    const doc = window.parent.document;
    const inputs = doc.querySelectorAll('input');
    inputs.forEach(input => {
        if (input.getAttribute('aria-label') === 'CNPJ da Empresa' && !input.dataset.masked) {
            input.dataset.masked = 'true';
            input.addEventListener('input', function(e) {
                let v = e.target.value.replace(/\\D/g, '');
                if (v.length > 14) v = v.substring(0, 14);
                v = v.replace(/^(\\d{2})(\\d)/, '$1.$2');
                v = v.replace(/^(\\d{2})\\.(\\d{3})(\\d)/, '$1.$2.$3');
                v = v.replace(/\\.(\\d{3})(\\d)/, '.$1/$2');
                v = v.replace(/(\\d{4})(\\d)/, '$1-$2');

                if (e.target.value !== v) {
                    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                    nativeInputValueSetter.call(input, v);
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                }
            });
        }
    });
    </script>
    """,
    height=0,
    width=0
)

arquivos = st.file_uploader(
    "Envie os balancetes no novo formato (.xls ou .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

# =============================================================================
# Seção BI Fiscal — extração direta do Alterdata
# =============================================================================
tem_bi_mappings = bool(modelo_config.get("bi_cfop_mappings"))
bi_state_prefix = f"bi_{modelo_key}"

usar_bi = False
bi_config = {}
bi_preview_params = None

if tem_bi_mappings:
    st.divider()
    st.subheader("BI Fiscal — Dif. Estoque (Gerencial vs Contábil)")

    if not BI_DISPONIVEL:
        st.warning(
            "Módulo de extração não disponível. "
            "Verifique se `pyodbc` e `numpy` estão instalados."
        )
    else:
        col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 1])
        with col_a:
            bi_company_code = st.text_input(
                "Código da Empresa",
                max_chars=5,
                placeholder="00768",
                key=f"{bi_state_prefix}_company_code",
            )
        with col_b:
            meses_lista = list(MESES_LABELS.keys())
            bi_mes_inicio = st.selectbox(
                "Mês inicial",
                meses_lista,
                index=0,
                format_func=lambda m: MESES_LABELS[m],
                key=f"{bi_state_prefix}_mes_inicio",
            )
        with col_c:
            idx_fim = max(meses_lista.index(bi_mes_inicio), 0)
            bi_mes_fim = st.selectbox(
                "Mês final",
                meses_lista,
                index=idx_fim,
                format_func=lambda m: MESES_LABELS[m],
                key=f"{bi_state_prefix}_mes_fim",
            )
        with col_d:
            bi_ano = st.number_input(
                "Ano",
                min_value=2020,
                max_value=2035,
                value=datetime.now().year,
                key=f"{bi_state_prefix}_ano",
            )

        usar_env = st.checkbox(
            "Usar credenciais do arquivo `.env` (Alterdata_BI/.env)",
            value=True,
            key=f"{bi_state_prefix}_usar_env",
        )

        if not usar_env:
            col1, col2 = st.columns(2)
            with col1:
                bi_host = st.text_input("Servidor (HOST\\INSTÂNCIA)", key=f"{bi_state_prefix}_host")
                bi_database = st.text_input("Banco de Dados", key=f"{bi_state_prefix}_database")
            with col2:
                bi_user = st.text_input("Usuário", key=f"{bi_state_prefix}_user")
                bi_password = st.text_input("Senha", type="password", key=f"{bi_state_prefix}_password")
        else:
            bi_host = bi_database = bi_user = bi_password = ""

        # Calcula a lista de meses no intervalo selecionado
        idx_inicio = meses_lista.index(bi_mes_inicio)
        idx_fim_sel = meses_lista.index(bi_mes_fim)
        if idx_fim_sel < idx_inicio:
            st.warning("O mês final não pode ser anterior ao mês inicial.")
            bi_meses_intervalo = []
        else:
            bi_meses_intervalo = meses_lista[idx_inicio: idx_fim_sel + 1]

        usar_bi = bool(bi_company_code) and bool(bi_meses_intervalo)
        bi_config = {
            "company_code": bi_company_code,
            "ano": int(bi_ano),
            "meses": bi_meses_intervalo,
            "usar_env": usar_env,
            "host": bi_host,
            "database": bi_database,
            "user": bi_user,
            "password": bi_password,
        }
        bi_preview_params = {
            "modelo_key": modelo_key,
            "company_code": bi_company_code,
            "ano": int(bi_ano),
            "meses": tuple(bi_meses_intervalo),
            "usar_env": usar_env,
            "host": bi_host,
            "database": bi_database,
            "user": bi_user,
            "password": bi_password,
        }

        preview_col, clear_col = st.columns([2, 1])
        with preview_col:
            preview_bi = st.button(
                "Pré-visualizar BI e CFOPs",
                disabled=not usar_bi,
                key=f"{bi_state_prefix}_preview_button",
            )
        with clear_col:
            limpar_preview = st.button("Limpar prévia", key=f"{bi_state_prefix}_preview_clear")

        if limpar_preview:
            st.session_state.pop(f"{bi_state_prefix}_preview_cache", None)

        if preview_bi:
            try:
                progress = st.progress(0, text="Extraindo dados do BI Fiscal...")
                resultados_por_mes, bi_valores_por_mes, resumo_df, detalhes_df = extrair_preview_bi(
                    config=config,
                    modelo_key=modelo_key,
                    bi_config=bi_config,
                    progress=progress,
                )
                progress.empty()
                st.session_state[f"{bi_state_prefix}_preview_cache"] = {
                    "params": bi_preview_params,
                    "resultados_por_mes": resultados_por_mes,
                    "bi_valores_por_mes": bi_valores_por_mes,
                    "resumo_df": resumo_df,
                    "detalhes_df": detalhes_df,
                }
            except ConnectionError as erro:
                st.warning(f"Não foi possível conectar ao BI Fiscal: {erro}")
            except Exception as erro:
                st.error(f"Erro ao pré-visualizar BI Fiscal: {erro}")

        preview_cache = st.session_state.get(f"{bi_state_prefix}_preview_cache")
        if preview_cache and preview_cache.get("params") == bi_preview_params:
            st.markdown("**Pré-visualização do BI**")
            st.dataframe(
                preview_cache["resumo_df"],
                use_container_width=True,
                hide_index=True,
            )
            with st.expander("Ver valores por CFOP"):
                st.dataframe(
                    preview_cache["detalhes_df"],
                    use_container_width=True,
                    hide_index=True,
                )

st.divider()

tab_processamento, tab_mapeamentos = st.tabs(["Processamento", "Mapeamentos"])

with tab_processamento:
    if not arquivos:
        st.info("Envie um ou mais balancetes para revisar os meses e processar.")
    else:
        st.subheader("Meses Detectados")
        meses = list(config["months"].keys())
        meses_por_arquivo = {}

        for arquivo in arquivos:
            mes_detectado = detectar_mes_por_nome(arquivo.name)
            indice = meses.index(mes_detectado) if mes_detectado in meses else 0
            mes = st.selectbox(
                arquivo.name,
                meses,
                index=indice,
                format_func=lambda item: MESES_LABELS[item],
                key=f"mes_{arquivo.file_id}_{arquivo.name}",
            )
            meses_por_arquivo[arquivo.name] = mes

        contagem_meses = Counter(meses_por_arquivo.values())
        duplicados = [MESES_LABELS[mes] for mes, qtd in contagem_meses.items() if qtd > 1]
        if duplicados:
            st.error(f"Há mais de um balancete para o mesmo mês: {', '.join(duplicados)}.")

        processar = st.button(
            "Processar Novo Formato",
            type="primary",
            disabled=bool(duplicados),
        )

        if processar:
            try:
                bi_valores_por_mes = {}

                if usar_bi and bi_config.get("company_code") and bi_config.get("meses"):
                    preview_cache = st.session_state.get(f"{bi_state_prefix}_preview_cache")
                    if preview_cache and preview_cache.get("params") == bi_preview_params:
                        bi_valores_por_mes = preview_cache["bi_valores_por_mes"]
                    else:
                        progress = st.progress(0, text="Extraindo dados do BI Fiscal...")
                        _, bi_valores_por_mes, resumo_df, detalhes_df = extrair_preview_bi(
                            config=config,
                            modelo_key=modelo_key,
                            bi_config=bi_config,
                            progress=progress,
                        )
                        progress.empty()
                        st.session_state[f"{bi_state_prefix}_preview_cache"] = {
                            "params": bi_preview_params,
                            "resultados_por_mes": {},
                            "bi_valores_por_mes": bi_valores_por_mes,
                            "resumo_df": resumo_df,
                            "detalhes_df": detalhes_df,
                        }

                arquivo_saida, resumo = preencher_workbook(
                    config=config,
                    modelo_key=modelo_key,
                    arquivos=arquivos,
                    meses_por_arquivo=meses_por_arquivo,
                    nome_empresa=nome_empresa,
                    cnpj_empresa=cnpj_empresa,
                    bi_valores_por_mes=bi_valores_por_mes if bi_valores_por_mes else None,
                )

                st.success("Processamento concluído.")
                st.dataframe(resumo, use_container_width=True, hide_index=True)

                if bi_valores_por_mes:
                    st.info(
                        f"BI Fiscal: {len(bi_valores_por_mes)} mês(es) extraído(s) do Alterdata — "
                        "Dif. Estoque aplicada."
                    )

                nome_arquivo = f"Acompanhamento Novo Formato - {modelo_label}.xlsx"
                st.download_button(
                    "Baixar planilha processada",
                    data=arquivo_saida,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except ConnectionError as erro:
                st.warning(f"Não foi possível conectar ao BI Fiscal: {erro}")
            except Exception as erro:
                st.error(f"Erro no processamento: {erro}")

with tab_mapeamentos:
    st.subheader(f"Mapeamentos - {modelo_label}")
    st.dataframe(
        montar_resumo_linhas(modelo_config),
        use_container_width=True,
        hide_index=True,
    )
