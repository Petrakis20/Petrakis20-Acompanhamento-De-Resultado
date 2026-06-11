"""
Módulo standalone de extração e processamento do BI Fiscal (Alterdata).

Dependências:
    pip install pandas openpyxl xlrd pyodbc python-dotenv numpy

Uso rápido:
    from bi_fiscal_extractor import extract_and_save, load_bi_multisheet

    # 1. Extrai do SQL e salva em Excel
    filepath, msg = extract_and_save("00768", datetime(2025,9,1), datetime(2025,10,1))

    # 2. Carrega o Excel gerado em DataFrames normalizados
    result_entrada, result_saida = load_bi_multisheet(filepath)
    # result_entrada = (df_out, cfop_series)  ou  None se sem movimento
    # result_saida   = (df_out, cfop_series)  ou  None se sem movimento
"""

from __future__ import annotations

import io
import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import pyodbc
from dotenv import load_dotenv


# =============================================================================
# Configuração — carrega .env se existir
# =============================================================================
_ENV_PATH = Path("Alterdata_BI/.env")
if _ENV_PATH.exists():
    load_dotenv(_ENV_PATH)


# =============================================================================
# Funções utilitárias
# =============================================================================
EMPTY_TOKENS_MAIN = {"", "nan", "none", "nao possui", "não possui", "x"}


def norm_text_main(s: str) -> str:
    s = str(s).strip().lower()
    s = (s.replace("ã","a").replace("á","a").replace("à","a").replace("â","a")
           .replace("ç","c").replace("é","e").replace("ê","e")
           .replace("í","i").replace("ó","o").replace("ô","o").replace("ú","u"))
    s = re.sub(r"[^a-z0-9 ]+"," ", s)
    s = re.sub(r"\s+"," ", s).strip()
    return s


def clean_code_main(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if s == "" or norm_text_main(s) in EMPTY_TOKENS_MAIN:
        return ""
    s = re.sub(r"\.0+$", "", s)
    s = re.sub(r"[^0-9]", "", s)
    while len(s) > 5 and s.endswith("0"):
        s = s[:-1]
    return s


def is_empty_code_main(x: str) -> bool:
    s = clean_code_main(x)
    return s == "" or set(s) == {"0"}


def to_number_br_main(v) -> float:
    if v is None:
        return 0.0
    s = str(v).strip()
    if s == "" or norm_text_main(s) in EMPTY_TOKENS_MAIN:
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    s = s.replace(".", "").replace(",", ".")
    try:
        val = float(s)
    except Exception:
        val = float(re.sub(r"[^0-9\.\-]", "", s) or 0)
    return -val if neg else val


def read_excel_best_main(file) -> pd.DataFrame:
    """Lê a planilha Excel com mais colunas (melhor estrutura)."""
    xls = pd.ExcelFile(file)
    best_df, best_cols = None, -1
    for sh in xls.sheet_names:
        df = xls.parse(sh)
        if df.shape[1] > best_cols:
            best_df, best_cols = df, df.shape[1]
    return best_df if best_df is not None else pd.DataFrame()


# =============================================================================
# Queries SQL (Alterdata — WFiscal)
# =============================================================================
_QUERY_BASE = """
SELECT DISTINCT
/*  1 */ CAST(CASE WHEN M.StCancelada = 'S' THEN 'Sim' ELSE 'Não' END AS VARCHAR(3)) AS [Cancelada],
/*  2 */ M.dtEscrituracao                                   AS [Dt. Escrituração],
/*  3 */ M.DtEmissao                                        AS [Data Emissão],
/*  4 */ M.IdCodFiscal                                      AS [CFOP],
/*  5 */ CAST(CASE
              WHEN CFOP.CdTipo = 'C' AND M.StTipo = 'E' THEN 'Compras Normais'
              WHEN CFOP.CdTipo = 'C' AND M.StTipo = 'S' THEN 'Vendas Normais'
              WHEN CFOP.CdTipo = 'T' THEN 'Transferências'
              WHEN CFOP.CdTipo = 'D' THEN 'Devoluções'
              WHEN CFOP.CdTipo = 'E' THEN 'Energia Elétrica'
              WHEN CFOP.CdTipo = 'U' THEN 'Uso e Consumo'
              WHEN CFOP.CdTipo = 'A' THEN 'Ativo Imobilizado'
              WHEN CFOP.CdTipo = 'M' THEN 'Comunicações'
              WHEN CFOP.CdTipo = 'R' THEN 'Transportes'
              WHEN CFOP.CdTipo = 'O' THEN 'Outros'
              WHEN CFOP.CdTipo = 'X' THEN 'Exportação'
              WHEN CFOP.CdTipo = 'I' THEN 'Importação'
              WHEN CFOP.CdTipo = 'S' THEN 'Subst. Tributária'
              WHEN CFOP.CdTipo = 'N' THEN 'Transf. Ativo'
              WHEN CFOP.CdTipo = 'F' THEN 'Transf. Uso Cons.'
              WHEN CFOP.CdTipo = '.' THEN 'Transf. Crédito'
              ELSE '' END AS VARCHAR(30))                   AS [Tipo CFOP],
/*  6 */ M.NmNumero                                         AS [Número],
/*  7 */ F.NmNome                                           AS [Nome Forn/Cliente],
/*  8 */ M.VlContabil                                       AS [Valor Contábil],
/*  9 */ ROUND(COALESCE(M.VlICMSValor,0), 2)                AS [Vl. ICMS],
/* 10 */ M.VlValorST                                        AS [Vl. ST],
/* 11 */ M.VlIPIValor                                       AS [Vl. IPI],
/* 12 */ M.IdTipoOperacao                                   AS [Cód. Oper. Contábil],
/* 13 */ M.IdLancContabil                                   AS [Lanc. Cont. Vl. Contábil],
/* 14 */ M.IdLancIcms                                       AS [Lanc. Cont. Vl. ICMS],
/* 15 */ M.IdLancIcmsST                                     AS [Lanc. Cont. Vl. Subst. Trib.],
/* 16 */ M.IdLancIpi                                        AS [Lanc. Cont. Vl. IPI],
/* 17 */ CAST(CASE
              WHEN COALESCE(CASE WHEN ZOP.Exportado = CAST(1 AS bit) THEN 'S' ELSE M.StExportado END,'N')='S'
              THEN 'Sim' ELSE 'Não' END AS VARCHAR(3))      AS [Exportado],
/* 18 */ M.VlBaseST                                         AS [Base ST],
/* 19 */ COALESCE(M.Total_Pis_Unidade_Medida,0)
       + COALESCE(M.Total_Pis_Cumulativo,0)
       + COALESCE(M.Total_Pis_Nao_Cumulativo,0)             AS [Total PIS],
/* 20 */ COALESCE(M.Total_Cofins_Unidade_Medida,0)
       + COALESCE(M.Total_Cofins_Cumulativo,0)
       + COALESCE(M.Total_Cofins_Nao_Cumulativo,0)          AS [Total CONFINS],
/* 21 */ M.VlIPIBase                                        AS [Vl. Base IPI],
/* 22 */ M.VlIPIAliquota                                    AS [% IPI],
/* 23 */ M.VlIPINaoAproveitado                              AS [IPI Não Aproveitado],
/* 24 */ M.VlICMSBase                                       AS [Vl. Base ICMS],
/* 25 */ M.VlICMSAliquota                                   AS [%ICMS],
/* 26 */ M.CSTICMS                                          AS [CST ICMS],
/* 27 */ M.informacao_complementar                          AS [Informações Complementares],
/* 28 */ CONVERT(DATETIME, M.data_importacao)               AS [Data da Importação],
/* 29 */ M.nome_usuario_importacao                          AS [Usuário Importador],
/* 30 */ F.CdCgc                                            AS [CNPJ/CPF forn/Cliente],
/* 31 */ M.IdModDocFiscal                                   AS [Mod.],
/* 32 */ M.chave_acesso_nota_eletronica                     AS [Chave de Acesso NFe/CF SAT]
FROM WFiscal.M{code5} M
LEFT JOIN WFISCAL.movimento_reducao_z Z
       ON Z.Data    = M.DtEscrituracao
      AND Z.Ecf_Id  = M.CodECF
LEFT JOIN WFISCAL.movimento_reducao_z_por_operacao ZOP
       ON ZOP.Movimento_Reducao_Z_Id = Z.Id
      AND ZOP.cfop                   = M.IdCodFiscal
      AND ZOP.aliquota_icms          = M.VlICMSAliquota
      AND RIGHT(REPLICATE('0',3) + CAST(ZOP.cst_icms AS VARCHAR(3)), 3)
          = RIGHT(REPLICATE('0',3) + CAST(M.CSTICMS   AS VARCHAR(3)), 3)
LEFT JOIN WFiscal.CadFisM CFOP ON M.IdCodFiscal = CFOP.IdCodigo
LEFT JOIN wfiscal.FORNEC F     ON M.IdCodForCli = F.CdFornecedor
LEFT JOIN wphd.MunicipiosIBGE MUN
       ON (MUN.IdMunicipio = F.IdMunicipio AND M.TpEmissaoNF <> 'S')
       OR (MUN.IdMunicipio = ?                 AND M.TpEmissaoNF  = 'S')
LEFT JOIN wfiscal.arquivos_xml_danfe X ON M.chave_acesso_nota_eletronica = X.id
LEFT JOIN WFiscal.MODDOC MD            ON M.IdModDocFiscal = MD.CdCodigo
WHERE M.dtEscrituracao >= ?
  AND M.dtEscrituracao <  ?
  AND M.StTipo = '{tipo}'
  AND ISNULL(NULLIF(LTRIM(RTRIM(M.StCancelada)), ''), 'N') = 'N'
"""

QUERY_SAIDA   = _QUERY_BASE.replace("{tipo}", "S")
QUERY_ENTRADA = _QUERY_BASE.replace("{tipo}", "E")


# =============================================================================
# Conexão SQL Server
# =============================================================================
def get_db_credentials() -> Tuple[str, str, str, str]:
    host     = os.getenv("DB_HOST", "")
    database = os.getenv("DB_NAME", "")
    user     = os.getenv("DB_USER", "")
    password = os.getenv("DB_PASS", "").strip("'\"")
    return host, database, user, password


def get_connection_string(host: str, database: str, user: str, password: str) -> str:
    drivers = [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
    ]
    available = [d for d in pyodbc.drivers() if "SQL Server" in d]
    driver = next((d for d in drivers if d in available), available[0] if available else drivers[0])

    if "\\" in host:
        server, instance = host.split("\\", 1)
        return (f"DRIVER={{{driver}}};SERVER={server}\\{instance};DATABASE={database};"
                f"UID={user};PWD={password};TrustServerCertificate=yes")
    return (f"DRIVER={{{driver}}};SERVER={host};DATABASE={database};"
            f"UID={user};PWD={password};TrustServerCertificate=yes")


def test_connection(host: str, database: str, user: str, password: str) -> Tuple[bool, str]:
    try:
        conn = pyodbc.connect(get_connection_string(host, database, user, password), timeout=10)
        conn.close()
        return True, "Conexão estabelecida com sucesso!"
    except Exception as e:
        return False, f"Erro ao conectar: {e}"


def test_connection_from_env() -> Tuple[bool, str]:
    host, database, user, password = get_db_credentials()
    if not all([host, database, user, password]):
        return False, "Credenciais não encontradas no arquivo Alterdata_BI/.env"
    return test_connection(host, database, user, password)


def _check_company_exists(conn: pyodbc.Connection, code5: str) -> bool:
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 1 FROM sys.objects o JOIN sys.schemas s ON s.schema_id = o.schema_id "
            "WHERE s.name = 'WFiscal' AND o.name = ? AND o.type IN ('U','V')",
            f"M{code5}"
        )
        return cursor.fetchone() is not None
    except Exception:
        return False


# =============================================================================
# Extração SQL → DataFrames brutos
# =============================================================================
def extract_bi_data(
    host: str,
    database: str,
    user: str,
    password: str,
    company_code: str,
    start_date: datetime,
    end_date: datetime,
    municipio_id: Optional[int] = None,
) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], str]:
    """
    Extrai dados de BI do Alterdata.

    Returns:
        (df_entrada, df_saida, mensagem)
    """
    try:
        code5 = str(company_code).zfill(5)
        conn = pyodbc.connect(get_connection_string(host, database, user, password), timeout=30)

        if not _check_company_exists(conn, code5):
            conn.close()
            return None, None, f"Empresa {code5} não encontrada (tabela WFiscal.M{code5} inexistente)"

        params = [municipio_id, start_date, end_date]
        df_saida   = pd.read_sql(QUERY_SAIDA.replace("{code5}", code5),   conn, params=params)
        df_entrada = pd.read_sql(QUERY_ENTRADA.replace("{code5}", code5), conn, params=params)
        conn.close()

        return df_entrada, df_saida, (
            f"Extração concluída: {len(df_entrada)} registros de Entrada, "
            f"{len(df_saida)} registros de Saída"
        )
    except Exception as e:
        return None, None, f"Erro ao extrair dados: {e}"


def extract_bi_data_from_env(
    company_code: str,
    start_date: datetime,
    end_date: datetime,
    municipio_id: Optional[int] = None,
) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], str]:
    """Extrai usando credenciais do .env."""
    host, database, user, password = get_db_credentials()
    if not all([host, database, user, password]):
        return None, None, "Credenciais não encontradas no arquivo Alterdata_BI/.env"
    return extract_bi_data(host, database, user, password, company_code, start_date, end_date, municipio_id)


# =============================================================================
# Geração do Excel (abas Entrada / Saída)
# =============================================================================
def generate_excel_file(
    df_entrada: pd.DataFrame,
    df_saida: pd.DataFrame,
    company_code: str,
    start_date: datetime,
    end_date: datetime,
    output_dir: str = "temp_bi",
) -> Tuple[Optional[str], str]:
    """
    Salva os DataFrames em um .xlsx com abas 'Entrada' e 'Saída'.

    Returns:
        (caminho_arquivo, mensagem)
    """
    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        code5    = str(company_code).zfill(5)
        today    = datetime.now()
        filename = (f"{code5}_{start_date.strftime('%Y-%m-%d')}_{end_date.strftime('%Y-%m-%d')}"
                    f"_{today.strftime('%d-%m-%Y')}.xlsx")
        filepath = Path(output_dir) / filename

        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            if df_entrada is None or df_entrada.empty:
                pd.DataFrame([["(sem registros no período)"]]).to_excel(
                    writer, sheet_name="Entrada", index=False, header=False)
            else:
                df_entrada.to_excel(writer, sheet_name="Entrada", index=False)

            if df_saida is None or df_saida.empty:
                pd.DataFrame([["(sem registros no período)"]]).to_excel(
                    writer, sheet_name="Saída", index=False, header=False)
            else:
                df_saida.to_excel(writer, sheet_name="Saída", index=False)

        return str(filepath), f"Arquivo gerado: {filename}"
    except Exception as e:
        return None, f"Erro ao gerar Excel: {e}"


def extract_and_save(
    company_code: str,
    start_date: datetime,
    end_date: datetime,
    municipio_id: Optional[int] = None,
    output_dir: str = "temp_bi",
) -> Tuple[Optional[str], str]:
    """
    Atalho: extrai do banco e salva em Excel num único passo.

    Returns:
        (caminho_arquivo, mensagem)
    """
    df_entrada, df_saida, msg = extract_bi_data_from_env(
        company_code, start_date, end_date, municipio_id)
    if df_entrada is None and df_saida is None:
        return None, msg
    return generate_excel_file(df_entrada, df_saida, company_code, start_date, end_date, output_dir)


# =============================================================================
# Constantes para processamento do Excel
# =============================================================================
REQUIRED_COLS_DISPLAY = [
    "CFOP",
    "Lanc. Cont. Vl. Contábil",
    "Lanc. Cont. Vl. ICMS",
    "Lanc. Cont. Vl. Subst. Trib.",
    "Lanc. Cont. Vl. IPI",
]

OPTIONAL_VALUE_COLS = ["Valor Contábil", "Vl. ICMS", "Vl. ST", "Vl. IPI", "Cancelada"]

INTERNAL_KEYS = {
    "CFOP":                         "CFOP",
    "Lanc. Cont. Vl. Contábil":     "contabil",
    "Lanc. Cont. Vl. ICMS":         "icms",
    "Lanc. Cont. Vl. Subst. Trib.": "icms_subst",
    "Lanc. Cont. Vl. IPI":          "ipi",
}

INTERNAL_VALUE_KEYS = {
    "Valor Contábil": "valor_contabil",
    "Vl. ICMS":       "vl_icms",
    "Vl. ST":         "vl_st",
    "Vl. IPI":        "vl_ipi",
    "Cancelada":      "cancelada",
}

SERV_COLS = [
    ("Lanc Cont. Valor Documento", ["Valor Documento"],                    "doc"),
    ("Lanc.Cont.Vl.Cofins",        ["Vl. Cofins"],                         "cofins"),
    ("Lanc.Cont.Vl.PIS",           ["Vl. PIS"],                            "pis"),
    ("Lanc Cont. Vl. ISS",         ["Vl. ISS"],                            "iss"),
    ("Lanc Cont. Vl. ISS Ret.",    ["Vl. ISS Ret.", "Vl ISS Ret"],         "iss_ret"),
    ("Lanc Cont. Vl. IRRF",        ["Vl. IRRF", "Vl IRRF"],               "irrf"),
    ("Lanc Cont. Vl. PIS Ret.",    ["Vl. PIS Ret.", "Vl PIS Ret"],        "pis_ret"),
    ("Lanc Cont. Vl. COFINS Ret.", ["Vl. COFINS Ret.", "Vl COFINS Ret"],  "cofins_ret"),
    ("Lanc Cont. Vl. INSS Ret.",   ["Vl. INSS Ret.", "Vl INSS Ret"],      "inss_ret"),
    ("Lanc Cont. Vl. CSLL Ret.",   ["Vl. CSLL Ret.", "Vl CSLL Ret"],      "csll_ret"),
]


# =============================================================================
# Leitura de Excel (internos)
# =============================================================================
def _read_excel_first_sheet(data: bytes, engine: str) -> Optional[pd.DataFrame]:
    bio = io.BytesIO(data)
    try:
        xls = pd.ExcelFile(bio, engine=engine)
        if not xls.sheet_names:
            return None
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str)
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        return None


def _try_read_as_excel(data: bytes) -> Optional[pd.DataFrame]:
    df = _read_excel_first_sheet(data, engine="openpyxl")
    if df is not None:
        return df
    try:
        df = _read_excel_first_sheet(data, engine="xlrd")
    except Exception:
        pass
    return df


def _try_read_as_csv(data: bytes) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_csv(io.BytesIO(data), dtype=str, sep=None, engine="python")
        df.columns = [str(c) for c in df.columns]
        return df
    except Exception:
        return None


def _open_excel_file(file) -> pd.ExcelFile:
    """Abre ExcelFile a partir de caminho (str/Path) ou objeto de arquivo."""
    if isinstance(file, (str, Path)):
        try:
            return pd.ExcelFile(file, engine="openpyxl")
        except Exception:
            return pd.ExcelFile(file, engine="xlrd")

    raw = file.read()
    bio = io.BytesIO(raw)
    try:
        return pd.ExcelFile(bio, engine="openpyxl")
    except Exception:
        bio = io.BytesIO(raw)
        return pd.ExcelFile(bio, engine="xlrd")


def _read_excel_hybrid(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """Lê Excel mantendo colunas de código como str e valores como float."""
    code_columns = [
        "CFOP", "Lanc. Cont. Vl. Contábil", "Lanc. Cont. Vl. ICMS",
        "Lanc. Cont. Vl. Subst. Trib.", "Lanc. Cont. Vl. IPI", "Cancelada",
    ]
    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
    dtype_dict = {c: str for c in code_columns if c in df_temp.columns}
    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=dtype_dict)
    df.columns = [str(c) for c in df.columns]
    return df


def _safe_to_number(val) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    return to_number_br_main(val)


# =============================================================================
# Filtros de limpeza
# =============================================================================
def _filter_cancelada(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "cancelada" not in df.columns:
        return df
    df = df.copy()
    empty_mask = (
        df["cancelada"].isna() |
        df["cancelada"].astype(str).str.strip().eq("") |
        df["cancelada"].astype(str).str.lower().isin(["nan", "none", "null"])
    )
    return df.loc[~empty_mask].reset_index(drop=True)


def _drop_garbage(df: pd.DataFrame, cfop_series: pd.Series) -> Tuple[pd.DataFrame, pd.Series]:
    """Remove linhas com CFOP vazio E todos os valores zerados."""
    req = ["v_cont", "v_icms", "v_st", "v_ipi"]
    if not all(c in df.columns for c in req):
        return df, cfop_series

    cfop_digits = (cfop_series.astype(str).fillna("").str.replace(r"\D+", "", regex=True))
    cfop_empty  = cfop_digits.str.len().eq(0)
    all_zero    = df[req].fillna(0.0).astype(float).abs().sum(axis=1).eq(0.0)
    keep        = ~(cfop_empty & all_zero)
    return df.loc[keep].reset_index(drop=True), cfop_series.loc[keep].reset_index(drop=True)


# =============================================================================
# Detecção automática de colunas
# =============================================================================
def detect_bi_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mapping = {c: norm_text_main(c) for c in df.columns}
    aliases = {
        "cfop":      ["cfop"],
        "la_cont":   ["lanc cont vl contabil", "lanc cont vl"],
        "la_icms":   ["lanc cont vl icms"],
        "la_st":     ["lanc cont vl subst trib", "lanc cont vl icms st", "lanc cont vl subst"],
        "la_ipi":    ["lanc cont vl ipi", "trib lanc cont vl ipi"],
        "v_cont":    ["valor contabil"],
        "v_icms":    ["vl icms"],
        "v_st":      ["vl subst trib", "vl icms st", "vl st"],
        "v_ipi":     ["vl ipi"],
        "cancelada": ["cancelada"],
    }
    cols: Dict[str, Optional[str]] = {k: None for k in aliases}
    for key, opts in aliases.items():
        for c, nm in mapping.items():
            if nm in opts:
                cols[key] = c
                break
        if cols[key] is None and key == "la_cont":
            for c, nm in mapping.items():
                if nm.startswith("lanc cont vl") and not any(x in nm for x in ["icms", "ipi", "st", "subst"]):
                    cols[key] = c
                    break

    missing = [k for k in ["la_cont","la_icms","la_st","la_ipi","v_cont","v_icms","v_st","v_ipi"] if cols[k] is None]
    if missing:
        raise ValueError(f"Colunas do BI ausentes: {missing}")
    return cols


def _find_col(df: pd.DataFrame, *alvos: str) -> Optional[str]:
    m = {c: norm_text_main(c) for c in df.columns}
    for alvo in alvos:
        a = norm_text_main(alvo)
        for c, n in m.items():
            if n == a:
                return c
    for alvo in alvos:
        a = norm_text_main(alvo)
        for c, n in m.items():
            if n.startswith(a):
                return c
    return None


# =============================================================================
# Processamento de abas Entrada / Saída (interno)
# =============================================================================
_EMPTY_SHEET_TOKEN = "(sem registros no período)"

_EMPTY_ES = (
    pd.DataFrame(columns=["la_cont","la_icms","la_st","la_ipi","v_cont","v_icms","v_st","v_ipi"]),
    pd.Series([], dtype="object"),
)


def _process_es_sheet(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.Series]:
    """Normaliza um DataFrame de BI Entrada ou Saída."""
    cols = detect_bi_columns(df)

    cfop_raw    = df[cols["cfop"]] if cols.get("cfop") else pd.Series([""] * len(df), index=df.index)
    cfop_series = cfop_raw.map(clean_code_main)

    out = pd.DataFrame({
        "la_cont": df[cols["la_cont"]].map(clean_code_main),
        "la_icms": df[cols["la_icms"]].map(clean_code_main),
        "la_st":   df[cols["la_st"]].map(clean_code_main),
        "la_ipi":  df[cols["la_ipi"]].map(clean_code_main),
        "v_cont":  df[cols["v_cont"]].map(_safe_to_number),
        "v_icms":  df[cols["v_icms"]].map(_safe_to_number),
        "v_st":    df[cols["v_st"]].map(_safe_to_number),
        "v_ipi":   df[cols["v_ipi"]].map(_safe_to_number),
    })

    if cols.get("cancelada"):
        out["cancelada"] = df[cols["cancelada"]]
        cancelada_empty = (
            out["cancelada"].isna() |
            out["cancelada"].astype(str).str.strip().eq("") |
            out["cancelada"].astype(str).str.lower().isin(["nan", "none", "null"])
        )
        keep_mask   = ~cancelada_empty
        out         = out.loc[keep_mask].reset_index(drop=True)
        cfop_series = cfop_series.loc[keep_mask].reset_index(drop=True)

    return _drop_garbage(out, cfop_series)


def _is_empty_sheet(df: pd.DataFrame) -> bool:
    cols = list(df.columns)
    return len(cols) == 1 and _EMPTY_SHEET_TOKEN in str(cols[0]).lower()


# =============================================================================
# API pública — carregamento de Excel
# =============================================================================
def load_bi_multisheet(
    file,
) -> Tuple[Optional[Tuple[pd.DataFrame, pd.Series]], Optional[Tuple[pd.DataFrame, pd.Series]]]:
    """
    Lê arquivo Excel com abas 'Entrada' e/ou 'Saída' e retorna DataFrames normalizados.

    Args:
        file: Caminho (str/Path) ou objeto de arquivo (UploadedFile do Streamlit)

    Returns:
        (result_entrada, result_saida)
        Cada result é (df_out, cfop_series) ou None se a aba não existir.
        df_out tem colunas: la_cont, la_icms, la_st, la_ipi, v_cont, v_icms, v_st, v_ipi
    """
    if file is None:
        return None, None

    xls              = _open_excel_file(file)
    available_sheets = xls.sheet_names

    entrada_sheet = next((s for s in available_sheets if s.lower().strip() == "entrada"), None)
    saida_sheet   = next((s for s in available_sheets if s.lower().strip() in ("saída", "saida")), None)

    result_entrada = result_saida = None

    if entrada_sheet:
        try:
            df = _read_excel_hybrid(xls, entrada_sheet)
            result_entrada = _EMPTY_ES if _is_empty_sheet(df) else _process_es_sheet(df)
        except Exception as e:
            raise ValueError(f"Erro ao processar aba 'Entrada': {e}")

    if saida_sheet:
        try:
            df = _read_excel_hybrid(xls, saida_sheet)
            result_saida = _EMPTY_ES if _is_empty_sheet(df) else _process_es_sheet(df)
        except Exception as e:
            raise ValueError(f"Erro ao processar aba 'Saída': {e}")

    if result_entrada is None and result_saida is None:
        raise ValueError(
            f"Nenhuma aba 'Entrada' ou 'Saída' encontrada. Abas disponíveis: {', '.join(available_sheets)}"
        )

    return result_entrada, result_saida


def load_bi_es(file) -> Tuple[pd.DataFrame, pd.Series]:
    """
    Lê um único arquivo BI de Entradas ou Saídas (sem abas separadas).

    Returns:
        (df_out, cfop_series)
    """
    df = read_excel_best_main(file)
    return _process_es_sheet(df)


def load_bi_servico(file) -> Tuple[pd.DataFrame, pd.Series, pd.DataFrame]:
    """
    Carrega BI de Serviços.

    Returns:
        (agg_lancamentos, cfop_series, missing_matrix_srv)
    """
    df = read_excel_best_main(file)

    cancelada_col = _find_col(df, "cancelada")
    if cancelada_col:
        cancelada_empty = (
            df[cancelada_col].isna() |
            df[cancelada_col].astype(str).str.strip().eq("") |
            df[cancelada_col].astype(str).str.lower().isin(["nan", "none", "null"])
        )
        df = df.loc[~cancelada_empty].reset_index(drop=True)

    cfop_col    = _find_col(df, "cfop")
    cfop_series = df[cfop_col].map(clean_code_main) if cfop_col else pd.Series([], dtype="object")

    stacks, pres_cols = [], {}
    for code_label, val_opts, lbl in SERV_COLS:
        c_cod = _find_col(df, code_label, code_label.replace(".", " ").replace("  ", " "))
        c_val = None
        for vname in val_opts:
            c_val = c_val or _find_col(df, vname, vname.replace(".", " ").replace("  ", " "))
        if c_cod and c_val:
            pres_cols[lbl] = ~df[c_cod].map(is_empty_code_main)
            tmp = pd.DataFrame({
                "lancamento": df[c_cod].map(clean_code_main),
                "valor":      df[c_val].map(to_number_br_main),
            })
            tmp = tmp[(tmp["lancamento"] != "") & (tmp["valor"].notna())]
            stacks.append(tmp)

    if not stacks:
        raise ValueError("Nenhuma dupla código/valor encontrada no BI de Serviços.")

    long      = pd.concat(stacks, ignore_index=True)
    long["valor"] = long["valor"].fillna(0.0).astype(float)
    agg       = long.groupby("lancamento", as_index=False)["valor"].sum().rename(columns={"valor": "valor_bi"})

    missing_matrix_srv = pd.DataFrame()
    if not cfop_series.empty and pres_cols:
        aux = pd.DataFrame({"CFOP": cfop_series.map(clean_code_main)})
        for lbl, ser in pres_cols.items():
            aux[lbl] = ser.fillna(False).astype(bool)
        aux = aux[aux["CFOP"] != ""]
        grp = aux.groupby("CFOP").agg({lbl: "any" for lbl in pres_cols}).reset_index()

        label_order = ["doc","cofins","pis","iss","iss_ret","irrf","pis_ret","cofins_ret","inss_ret","csll_ret"]
        col_map     = {
            "doc": "Documento", "cofins": "Cofins", "pis": "PIS", "iss": "ISS",
            "iss_ret": "ISS Ret.", "irrf": "IRRF", "pis_ret": "PIS Ret.",
            "cofins_ret": "COFINS Ret.", "inss_ret": "INSS Ret.", "csll_ret": "CSLL Ret.",
        }
        out = grp[["CFOP"]].copy()
        for lbl in label_order:
            out[col_map[lbl]] = np.where(grp.get(lbl, False), "OK", "FALTA")
        mask_any           = (out.drop(columns=["CFOP"]) == "FALTA").any(axis=1)
        missing_matrix_srv = out[mask_any].sort_values("CFOP")

    return agg, cfop_series, missing_matrix_srv


# =============================================================================
# Funções de agregação
# =============================================================================
def aggregate_bi_all(bi: pd.DataFrame) -> pd.DataFrame:
    """Agrega valores do BI por lançamento contábil."""
    stacks = []
    for c_l, c_v in [("la_cont","v_cont"),("la_icms","v_icms"),("la_st","v_st"),("la_ipi","v_ipi")]:
        tmp = bi[[c_l, c_v]].copy()
        tmp.columns = ["lancamento", "valor"]
        tmp["lancamento"] = tmp["lancamento"].map(clean_code_main)
        tmp = tmp[tmp["lancamento"] != ""]
        stacks.append(tmp)
    long = pd.concat(stacks, ignore_index=True) if stacks else pd.DataFrame(columns=["lancamento","valor"])
    long["valor"] = long["valor"].fillna(0.0).astype(float)
    return long.groupby("lancamento", as_index=False)["valor"].sum().rename(columns={"valor": "valor_bi"})


def cfop_missing_matrix_es(bi_df: pd.DataFrame, cfop_series: pd.Series) -> pd.DataFrame:
    """Matriz de lacunas de lançamentos por CFOP (Entradas/Saídas)."""
    if cfop_series is None or cfop_series.empty:
        return pd.DataFrame()
    aux = pd.DataFrame({
        "CFOP":     cfop_series.map(clean_code_main),
        "has_cont": ~bi_df["la_cont"].map(is_empty_code_main),
        "has_icms": ~bi_df["la_icms"].map(is_empty_code_main),
        "has_st":   ~bi_df["la_st"].map(is_empty_code_main),
        "has_ipi":  ~bi_df["la_ipi"].map(is_empty_code_main),
    })
    aux = aux[aux["CFOP"] != ""]
    grp = aux.groupby("CFOP").agg({k:"any" for k in ["has_cont","has_icms","has_st","has_ipi"]}).reset_index()
    out = grp[["CFOP"]].copy()
    out["Contábil"] = np.where(grp["has_cont"], "OK", "FALTA")
    out["ICMS"]     = np.where(grp["has_icms"], "OK", "FALTA")
    out["ST"]       = np.where(grp["has_st"],   "OK", "FALTA")
    out["IPI"]      = np.where(grp["has_ipi"],  "OK", "FALTA")
    mask = (out[["Contábil","ICMS","ST","IPI"]] == "FALTA").any(axis=1)
    return out[mask].sort_values("CFOP")
