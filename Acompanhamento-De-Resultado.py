import streamlit as st
import pandas as pd
import openpyxl
import xlrd
import io

# --- Mapeamento LUCRO REAL (exemplo) ---
mapeamento_acomp_lucro_real = {
    994:  {"Janeiro": "D17", "Fevereiro": "E17", "Março": "F17", "Abril": "D54", "Maio": "E54", "Junho": "F54",
           "Julho": "D91", "Agosto": "E91", "Setembro": "F91", "Outubro": "D128", "Novembro": "E128", "Dezembro": "F128"},
    1022: {"Janeiro": "D18", "Fevereiro": "E18", "Março": "F18", "Abril": "D55", "Maio": "E55", "Junho": "F55",
           "Julho": "D92", "Agosto": "E92", "Setembro": "F92", "Outubro": "D129", "Novembro": "E129", "Dezembro": "F129"},
    1085: {"Janeiro": "D19", "Fevereiro": "E19", "Março": "F19", "Abril": "D56", "Maio": "E56", "Junho": "F56",
           "Julho": "D93", "Agosto": "E93", "Setembro": "F93", "Outubro": "D130", "Novembro": "E130", "Dezembro": "F130"},
    1176: {"Janeiro": "D20", "Fevereiro": "E20", "Março": "F20", "Abril": "D57", "Maio": "E57", "Junho": "F57",
           "Julho": "D94", "Agosto": "E94", "Setembro": "F94", "Outubro": "D131", "Novembro": "E131", "Dezembro": "F131"},
    5139: {"Janeiro": "D21", "Fevereiro": "E21", "Março": "F21", "Abril": "D58", "Maio": "E58", "Junho": "F58",
           "Julho": "D95", "Agosto": "E95", "Setembro": "F95", "Outubro": "D132", "Novembro": "E132", "Dezembro": "F132"},
    1197: {"Janeiro": "D22", "Fevereiro": "E22", "Março": "F22", "Abril": "D59", "Maio": "E59", "Junho": "F59",
           "Julho": "D96", "Agosto": "E96", "Setembro": "F96", "Outubro": "D133", "Novembro": "E133", "Dezembro": "F133"},
    3276: {"Janeiro": "D23", "Fevereiro": "E23", "Março": "F23", "Abril": "D60", "Maio": "E60", "Junho": "F60",
           "Julho": "D97", "Agosto": "E97", "Setembro": "F97", "Outubro": "D134", "Novembro": "E134", "Dezembro": "F134"},
    266:  {"Janeiro": "D31", "Fevereiro": "E31", "Março": "F31", "Abril": "D68", "Maio": "E68", "Junho": "F68",
           "Julho": "D105", "Agosto": "E105", "Setembro": "F105", "Outubro": "D142", "Novembro": "E142", "Dezembro": "F142"},
    2079: {"Janeiro": "D39", "Fevereiro": "E39", "Março": "F39", "Abril": "D76", "Maio": "E76", "Junho": "F76",
           "Julho": "D113", "Agosto": "E113", "Setembro": "F113", "Outubro": "D150", "Novembro": "E150", "Dezembro": "F150"},
    2849: {"Janeiro": "D41", "Fevereiro": "E41", "Março": "F41", "Abril": "D78", "Maio": "E78", "Junho": "F78",
           "Julho": "D115", "Agosto": "E115", "Setembro": "F115", "Outubro": "D152", "Novembro": "E152", "Dezembro": "F152"}
}

# --- Mapeamento LUCRO PRESUMIDO (exemplo; ajuste conforme sua imagem) ---
mapeamento_acomp_lucro_presumido = {
    994: {
        "Janeiro":   "D17", "Fevereiro": "E17", "Março":     "F17",
        "Abril":     "D56", "Maio":      "E56", "Junho":     "F56",
        "Julho":     "D95", "Agosto":    "E95", "Setembro":  "F95",
        "Outubro":   "D134","Novembro":  "E134","Dezembro":  "F134"
    },
    1022: {
        "Janeiro":   "D18", "Fevereiro": "E18", "Março":     "F18",
        "Abril":     "D57", "Maio":      "E57", "Junho":     "F57",
        "Julho":     "D96", "Agosto":    "E96", "Setembro":  "F96",
        "Outubro":   "D135","Novembro":  "E135","Dezembro":  "F135"
    },
    1176: {
        "Janeiro":   "D19", "Fevereiro": "E19", "Março":     "F19",
        "Abril":     "D58", "Maio":      "E58", "Junho":     "F58",
        "Julho":     "D97", "Agosto":    "E97", "Setembro":  "F97",
        "Outubro":   "D136","Novembro":  "E136","Dezembro":  "F136"
    },
    1085: {
        "Janeiro":   "D20", "Fevereiro": "E20", "Março":     "F20",
        "Abril":     "D59", "Maio":      "E59", "Junho":     "F59",
        "Julho":     "D98", "Agosto":    "E98", "Setembro":  "F98",
        "Outubro":   "D137","Novembro":  "E137","Dezembro":  "F137"
    },
    1197: {
        "Janeiro":   "D21", "Fevereiro": "E21", "Março":     "F21",
        "Abril":     "D60", "Maio":      "E60", "Junho":     "F60",
        "Julho":     "D99", "Agosto":    "E99", "Setembro":  "D99",
        "Outubro":   "D138","Novembro":  "E138","Dezembro":  "F138"
    },
    3276: {
        "Janeiro":   "D22", "Fevereiro": "E22", "Março":     "F22",
        "Abril":     "D61", "Maio":      "E61", "Junho":     "F61",
        "Julho":     "D100", "Agosto":    "E100", "Setembro":  "F100",
        "Outubro":   "D139","Novembro":  "E139","Dezembro":  "F139"
    },
    266: {
        "Janeiro":   "D31", "Fevereiro": "E31", "Março":     "F31",
        "Abril":     "D70", "Maio":      "E70", "Junho":     "F70",
        "Julho":     "D109","Agosto":    "E109","Setembro":  "F109",
        "Outubro":   "D148","Novembro":  "E148","Dezembro":  "F148"
    },
    2079: {
        "Janeiro":   "D39", "Fevereiro": "E39", "Março":     "F39",
        "Abril":     "D78", "Maio":      "E78", "Junho":     "F78",
        "Julho":     "D117","Agosto":    "E117","Setembro":  "F117",
        "Outubro":   "D156","Novembro":  "E156","Dezembro":  "F156"
    },
    2849: {
        "Janeiro":   "D41", "Fevereiro": "E41", "Março":     "F41",
        "Abril":     "D80", "Maio":      "E80", "Junho":     "F80",
        "Julho":     "D119","Agosto":    "E119","Setembro":  "F119",
        "Outubro":   "D158","Novembro":  "E158","Dezembro":  "F158"
    }
}

mapeamento_acomp_lucro_real_estimativo = {
    994: {
        "Janeiro":   "C8",  "Fevereiro": "D8",  "Março":     "E8",
        "Abril":     "F8",  "Maio":      "G8",  "Junho":     "H8",
        "Julho":     "I8",  "Agosto":    "J8",  "Setembro":  "K8",
        "Outubro":   "L8",  "Novembro":  "M8",  "Dezembro":  "N8"
    },
    1022: {
        "Janeiro":   "C9",  "Fevereiro": "D9",  "Março":     "E9",
        "Abril":     "F9",  "Maio":      "G9",  "Junho":     "H9",
        "Julho":     "I9",  "Agosto":    "J9",  "Setembro":  "K9",
        "Outubro":   "L9",  "Novembro":  "M9",  "Dezembro":  "N9"
    },
    1085: {
        "Janeiro":   "C10", "Fevereiro": "D10", "Março":     "E10",
        "Abril":     "F10", "Maio":      "G10", "Junho":     "H10",
        "Julho":     "I10", "Agosto":    "J10", "Setembro":  "K10",
        "Outubro":   "L10", "Novembro":  "M10", "Dezembro":  "N10"
    },
    1176: {
        "Janeiro":   "C11", "Fevereiro": "D11", "Março":     "E11",
        "Abril":     "F11", "Maio":      "G11", "Junho":     "H11",
        "Julho":     "I11", "Agosto":    "J11", "Setembro":  "K11",
        "Outubro":   "L11", "Novembro":  "M11", "Dezembro":  "N11"
    },
    5139: {
        "Janeiro":   "C12", "Fevereiro": "D12", "Março":     "E12",
        "Abril":     "F12", "Maio":      "G12", "Junho":     "H12",
        "Julho":     "I12", "Agosto":    "J12", "Setembro":  "K12",
        "Outubro":   "L12", "Novembro":  "M12", "Dezembro":  "N12"
    },
    1197: {
        "Janeiro":   "C13", "Fevereiro": "D13", "Março":     "E13",
        "Abril":     "F13", "Maio":      "G13", "Junho":     "H13",
        "Julho":     "I13", "Agosto":    "J13", "Setembro":  "K13",
        "Outubro":   "L13", "Novembro":  "M13", "Dezembro":  "N13"
    },
    3276: {
        "Janeiro":   "C14", "Fevereiro": "D14", "Março":     "E14",
        "Abril":     "F14", "Maio":      "G14", "Junho":     "H14",
        "Julho":     "I14", "Agosto":    "J14", "Setembro":  "K14",
        "Outubro":   "L14", "Novembro":  "M14", "Dezembro":  "N14"
    }
}



# Meses finais de cada trimestre (para o cálculo 1785 - 1197)
MESES_FINAIS_TRIMESTRE = ["Março", "Junho", "Setembro", "Dezembro"]

# Lista completa de meses
months_list = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# Mapeia cabeçalhos no formato "MM/2024" -> nome do mês
col_to_month = {
    "01/2024": "Janeiro",  "02/2024": "Fevereiro", "03/2024": "Março",
    "04/2024": "Abril",    "05/2024": "Maio",      "06/2024": "Junho",
    "07/2024": "Julho",    "08/2024": "Agosto",    "09/2024": "Setembro",
    "10/2024": "Outubro",  "11/2024": "Novembro",  "12/2024": "Dezembro"
}

def extrair_dados_balancete(balancete_file):
    """
    Lê a planilha 'Balancete' e converte a coluna 'Código' para numérico.
    Aceita .xls (xlrd) e .xlsx (openpyxl).
    """
    extension = balancete_file.name.split('.')[-1].lower()
    if extension == 'xls':
        # Necessário xlrd==1.2.0 para ler .xls
        df = pd.read_excel(balancete_file, sheet_name='Balancete', engine='xlrd')
    else:
        df = pd.read_excel(balancete_file, sheet_name='Balancete', engine='openpyxl')
    
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
    return df

# -----------------------------------------------------
# INÍCIO DA INTERFACE STREAMLIT
# -----------------------------------------------------
st.title("Processamento de Balancetes - Múltiplos Modelos")

# Agora, oferecemos três opções:
opcao_regime = st.radio(
    "Selecione o Modelo/Regime:",
    ["Lucro Real", "Lucro Presumido", "Lucro Real Estimativo"]
)

# Campos para inserir Nome da Empresa e CNPJ
nome_empresa = st.text_input("Nome da Empresa")
cnpj_empresa = st.text_input("CNPJ da Empresa")

# Permite arquivos .xls e .xlsx
balancete_files = st.file_uploader(
    "Faça upload dos arquivos de Balancete", 
    type=['xls', 'xlsx'], 
    accept_multiple_files=True
)
# Modelo de planilha
modelo_file = st.file_uploader("Faça upload do modelo de planilha (apenas .xlsx)", type=['xlsx'])

if st.button("Processar"):
    if balancete_files and modelo_file and nome_empresa and cnpj_empresa:
        try:
            # Escolhe o dicionário de mapeamento e a aba-alvo
            if opcao_regime == "Lucro Real":
                mapeamento_acomp = mapeamento_acomp_lucro_real
                sheet_name = "Acomp.Resultado_2024"
                usar_ajuste_1785 = True  # Se ainda quiser manter a lógica 1785→1197
            elif opcao_regime == "Lucro Presumido":
                mapeamento_acomp = mapeamento_acomp_lucro_presumido
                sheet_name = "Acomp.Resultado_2024"
                usar_ajuste_1785 = True  # Se ainda quiser manter a lógica 1785→1197
            else:  # Lucro Real Estimativo
                mapeamento_acomp = mapeamento_acomp_lucro_real_estimativo
                sheet_name = "Pós retificação Final"
                usar_ajuste_1785 = False  # Não aplicar lógica 1785→1197
            
            # Cria lista de códigos a serem acumulados (chaves do mapeamento)
            acomp_codes = list(mapeamento_acomp.keys())
            
            # Inicializa os acumuladores de valores
            acomp_values = {codigo: {mes: 0 for mes in months_list} for codigo in acomp_codes}
            code1785_values = {mes: 0 for mes in months_list}  # Usado apenas se usar_ajuste_1785 for True

            # Lê cada arquivo e acumula
            for balancete_file in balancete_files:
                df = extrair_dados_balancete(balancete_file)
                st.write(f"Dados extraídos do arquivo {balancete_file.name}:")
                st.write(df)
                
                for _, row in df.iterrows():
                    codigo = row['Código']
                    if pd.isna(codigo):
                        continue
                    for col in df.columns:
                        if col in col_to_month:  # coluna de mês
                            mes_nome = col_to_month[col]
                            valor = row[col]
                            valor = abs(valor) if pd.notna(valor) else 0
                            
                            if codigo in acomp_codes:
                                acomp_values[codigo][mes_nome] += valor
                            if usar_ajuste_1785 and codigo == 1785:
                                code1785_values[mes_nome] += valor
            
            # Ajuste 1785→1197 somente se usar_ajuste_1785 == True
            if usar_ajuste_1785 and 1197 in acomp_codes:
                MESES_FINAIS_TRIMESTRE = ["Março", "Junho", "Setembro", "Dezembro"]
                for mes in MESES_FINAIS_TRIMESTRE:
                    novo_valor = acomp_values[1197][mes] - code1785_values[mes]
                    st.write(f"Ajuste 1197 ({mes}): 1197({acomp_values[1197][mes]}) -  1785({code1785_values[mes]}) = {novo_valor}")
                    acomp_values[1197][mes] = novo_valor

            # Carrega a planilha modelo
            workbook = openpyxl.load_workbook(modelo_file)

            # Verifica se a aba existe
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"A aba '{sheet_name}' não foi encontrada na planilha modelo.")
            
            sheet_acomp = workbook[sheet_name]
            
            # Preenche as células
            for codigo, mapping in mapeamento_acomp.items():
                for mes, celula in mapping.items():
                    valor = acomp_values[codigo][mes]
                    st.write(f"[{opcao_regime}] Preenchendo célula {celula} com {valor} (cód {codigo}, {mes})")
                    sheet_acomp[celula].value = valor

            # Preenche nome e CNPJ
            # Ajuste as células conforme seu template. Exemplo:
            sheet_acomp["F7"].value = nome_empresa
            sheet_acomp["F8"].value = cnpj_empresa

            # Salva em buffer
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)

            # Nome final do arquivo
            file_name = f"Acompanhamento de Resultado - {nome_empresa}.xlsx"
            st.success("Processamento concluído com sucesso!")
            st.download_button(
                label="Baixar Arquivo Processado",
                data=output,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
    else:
        st.error("Por favor, carregue os arquivos de balancete, o modelo de planilha e preencha o Nome/CNPJ da Empresa.")