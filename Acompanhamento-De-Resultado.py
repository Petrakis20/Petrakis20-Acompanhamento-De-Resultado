import streamlit as st
import pandas as pd
import openpyxl
import xlrd
import io

# Monkey-patch para "enganar" o Pandas quanto à versão do xlrd
if xlrd.__version__ == "1.2.0":
    xlrd.__version__ = "2.0.1"

from mapeamentos import (
    mapeamento_acomp_lucro_real,
    mapeamento_acomp_lucro_presumido,
    mapeamento_extra_presumido,
    mapeamento_acomp_lucro_real_estimativo,
    month_to_quarter,
    MESES_FINAIS_TRIMESTRE,
    months_list,
    col_to_month,
)

def extrair_dados_balancete(balancete_file):
    extension = balancete_file.name.split('.')[-1].lower()
    if extension == 'xls':
        df = pd.read_excel(balancete_file, sheet_name='Balancete', engine='xlrd')
    else:
        df = pd.read_excel(balancete_file, sheet_name='Balancete', engine='openpyxl')
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
    return df

st.image("logo.png", width=100)
st.title("Acompanhamento de Resultado JCA Contadores")

# Botão para baixar o Manual em PDF (botão estilizado em vermelho)
st.markdown(
    """
    <style>
    /* Alvo do botão de download do manual */
    div.stDownloadButton > button {
         background-color: #FF0000;
         color: white;
         border: 1px solid red;
         padding: 8px 16px;
         font-size: 16px;
         cursor: pointer;
         border-radius: 4px;
    }
    div.stDownloadButton > button:hover {
        color: #F00;
        background-color: #FFF;
    }
    </style>
    """,
    unsafe_allow_html=True
)
try:
    with open("./Passo a Passo Acompanhamento de resultado.pdf", "rb") as pdf_file:
        manual_data = pdf_file.read()
    st.download_button(
        label="Manual",
        data=manual_data,
        file_name="Passo a Passo Acompanhamento de resultado.pdf",
        mime="application/pdf"
    )
except Exception as e:
    st.error("Manual não disponível no momento.")

# Novo campo para selecionar o ano da operação
ano_operacao = st.radio("Selecione o Ano da Operação:", ["2024", "2025"])

opcao_regime = st.radio(
    "Selecione o Modelo/Regime:",
    ["Lucro Real", "Lucro Presumido", "Lucro Real Estimativo"]
)

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

balancete_files = st.file_uploader(
    "Faça upload dos arquivos de Balancete", type=['xls', 'xlsx'], accept_multiple_files=True
)
modelo_file = st.file_uploader("Faça upload do modelo de planilha (apenas .xlsx)", type=['xlsx'])

if st.button("Processar"):
    if balancete_files and modelo_file and nome_empresa and cnpj_empresa:
        try:
            # Define variáveis de mapeamento e aba conforme o regime escolhido
            if opcao_regime == "Lucro Real":
                mapeamento_acomp = mapeamento_acomp_lucro_real
                usar_ajuste_1785 = True   # Aplica ajuste somente para Lucro Real
                processar_extra = False
            elif opcao_regime == "Lucro Presumido":
                mapeamento_acomp = mapeamento_acomp_lucro_presumido
                usar_ajuste_1785 = False  # Não aplica ajuste para Lucro Presumido
                processar_extra = True   # Processa os códigos extras
            else:  # Lucro Real Estimativo
                mapeamento_acomp = mapeamento_acomp_lucro_real_estimativo
                usar_ajuste_1785 = False
                processar_extra = False

            # Independente do regime, a sheet a ser aberta dependerá do ano selecionado
            if ano_operacao == "2024":
                sheet_name = "Acomp.Resultado_2024"
            else:
                sheet_name = "Acomp.Resultado_2025"

            acomp_codes = list(mapeamento_acomp.keys())
            acomp_values = {codigo: {mes: 0 for mes in months_list} for codigo in acomp_codes}
            code1785_values = {mes: 0 for mes in months_list}
            extra_values = {
                "1015_1981": {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0},
                1043: {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0},
                2919: {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0},
                4429: {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0},
                1904: {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0},
                6196: {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0}
            }

            for balancete_file in balancete_files:
                df = extrair_dados_balancete(balancete_file)
                st.write(f"Dados extraídos do arquivo {balancete_file.name}:")
                st.write(df)
                for _, row in df.iterrows():
                    codigo = row['Código']
                    if pd.isna(codigo):
                        continue
                    for col in df.columns:
                        if col in col_to_month:
                            mes_nome = col_to_month[col]
                            valor = row[col]
                            valor = abs(valor) if pd.notna(valor) else 0
                            if codigo in acomp_codes:
                                acomp_values[codigo][mes_nome] += valor
                            if usar_ajuste_1785 and codigo == 1785:
                                code1785_values[mes_nome] += valor
                            if processar_extra:
                                qtr = month_to_quarter[mes_nome]
                                if codigo in [1015, 1981]:
                                    extra_values["1015_1981"][qtr] += valor
                                elif codigo in [1043, 2919, 4429, 1904, 6196]:
                                    extra_values[codigo][qtr] += valor

            # Ajuste para o código 1197: Inverter a operação para evitar números negativos,
            # aplicado somente para Lucro Real
            if usar_ajuste_1785 and 1197 in acomp_codes:
                for mes in ["Março", "Junho", "Setembro", "Dezembro"]:
                    novo_valor = acomp_values[1197][mes] - code1785_values[mes]
                    st.write(f"Ajuste 1197 ({mes}): {acomp_values[1197][mes]} - 1785({code1785_values[mes]}) = {novo_valor}")
                    acomp_values[1197][mes] = novo_valor

            workbook = openpyxl.load_workbook(modelo_file)
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"A aba '{sheet_name}' não foi encontrada na planilha modelo.")
            sheet_acomp = workbook[sheet_name]
            for codigo, mapping in mapeamento_acomp.items():
                for mes, celula in mapping.items():
                    valor = acomp_values[codigo].get(mes, 0)
                    st.write(f"[{opcao_regime}] Preenchendo célula {celula} com {valor} (cód {codigo}, {mes})")
                    sheet_acomp[celula].value = valor

            if opcao_regime == "Lucro Presumido" and processar_extra:
                for key, mapping in mapeamento_extra_presumido.items():
                    for trimestre, celula in mapping.items():
                        valor = extra_values[key][trimestre]
                        st.write(f"[Lucro Presumido Extra] Preenchendo célula {celula} com {valor} (chave {key}, {trimestre})")
                        sheet_acomp[celula].value = valor

            sheet_acomp["F7"].value = nome_empresa
            sheet_acomp["F8"].value = cnpj_empresa

            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)
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
