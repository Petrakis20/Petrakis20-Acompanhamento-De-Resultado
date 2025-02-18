import streamlit as st
import pandas as pd
import openpyxl
import io

# Adicione o caminho para o arquivo da logo
logo_path = "logo JCA.png"

# Mapeamento para a aba "Acomp.Resultado_2024"
mapeamento_acomp = {
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

# Mapeamento para a aba "Adições 2024"
mapeamento_adicoes = {
    6250: {"Janeiro": "C10", "Fevereiro": "D10", "Março": "E10", "Abril": "C24", "Maio": "D24", "Junho": "E24",
           "Julho": "C38", "Agosto": "D38", "Setembro": "E38", "Outubro": "C52", "Novembro": "D52", "Dezembro": "E52"},
    6109: {"Janeiro": "C11", "Fevereiro": "D11", "Março": "E11", "Abril": "C25", "Maio": "D25", "Junho": "E25",
           "Julho": "C39", "Agosto": "D39", "Setembro": "E39", "Outubro": "C53", "Novembro": "D53", "Dezembro": "E53"},
    3325: {"Janeiro": "C12", "Fevereiro": "D12", "Março": "E12", "Abril": "C26", "Maio": "D26", "Junho": "E26",
           "Julho": "C40", "Agosto": "D40", "Setembro": "E40", "Outubro": "C54", "Novembro": "D54", "Dezembro": "E54"},
    6257: {"Janeiro": "C13", "Fevereiro": "D13", "Março": "E13", "Abril": "C27", "Maio": "D27", "Junho": "E27",
           "Julho": "C41", "Agosto": "D41", "Setembro": "E41", "Outubro": "C55", "Novembro": "D55", "Dezembro": "E55"},
    6119: {"Janeiro": "C14", "Fevereiro": "D14", "Março": "E14", "Abril": "C28", "Maio": "D28", "Junho": "E28",
           "Julho": "C42", "Agosto": "D42", "Setembro": "E42", "Outubro": "C56", "Novembro": "D56", "Dezembro": "E56"}
}

# Lista dos meses (utilizada nas chaves dos dicionários de acumulação e mapeamentos)
months_list = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Dicionário que converte o cabeçalho da coluna (formato "MM/2024") para o nome do mês
col_to_month = {
    "01/2024": "Janeiro",
    "02/2024": "Fevereiro",
    "03/2024": "Março",
    "04/2024": "Abril",
    "05/2024": "Maio",
    "06/2024": "Junho",
    "07/2024": "Julho",
    "08/2024": "Agosto",
    "09/2024": "Setembro",
    "10/2024": "Outubro",
    "11/2024": "Novembro",
    "12/2024": "Dezembro"
}

# Listas dos códigos para cada aba
acomp_codes = list(mapeamento_acomp.keys())
adicoes_codes = list(mapeamento_adicoes.keys())

# Inicializa os acumuladores de valores para cada código e mês
acomp_values = {codigo: {mes: 0 for mes in months_list} for codigo in acomp_codes}
adicoes_values = {codigo: {mes: 0 for mes in months_list} for codigo in adicoes_codes}

def extrair_dados_balancete(balancete_path):
    """
    Lê a planilha 'Balancete' e normaliza os nomes das colunas.
    Espera-se que haja uma coluna 'Código' e colunas com cabeçalhos no formato "MM/2024".
    """
    df = pd.read_excel(balancete_path, sheet_name='Balancete dinâmico', engine='openpyxl')
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    # Converte a coluna 'Código' para numérico (assumindo que já está em formato simples)
    df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
    return df

st.image(logo_path, width=150)

st.title("Acompanhamento de Resultados Lucro Real - JCA Contadores")

# Campos para inserir Nome da Empresa e CNPJ
nome_empresa = st.text_input("Nome da Empresa")
cnpj_empresa = st.text_input("CNPJ da Empresa")

balancete_files = st.file_uploader("Faça upload dos arquivos de Balancete", type=['xlsx'], accept_multiple_files=True)
modelo_file = st.file_uploader("Faça upload do modelo de planilha", type=['xlsx'])

if st.button("Processar"):
    if balancete_files and modelo_file and nome_empresa and cnpj_empresa:
        try:
            # Processa cada arquivo de balancete
            for balancete_file in balancete_files:
                df = extrair_dados_balancete(balancete_file)
                st.write(f"Dados extraídos do arquivo {balancete_file.name}:")
                st.write(df)
                # Itera por cada linha e para cada coluna que corresponde a um mês
                for _, row in df.iterrows():
                    codigo = row['Código']
                    if pd.isna(codigo):
                        continue
                    for col in df.columns:
                        if col in col_to_month:
                            month_name = col_to_month[col]
                            valor = row[col]
                            if pd.notna(valor):
                                valor = abs(valor)
                            else:
                                valor = 0
                            if codigo in acomp_codes:
                                acomp_values[codigo][month_name] += valor
                            if codigo in adicoes_codes:
                                adicoes_values[codigo][month_name] += valor

            # Carrega a planilha modelo
            workbook = openpyxl.load_workbook(modelo_file)
            
            # Preenche a aba "Acomp.Resultado_2024"
            if "Acomp.Resultado_2024" not in workbook.sheetnames:
                raise ValueError("A aba 'Acomp.Resultado_2024' não foi encontrada na planilha modelo.")
            sheet_acomp = workbook["Acomp.Resultado_2024"]
            for codigo, mapping in mapeamento_acomp.items():
                for mes, celula in mapping.items():
                    valor = acomp_values[codigo][mes]
                    st.write(f"Acomp.Resultado_2024: Preenchendo célula {celula} com o valor {valor} para o código {codigo} no mês {mes}")
                    sheet_acomp[celula].value = valor

            # Insere Nome da Empresa e CNPJ nas células F7 e F8, respectivamente, na aba Acomp.Resultado_2024
            sheet_acomp["F7"].value = nome_empresa
            sheet_acomp["F8"].value = cnpj_empresa

            # Preenche a aba "Adições 2024"
            if "Adições 2024" not in workbook.sheetnames:
                raise ValueError("A aba 'Adições 2024' não foi encontrada na planilha modelo.")
            sheet_adicoes = workbook["Adições 2024"]
            for codigo, mapping in mapeamento_adicoes.items():
                for mes, celula in mapping.items():
                    valor = adicoes_values[codigo][mes]
                    st.write(f"Adições 2024: Preenchendo célula {celula} com o valor {valor} para o código {codigo} no mês {mes}")
                    sheet_adicoes[celula].value = valor

            # Salva o arquivo processado em um buffer para download
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)
            
            # Nome do arquivo de download conforme solicitado
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
        st.error("Por favor, carregue os arquivos de balancete, o modelo de planilha e preencha o Nome e CNPJ da Empresa.")
