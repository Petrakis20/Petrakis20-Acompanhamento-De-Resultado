import streamlit as st
import pandas as pd
import openpyxl
import io

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

# Lista dos meses (usada como chave para acumulação)
months_list = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Mapeia cabeçalhos no formato "MM/2024" para o nome do mês
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

# Lista dos códigos que serão preenchidos na aba Acomp.Resultado_2024
# (conforme mapeamento_acomp; note que o código 1785 não está aqui)
acomp_codes = list(mapeamento_acomp.keys())

# Inicializa os acumuladores de valores para cada código e mês
acomp_values = {codigo: {mes: 0 for mes in months_list} for codigo in acomp_codes}
# Inicializa um acumulador para os valores do código 1785 (que será utilizado para os meses específicos)
code1785_values = {mes: 0 for mes in months_list}

def extrair_dados_balancete(balancete_path):
    """
    Lê a planilha 'Balancete' e normaliza os nomes das colunas.
    Espera-se que haja uma coluna 'Código' e colunas com cabeçalhos no formato "MM/2024".
    """
    df = pd.read_excel(balancete_path, sheet_name='Balancete dinâmico', engine='openpyxl')
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    # Converte a coluna 'Código' para numérico
    df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
    return df

st.title("Processamento de Balancetes - Mês nas Colunas (Formato MM/2024)")

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
                        # Se a coluna tiver um cabeçalho no formato "MM/2024"
                        if col in col_to_month:
                            month_name = col_to_month[col]
                            valor = row[col]
                            if pd.notna(valor):
                                valor = abs(valor)
                            else:
                                valor = 0
                            # Acumula os valores para os códigos que estão no mapeamento
                            if codigo in acomp_codes:
                                acomp_values[codigo][month_name] += valor
                            # Acumula os valores para o código 1785 (mesmo que não esteja em acomp_codes)
                            if codigo == 1785:
                                code1785_values[month_name] += valor
            
            # Para os meses finais de cada trimestre, atualiza o valor do código 1197
            for mes in ["Março", "Junho", "Setembro", "Dezembro"]:
                # Novo valor = (valor do código 1785) - (valor do código 1197 acumulado)
                novo_valor = acomp_values[1197][mes] - code1785_values[mes]
                st.write(f"Atualizando código 1197 para {mes}:  {acomp_values[1197][mes]} - {code1785_values[mes]} = {novo_valor}")
                acomp_values[1197][mes] = novo_valor
            
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

            # Insere Nome da Empresa e CNPJ nas células F7 e F8, respectivamente
            sheet_acomp["F7"].value = nome_empresa
            sheet_acomp["F8"].value = cnpj_empresa

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
