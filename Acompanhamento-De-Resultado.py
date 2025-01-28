import streamlit as st
import pandas as pd
import openpyxl
import re
import streamlit_authenticator as stauth

# Mapeamento de células para cada código e mês
mapeamento = {
994: {
        "Janeiro": "D17", 
        "Fevereiro": "E17", 
        "Março": "F17", 
        "Abril": "D54", 
        "Maio": "E54", 
        "Junho": "F54",
        "Julho": "D91", 
        "Agosto": "E91", 
        "Setembro": "F91", 
        "Outubro": "D128", 
        "Novembro": "E128", 
        "Dezembro": "F128"},

1022: {
        "Janeiro": "D18", 
        "Fevereiro": "E18", 
        "Março": "F18", 
        "Abril": "D55", 
        "Maio": "E55", 
        "Junho": "F55",
        "Julho": "D92",
        "Agosto": "E92", 
        "Setembro": "F92", 
        "Outubro": "D129", 
        "Novembro": "E129", 
        "Dezembro": "F129"},

1085: {
        "Janeiro": "D19", 
        "Fevereiro": "E19", 
        "Março": "F19", 
        "Abril": "D56", 
        "Maio": "E56", 
        "Junho": "F56",
        "Julho": "D93", 
        "Agosto": "E93", 
        "Setembro": "F93", 
        "Outubro": "D130", 
        "Novembro": "E130", 
        "Dezembro": "F130"},

1176: {
        "Janeiro": "D20", 
        "Fevereiro": "E20", 
        "Março": "F20", 
        "Abril": "D57", 
        "Maio": "E57", 
        "Junho": "F57",
        "Julho": "D94", 
        "Agosto": "E94", 
        "Setembro": "F94", 
        "Outubro": "D131", 
        "Novembro": "E131", 
        "Dezembro": "F131"},

5139: {
    "Janeiro": "D21",
    "Fevereiro": "E21",
    "Março": "F21",
    "Abril": "D58",
    "Maio": "E58",
    "Junho": "F58",
    "Julho": "D95",
    "Agosto": "E95",
    "Setembro": "F95",
    "Outubro": "D132",
    "Novembro": "E132",
    "Dezembro": "F132"},

1197: {
    "Janeiro": "D22",
    "Fevereiro": "E22",
    "Março": "F22",
    "Abril": "D59",
    "Maio": "E59",
    "Junho": "F59",
    "Julho": "D96",
    "Agosto": "E96",
    "Setembro": "F96",
    "Outubro": "D133",
    "Novembro": "E133",
    "Dezembro": "F133"},

3276: {
    "Janeiro": "D23",
    "Fevereiro": "E23",
    "Março": "F23",
    "Abril": "D60",
    "Maio": "E60",
    "Junho": "F60",
    "Julho": "D97",
    "Agosto": "E97",
    "Setembro": "F97",
    "Outubro": "D134",
    "Novembro": "E134",
    "Dezembro": "F134"},

266: {
    "Janeiro": "D31",
    "Fevereiro": "E31",
    "Março": "F31",
    "Abril": "D68",
    "Maio": "E68",
    "Junho": "F68",
    "Julho": "D105",
    "Agosto": "E105",
    "Setembro": "F105",
    "Outubro": "D142",
    "Novembro": "E142",
    "Dezembro": "F142"},

2079: {
    "Janeiro": "D39",
    "Fevereiro": "E39",
    "Março": "F39",
    "Abril": "D76",
    "Maio": "E76",
    "Junho": "F76",
    "Julho": "D113",
    "Agosto": "E113",
    "Setembro": "F113",
    "Outubro": "D150",
    "Novembro": "E150",
    "Dezembro": "F150"},

2849: {
    "Janeiro": "D41",
    "Fevereiro": "E41",
    "Março": "F41",
    "Abril": "D78",
    "Maio": "E78",
    "Junho": "F78",
    "Julho": "D115",
    "Agosto": "E115",
    "Setembro": "F115",
    "Outubro": "D152",
    "Novembro": "E152",
    "Dezembro": "F152"},

231: {
    "Janeiro": "J29",
    "Abril": "J66",
    "Julho": "J103",
    "Outubro": "J140",},

}

def identificar_mes(nome_arquivo):
    """Identifica o mês no nome do arquivo."""
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    for mes in meses:
        if mes.lower() in nome_arquivo.lower():
            return mes
    return None

def extrair_dados_balancete(balancete_path):
    """Extrai os dados relevantes do balancete."""
    def extract_code_number(code):
        match = re.search(r'\[(\d+)\]', str(code))
        return int(match.group(1)) if match else None

    balancete_data = pd.read_excel(balancete_path, sheet_name='Balancete', engine='openpyxl')
    balancete_data['Código'] = balancete_data['Código'].apply(extract_code_number)
    balancete_data['Movimento'] = balancete_data['Movimento'].abs()
    balancete_data = balancete_data[['Código', 'Movimento', 'Saldo Atual']].dropna()
    balancete_data['Valor Final'] = balancete_data.apply(
        lambda row: row['Saldo Atual'] if row['Código'] == 266 else row['Movimento'], axis=1
    )
    return balancete_data

def preencher_planilha_modelo(balancete_data, workbook, mes):
    """Preenche a planilha modelo para o mês especificado."""
    sheet_name = 'Acomp.Resultado_2024'

    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"A aba '{sheet_name}' não foi encontrada na planilha modelo.")
    
    sheet = workbook[sheet_name]

    for _, row in balancete_data.iterrows():
        codigo = row['Código']
        valor = row['Valor Final']
        if codigo in mapeamento and mes in mapeamento[codigo]:
            celula = mapeamento[codigo][mes]
            st.write(f"Preenchendo célula {celula} com o valor {valor} para o código {codigo}")
            # Garantir que a célula é preenchida corretamente
            sheet[celula].value = valor

# Interface de processamento
st.title("Interface Interativa para Processamento de Planilhas")

balancete_files = st.file_uploader("Faça upload dos arquivos de Balancete", type=['xlsx'], accept_multiple_files=True)
modelo_file = st.file_uploader("Faça upload do modelo de planilha", type=['xlsx'])

if st.button("Processar"):
    if balancete_files and modelo_file:
        try:
            # Carregar o modelo de planilha apenas uma vez
            workbook = openpyxl.load_workbook(modelo_file)

            # Iterar sobre todos os balancetes enviados
            for balancete_file in balancete_files:
                mes = identificar_mes(balancete_file.name)
                if mes:
                    st.write(f"Processando o balancete: {balancete_file.name} para o mês: {mes}")
                    
                    # Verificar e exibir os dados extraídos
                    balancete_data = extrair_dados_balancete(balancete_file)
                    st.write(f"Dados extraídos do balancete ({mes}):")
                    st.write(balancete_data)
                    
                    # Preencher a planilha para o mês identificado
                    preencher_planilha_modelo(balancete_data, workbook, mes)
                else:
                    st.warning(f"Mês não identificado no arquivo: {balancete_file.name}")

            # Salvar o arquivo preenchido após processar todos os balancetes
            output_path = 'modelo_preenchido.xlsx'
            workbook.save(output_path)

            st.success("Processamento concluído com sucesso para todos os balancetes!")
            with open(output_path, "rb") as file:
                st.download_button(label="Baixar Arquivo Processado", data=file, file_name="modelo_preenchido.xlsx")
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
    else:
        st.error("Por favor, carregue os arquivos de balancete e o modelo de planilha antes de processar.")