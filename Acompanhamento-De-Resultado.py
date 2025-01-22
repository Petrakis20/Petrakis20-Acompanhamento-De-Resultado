import streamlit as st
import pandas as pd
import openpyxl
import re

# Função de validação de e-mails
def validar_email(email):
    return email.endswith("@jcacontadores.com.br")

# Função de autenticação
def autenticar_usuario(email, senha, usuarios):
    if email in usuarios and usuarios[email] == senha:
        return True
    return False

# Lista de usuários e senhas (recomenda-se usar hashes para segurança em produção)
usuarios = {
    "admin@jcacontadores.com.br": "admin123",
    "user@jcacontadores.com.br": "user123",
    "felipe.reis@jcacontadores.com.br": "felipe123",
}

# Tela de login
st.title("Sistema de Processamento de Planilhas - JCA Contadores")
st.sidebar.title("Login")
email = st.sidebar.text_input("E-mail")
senha = st.sidebar.text_input("Senha", type="password")
login_button = st.sidebar.button("Login")

# Autenticação
if login_button:
    if not validar_email(email):
        st.sidebar.error("Apenas e-mails do domínio jcacontadores.com.br são permitidos.")
    elif autenticar_usuario(email, senha, usuarios):
        st.sidebar.success("Login realizado com sucesso!")
        
        # Sistema principal
        st.title("Acompanhamento de Resultado")

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
            meses = [
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            ]
            for mes in meses:
                if mes.lower() in nome_arquivo.lower():
                    return mes
            raise ValueError("Mês não identificado no nome do arquivo.")

        def extrair_dados_balancete(balancete_path):
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

        def preencher_planilha_modelo(balancete_data, modelo_path, output_path, mes):
            workbook = openpyxl.load_workbook(modelo_path)
            sheet = workbook['Acomp.Resultado_2024']
            for _, row in balancete_data.iterrows():
                codigo = row['Código']
                valor = row['Valor Final']
                if codigo in mapeamento and mes in mapeamento[codigo]:
                    celula = mapeamento[codigo][mes]
                    sheet[celula].value = valor
            workbook.save(output_path)

        balancete_file = st.file_uploader("Faça upload do arquivo Balancete", type=['xlsx'])
        modelo_file = st.file_uploader("Faça upload do modelo de planilha", type=['xlsx'])

        if st.button("Processar"):
            if balancete_file and modelo_file:
                try:
                    mes = identificar_mes(balancete_file.name)
                    balancete_data = extrair_dados_balancete(balancete_file)
                    output_path = 'modelo_preenchido.xlsx'
                    preencher_planilha_modelo(balancete_data, modelo_file, output_path, mes)
                    st.success("Processamento concluído com sucesso!")
                    with open(output_path, "rb") as file:
                        st.download_button(label="Baixar Arquivo Processado", data=file, file_name="modelo_preenchido.xlsx")
                except Exception as e:
                    st.error(f"Erro no processamento: {e}")
            else:
                st.error("Por favor, carregue ambos os arquivos antes de processar.")
    else:
        st.sidebar.error("Credenciais inválidas. Por favor, tente novamente.")
