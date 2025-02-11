import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import streamlit_authenticator as stauth

# Mapeamento para a aba "Acomp.Resultado_2024"
mapeamento_acomp = {
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
        "Dezembro": "F128"
    },
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
        "Dezembro": "F129"
    },
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
        "Dezembro": "F130"
    },
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
        "Dezembro": "F131"
    },
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
        "Dezembro": "F132"
    },
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
        "Dezembro": "F133"
    },
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
        "Dezembro": "F134"
    },
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
        "Dezembro": "F142"
    },
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
        "Dezembro": "F150"
    },
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
        "Dezembro": "F152"
    },
    231: {
        "Janeiro": "J29",
        "Abril": "J66",
        "Julho": "J103",
        "Outubro": "J140",
    },
}

# Mapeamento para a aba "Adições 2024" (novos códigos)
mapeamento_adicoes = {
    6250: {
        "Janeiro": "C10", "Fevereiro": "D10", "Março": "E10",
        "Abril": "C24", "Maio": "D24", "Junho": "E24",
        "Julho": "C38", "Agosto": "D38", "Setembro": "E38",
        "Outubro": "C52", "Novembro": "D52", "Dezembro": "E52"
    },
    6109: {
        "Janeiro": "C11", "Fevereiro": "D11", "Março": "E11",
        "Abril": "C25", "Maio": "D25", "Junho": "E25",
        "Julho": "C39", "Agosto": "D39", "Setembro": "E39",
        "Outubro": "C53", "Novembro": "D53", "Dezembro": "E53"
    },
    3325: {
        "Janeiro": "C12", "Fevereiro": "D12", "Março": "E12",
        "Abril": "C26", "Maio": "D26", "Junho": "E26",
        "Julho": "C40", "Agosto": "D40", "Setembro": "E40",
        "Outubro": "C54", "Novembro": "D54", "Dezembro": "E54"
    },
    6257: {
        "Janeiro": "C13", "Fevereiro": "D13", "Março": "E13",
        "Abril": "C27", "Maio": "D27", "Junho": "E27",
        "Julho": "C41", "Agosto": "D41", "Setembro": "E41",
        "Outubro": "C55", "Novembro": "D55", "Dezembro": "E55"
    },
    6119: {
        "Janeiro": "C14", "Fevereiro": "D14", "Março": "E14",
        "Abril": "C28", "Maio": "D28", "Junho": "E28",
        "Julho": "C42", "Agosto": "D42", "Setembro": "E42",
        "Outubro": "C56", "Novembro": "D56", "Dezembro": "E56"
    },
}

def identificar_mes(nome_arquivo):
    """
    Identifica o mês no nome do arquivo.
    Para garantir a compatibilidade com os mapeamentos, retornamos os nomes com acentuação.
    """
    mapping_meses = {
        "janeiro": "Janeiro",
        "fevereiro": "Fevereiro",
        "março": "Março",
        "marco": "Março",
        "abril": "Abril",
        "maio": "Maio",
        "junho": "Junho",
        "julho": "Julho",
        "agosto": "Agosto",
        "setembro": "Setembro",
        "outubro": "Outubro",
        "novembro": "Novembro",
        "dezembro": "Dezembro"
    }
    nome_lower = nome_arquivo.lower()
    for key, mes in mapping_meses.items():
        if key in nome_lower:
            return mes
    return None

def extrair_dados_balancete(balancete_path):
    """Extrai os dados relevantes do balancete."""
    def extract_code_number(code):
        try:
            return int(code)
        except:
            match = re.search(r'\[(\d+)\]', str(code))
            return int(match.group(1)) if match else None

    df = pd.read_excel(balancete_path, sheet_name='Balancete', engine='openpyxl')
    # Seleciona as colunas necessárias
    colunas = ['Código', 'Movimento', 'Saldo Atual']
    if "Débito" in df.columns:
        colunas.append("Débito")
    df = df[colunas].dropna(subset=['Código'])
    df['Código'] = df['Código'].apply(extract_code_number)
    df['Movimento'] = df['Movimento'].abs()
    # Para os códigos diferentes de 266, usa a regra atual
    df['Valor Final'] = df.apply(
        lambda row: row['Saldo Atual'] if row['Código'] == 266 else row['Movimento'], axis=1
    )
    return df

def preencher_acomp(balancete_data, workbook, mes):
    """Preenche a aba 'Acomp.Resultado_2024' para os códigos (exceto 231 e os de Adições)."""
    sheet_name = 'Acomp.Resultado_2024'
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"A aba '{sheet_name}' não foi encontrada na planilha modelo.")
    sheet = workbook[sheet_name]
    for _, row in balancete_data.iterrows():
        codigo = row['Código']
        if codigo in mapeamento_acomp and mes in mapeamento_acomp[codigo]:
            celula = mapeamento_acomp[codigo][mes]
            st.write(f"Preenchendo célula {celula} com o valor {row['Valor Final']} para o código {codigo}")
            sheet[celula].value = row['Valor Final']

# Interface de processamento
st.title("Interface Interativa para Processamento de Planilhas")

balancete_files = st.file_uploader("Faça upload dos arquivos de Balancete", type=['xlsx'], accept_multiple_files=True)
modelo_file = st.file_uploader("Faça upload do modelo de planilha", type=['xlsx'])

if st.button("Processar"):
    if balancete_files and modelo_file:
        try:
            workbook = openpyxl.load_workbook(modelo_file)
            
            # Dicionário para acumular os valores do código 231 (usando a coluna Débito) por trimestre
            soma_231 = {"Janeiro": 0, "Abril": 0, "Julho": 0, "Outubro": 0}
            
            # Para os códigos da aba Adições 2024
            codigos_adicoes = [6250, 6109, 3325, 6257, 6119]
            todos_os_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                              "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
            adicoes_valores = {codigo: {mes: 0 for mes in todos_os_meses} for codigo in codigos_adicoes}
            
            # Processa cada arquivo de balancete
            for balancete_file in balancete_files:
                mes = identificar_mes(balancete_file.name)
                if mes:
                    st.write(f"Processando o balancete: {balancete_file.name} para o mês: {mes}")
                    df_balancete = extrair_dados_balancete(balancete_file)
                    st.write(f"Dados extraídos do balancete ({mes}):")
                    st.write(df_balancete)
                    
                    # Código 231: soma os valores da coluna 'Débito'
                    df_231 = df_balancete[df_balancete['Código'] == 231]
                    if not df_231.empty and "Débito" in df_231.columns:
                        debito_total = df_231["Débito"].sum()
                        # Define o trimestre de acordo com o mês
                        if mes in ["Janeiro", "Fevereiro", "Março"]:
                            trimestre = "Janeiro"
                        elif mes in ["Abril", "Maio", "Junho"]:
                            trimestre = "Abril"
                        elif mes in ["Julho", "Agosto", "Setembro"]:
                            trimestre = "Julho"
                        elif mes in ["Outubro", "Novembro", "Dezembro"]:
                            trimestre = "Outubro"
                        else:
                            trimestre = None
                        if trimestre:
                            soma_231[trimestre] += debito_total
                    
                    # Processa os demais códigos para a aba Acomp (exceto 231 e os de Adições)
                    df_outros = df_balancete[~df_balancete['Código'].isin([231] + codigos_adicoes)]
                    preencher_acomp(df_outros, workbook, mes)
                    
                    # Processa os códigos para Adições 2024 (acumula os valores da coluna 'Movimento')
                    df_adicoes = df_balancete[df_balancete['Código'].isin(codigos_adicoes)]
                    if not df_adicoes.empty:
                        for _, row in df_adicoes.iterrows():
                            codigo = row['Código']
                            valor = row['Movimento']
                            adicoes_valores[codigo][mes] += valor
                else:
                    st.warning(f"Mês não identificado no arquivo: {balancete_file.name}")
            
            # Preenche a aba Acomp para o código 231
            sheet_acomp = workbook['Acomp.Resultado_2024']
            if 231 in mapeamento_acomp:
                for trimestre, soma in soma_231.items():
                    celula = mapeamento_acomp[231].get(trimestre)
                    if celula:
                        st.write(f"Preenchendo célula {celula} com a soma {soma} para o código 231 no trimestre iniciado em {trimestre}")
                        sheet_acomp[celula].value = soma
            
            # Preenche a aba "Adições 2024" para os códigos 6250, 6109, 3325, 6257, 6119
            if "Adições 2024" not in workbook.sheetnames:
                raise ValueError("A aba 'Adições 2024' não foi encontrada na planilha modelo.")
            sheet_adicoes = workbook["Adições 2024"]
            for codigo, mapping in mapeamento_adicoes.items():
                for mes, celula in mapping.items():
                    valor = adicoes_valores[codigo][mes]
                    st.write(f"Preenchendo célula {celula} com o valor {valor} para o código {codigo} no mês {mes}")
                    sheet_adicoes[celula].value = valor
            
            # Salva o arquivo processado em um buffer para download
            output = io.BytesIO()
            workbook.save(output)
            output.seek(0)
            
            st.success("Processamento concluído com sucesso para todos os balancetes!")
            st.download_button(
                label="Baixar Arquivo Processado",
                data=output,
                file_name="modelo_preenchido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
    else:
        st.error("Por favor, carregue os arquivos de balancete e o modelo de planilha antes de processar.")
