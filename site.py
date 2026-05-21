import streamlit as st
import pdfplumber
import re
import pandas as pd
import numpy as np
from io import BytesIO

# ------------------------------------------------------------
# FUNÇÕES COMPARTILHADAS (cache e processamento RIO PAX)
# ------------------------------------------------------------
@st.cache_data
def extrair_texto_pdfs(uploaded_pdfs):
    texto_completo = ""
    for nome_arquivo, conteudo in uploaded_pdfs.items():
        if nome_arquivo.lower().endswith('.pdf'):
            with pdfplumber.open(BytesIO(conteudo)) as pdf:
                texto_completo += f"\n\n{'='*50}\n"
                texto_completo += f"ARQUIVO: {nome_arquivo}\n"
                texto_completo += f"{'='*50}\n\n"
                for i, pagina in enumerate(pdf.pages):
                    texto = pagina.extract_text(layout=True)
                    if texto:
                        texto_completo += f"===== Página {i+1} =====\n{texto}\n\n"
    return texto_completo

@st.cache_data
def parse_extrato(texto):
    lines = [line.strip() for line in texto.split('\n') if line.strip()]

    data_line_pattern = re.compile(
        r'^(\S+)\s+(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(\d{2}/\d{2}/\d{4})\s+'
        r'(.*?)\s+(\d+\s*/\s*\d+)\s+(.*?)\s+(\S+)\s+(\S+)\s+([\d\.,]+)\s+(\d+)$'
    )
    section_pattern = re.compile(r'^(\w+)\s*-\s*(CREDITO|DEBITO|FATURADO|PIX)')
    total_pattern = re.compile(r'^Total\s+_______|^Total Geral')
    ignore_patterns = [
        re.compile(r'Extrato de caixa'), re.compile(r'Baixa Inicial|Baixa Final'),
        re.compile(r'MS Consultoria'), re.compile(r'Página'),
        re.compile(r'^\d{2}/\d{2}/\d{4}'), re.compile(r'RIO PAX'), re.compile(r'JULIAC'),
    ]

    def is_ignored(line):
        return any(p.search(line) for p in ignore_patterns)

    transactions = []
    current_section = None

    for i, line in enumerate(lines):
        if section_pattern.match(line):
            current_section = line
            continue
        if is_ignored(line) or total_pattern.match(line):
            continue
        match = data_line_pattern.match(line)
        if not match:
            continue

        desc = None
        suffix = None
        if i > 0 and not is_ignored(lines[i-1]) and not data_line_pattern.match(lines[i-1]):
            desc = lines[i-1]
        if i+1 < len(lines) and not is_ignored(lines[i+1]) and not data_line_pattern.match(lines[i+1]):
            suffix = lines[i+1]

        groups = match.groups()
        transaction = {
            'secao': current_section,
            'descricao': desc,
            'sufixo': suffix,
            'deonde': groups[0],
            'dt_baixa': groups[1],
            'usuario': groups[2],
            'dt_pgto': groups[3],
            'cobrador': groups[4].strip(),
            'pacote_os_contrato': groups[5].strip(),
            'nome': groups[6].strip(),
            'referencia': groups[7],
            'documento': groups[8],
            'valor': groups[9].replace('.', '').replace(',', '.'),
            'nfs_e': groups[10],
        }
        transactions.append(transaction)
    return transactions

def separar_pacote_os(df):
    split_df = df['pacote_os_contrato'].str.split(' / ', expand=True)
    split_df.columns = ['contrato', 'os']
    split_df['contrato'] = split_df['contrato'].astype(str).str.lstrip('0')
    split_df['os'] = split_df['os'].astype(str).str.lstrip('0')
    df = df.drop(columns=['pacote_os_contrato'])
    col_index = df.columns.get_loc('cobrador') + 1
    df.insert(col_index, 'contrato', split_df['contrato'])
    df.insert(col_index + 1, 'os', split_df['os'])
    return df

def merge_control_desk(df, control_desk_df):
    col_chave = 'COD EXTERNO'
    col_valor = 'PROCESSO'
    if control_desk_df is not None and col_chave in control_desk_df.columns and col_valor in control_desk_df.columns:
        mapa = control_desk_df.set_index(col_chave)[col_valor].to_dict()
        df['control_desk_info'] = df['os'].map(mapa).fillna('NA')
        st.success("✅ Merge realizado com sucesso!")
    else:
        st.warning(f"Arquivo Control Desk não fornecido ou colunas '{col_chave}'/'{col_valor}' não encontradas. 'control_desk_info' = 'NA'")
        df['control_desk_info'] = 'NA'
    return df

def limpar_df(df):
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.fillna('')
    for col in df.select_dtypes(include=['float', 'int']).columns:
        df[col] = df[col].apply(lambda x: '' if pd.isna(x) else str(x) if not np.isfinite(x) else x)
    return df

# ------------------------------------------------------------
# FUNÇÃO ESPECÍFICA PARA RIO PAX
# ------------------------------------------------------------
def rio_pax_interface():
    st.header(" RIO PAX - Extração e estruturação de extratos PDF")
    
    pdf_files = st.file_uploader("Selecione um ou mais PDFs do RIO PAX", type="pdf", accept_multiple_files=True)
    control_file = st.file_uploader("Arquivo Control Desk (Excel com colunas 'COD EXTERNO' e 'PROCESSO')", type=["xlsx", "xls"])
    
    processar = st.button("🚀 Processar PDFs e gerar Excel", type="primary")
    
    if processar:
        if not pdf_files:
            st.error("Envie pelo menos um arquivo PDF.")
        else:
            pdf_dict = {pdf.name: pdf.read() for pdf in pdf_files}
            
            with st.status("Extraindo texto dos PDFs...", expanded=True) as status:
                texto_consolidado = extrair_texto_pdfs(pdf_dict)
                status.update(label="Texto extraído com sucesso!", state="complete")
            
            with st.spinner("Parseando transações..."):
                transacoes = parse_extrato(texto_consolidado)
                if not transacoes:
                    st.error("Nenhuma transação encontrada. Verifique o formato do extrato.")
                    st.stop()
                df = pd.DataFrame(transacoes)
                st.info(f"Total de transações encontradas: {len(df)}")
            
            with st.spinner("Separando coluna 'pacote_os_contrato'..."):
                df = separar_pacote_os(df)
            
            if control_file:
                with st.spinner("Lendo Control Desk e aplicando merge..."):
                    control_desk_df = pd.read_excel(control_file)
                    df = merge_control_desk(df, control_desk_df)
            else:
                df['control_desk_info'] = 'NA'
            
            df = limpar_df(df)
            
            st.subheader("Prévia dos dados processados")
            st.dataframe(df.head(100))
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Extrato')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Baixar Excel (rio_pax_extrato.xlsx)",
                data=excel_data,
                file_name="rio_pax_extrato.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Processamento concluído!")

# ------------------------------------------------------------
# FUNÇÃO ESPECÍFICA PARA REVIVER (baseada no merge.py)
# ------------------------------------------------------------
def reviver_interface():
    st.header("REVIVER - Merge entre Relatório e Control Desk")
    
    
    col1, col2 = st.columns(2)
    with col1:
        relatorio_file = st.file_uploader("📁 Arquivo do Relatório (Excel)", type=["xls", "xlsx"], key="reviver_relatorio")
    with col2:
        control_file = st.file_uploader("📁 Arquivo Control Desk (Excel)", type=["xls", "xlsx"], key="reviver_control")
    
    processar = st.button("🔄 Realizar Merge e gerar Excel", type="primary")
    
    if processar:
        if relatorio_file is None or control_file is None:
            st.error("É necessário enviar ambos os arquivos (Relatório e Control Desk).")
        else:
            try:
                # 1. Carregar as planilhas (primeira aba)
                df_relatorio = pd.read_excel(relatorio_file, sheet_name=0)
                df_control = pd.read_excel(control_file, sheet_name=0)
                
                # Mostrar prévia das primeiras linhas
                with st.expander("Prévia do Relatório"):
                    st.dataframe(df_relatorio.head())
                with st.expander("Prévia do Control Desk"):
                    st.dataframe(df_control.head())
                
                # 2. Identificar as colunas pelas posições
                col_chave_relatorio = df_relatorio.columns[0]          # primeira coluna
                col_chave_control = df_control.columns[6]              # sétima coluna (índice 6)
                col_valor_control = df_control.columns[0]              # primeira coluna
                
                st.info(f"Chave no relatório: **{col_chave_relatorio}**")
                st.info(f"Chave no control desk: **{col_chave_control}**")
                st.info(f"Valor a ser trazido: **{col_valor_control}**")
                
                # 3. Renomear temporariamente para facilitar o merge
                df_relatorio_ren = df_relatorio.rename(columns={col_chave_relatorio: 'chave_relatorio'})
                df_control_ren = df_control.rename(columns={col_chave_control: 'chave_controle',
                                                            col_valor_control: 'valor_retornar'})
                
                # 4. Left join
                resultado = df_relatorio_ren.merge(
                    df_control_ren[['chave_controle', 'valor_retornar']],
                    left_on='chave_relatorio',
                    right_on='chave_controle',
                    how='left'
                )
                
                # 5. Remover coluna auxiliar e renomear a coluna de valor
                resultado.drop(columns=['chave_controle'], inplace=True)
                resultado.rename(columns={'valor_retornar': 'informacao_control_desk'}, inplace=True)
                
                # 6. Restaurar nome original da coluna chave do relatório
                resultado.rename(columns={'chave_relatorio': col_chave_relatorio}, inplace=True)
                
                # 7. Mostrar resultado
                st.subheader("Resultado do Merge")
                st.dataframe(resultado.head(100))
                st.success(f"Total de linhas no resultado: {len(resultado)}")
                
                # 8. Download do Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    resultado.to_excel(writer, index=False, sheet_name='Merge_Resultado')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Baixar Excel (relatorio_com_control_desk.xlsx)",
                    data=excel_data,
                    file_name="relatorio_com_control_desk.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Erro durante o processamento: {str(e)}")
                st.exception(e)

# ------------------------------------------------------------
# MAIN PAGE
# ------------------------------------------------------------
st.set_page_config(page_title="Real Cobrança - Relatórios de Baixa", layout="wide")

# Título principal
st.title(" 💲 REAL COBRANÇA RELATÓRIOS DE BAIXA")
st.markdown("---")
st.markdown("Selecione qual dos credores deseja realizar a baixa de títulos, serão gerados arquivos formato excel para facilitação do processo")

# Seleção do cliente (botões lado a lado)
col1, col2 = st.columns(2)
with col1:
    if st.button("RIO PAX", use_container_width=True):
        st.session_state["cliente"] = "RIO_PAX"
with col2:
    if st.button("REVIVER", use_container_width=True):
        st.session_state["cliente"] = "REVIVER"

# Define cliente padrão (primeira execução)
if "cliente" not in st.session_state:
    st.session_state["cliente"] = None

st.markdown("---")

# Exibe a interface do cliente selecionado
if st.session_state["cliente"] == "RIO_PAX":
    rio_pax_interface()
elif st.session_state["cliente"] == "REVIVER":
    reviver_interface()
else:
    st.info("Selecione um cliente em cima para começar.")
