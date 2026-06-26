import streamlit as st
import pdfplumber
import re
import pandas as pd
import numpy as np
from io import BytesIO
import os

# ------------------------------------------------------------
# FUNÇÕES COMPARTILHADAS
# ------------------------------------------------------------

def extrair_numero_apos_barra(os_contrato):
    """
    Extrai o número após a barra em os_contrato
    Exemplo: '422633 / 244906' -> '244906'
    """
    if pd.isna(os_contrato):
        return None
    
    # Converte para string
    os_contrato = str(os_contrato)
    
    # Busca padrão: número após a barra
    padrao = r'(?:/\s*)(\d+)'
    match = re.search(padrao, os_contrato)
    
    if match:
        return match.group(1)
    return None

def ler_planilha_com_engine(arquivo_bytes, nome_arquivo):
    """
    Tenta ler a planilha usando diferentes engines
    Suporta arquivos .XLS e .XLSX
    """
    # Lista de engines para tentar
    engines = ['xlrd', 'openpyxl', 'odf', None]
    
    # Salva temporariamente o arquivo para leitura
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(nome_arquivo)[1]) as tmp_file:
        tmp_file.write(arquivo_bytes)
        tmp_path = tmp_file.name
    
    try:
        for engine in engines:
            try:
                if engine is None:
                    df = pd.read_excel(tmp_path, sheet_name=0)
                else:
                    df = pd.read_excel(tmp_path, sheet_name=0, engine=engine)
                return df
            except:
                continue
        
        # Último recurso: tenta com xlrd manual
        try:
            import xlrd
            workbook = xlrd.open_workbook(tmp_path, encoding_override='latin-1')
            sheet = workbook.sheet_by_index(0)
            data = []
            for row_idx in range(sheet.nrows):
                row_data = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell_value(row_idx, col_idx)
                    row_data.append(cell)
                data.append(row_data)
            if data:
                headers = data[0]
                values = data[1:]
                df = pd.DataFrame(values, columns=headers)
                return df
        except:
            pass
        
        raise Exception("Não foi possível ler o arquivo com nenhum engine")
    finally:
        # Limpa o arquivo temporário
        try:
            os.unlink(tmp_path)
        except:
            pass

def consolidar_multiplos_excel(arquivos_upload):
    """
    Lê múltiplos arquivos Excel e consolida em um único DataFrame
    """
    dfs = []
    nomes_arquivos = []
    
    for arquivo in arquivos_upload:
        try:
            # Lê o arquivo
            df = ler_planilha_com_engine(arquivo.read(), arquivo.name)
            
            # Adiciona coluna com nome do arquivo de origem
            df['arquivo_origem'] = arquivo.name
            
            dfs.append(df)
            nomes_arquivos.append(arquivo.name)
            
            # Reset do ponteiro do arquivo
            arquivo.seek(0)
            
        except Exception as e:
            st.warning(f"⚠️ Erro ao ler o arquivo {arquivo.name}: {str(e)}")
            continue
    
    if not dfs:
        raise Exception("Nenhum arquivo pôde ser lido com sucesso.")
    
    # Concatena todos os DataFrames
    df_consolidado = pd.concat(dfs, ignore_index=True)
    
    return df_consolidado, nomes_arquivos

def merge_control_desk_melhorado(df, control_desk_df, nome_cliente="Relatório"):
    """
    Faz o merge entre o relatório e o control desk usando a lógica:
    - Extrai número após a barra em 'os_contrato'
    - Compara com 'COD EXTERNO'
    - Traz o 'PROCESSO'
    """
    # Verifica se as colunas necessárias existem
    if 'os_contrato' not in df.columns:
        st.error(f"❌ Coluna 'os_contrato' não encontrada no {nome_cliente}!")
        st.info(f"Colunas disponíveis: {', '.join(df.columns.tolist())}")
        return df
    
    if 'COD EXTERNO' not in control_desk_df.columns or 'PROCESSO' not in control_desk_df.columns:
        st.error("❌ Colunas 'COD EXTERNO' e/ou 'PROCESSO' não encontradas no Control Desk!")
        st.info(f"Colunas disponíveis: {', '.join(control_desk_df.columns.tolist())}")
        return df
    
    with st.spinner("🔍 Extraindo números após a barra..."):
        # Extrai o número após a barra
        df['numero_extraido'] = df['os_contrato'].apply(extrair_numero_apos_barra)
        
        # Mostra estatísticas
        total_com_numero = df['numero_extraido'].notna().sum()
        st.info(f"📊 Registros com número extraído: {total_com_numero} de {len(df)}")
        
        # Mostra exemplos
        if total_com_numero > 0:
            with st.expander("📋 Exemplos de números extraídos"):
                exemplos = df[df['numero_extraido'].notna()][['os_contrato', 'numero_extraido']].head(10)
                if 'arquivo_origem' in df.columns:
                    exemplos = df[df['numero_extraido'].notna()][['os_contrato', 'numero_extraido', 'arquivo_origem']].head(10)
                st.dataframe(exemplos)
    
    with st.spinner("🔄 Realizando merge com Control Desk..."):
        # Converte para string para comparação
        control_desk_df['COD_EXTERNO_STR'] = control_desk_df['COD EXTERNO'].astype(str).str.strip()
        df['numero_extraido'] = df['numero_extraido'].astype(str).str.strip()
        
        # Cria dicionário de mapeamento
        mapa_processo = dict(zip(control_desk_df['COD_EXTERNO_STR'], control_desk_df['PROCESSO']))
        
        # Adiciona a informação do Control Desk
        df['informacao_control_desk'] = df['numero_extraido'].map(mapa_processo)
        
        # Preenche com 'NA' onde não houve match
        df['informacao_control_desk'] = df['informacao_control_desk'].fillna('NA')
        
        # Estatísticas do merge
        total_matches = (df['informacao_control_desk'] != 'NA').sum()
        st.success(f"✅ Matches encontrados: {total_matches} de {total_com_numero}")
        
        if total_com_numero > 0:
            taxa_match = (total_matches / total_com_numero) * 100
            st.info(f"📈 Taxa de match: {taxa_match:.2f}%")
        
        # Mostra exemplos de matches
        if total_matches > 0:
            with st.expander("🎯 Exemplos de matches encontrados"):
                cols = ['os_contrato', 'numero_extraido', 'informacao_control_desk']
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                matches = df[df['informacao_control_desk'] != 'NA'][cols].head(10)
                st.dataframe(matches)
        
        # Mostra exemplos sem match
        sem_match = df[(df['numero_extraido'].notna()) & (df['informacao_control_desk'] == 'NA')]
        if len(sem_match) > 0:
            with st.expander("⚠️ Exemplos sem match"):
                cols = ['os_contrato', 'numero_extraido']
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                st.dataframe(sem_match[cols].head(10))
    
    return df

def limpar_df(df):
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.fillna('')
    for col in df.select_dtypes(include=['float', 'int']).columns:
        df[col] = df[col].apply(lambda x: '' if pd.isna(x) else str(x) if not np.isfinite(x) else x)
    return df

# ------------------------------------------------------------
# FUNÇÃO RIO PAX - VERSÃO COM MÚLTIPLOS ARQUIVOS
# ------------------------------------------------------------
def rio_pax_interface():
    st.header("🏦 RIO PAX - Processamento de Múltiplos Relatórios Excel")
    
    st.markdown("""
    **Como funciona:**
    1. Faça upload de **um ou mais** relatórios Excel do RIO PAX (todos com a coluna `os_contrato`)
    2. Faça upload do arquivo Control Desk (contém `COD EXTERNO` e `PROCESSO`)
    3. O script consolida todos os relatórios em um único DataFrame
    4. Extrai o número **após a barra** em `os_contrato` (ex: `422633 / 244906` → `244906`)
    5. Compara com `COD EXTERNO` do Control Desk
    6. Quando encontra match, traz a informação da coluna `PROCESSO`
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        relatorios_files = st.file_uploader(
            "📁 Arquivos do Relatório RIO PAX (Excel)", 
            type=["xls", "xlsx"], 
            key="rio_pax_relatorios",
            accept_multiple_files=True,  # Permite múltiplos arquivos
            help="Selecione um ou mais arquivos Excel com a coluna 'os_contrato'"
        )
    with col2:
        control_file = st.file_uploader(
            "📁 Arquivo Control Desk (Excel)", 
            type=["xls", "xlsx"], 
            key="rio_pax_control",
            help="Arquivo que contém 'COD EXTERNO' e 'PROCESSO'"
        )
    
    # Opção para escolher qual coluna usar no relatório
    if relatorios_files and len(relatorios_files) > 0:
        try:
            # Tenta ler o primeiro arquivo para identificar colunas
            df_temp = ler_planilha_com_engine(relatorios_files[0].read(), relatorios_files[0].name)
            relatorios_files[0].seek(0)  # Reset do ponteiro do arquivo
            colunas_relatorio = df_temp.columns.tolist()
            
            coluna_os = st.selectbox(
                "🔍 Selecione a coluna que contém 'os_contrato' nos relatórios:",
                options=colunas_relatorio,
                index=colunas_relatorio.index('os_contrato') if 'os_contrato' in colunas_relatorio else 0,
                help="Esta coluna deve conter valores como '422633 / 244906'"
            )
        except:
            coluna_os = 'os_contrato'
            st.warning("Não foi possível ler o arquivo para identificar colunas. Usando 'os_contrato' como padrão.")
    else:
        coluna_os = 'os_contrato'
    
    processar = st.button("🚀 Processar Múltiplos Arquivos e gerar Excel", type="primary")
    
    if processar:
        if not relatorios_files:
            st.error("❌ Envie pelo menos um arquivo do Relatório RIO PAX.")
        elif control_file is None:
            st.error("❌ Envie o arquivo Control Desk.")
        else:
            try:
                # 1. Carregar e consolidar os relatórios
                with st.spinner(f"📂 Carregando e consolidando {len(relatorios_files)} arquivos..."):
                    df_relatorio, nomes_arquivos = consolidar_multiplos_excel(relatorios_files)
                    
                # 2. Carregar Control Desk
                with st.spinner("📂 Carregando Control Desk..."):
                    df_control = ler_planilha_com_engine(control_file.read(), control_file.name)
                
                # Mostrar informações
                st.success(f"✅ {len(relatorios_files)} arquivos consolidados com sucesso!")
                st.info(f"📊 Total de registros consolidados: {len(df_relatorio)}")
                
                with st.expander("📋 Arquivos processados"):
                    for nome in nomes_arquivos:
                        st.write(f"- {nome}")
                
                # Mostrar prévia
                with st.expander("📊 Prévia do Relatório Consolidado"):
                    st.dataframe(df_relatorio.head())
                with st.expander("📊 Prévia do Control Desk"):
                    st.dataframe(df_control.head())
                
                # Verifica se a coluna existe
                if coluna_os not in df_relatorio.columns:
                    st.error(f"❌ Coluna '{coluna_os}' não encontrada no relatório!")
                    st.info(f"Colunas disponíveis: {', '.join(df_relatorio.columns.tolist())}")
                    return
                
                # Renomeia a coluna para padronizar
                if coluna_os != 'os_contrato':
                    df_relatorio.rename(columns={coluna_os: 'os_contrato'}, inplace=True)
                
                # 3. Aplica o merge melhorado
                df_resultado = merge_control_desk_melhorado(df_relatorio, df_control, "Relatórios RIO PAX")
                
                # 4. Limpeza final
                df_resultado = limpar_df(df_resultado)
                
                # 5. Mostrar resultado
                st.subheader("📊 Resultado do Processamento")
                st.dataframe(df_resultado.head(100))
                
                # Estatísticas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total de registros", len(df_resultado))
                with col2:
                    st.metric("Arquivos processados", len(relatorios_files))
                with col3:
                    total_com_numero = df_resultado['numero_extraido'].notna().sum()
                    st.metric("Com número extraído", total_com_numero)
                with col4:
                    total_matches = (df_resultado['informacao_control_desk'] != 'NA').sum()
                    st.metric("Matches encontrados", total_matches)
                
                # 6. Download do Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Sheet com resultado completo
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultado_Completo')
                    
                    # Sheet com apenas os matches
                    matches = df_resultado[df_resultado['informacao_control_desk'] != 'NA']
                    if not matches.empty:
                        matches.to_excel(writer, index=False, sheet_name='Apenas_Matches')
                    
                    # Sheet com apenas sem match
                    sem_match = df_resultado[(df_resultado['numero_extraido'].notna()) & 
                                            (df_resultado['informacao_control_desk'] == 'NA')]
                    if not sem_match.empty:
                        sem_match.to_excel(writer, index=False, sheet_name='Sem_Match')
                    
                    # Resumo por arquivo (se a coluna existir)
                    if 'arquivo_origem' in df_resultado.columns:
                        resumo_arquivos = df_resultado.groupby('arquivo_origem').agg({
                            'numero_extraido': lambda x: x.notna().sum(),
                            'informacao_control_desk': lambda x: (x != 'NA').sum()
                        }).reset_index()
                        resumo_arquivos.columns = ['Arquivo', 'Com_Número_Extraído', 'Matches']
                        resumo_arquivos.to_excel(writer, index=False, sheet_name='Resumo_por_Arquivo')
                
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Baixar Excel (rio_pax_consolidado.xlsx)",
                    data=excel_data,
                    file_name="rio_pax_consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ Processamento concluído com sucesso!")
                
            except Exception as e:
                st.error(f"❌ Erro durante o processamento: {str(e)}")
                st.exception(e)

# ------------------------------------------------------------
# FUNÇÃO REVIVER - MANTIDA IGUAL
# ------------------------------------------------------------
def reviver_interface():
    st.header("🔄 REVIVER - Merge entre Relatório e Control Desk")
    
    st.markdown("""
    **Como funciona:**
    1. O script extrai o número **após a barra** em `os_contrato` (ex: `422633 / 244906` → `244906`)
    2. Compara com a coluna `COD EXTERNO` do Control Desk
    3. Quando encontra match, traz a informação da coluna `PROCESSO`
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        relatorio_file = st.file_uploader(
            "📁 Arquivo do Relatório (Excel)", 
            type=["xls", "xlsx"], 
            key="reviver_relatorio",
            help="Arquivo que contém a coluna 'os_contrato'"
        )
    with col2:
        control_file = st.file_uploader(
            "📁 Arquivo Control Desk (Excel)", 
            type=["xls", "xlsx"], 
            key="reviver_control",
            help="Arquivo que contém 'COD EXTERNO' e 'PROCESSO'"
        )
    
    # Opção para escolher qual coluna usar no relatório
    if relatorio_file:
        try:
            df_temp = pd.read_excel(relatorio_file, sheet_name=0, nrows=5)
            colunas_relatorio = df_temp.columns.tolist()
            
            coluna_os = st.selectbox(
                "🔍 Selecione a coluna que contém 'os_contrato' no relatório:",
                options=colunas_relatorio,
                index=colunas_relatorio.index('os_contrato') if 'os_contrato' in colunas_relatorio else 0,
                help="Esta coluna deve conter valores como '422633 / 244906'"
            )
        except:
            coluna_os = 'os_contrato'
            st.warning("Não foi possível ler o arquivo para identificar colunas. Usando 'os_contrato' como padrão.")
    else:
        coluna_os = 'os_contrato'
    
    processar = st.button("🔄 Realizar Merge e gerar Excel", type="primary")
    
    if processar:
        if relatorio_file is None or control_file is None:
            st.error("❌ É necessário enviar ambos os arquivos (Relatório e Control Desk).")
        else:
            try:
                # 1. Carregar as planilhas
                with st.spinner("📂 Carregando arquivos..."):
                    df_relatorio = pd.read_excel(relatorio_file, sheet_name=0)
                    df_control = pd.read_excel(control_file, sheet_name=0)
                
                # Mostrar prévia
                with st.expander("📊 Prévia do Relatório"):
                    st.dataframe(df_relatorio.head())
                with st.expander("📊 Prévia do Control Desk"):
                    st.dataframe(df_control.head())
                
                # Verifica se a coluna existe
                if coluna_os not in df_relatorio.columns:
                    st.error(f"❌ Coluna '{coluna_os}' não encontrada no relatório!")
                    st.info(f"Colunas disponíveis: {', '.join(df_relatorio.columns.tolist())}")
                    return
                
                # Renomeia a coluna para padronizar
                if coluna_os != 'os_contrato':
                    df_relatorio.rename(columns={coluna_os: 'os_contrato'}, inplace=True)
                
                # 2. Aplica o merge melhorado
                df_resultado = merge_control_desk_melhorado(df_relatorio, df_control, "Relatório REVIVER")
                
                # 3. Limpeza final
                df_resultado = limpar_df(df_resultado)
                
                # 4. Mostrar resultado
                st.subheader("📊 Resultado do Merge")
                st.dataframe(df_resultado.head(100))
                
                # Estatísticas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de linhas", len(df_resultado))
                with col2:
                    total_com_numero = df_resultado['numero_extraido'].notna().sum()
                    st.metric("Com número extraído", total_com_numero)
                with col3:
                    total_matches = (df_resultado['informacao_control_desk'] != 'NA').sum()
                    st.metric("Matches encontrados", total_matches)
                
                # 5. Download do Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Merge_Resultado')
                    
                    # Adiciona sheet com apenas os matches
                    matches = df_resultado[df_resultado['informacao_control_desk'] != 'NA']
                    if not matches.empty:
                        matches.to_excel(writer, index=False, sheet_name='Apenas_Matches')
                    
                    # Adiciona sheet com apenas sem match
                    sem_match = df_resultado[(df_resultado['numero_extraido'].notna()) & 
                                            (df_resultado['informacao_control_desk'] == 'NA')]
                    if not sem_match.empty:
                        sem_match.to_excel(writer, index=False, sheet_name='Sem_Match')
                
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Baixar Excel (relatorio_com_control_desk.xlsx)",
                    data=excel_data,
                    file_name="relatorio_com_control_desk.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("✅ Processamento concluído com sucesso!")
                
            except Exception as e:
                st.error(f"❌ Erro durante o processamento: {str(e)}")
                st.exception(e)

# ------------------------------------------------------------
# MAIN PAGE
# ------------------------------------------------------------
st.set_page_config(page_title="Real Cobrança - Relatórios de Baixa", layout="wide")

# Título principal
st.title("💰 REAL COBRANÇA RELATÓRIOS DE BAIXA")
st.markdown("---")
st.markdown("Selecione qual dos credores deseja realizar a baixa de títulos, serão gerados arquivos formato excel para facilitação do processo")

# Seleção do cliente (botões lado a lado)
col1, col2 = st.columns(2)
with col1:
    if st.button("🏦 RIO PAX", use_container_width=True):
        st.session_state["cliente"] = "RIO_PAX"
with col2:
    if st.button("🔄 REVIVER", use_container_width=True):
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
    st.info("👈 Selecione um cliente ao lado para começar.")
