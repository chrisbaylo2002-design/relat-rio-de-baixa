import streamlit as st
import pdfplumber
import re
import pandas as pd
import numpy as np
import os
import tempfile
from io import BytesIO

# ==============================================================
# FUNÇÕES COMPARTILHADAS GERAIS
# ==============================================================

def limpar_df(df):
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.fillna('')
    for col in df.select_dtypes(include=['float', 'int']).columns:
        df[col] = df[col].apply(lambda x: '' if pd.isna(x) else str(x) if not np.isfinite(x) else x)
    return df


# ==============================================================
# FUNÇÕES PARA LEITURA DE PDF (RIO PAX - modo PDF)
# ==============================================================

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


def merge_control_desk_pdf(df, control_desk_df):
    """Merge simples usado no fluxo PDF: compara a coluna 'os' com 'COD EXTERNO'."""
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


# ==============================================================
# FUNÇÕES PARA LEITURA DE EXCEL (RIO PAX - modo Excel / REVIVER)
# ==============================================================

def extrair_numero_apos_barra(os_contrato):
    """
    Extrai o número após a barra em os_contrato
    Exemplo: '422633 / 244906' -> '244906'
    Também funciona com: '422633/244906', '422633 - 244906', '000000 / 196933' -> '196933'
    """
    if pd.isna(os_contrato):
        return None
    
    os_contrato = str(os_contrato).strip()
    
    # Padrão 1: espaço barra espaço (formato mais comum)
    padrao1 = r'(?:/\s*)(\d+)'
    match = re.search(padrao1, os_contrato)
    if match:
        return match.group(1).lstrip('0')
    
    # Padrão 2: apenas barra sem espaços
    padrao2 = r'/(\d+)'
    match = re.search(padrao2, os_contrato)
    if match:
        return match.group(1).lstrip('0')
    
    # Padrão 3: hífen com espaços
    padrao3 = r'(?:-\s*)(\d+)'
    match = re.search(padrao3, os_contrato)
    if match:
        return match.group(1).lstrip('0')
    
    # Padrão 4: se não tiver separador, pega o último número
    numeros = re.findall(r'\d+', os_contrato)
    if len(numeros) >= 2:
        return numeros[-1].lstrip('0')
    
    # Padrão 5: tenta pegar qualquer número no final
    match = re.search(r'(\d+)$', os_contrato)
    if match:
        return match.group(1).lstrip('0')
    
    return None


def ler_planilha_com_engine(arquivo_bytes, nome_arquivo):
    """
    Tenta ler a planilha usando diferentes engines.
    Suporta arquivos .XLS e .XLSX
    """
    engines = ['xlrd', 'openpyxl', 'odf', None]

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
            except Exception:
                continue

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
        except Exception:
            pass

        raise Exception("Não foi possível ler o arquivo com nenhum engine")
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


def consolidar_multiplos_excel(arquivos_upload):
    """Lê múltiplos arquivos Excel e consolida em um único DataFrame."""
    dfs = []
    nomes_arquivos = []

    for arquivo in arquivos_upload:
        try:
            df = ler_planilha_com_engine(arquivo.read(), arquivo.name)
            df['arquivo_origem'] = arquivo.name
            dfs.append(df)
            nomes_arquivos.append(arquivo.name)
            arquivo.seek(0)
        except Exception as e:
            st.warning(f"⚠️ Erro ao ler o arquivo {arquivo.name}: {str(e)}")
            continue

    if not dfs:
        raise Exception("Nenhum arquivo pôde ser lido com sucesso.")

    df_consolidado = pd.concat(dfs, ignore_index=True)
    return df_consolidado, nomes_arquivos


def merge_control_desk_melhorado(df, control_desk_df, nome_cliente="Relatório"):
    """
    Faz o merge entre o relatório e o control desk usando a lógica:
    - Extrai número após a barra em 'os_contrato'
    - Compara com 'COD EXTERNO'
    - Traz o 'PROCESSO'
    """
    if 'os_contrato' not in df.columns:
        st.error(f"❌ Coluna 'os_contrato' não encontrada no {nome_cliente}!")
        st.info(f"Colunas disponíveis: {', '.join(df.columns.tolist())}")
        return df

    if 'COD EXTERNO' not in control_desk_df.columns or 'PROCESSO' not in control_desk_df.columns:
        st.error("❌ Colunas 'COD EXTERNO' e/ou 'PROCESSO' não encontradas no Control Desk!")
        st.info(f"Colunas disponíveis: {', '.join(control_desk_df.columns.tolist())}")
        return df

    with st.spinner("🔍 Extraindo números após a barra..."):
        df['numero_extraido'] = df['os_contrato'].apply(extrair_numero_apos_barra)
        total_com_numero = df['numero_extraido'].notna().sum()
        st.info(f"📊 Registros com número extraído: {total_com_numero} de {len(df)}")

        if total_com_numero > 0:
            with st.expander("📋 Exemplos de números extraídos"):
                exemplos = df[df['numero_extraido'].notna()][['os_contrato', 'numero_extraido']].head(10)
                if 'arquivo_origem' in df.columns:
                    exemplos = df[df['numero_extraido'].notna()][['os_contrato', 'numero_extraido', 'arquivo_origem']].head(10)
                st.dataframe(exemplos)

    with st.spinner("🔄 Realizando merge com Control Desk..."):
        control_desk_df['COD_EXTERNO_STR'] = control_desk_df['COD EXTERNO'].astype(str).str.strip()
        control_desk_df['COD_EXTERNO_STR'] = control_desk_df['COD_EXTERNO_STR'].str.lstrip('0')
        df['numero_extraido'] = df['numero_extraido'].astype(str).str.strip()
        df['numero_extraido'] = df['numero_extraido'].str.lstrip('0')

        mapa_processo = dict(zip(control_desk_df['COD_EXTERNO_STR'], control_desk_df['PROCESSO']))
        df['informacao_control_desk'] = df['numero_extraido'].map(mapa_processo)
        df['informacao_control_desk'] = df['informacao_control_desk'].fillna('NA')

        total_matches = (df['informacao_control_desk'] != 'NA').sum()
        st.success(f"✅ Matches encontrados: {total_matches} de {total_com_numero}")

        if total_com_numero > 0:
            taxa_match = (total_matches / total_com_numero) * 100
            st.info(f"📈 Taxa de match: {taxa_match:.2f}%")

        if total_matches > 0:
            with st.expander("🎯 Exemplos de matches encontrados"):
                cols = ['os_contrato', 'numero_extraido', 'informacao_control_desk']
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                matches = df[df['informacao_control_desk'] != 'NA'][cols].head(10)
                st.dataframe(matches)

        sem_match = df[(df['numero_extraido'].notna()) & (df['informacao_control_desk'] == 'NA')]
        if len(sem_match) > 0:
            with st.expander("⚠️ Exemplos sem match"):
                cols = ['os_contrato', 'numero_extraido']
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                st.dataframe(sem_match[cols].head(10))

    return df


def merge_control_desk_direto(df, control_desk_df, coluna_chave_relatorio='os_contrato', nome_cliente="Relatório"):
    """
    Merge DIRETO (sem extrair número após a barra), usado no fluxo REVIVER:
    - Compara a coluna-chave do relatório (ex: 'Contrato/OS') diretamente com 'COD EXTERNO'
    - Traz o 'PROCESSO'
    """
    if coluna_chave_relatorio not in df.columns:
        st.error(f"❌ Coluna '{coluna_chave_relatorio}' não encontrada no {nome_cliente}!")
        st.info(f"Colunas disponíveis: {', '.join(df.columns.tolist())}")
        return df

    if 'COD EXTERNO' not in control_desk_df.columns or 'PROCESSO' not in control_desk_df.columns:
        st.error("❌ Colunas 'COD EXTERNO' e/ou 'PROCESSO' não encontradas no Control Desk!")
        st.info(f"Colunas disponíveis: {', '.join(control_desk_df.columns.tolist())}")
        return df

    with st.spinner("🔄 Realizando merge direto com Control Desk..."):
        control_desk_df['COD_EXTERNO_STR'] = control_desk_df['COD EXTERNO'].astype(str).str.strip()
        control_desk_df['COD_EXTERNO_STR'] = control_desk_df['COD_EXTERNO_STR'].str.lstrip('0')

        df['chave_comparacao'] = df[coluna_chave_relatorio].astype(str).str.strip()
        df['chave_comparacao'] = df['chave_comparacao'].str.lstrip('0')

        mapa_processo = dict(zip(control_desk_df['COD_EXTERNO_STR'], control_desk_df['PROCESSO']))
        df['informacao_control_desk'] = df['chave_comparacao'].map(mapa_processo)
        df['informacao_control_desk'] = df['informacao_control_desk'].fillna('NA')
        df = df.drop(columns=['chave_comparacao'])

        total_matches = (df['informacao_control_desk'] != 'NA').sum()
        st.success(f"✅ Matches encontrados: {total_matches} de {len(df)}")

        if len(df) > 0:
            taxa_match = (total_matches / len(df)) * 100
            st.info(f"📈 Taxa de match: {taxa_match:.2f}%")

        if total_matches > 0:
            with st.expander("🎯 Exemplos de matches encontrados"):
                cols = [coluna_chave_relatorio, 'informacao_control_desk']
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                matches = df[df['informacao_control_desk'] != 'NA'][cols].head(10)
                st.dataframe(matches)

        sem_match = df[df['informacao_control_desk'] == 'NA']
        if len(sem_match) > 0:
            with st.expander("⚠️ Exemplos sem match"):
                cols = [coluna_chave_relatorio]
                if 'arquivo_origem' in df.columns:
                    cols.append('arquivo_origem')
                st.dataframe(sem_match[cols].head(10))

    return df


# ==============================================================
# RIO PAX - SUB-FLUXO: PDF
# ==============================================================

def rio_pax_pdf_fluxo():
    st.markdown("Faça upload de um ou mais PDFs do RIO PAX. O texto será extraído e as transações serão estruturadas automaticamente.")

    pdf_files = st.file_uploader("Selecione um ou mais PDFs do RIO PAX", type="pdf", accept_multiple_files=True, key="rio_pax_pdf_files")
    control_file = st.file_uploader("Arquivo Control Desk (Excel com colunas 'COD EXTERNO' e 'PROCESSO')", type=["xlsx", "xls"], key="rio_pax_pdf_control")

    processar = st.button("🚀 Processar PDFs e gerar Excel", type="primary", key="rio_pax_pdf_btn")

    if processar:
        if not pdf_files:
            st.error("Envie pelo menos um arquivo PDF.")
            return

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
                df = merge_control_desk_pdf(df, control_desk_df)
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


# ==============================================================
# RIO PAX - SUB-FLUXO: EXCEL
# ==============================================================

def rio_pax_excel_fluxo():
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
            key="rio_pax_excel_relatorios",
            accept_multiple_files=True,
            help="Selecione um ou mais arquivos Excel com a coluna 'os_contrato'"
        )
    with col2:
        control_file = st.file_uploader(
            "📁 Arquivo Control Desk (Excel)",
            type=["xls", "xlsx"],
            key="rio_pax_excel_control",
            help="Arquivo que contém 'COD EXTERNO' e 'PROCESSO'"
        )

    if relatorios_files and len(relatorios_files) > 0:
        try:
            df_temp = ler_planilha_com_engine(relatorios_files[0].read(), relatorios_files[0].name)
            relatorios_files[0].seek(0)
            colunas_relatorio = df_temp.columns.tolist()

            coluna_os = st.selectbox(
                "🔍 Selecione a coluna que contém 'os_contrato' nos relatórios:",
                options=colunas_relatorio,
                index=colunas_relatorio.index('os_contrato') if 'os_contrato' in colunas_relatorio else 0,
                help="Esta coluna deve conter valores como '422633 / 244906'",
                key="rio_pax_excel_coluna_os"
            )
        except Exception:
            coluna_os = 'os_contrato'
            st.warning("Não foi possível ler o arquivo para identificar colunas. Usando 'os_contrato' como padrão.")
    else:
        coluna_os = 'os_contrato'

    processar = st.button("🚀 Processar Múltiplos Arquivos e gerar Excel", type="primary", key="rio_pax_excel_btn")

    if processar:
        if not relatorios_files:
            st.error("❌ Envie pelo menos um arquivo do Relatório RIO PAX.")
        elif control_file is None:
            st.error("❌ Envie o arquivo Control Desk.")
        else:
            try:
                with st.spinner(f"📂 Carregando e consolidando {len(relatorios_files)} arquivos..."):
                    df_relatorio, nomes_arquivos = consolidar_multiplos_excel(relatorios_files)

                with st.spinner("📂 Carregando Control Desk..."):
                    df_control = ler_planilha_com_engine(control_file.read(), control_file.name)

                st.success(f"✅ {len(relatorios_files)} arquivos consolidados com sucesso!")
                st.info(f"📊 Total de registros consolidados: {len(df_relatorio)}")

                with st.expander("📋 Arquivos processados"):
                    for nome in nomes_arquivos:
                        st.write(f"- {nome}")

                with st.expander("📊 Prévia do Relatório Consolidado"):
                    st.dataframe(df_relatorio.head())
                with st.expander("📊 Prévia do Control Desk"):
                    st.dataframe(df_control.head())

                if coluna_os not in df_relatorio.columns:
                    st.error(f"❌ Coluna '{coluna_os}' não encontrada no relatório!")
                    st.info(f"Colunas disponíveis: {', '.join(df_relatorio.columns.tolist())}")
                    return

                if coluna_os != 'os_contrato':
                    df_relatorio.rename(columns={coluna_os: 'os_contrato'}, inplace=True)

                df_resultado = merge_control_desk_melhorado(df_relatorio, df_control, "Relatórios RIO PAX")
                df_resultado = limpar_df(df_resultado)

                st.subheader("📊 Resultado do Processamento")
                st.dataframe(df_resultado.head(100))

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

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultado_Completo')

                    matches = df_resultado[df_resultado['informacao_control_desk'] != 'NA']
                    if not matches.empty:
                        matches.to_excel(writer, index=False, sheet_name='Apenas_Matches')

                    sem_match = df_resultado[(df_resultado['numero_extraido'].notna()) &
                                              (df_resultado['informacao_control_desk'] == 'NA')]
                    if not sem_match.empty:
                        sem_match.to_excel(writer, index=False, sheet_name='Sem_Match')

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


# ==============================================================
# RIO PAX - INTERFACE PRINCIPAL (escolhe PDF ou Excel)
# ==============================================================

def rio_pax_interface():
    st.header("🏦 RIO PAX - Extração e estruturação de relatórios")

    tipo_arquivo = st.radio(
        "Qual o formato do arquivo do RIO PAX que você vai enviar?",
        options=["PDF", "Excel"],
        horizontal=True,
        key="rio_pax_tipo_arquivo"
    )

    st.markdown("---")

    if tipo_arquivo == "PDF":
        rio_pax_pdf_fluxo()
    else:
        rio_pax_excel_fluxo()


# ==============================================================
# REVIVER - Merge entre Relatório e Control Desk
# ==============================================================

def reviver_interface():
    st.header("🔄 REVIVER - Merge entre Relatório e Control Desk")

    st.markdown("""
    **Como funciona:**
    1. Compara diretamente a coluna `Contrato/OS` do relatório com a coluna `COD EXTERNO` do Control Desk
    2. Quando encontra match, traz a informação da coluna `PROCESSO`
    """)

    col1, col2 = st.columns(2)
    with col1:
        relatorio_file = st.file_uploader(
            "📁 Arquivo do Relatório (Excel)",
            type=["xls", "xlsx"],
            key="reviver_relatorio",
            help="Arquivo que contém a coluna 'Contrato/OS'"
        )
    with col2:
        control_file = st.file_uploader(
            "📁 Arquivo Control Desk (Excel)",
            type=["xls", "xlsx"],
            key="reviver_control",
            help="Arquivo que contém 'COD EXTERNO' e 'PROCESSO'"
        )

    if relatorio_file:
        try:
            df_temp = pd.read_excel(relatorio_file, sheet_name=0, nrows=5)
            colunas_relatorio = df_temp.columns.tolist()

            coluna_os = st.selectbox(
                "🔍 Selecione a coluna que contém 'Contrato/OS' no relatório:",
                options=colunas_relatorio,
                index=_sugerir_coluna(colunas_relatorio, ['contrato/os', 'contrato / os', 'os_contrato']),
                help="Esta coluna será comparada diretamente com 'COD EXTERNO'",
                key="reviver_coluna_os"
            )
        except Exception:
            coluna_os = 'os_contrato'
            st.warning("Não foi possível ler o arquivo para identificar colunas. Usando 'os_contrato' como padrão.")
    else:
        coluna_os = 'os_contrato'

    processar = st.button("🔄 Realizar Merge e gerar Excel", type="primary", key="reviver_btn")

    if processar:
        if relatorio_file is None or control_file is None:
            st.error("❌ É necessário enviar ambos os arquivos (Relatório e Control Desk).")
        else:
            try:
                with st.spinner("📂 Carregando arquivos..."):
                    df_relatorio = pd.read_excel(relatorio_file, sheet_name=0)
                    df_control = pd.read_excel(control_file, sheet_name=0)

                with st.expander("📊 Prévia do Relatório"):
                    st.dataframe(df_relatorio.head())
                with st.expander("📊 Prévia do Control Desk"):
                    st.dataframe(df_control.head())

                if coluna_os not in df_relatorio.columns:
                    st.error(f"❌ Coluna '{coluna_os}' não encontrada no relatório!")
                    st.info(f"Colunas disponíveis: {', '.join(df_relatorio.columns.tolist())}")
                    return

                df_resultado = merge_control_desk_direto(df_relatorio, df_control, coluna_os, "Relatório REVIVER")
                df_resultado = limpar_df(df_resultado)

                st.subheader("📊 Resultado do Merge")
                st.dataframe(df_resultado.head(100))

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total de linhas", len(df_resultado))
                with col2:
                    total_matches = (df_resultado['informacao_control_desk'] != 'NA').sum()
                    st.metric("Matches encontrados", total_matches)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Merge_Resultado')

                    matches = df_resultado[df_resultado['informacao_control_desk'] != 'NA']
                    if not matches.empty:
                        matches.to_excel(writer, index=False, sheet_name='Apenas_Matches')

                    sem_match = df_resultado[df_resultado['informacao_control_desk'] == 'NA']
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


# ==============================================================
# COMPENSADO - Unir relatórios e comparar com títulos cadastrados
# ==============================================================

def _sugerir_coluna(colunas, candidatos):
    """Retorna o índice da primeira coluna que bate (case-insensitive) com algum candidato."""
    colunas_lower = [c.lower().strip() for c in colunas]
    for candidato in candidatos:
        if candidato.lower() in colunas_lower:
            return colunas_lower.index(candidato.lower())
    return 0


def compensado_interface():
    st.header("✅ COMPENSADO - Unir relatórios e comparar com Títulos Cadastrados")

    st.markdown("""
    **Como funciona:**
    1. Faça upload de **todos** os relatórios (Aquisição, Columbário, Exumação, Gaveta Perp., Manutenção, etc.) — quantos forem necessários
    2. Faça upload de **um ou mais** arquivos de **Títulos Cadastrados** — eles também serão unidos entre si
    3. Escolha as colunas-chave em cada arquivo — os nomes podem variar entre os relatórios
    4. O script une todos os relatórios em um único DataFrame, e todos os títulos cadastrados em outro
    5. Na coluna `os_contrato` dos relatórios (formato `000000/00000`), o script extrai o **número após a barra** e compara com a coluna `COD EXTERNO` dos títulos
    6. Também compara a coluna de Referência dos relatórios com a de Parcela dos títulos
    7. Gera o arquivo `compensado.xlsx` apenas com os registros que deram match em ambas as comparações (merge `inner`)
    """)

    col1, col2 = st.columns(2)
    with col1:
        relatorios_files = st.file_uploader(
            "📁 Relatórios (Aquisição, Columbário, Exumação, Gaveta Perp., Manutenção...)",
            type=["xls", "xlsx"],
            key="compensado_relatorios",
            accept_multiple_files=True,
            help="Selecione todos os arquivos de relatório que devem ser unidos"
        )
    with col2:
        titulos_files = st.file_uploader(
            "📁 Arquivo(s) de Títulos Cadastrados (Excel)",
            type=["xls", "xlsx"],
            key="compensado_titulos",
            accept_multiple_files=True,
            help="Selecione um ou mais arquivos de Títulos Cadastrados — eles serão unidos entre si"
        )

    # ----------------------------------------------------------
    # Seleção das colunas-chave (assim que os arquivos forem enviados)
    # ----------------------------------------------------------
    colunas_relatorio = []
    colunas_titulos = []

    if relatorios_files:
        try:
            df_preview_rel = ler_planilha_com_engine(relatorios_files[0].read(), relatorios_files[0].name)
            relatorios_files[0].seek(0)
            colunas_relatorio = df_preview_rel.columns.tolist()
        except Exception:
            st.warning("Não foi possível ler o primeiro relatório para identificar as colunas.")

    if titulos_files:
        try:
            df_preview_tit = ler_planilha_com_engine(titulos_files[0].read(), titulos_files[0].name)
            titulos_files[0].seek(0)
            colunas_titulos = df_preview_tit.columns.tolist()
        except Exception:
            st.warning("Não foi possível ler o primeiro arquivo de Títulos Cadastrados para identificar as colunas.")

    if colunas_relatorio or colunas_titulos:
        st.markdown("#### 🔍 Colunas-chave para o merge")
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("**Nos relatórios:**")
            if colunas_relatorio:
                col_chave_rel = st.selectbox(
                    "Coluna de os_contrato (formato 000000/00000) (relatórios)",
                    options=colunas_relatorio,
                    index=_sugerir_coluna(colunas_relatorio, ['os_contrato', 'Contrato/OS', 'contrato/os']),
                    help="Coluna com valores tipo '422633 / 244906' — será extraído o número após a barra",
                    key="compensado_col_chave_rel"
                )
                col_venc_rel = st.selectbox(
                    "Coluna de Referência (relatórios)",
                    options=colunas_relatorio,
                    index=_sugerir_coluna(colunas_relatorio, ['refer', 'referencia', 'parcela']),
                    key="compensado_col_venc_rel"
                )
            else:
                col_chave_rel = col_venc_rel = None
                st.info("Envie os relatórios para escolher as colunas.")

        with c2:
            st.markdown("**No arquivo de Títulos Cadastrados:**")
            if colunas_titulos:
                col_chave_tit = st.selectbox(
                    "Coluna de Código Externo (títulos)",
                    options=colunas_titulos,
                    index=_sugerir_coluna(colunas_titulos, ['COD EXTERNO', 'cod externo', 'codigo externo']),
                    help="Comparada com o número após a barra em 'os_contrato'",
                    key="compensado_col_chave_tit"
                )
                col_venc_tit = st.selectbox(
                    "Coluna de Parcela (títulos)",
                    options=colunas_titulos,
                    index=_sugerir_coluna(colunas_titulos, ['parcela', 'refer', 'referencia']),
                    key="compensado_col_venc_tit"
                )
            else:
                col_chave_tit = col_venc_tit = None
                st.info("Envie o arquivo de Títulos Cadastrados para escolher as colunas.")
    else:
        col_chave_rel = col_venc_rel = col_chave_tit = col_venc_tit = None

    processar = st.button("✅ Unir relatórios e gerar compensado.xlsx", type="primary", key="compensado_btn")

    if processar:
        if not relatorios_files:
            st.error("❌ Envie pelo menos um arquivo de relatório.")
            return
        if not titulos_files:
            st.error("❌ Envie pelo menos um arquivo de Títulos Cadastrados.")
            return
        if not all([col_chave_rel, col_venc_rel, col_chave_tit, col_venc_tit]):
            st.error("❌ Selecione todas as colunas-chave antes de processar.")
            return

        try:
            # 1. Lê e une todos os relatórios enviados
            with st.spinner(f"📂 Lendo e unindo {len(relatorios_files)} relatório(s)..."):
                dfs_relatorios = []
                nomes_arquivos = []
                for arquivo in relatorios_files:
                    df_temp = ler_planilha_com_engine(arquivo.read(), arquivo.name)
                    dfs_relatorios.append(df_temp)
                    nomes_arquivos.append(arquivo.name)
                    arquivo.seek(0)

                df_unido = pd.concat(dfs_relatorios, ignore_index=True)

            st.success(f"✅ {len(relatorios_files)} relatório(s) unidos com sucesso!")
            st.info(f"📊 Total de linhas após união: {len(df_unido)}")

            with st.expander("📋 Arquivos de relatório unidos"):
                for nome in nomes_arquivos:
                    st.write(f"- {nome}")

            with st.expander("📊 Prévia dos relatórios unidos"):
                st.dataframe(df_unido.head())

            # 2. Lê e une todos os arquivos de títulos cadastrados
            with st.spinner(f"📂 Lendo e unindo {len(titulos_files)} arquivo(s) de Títulos Cadastrados..."):
                dfs_titulos = []
                nomes_titulos = []
                for arquivo in titulos_files:
                    df_temp = ler_planilha_com_engine(arquivo.read(), arquivo.name)
                    dfs_titulos.append(df_temp)
                    nomes_titulos.append(arquivo.name)
                    arquivo.seek(0)

                df_titulos = pd.concat(dfs_titulos, ignore_index=True)

            st.success(f"✅ {len(titulos_files)} arquivo(s) de Títulos Cadastrados unidos com sucesso!")
            st.info(f"📊 Total de linhas após união: {len(df_titulos)}")

            with st.expander("📋 Arquivos de Títulos Cadastrados unidos"):
                for nome in nomes_titulos:
                    st.write(f"- {nome}")

            with st.expander("📊 Prévia dos Títulos Cadastrados unidos"):
                st.dataframe(df_titulos.head())

            # 3. Remove espaços dos nomes das colunas
            df_unido.columns = df_unido.columns.str.strip()
            df_titulos.columns = df_titulos.columns.str.strip()

            # 4. Verifica se as colunas escolhidas ainda existem (por segurança)
            if col_chave_rel not in df_unido.columns or col_venc_rel not in df_unido.columns:
                st.error(f"❌ Coluna(s) selecionada(s) não encontrada(s) nos relatórios unidos!")
                st.info(f"Colunas disponíveis: {', '.join(df_unido.columns.tolist())}")
                return
            if col_chave_tit not in df_titulos.columns or col_venc_tit not in df_titulos.columns:
                st.error(f"❌ Coluna(s) selecionada(s) não encontrada(s) no arquivo de Títulos Cadastrados!")
                st.info(f"Colunas disponíveis: {', '.join(df_titulos.columns.tolist())}")
                return

            # 5. Extrai o número após a barra em os_contrato (ex: '422633 / 244906' -> '244906')
            with st.spinner("🔍 Extraindo números após a barra em os_contrato..."):
                df_unido['numero_extraido'] = df_unido[col_chave_rel].apply(extrair_numero_apos_barra)
                total_com_numero = df_unido['numero_extraido'].notna().sum()
                st.info(f"📊 Registros com número extraído: {total_com_numero} de {len(df_unido)}")

                with st.expander("📋 Exemplos de números extraídos"):
                    exemplos = df_unido[df_unido['numero_extraido'].notna()][[col_chave_rel, 'numero_extraido']].head(10)
                    st.dataframe(exemplos)

            # 6. Padroniza os nomes das colunas-chave para o merge
            df_unido = df_unido.rename(columns={col_venc_rel: 'Parcela'})
            df_titulos = df_titulos.rename(columns={col_chave_tit: 'COD_EXTERNO', col_venc_tit: 'Parcela'})
            
            # 7. Limpa espaços nos valores de texto usados na comparação
            with st.spinner("🧹 Normalizando colunas-chave..."):
                # Converter para string e remover espaços
                df_unido['numero_extraido'] = df_unido['numero_extraido'].astype(str).str.strip()
                df_titulos['COD_EXTERNO'] = df_titulos['COD_EXTERNO'].astype(str).str.strip()
                
                # REMOVER ZEROS À ESQUERDA para padronizar a comparação
                df_unido['numero_extraido'] = df_unido['numero_extraido'].str.lstrip('0')
                df_titulos['COD_EXTERNO'] = df_titulos['COD_EXTERNO'].str.lstrip('0')
                
                # Remover pontos e outros caracteres especiais se necessário
                df_unido['numero_extraido'] = df_unido['numero_extraido'].str.replace('.', '', regex=False)
                df_unido['numero_extraido'] = df_unido['numero_extraido'].str.replace(',', '', regex=False)
                df_titulos['COD_EXTERNO'] = df_titulos['COD_EXTERNO'].str.replace('.', '', regex=False)
                df_titulos['COD_EXTERNO'] = df_titulos['COD_EXTERNO'].str.replace(',', '', regex=False)
                
                # Padronizar Parcela (remover espaços extras)
                df_unido['Parcela'] = df_unido['Parcela'].astype(str).str.strip()
                df_titulos['Parcela'] = df_titulos['Parcela'].astype(str).str.strip()
                
                # 🔍 DEBUG: Mostrar valores para verificar
                st.write("### 🔍 Verificação dos dados após normalização")
                col_debug1, col_debug2 = st.columns(2)
                with col_debug1:
                    st.write("**Relatório (numero_extraido):**")
                    st.dataframe(df_unido[['numero_extraido']].head(10))
                with col_debug2:
                    st.write("**Títulos (COD_EXTERNO):**")
                    st.dataframe(df_titulos[['COD_EXTERNO']].head(10))
                
                # Verificar quantos valores se sobrepõem
                valores_rel = set(df_unido['numero_extraido'].astype(str))
                valores_tit = set(df_titulos['COD_EXTERNO'].astype(str))
                interseccao = valores_rel.intersection(valores_tit)
                
                st.info(f"📊 Valores em comum entre numero_extraido e COD_EXTERNO: {len(interseccao)}")
                if len(interseccao) > 0:
                    st.success(f"✅ Exemplos de matches: {list(interseccao)[:10]}")
                else:
                    st.warning("⚠️ NENHUM match encontrado entre numero_extraido e COD_EXTERNO!")
                    st.write("Exemplos do relatório (numero_extraido):", list(valores_rel)[:10])
                    st.write("Exemplos do título (COD_EXTERNO):", list(valores_tit)[:10])

            st.info(f"📊 Relatórios unidos: {len(df_unido)} linhas | Títulos Cadastrados: {len(df_titulos)} linhas")

            # 8. Merge: número após a barra == COD EXTERNO, e Referência == Parcela
            with st.spinner("🔄 Realizando o merge (número após a barra + Parcela)..."):
                resultado = pd.merge(
                    df_unido, df_titulos,
                    left_on=['numero_extraido', 'Parcela'],
                    right_on=['COD_EXTERNO', 'Parcela'],
                    how='inner'
                )

            st.success(f"✅ Merge concluído com {len(resultado)} linhas")

            resultado = limpar_df(resultado)

            st.subheader("📊 Resultado do COMPENSADO")
            st.dataframe(resultado.head(100))

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Relatórios unidos", len(df_unido))
            with col2:
                st.metric("Títulos Cadastrados", len(df_titulos))
            with col3:
                st.metric("Compensados (match)", len(resultado))

            # 9. Download do Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Compensado')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Excel (compensado.xlsx)",
                data=excel_data,
                file_name="compensado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("✅ Processamento concluído com sucesso!")

        except Exception as e:
            st.error(f"❌ Erro durante o processamento: {str(e)}")
            st.exception(e)


# ==============================================================
# MAIN PAGE
# ==============================================================

st.set_page_config(page_title="Real Cobrança - Relatórios de Baixa", layout="wide")

st.title("💲 REAL COBRANÇA RELATÓRIOS DE BAIXA")
st.markdown("---")
st.markdown("Selecione qual dos credores deseja realizar a baixa de títulos, serão gerados arquivos formato excel para facilitação do processo")

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("🏦 RIO PAX", use_container_width=True):
        st.session_state["cliente"] = "RIO_PAX"
with col2:
    if st.button("🔄 REVIVER", use_container_width=True):
        st.session_state["cliente"] = "REVIVER"
with col3:
    if st.button("✅ COMPENSADO", use_container_width=True):
        st.session_state["cliente"] = "COMPENSADO"

if "cliente" not in st.session_state:
    st.session_state["cliente"] = None

st.markdown("---")

if st.session_state["cliente"] == "RIO_PAX":
    rio_pax_interface()
elif st.session_state["cliente"] == "REVIVER":
    reviver_interface()
elif st.session_state["cliente"] == "COMPENSADO":
    compensado_interface()
else:
    st.info("👈 Selecione um cliente acima para começar.")
