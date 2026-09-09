import streamlit as st
import pandas as pd
import requests
import re
import time
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURAÇÃO DA PÁGINA (Interface Streamlit)
# ==========================================
st.set_page_config(page_title="Buscador PNCP", page_icon="🔍", layout="wide")

# ==========================================
# FUNÇÕES DE API E CACHE (IBGE)
# ==========================================
@st.cache_data(show_spinner=False)
def buscar_estados():
    try:
        url = "https://servicodados.ibge.gov.br/api/v1/localidades/estados"
        resp = requests.get(url, timeout=10).json()
        estados = {est['sigla']: est['nome'] for est in resp}
        return dict(sorted(estados.items()))
    except:
        return {}

@st.cache_data(show_spinner=False)
def buscar_municipios(uf):
    try:
        url = f"https://servicodados.ibge.gov.br/api/v1/localidades/estados/{uf}/municipios"
        resp = requests.get(url, timeout=10).json()
        municipios = {mun['nome']: mun['id'] for mun in resp}
        return dict(sorted(municipios.items()))
    except:
        return {}

# ==========================================
# FUNÇÕES DE DADOS (Processamento)
# ==========================================
def limpar_dados(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
        
    for col in df.select_dtypes(include=['object', 'string', 'str']).columns:
        df[col] = df[col].apply(
            lambda x: re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', str(x)) if pd.notnull(x) else x
        )
        
    df['valor_estimado'] = pd.to_numeric(df['valor_estimado'], errors='coerce')
    df['valor_homologado'] = pd.to_numeric(df['valor_homologado'], errors='coerce')
    
    colunas_data = ['abertura', 'inclusao', 'encerramento']
    for col in colunas_data:
        df[col] = pd.to_datetime(df[col], format='%Y-%m-%dT%H:%M:%S', errors='coerce').dt.tz_localize(None)
        
    return df

def aplicar_filtro_data_estrito(df: pd.DataFrame, dt_ini, dt_fim, coluna: str) -> pd.DataFrame:
    if df.empty or coluna not in df.columns:
        return df
    
    dt_fim_ajustada = pd.to_datetime(dt_fim) + pd.Timedelta(days=1, seconds=-1)
    filtro = (df[coluna] >= pd.to_datetime(dt_ini)) & (df[coluna] <= dt_fim_ajustada)
    df = df.dropna(subset=[coluna])
    
    return df[filtro].reset_index(drop=True)

def filtrar_por_palavras(df: pd.DataFrame, palavras_str: str) -> pd.DataFrame:
    if df.empty or 'objeto' not in df.columns or not palavras_str.strip():
        return pd.DataFrame()
        
    palavras = [p.strip().lower() for p in palavras_str.split(',')]
    padrao_regex = '|'.join(palavras)
    
    filtro = df['objeto'].str.lower().str.contains(padrao_regex, na=False)
    return df[filtro].reset_index(drop=True)

def gerar_excel(df_completo: pd.DataFrame, df_filtrado: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_completo.to_excel(writer, sheet_name='Todos', index=False)
        
        if not df_filtrado.empty:
            df_filtrado.to_excel(writer, sheet_name='Filtrados', index=False)
            
        workbook = writer.book
        
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        link_font = Font(color="0563C1", underline="single")
        border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
        alignment_padrao = Alignment(vertical='top', wrap_text=True)
        
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = "A2"
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            for col_idx, col in enumerate(ws.columns, 1):
                col_letter = get_column_letter(col_idx)
                header_value = col[0].value
                
                if header_value in ['objeto', 'amparo_legal']: ws.column_dimensions[col_letter].width = 60
                elif header_value == 'link':
                    ws.column_dimensions[col_letter].width = 30
                    for cell in col[1:]:
                        if cell.value:
                            try:
                                cell.hyperlink = str(cell.value)
                                cell.font = link_font
                            except: pass
                elif header_value in ['orgao', 'plataforma', 'disputa']: ws.column_dimensions[col_letter].width = 30
                elif header_value in ['valor_estimado', 'valor_homologado']:
                    ws.column_dimensions[col_letter].width = 18
                    for cell in col[1:]:
                        if isinstance(cell.value, (int, float)): cell.number_format = 'R$ #,##0.00'
                elif header_value in ['abertura', 'inclusao', 'encerramento']:
                    ws.column_dimensions[col_letter].width = 18
                    for cell in col[1:]:
                        if cell.value: cell.number_format = 'DD/MM/YYYY HH:MM'
                else: ws.column_dimensions[col_letter].width = 15
                
                for cell in col[1:]:
                    cell.border = border
                    cell.alignment = alignment_padrao
                    
    return output.getvalue()

# ==========================================
# INTERFACE DO USUÁRIO (Sidebar / Menus)
# ==========================================
st.sidebar.title("⚙️ Configurações de Busca")

st.sidebar.header("🏢 Localidade")
estados = buscar_estados()
opcoes_estados = ["Brasil (Todos)"] + list(estados.keys())
uf_selecionada = st.sidebar.selectbox("Estado (UF)", opcoes_estados)

codigo_municipio = ""
uf_api = ""

if uf_selecionada != "Brasil (Todos)":
    uf_api = uf_selecionada
    municipios = buscar_municipios(uf_selecionada)
    opcoes_mun = ["Estado Inteiro (Todos)"] + list(municipios.keys())
    mun_selecionado = st.sidebar.selectbox("Município", opcoes_mun)
    
    if mun_selecionado != "Estado Inteiro (Todos)":
        codigo_municipio = str(municipios[mun_selecionado])

st.sidebar.header("📅 Período")
tipo_data = st.sidebar.radio("Essas datas se referem a:", ["Data de Abertura (Recomendado)", "Data de Publicação no PNCP"])
coluna_filtro = 'abertura' if "Abertura" in tipo_data else 'inclusao'

col1, col2 = st.sidebar.columns(2)
hoje = datetime.now()
data_ini = col1.date_input("Início", hoje)
data_fim = col2.date_input("Fim", datetime(hoje.year, 12, 31))

st.sidebar.header("📋 Modalidade")
modalidades = {
    "1 - Leilão": 1,
    "2 - Diálogo Competitivo": 2,
    "3 - Concurso": 3,
    "4 - Concorrência": 4,
    "5 - Pregão Eletrônico": 5,
    "6 - Dispensa de Licitação": 6,
    "7 - Inexigibilidade": 7,
    "8 - Manifestação de Interesse": 8
}
mod_nome = st.sidebar.selectbox("Selecione a Modalidade", list(modalidades.keys()), index=4)
mod_id = modalidades[mod_nome]

st.sidebar.header("🔑 Filtros Extras")
palavras_chave = st.sidebar.text_input("Palavras-chave (separadas por vírgula)", placeholder="Ex: alimento, comida")
max_paginas = st.sidebar.number_input("Máximo de Páginas para buscar", min_value=1, max_value=200, value=10, step=5)

buscar_btn = st.sidebar.button("🚀 Buscar Licitações", type="primary", use_container_width=True)

# ==========================================
# ÁREA PRINCIPAL DA TELA
# ==========================================
st.title("🔎 Buscador de Licitações PNCP")
st.markdown("Encontre licitações de todo o Brasil, visualize os dados aqui mesmo e baixe uma planilha Excel formatada e pronta para uso.")

if buscar_btn:
    if coluna_filtro == 'abertura':
        dt_ini_api = data_ini - timedelta(days=45)
        data_inicial_api_str = dt_ini_api.strftime("%Y%m%d")
    else:
        data_inicial_api_str = data_ini.strftime("%Y%m%d")
        
    data_final_api_str = data_fim.strftime("%Y%m%d")
    
    processos = []
    base_url = 'https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao'
    
    progress_bar = st.progress(0, text="Iniciando busca na API...")
    houve_erro = False
    
    for pagina in range(1, max_paginas + 1):
        # ====================================================
        # CORREÇÃO DEFINITIVA: O PARÂMETRO É 'modalidadeId' 
        # ====================================================
        parametros = {
            'dataInicial': data_inicial_api_str,
            'dataFinal': data_final_api_str,
            'modalidadeId': mod_id, 
            'tamanhoPagina': 50,
            'pagina': pagina
        }
        if uf_api: parametros['uf'] = uf_api
        if codigo_municipio: parametros['codigoMunicipioIbge'] = codigo_municipio

        sucesso_pagina = False
        tentativas = 0
        dados_dict = []
        
        while not sucesso_pagina and tentativas < 3:
            try:
                msg_progresso = f"Baixando página {pagina} de {max_paginas}..." if tentativas == 0 else f"Site lento. Re-tentando página {pagina} (Tentativa {tentativas+1}/3)..."
                progress_bar.progress(pagina / max_paginas, text=msg_progresso)
                
                resp = requests.get(base_url, params=parametros, timeout=45)
                
                if resp.status_code == 200:
                    dados_dict = resp.json().get('data', [])
                    sucesso_pagina = True
                elif resp.status_code == 204:
                    # 204 significa "Sem Conteúdo" (zero licitações).
                    dados_dict = []
                    sucesso_pagina = True
                else:
                    st.error(f"❌ O Servidor do Governo recusou a pesquisa (Status {resp.status_code}). Detalhes: {resp.text}")
                    houve_erro = True
                    break
                    
            except requests.exceptions.Timeout:
                tentativas += 1
                if tentativas < 3:
                    st.toast(f"⏳ O site do governo está muito lento. Tentando novamente... ({tentativas}/3)", icon="🔄")
                    time.sleep(3) 
                else:
                    st.error(f"❌ O site do governo não respondeu após 3 tentativas (Timeout). Tente buscar um período menor ou um Estado específico.")
                    houve_erro = True
                    
            except Exception as e:
                st.error(f"❌ Erro inesperado ao conectar com o governo na página {pagina}: {e}")
                houve_erro = True
                break
                
        if houve_erro:
            break
            
        if not dados_dict and sucesso_pagina:
            break # Fim das páginas
            
        if sucesso_pagina:
            for proc in dados_dict:
                processos.append({
                    'orgao': proc.get('orgaoEntidade', {}).get('razaoSocial'),
                    'uf': proc.get('unidadeOrgao', {}).get('ufSigla'),
                    'inclusao': proc.get('dataInclusao'),
                    'abertura': proc.get('dataAberturaProposta'),
                    'objeto': proc.get('objetoCompra'),
                    'valor_estimado': proc.get('valorTotalEstimado'),
                    'situacao': proc.get('situacaoCompraNome'),
                    'link': proc.get('linkSistemaOrigem'),
                    'disputa': proc.get('modoDisputaNome'),
                    'plataforma': proc.get('usuarioNome'),
                    'amparo_legal': proc.get('amparoLegal', {}).get('descricao'),
                    'encerramento': proc.get('dataEncerramentoProposta'),
                    'sequencial': proc.get('sequencialCompra'),
                    'n_processo': proc.get('processo'),
                    'valor_homologado': proc.get('valorTotalHomologado'),
                    'srp': proc.get('srp')
                })

    progress_bar.empty() 
    
    if not houve_erro:
        if not processos:
            st.warning("Nenhuma licitação encontrada com os filtros selecionados.")
        else:
            df_geral = pd.DataFrame(processos)
            df_geral = limpar_dados(df_geral)
            df_geral = aplicar_filtro_data_estrito(df_geral, data_ini, data_fim, coluna_filtro)
            df_filtrado = filtrar_por_palavras(df_geral, palavras_chave)
            
            df_para_exibir = df_filtrado if palavras_chave.strip() and not df_filtrado.empty else df_geral
            
            if df_para_exibir.empty:
                st.warning("Dados encontrados, mas nenhum sobreviveu ao filtro de Data Exata ou Palavra-Chave.")
            else:
                st.success(f"🎉 Busca concluída! {len(df_para_exibir)} licitações encontradas prontas para análise.")
                
                st.dataframe(
                    df_para_exibir[['orgao', 'uf', 'objeto', 'abertura', 'valor_estimado', 'link']],
                    column_config={
                        "link": st.column_config.LinkColumn("Link Oficial", display_text="Abrir Edital"),
                        "valor_estimado": st.column_config.NumberColumn("Valor Estimado", format="R$ %.2f"),
                        "abertura": st.column_config.DatetimeColumn("Abertura", format="DD/MM/YYYY HH:mm"),
                    },
                    hide_index=True,
                    use_container_width=True
                )
                
                arquivo_excel = gerar_excel(df_geral, df_filtrado)
                st.download_button(
                    label="📥 Baixar Planilha Completa (Excel)",
                    data=arquivo_excel,
                    file_name=f"licitacoes_pncp_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )