import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from dotenv import load_dotenv
import os
from PIL import Image
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from supabase import create_client, Client
from unidecode import unidecode

load_dotenv()

# === CONEXÃO SUPABASE ===
SUPABASE_URL = os.getenv('SUPABASE_URL')
SUPABASE_KEY = os.getenv('SUPABASE_KEY')
SENHA_EDICAO = os.getenv('SENHA_EDICAO')

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

url = SUPABASE_URL
key = SUPABASE_KEY 

LOGO_PATH="LogoCSN_Cinza.png"
FAVICON_PATH="Favicon.png"

# --- NOVA SEÇÃO: NORMALIZAÇÃO E GERENCIAMENTO DE NOMES DE ÁREA ---

# Função auxiliar para normalizar o texto (remove acentos, espaços extras e converte para maiúsculas)
def normalizar_texto(texto):
    if texto is None:
        return ""
    return unidecode(str(texto)).strip().upper()

# 1. Lista principal para EXIBIÇÃO (com acentos e grafia correta)
ordem_areas_display = [
    "CPIN - COORDENAÇÃO DE PINTURA",
    "CQP/GS - COORDENAÇÃO DE QUALIDADE E PCP",
    "CQP-LAB - SUPERVISAO DE LABORATORIO",
    "GCZ- LCI/LIN - SUPERVISAO DE CORTE LONGITUDINAL",
    "GCZ- LCT/LPR/LBT - SUPERVISÃO DE CORTE TRANSVERSAL",
    "GCZ- LLB - SUPERVISAO DE LAVADORA",
    "GCZ- LSL/LGT - SUPERVISÃO DE SOLDA A LASER",
    "GCZ/GS - GERENCIA CENTRO DE SERVICO E PINTURA",
    "GCZ-CS - SUPERVISÃO DE CENTRO DE SERVIÇOS",
    "GDM/GS - GERENCIA DE MANUTENCAO",
    "GDM-INSPELE - SUPERVISAO DE INSPECAO ELETRICA",
    "GDM-INSPMEC - SUPERVISAO DE INSPECAO MECANICA",
    "GGOP/GS - GERENCIA GERAL DE OPERACOES PORTO REAL",
    "GPR-PLANPROG - SUPERVISAO DE PLANEJAMENTO E PROGRAMACAO",
    "GZL/GS - GERENCIA DE ZINCAGEM E LOGÍSTICA",
    "GZL-EMB - SUPERVISAO DE EMBALAGEM",
    "GZL-LOG - SUPERVISAO DE LOGISTICA",
    "GZL-ZINCAGEM - SUPERVISAO DA ZINCAGEM",
    "TOTAL GERAL"
]

# 2. Lista derivada para PROCESSAMENTO (sem acentos, usada pela lógica interna)
ordem_areas_processamento = [normalizar_texto(area) for area in ordem_areas_display]

# 3. Dicionário para mapear do nome de processamento de volta para o nome de exibição
mapa_nomes = {normalizar_texto(area): area for area in ordem_areas_display}


# === CONFIGURAÇÃO STREAMLIT ===
st.set_page_config(
    layout="wide",
    page_title="Treinamentos")

st.title("Painel de Treinamentos Pendentes")
if st.session_state.get('importacao_sucesso', False):
    st.success("Dados do Excel importados e aplicados com sucesso!")
    st.session_state['importacao_sucesso'] = False

# === VARIÁVEL DE SESSÃO ===
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

# === FUNÇÕES ===
def carregar_dados(mes1, mes2, area=None):
    response = supabase.table("treinamentos").select("*").execute()
    dados = response.data

    df_raw = pd.DataFrame(dados)
    if df_raw.empty:
        st.warning("Nenhum dado encontrado na tabela 'treinamentos'.")
        return pd.DataFrame()

    df_raw.columns = df_raw.columns.str.lower()
    df_raw["mes"] = df_raw["mes"].str.strip().str.upper()
    
    # Normaliza a coluna 'area' para processamento
    df_raw["area_proc"] = df_raw["area"].apply(normalizar_texto)

    if area and area != "Todas":
        area_proc_sel = normalizar_texto(area)
        df_raw = df_raw[df_raw["area_proc"] == area_proc_sel]

    df_raw = df_raw[df_raw["mes"].isin([mes1.upper(), mes2.upper()])]

    if df_raw.empty:
        st.warning("Nenhum dado correspondente aos meses selecionados.")
        return pd.DataFrame()

    # Usa a coluna normalizada 'area_proc' para o pivot
    df = df_raw.pivot_table(index='area_proc', columns='mes', values=['em_dia', 'vencido'], fill_value=0)

    # Garante que as colunas fiquem na ordem mes1 -> mes2
    df = df[[('em_dia', mes1.upper()), ('vencido', mes1.upper()), ('em_dia', mes2.upper()), ('vencido', mes2.upper())]]

    df.columns = [f"{mes1.title()} (Em Dia)", f"{mes1.title()} (Vencido)", f"{mes2.title()} (Em Dia)", f"{mes2.title()} (Vencido)"]
    df.reset_index(inplace=True)

    # Usa a lista de processamento para categorizar e ordenar
    df['area_proc'] = pd.Categorical(df['area_proc'], categories=ordem_areas_processamento, ordered=True)
    df = df.sort_values('area_proc')

    # Mapeia de volta para os nomes com acento para exibição
    df['area'] = df['area_proc'].map(mapa_nomes)

    # Reordena e limpa as colunas para a exibição final
    colunas_finais = ['area'] + [col for col in df.columns if col not in ['area', 'area_proc']]
    df = df[colunas_finais]

    if "area" not in df.columns:
        st.error("Erro: coluna 'area' ausente no DataFrame final.")
        st.stop()

    return df

def salvar_registro(area, mes, em_dia, vencido):
    # Garante que a área seja salva com a grafia correta (com acentos)
    area_normalizada = normalizar_texto(area)
    area_correta = mapa_nomes.get(area_normalizada, area.strip().upper())
    
    mes_proc = mes.strip().upper()

    resultado = supabase.table("treinamentos").select("id").match({"area": area_correta, "mes": mes_proc}).execute()

    if resultado.data:
        id_existente = resultado.data[0]['id']
        supabase.table("treinamentos").update({"em_dia": em_dia, "vencido": vencido}).eq("id", id_existente).execute()
    else:
        supabase.table("treinamentos").insert({"area": area_correta, "mes": mes_proc, "em_dia": em_dia, "vencido": vencido}).execute()

def gerar_grafico(df, mes1, mes2):
    # Renomeia temporariamente a coluna 'area' para o plot não cortar nomes longos
    df_plot = df.rename(columns={'area': 'Área'})[::-1].copy()

    fig, ax = plt.subplots(figsize=(16, len(df_plot) * 0.35))
    y = np.arange(len(df_plot))
    bar_h = 0.4

    for idx, (mes, cor1, cor2) in enumerate([(mes1, "#c5e0b4", "#ff7357"), (mes2, "#759a64", "#af2d11")]):
        col_em_dia = f"{mes.title()} (Em Dia)"
        col_vencido = f"{mes.title()} (Vencido)"
        offset = (1 - idx) * bar_h

        if col_em_dia not in df_plot.columns or col_vencido not in df_plot.columns:
            st.warning(f"Colunas para o mês '{mes}' não foram encontradas.")
            continue

        em_dia_vals = df_plot[col_em_dia]
        vencido_vals = df_plot[col_vencido]

        ax.barh(y + offset, em_dia_vals, height=bar_h, label=f"{mes.title()} Em Dia", color=cor1)
        ax.barh(y + offset, vencido_vals, height=bar_h, left=em_dia_vals, label=f"{mes.title()} Vencido", color=cor2)

        for i, (em, ven) in enumerate(zip(em_dia_vals, vencido_vals)):
            if em > 0:
                ax.text(em / 2, i + offset, str(int(em)), ha="center", va="center", fontsize=7, color='black')
            if ven > 0:
                ax.text(em + ven / 2, i + offset, str(int(ven)), ha="center", va="center", fontsize=7, color='black')

    ax.set_yticks(y + bar_h / 2)
    ax.set_yticklabels(df_plot["Área"], fontsize=8)
    ax.set_xticks([])
    ax.legend(fontsize=8)
    ax.grid(axis="x", linestyle="--", alpha=0.7)
    plt.tight_layout()
    return fig

def exportar_para_excel_com_grafico(fig):
    # --- Início: Lógica para gerar a tabela de dados (igual à sua função original) ---
    response = supabase.table("treinamentos").select("*").execute()
    dados = response.data
    
    df_raw = pd.DataFrame(dados)
    if df_raw.empty:
        st.error("Nenhum dado encontrado na tabela 'treinamentos' para exportar.")
        return None

    df_raw.columns = df_raw.columns.str.lower()
    df_raw["mes"] = df_raw["mes"].str.strip().str.title()
    df_raw["area_proc"] = df_raw["area"].apply(normalizar_texto)

    df_pivo = df_raw.pivot_table(index='area_proc', columns='mes', values=['em_dia', 'vencido'], fill_value=0)
    df_pivo = df_pivo.sort_index(axis=1, level=1)

    df_pivo.columns = [f"{mes} Em Dia" if tipo == 'em_dia' else f"{mes} Vencido" for tipo, mes in df_pivo.columns]
    df_pivo.reset_index(inplace=True)
    
    df_pivo['area_proc'] = pd.Categorical(df_pivo['area_proc'], categories=ordem_areas_processamento, ordered=True)
    df_pivo = df_pivo.sort_values('area_proc')
    
    df_pivo['area'] = df_pivo['area_proc'].map(mapa_nomes)
    
    meses_ordem = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    colunas_ordenadas = ['area']
    for mes in meses_ordem:
        if f"{mes} Em Dia" in df_pivo.columns:
            colunas_ordenadas.append(f"{mes} Em Dia")
        if f"{mes} Vencido" in df_pivo.columns:
            colunas_ordenadas.append(f"{mes} Vencido")

    # Garante que apenas colunas existentes sejam selecionadas
    colunas_finais = [col for col in colunas_ordenadas if col in df_pivo.columns]
    df_pivo = df_pivo[colunas_finais]

    # Converte o DataFrame para um objeto Excel na memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pivo.to_excel(writer, index=False, sheet_name='Base Completa')
        
        # Carrega o workbook para adicionar a imagem
        workbook = writer.book
        worksheet = workbook.create_sheet(title='Gráfico')

        # Salva o gráfico que foi passado como argumento para um buffer de imagem
        img_buffer = BytesIO()
        fig.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)
        
        # Cria um objeto de imagem Excel a partir do buffer
        img = ExcelImage(img_buffer)

        # Define o tamanho da imagem para ocupar aproximadamente 20x15 células
        # Ajustes podem ser necessários dependendo da largura/altura padrão das colunas/linhas
        img.height = 300  # Altura em pixels (aprox. 20 linhas * 15 pontos/linha)
        img.width = 900   # Largura em pixels (aprox. 15 colunas * 60 pontos/coluna)
        
        # Adiciona a imagem à planilha 'Gráfico', ancorada na célula A1
        worksheet.add_image(img, 'A1')

    output.seek(0)
    return output

# === INTERFACE ===
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

col1, col2, col3 = st.columns(3)
mes1 = col1.selectbox("Escolha o 1º mês", meses, index=datetime.now().month - 2)
mes2 = col2.selectbox("Escolha o 2º mês", meses, index=datetime.now().month - 1)

# A lista de áreas para o filtro continua vindo do banco, mas a lógica de filtro é normalizada
dados_supabase = supabase.table("treinamentos").select("area").execute().data
areas_distintas = sorted(list(set(d['area'].strip() for d in dados_supabase if d['area'])))
areas = ["Todas"] + areas_distintas
area_sel = col3.selectbox("Filtrar por área", areas)

if mes1 == mes2:
    st.warning("Selecione dois meses diferentes.")
else:
    df = carregar_dados(mes1, mes2, area_sel)
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.markdown("### Gráfico Comparativo")
        fig = gerar_grafico(df, mes1, mes2)
        st.pyplot(fig)

        col1_exp, col2_exp = st.columns(2)
        with col1_exp:
            try:
                # A variável 'fig' já foi criada por gerar_grafico() e exibida com st.pyplot()
                # Agora passamos a mesma 'fig' para a função de exportação
                excel_data_com_grafico = exportar_para_excel_com_grafico(fig) 
                if excel_data_com_grafico:
                    st.download_button(
                        label="Baixar Excel", # NOVO LABEL
                        data=excel_data_com_grafico,
                        file_name="relatorio_treinamentos_com_grafico.xlsx", # NOVO NOME
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Erro ao gerar Excel com gráfico: {e}")

# === ÁREA PROTEGIDA PARA EDIÇÃO ===
with st.expander("Editar dados (restrito)", expanded=st.session_state["autenticado"]):
    if not st.session_state["autenticado"]:
        senha = st.text_input("Senha de edição", type="password", key="senha_login")
        if senha == SENHA_EDICAO:
            st.session_state["autenticado"] = True
            st.rerun()
        elif senha:
            st.error("Senha incorreta")
    else:
        st.info("Sessão de edição iniciada.")
        
        # === EDIÇÃO MANUAL ===
        st.markdown("#### Edição Manual")
        col_ed1, col_ed2 = st.columns(2)
        with col_ed1:
            mes_edicao = st.selectbox("Mês para editar", meses, key="mes_edicao")
            area_edicao = st.selectbox("Nome da área", areas_distintas, key="area_edicao")
        with col_ed2:
            em_dia = st.number_input("Em dia", min_value=0, step=1, key="em_dia_edicao")
            vencido = st.number_input("Vencido", min_value=0, step=1, key="vencido_edicao")

        if st.button("Salvar dados", key="botao_salvar"):
            salvar_registro(area_edicao, mes_edicao, em_dia, vencido)
            st.success(f"Dados de '{area_edicao}' para o mês '{mes_edicao}' atualizados com sucesso!")

        st.divider()
        
        # === IMPORTAÇÃO EXCEL ===
        st.markdown("#### Importar arquivo Excel (.xlsx)")
        uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=["xlsx"])

        if uploaded_file is not None:
            df_importado = pd.read_excel(uploaded_file)
            df_importado.columns = [col.strip().lower() for col in df_importado.columns]

            if "area" in df_importado.columns:
                meses_base = sorted(list(set([col.split()[0] for col in df_importado.columns if "em dia" in col or "vencido" in col])))
                
                with st.spinner('Importando dados...'):
                    for _, row in df_importado.iterrows():
                        area = str(row["area"])
                        for mes in meses_base:
                            em_dia_col = f"{mes.lower()} em dia"
                            vencido_col = f"{mes.lower()} vencido"

                            if em_dia_col in df_importado.columns and vencido_col in df_importado.columns:
                                em_dia = int(row[em_dia_col]) if pd.notna(row[em_dia_col]) else 0
                                vencido = int(row[vencido_col]) if pd.notna(row[vencido_col]) else 0
                                salvar_registro(area, mes.title(), em_dia, vencido)
                
                st.session_state['importacao_sucesso'] = True
                st.session_state["autenticado"] = False # Desloga por segurança após importação
                st.rerun()
            else:
                st.error("O arquivo Excel deve conter uma coluna chamada 'area'.")

        if st.button("Encerrar sessão de edição", type="primary"):
            st.session_state["autenticado"] = False
            st.rerun()
