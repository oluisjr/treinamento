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

# Ordem correta das áreas
ordem_areas = [
    "CPIN - COORDENACAO DE PINTURA",
    "CQP/GS - COORDENAÇÃO DE QUALIDADE E PCP",
    "CQP-LAB - SUPERVISAO DE LABORATORIO",
    "GCZ- LCI/LIN - SUPERVISAO DE CORTE LONGITUDINAL",
    "GCZ- LCT/LPR/LBT - SUPERVISAO DE CORTE TRANSVERSAL",
    "GCZ- LLB - SUPERVISAO DE LAVADORA",
    "GCZ- LSL/LGT - SUPERVISAO DE SOLDA A LASER",
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

# === CONFIGURAÇÃO STREAMLIT ===
favicon = Image.open(FAVICON_PATH).resize((48, 24))
st.set_page_config(
    layout="wide",
    page_title="CSN - Treinamentos",
    page_icon=favicon
)
st.image(LOGO_PATH, width=180)
st.title("Painel de Treinamentos Pendentes")
if st.session_state.get('importacao_sucesso', False):
    st.success("Dados do Excel importados e aplicados com sucesso!")
    st.session_state['importacao_sucesso'] = False  # Limpa a flag para futuras operações

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
    df_raw["area"] = df_raw["area"].str.strip().str.upper()

    if area and area != "Todas":
        area = area.strip().upper()
        df_raw = df_raw[df_raw["area"] == area]

    df_raw = df_raw[df_raw["mes"].isin([mes1.upper(), mes2.upper()])]

    if df_raw.empty:
        st.warning("Nenhum dado correspondente aos meses selecionados.")
        return pd.DataFrame()

    df = df_raw.pivot_table(index='area', columns='mes', values=['em_dia', 'vencido'], fill_value=0)

    # Garante que as colunas fiquem na ordem mes1 -> mes2
    df = df[[('em_dia', mes1.upper()), ('vencido', mes1.upper()), ('em_dia', mes2.upper()), ('vencido', mes2.upper())]]

    df.columns = [f"{mes1.title()} (Em Dia)", f"{mes1.title()} (Vencido)", f"{mes2.title()} (Em Dia)", f"{mes2.title()} (Vencido)"]
    df.reset_index(inplace=True)

    df['area'] = pd.Categorical(df['area'], categories=ordem_areas, ordered=True)
    df = df.sort_values('area')

    if "area" not in df.columns:
        st.error("Erro: coluna 'area' ausente no DataFrame final.")
        st.stop()

    return df

def salvar_registro(area, mes, em_dia, vencido):
    area = area.strip().upper()
    mes = mes.strip().upper()

    resultado = supabase.table("treinamentos").select("id").match({"area": area, "mes": mes}).execute()

    if resultado.data:
        id_existente = resultado.data[0]['id']
        supabase.table("treinamentos").update({"em_dia": em_dia, "vencido": vencido}).eq("id", id_existente).execute()
    else:
        supabase.table("treinamentos").insert({"area": area, "mes": mes, "em_dia": em_dia, "vencido": vencido}).execute()

def gerar_grafico(df, mes1, mes2):
    df_plot = df[::-1].copy()
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
                ax.text(em / 2, i + offset, str(int(em)), ha="center", va="center", fontsize=7)
            if ven > 0:
                ax.text(em + ven / 2, i + offset, str(int(ven)), ha="center", va="center", fontsize=7)

    ax.set_yticks(y + bar_h / 2)
    ax.set_yticklabels(df_plot["area"], fontsize=8)
    ax.set_xticks([])
    ax.legend(fontsize=8)
    ax.grid(axis="x", linestyle="--", alpha=0.7)
    plt.tight_layout()
    return fig

def exportar_para_excel_pivo():
    response = supabase.table("treinamentos").select("*").execute()
    dados = response.data
    
    df_raw = pd.DataFrame(dados)
    if df_raw.empty:
        st.error("Nenhum dado encontrado na tabela 'treinamentos'.")
        return None

    df_raw.columns = df_raw.columns.str.lower()
    df_raw["mes"] = df_raw["mes"].str.strip().str.title()
    df_raw["area"] = df_raw["area"].str.strip().str.upper()

    df_pivo = df_raw.pivot_table(index='area', columns='mes', values=['em_dia', 'vencido'], fill_value=0)
    df_pivo = df_pivo.sort_index(axis=1, level=1)

    df_pivo.columns = [f"{mes} Em Dia" if tipo == 'em_dia' else f"{mes} Vencido" for tipo, mes in df_pivo.columns]
    df_pivo.reset_index(inplace=True)
    
    # Reordenar a coluna 'area'
    df_pivo['area'] = pd.Categorical(df_pivo['area'], categories=ordem_areas, ordered=True)
    df_pivo = df_pivo.sort_values('area')

    meses_ordem = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    colunas_ordenadas = ['area']
    for mes in meses_ordem:
        if f"{mes} Em Dia" in df_pivo.columns:
            colunas_ordenadas.append(f"{mes} Em Dia")
        if f"{mes} Vencido" in df_pivo.columns:
            colunas_ordenadas.append(f"{mes} Vencido")

    df_pivo = df_pivo[colunas_ordenadas]

    # Salvar gráfico como imagem temporária
    grafico_temp = "grafico_temp.png"
    fig.savefig(grafico_temp, bbox_inches='tight', dpi=300)
    plt.close(fig)

    # Criar arquivo Excel com gráfico
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pivo.to_excel(writer, index=False, sheet_name='Base Completa')

    # Reabrir para adicionar a imagem
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    img = ExcelImage(grafico_temp)
    img.width = 600
    img.height = 420
    img.anchor = "A2"
    ws.add_image(img)

    # Salvar novamente no buffer
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    os.remove(grafico_temp)

    return final_output
# === INTERFACE ===
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

col1, col2, col3 = st.columns(3)
mes1 = col1.selectbox("Escolha o 1º mês", meses, index=4)
mes2 = col2.selectbox("Escolha o 2º mês", meses, index=5)

dados_supabase = supabase.table("treinamentos").select("area").execute().data
areas_distintas = sorted(set(d['area'].strip().upper() for d in dados_supabase))
areas = ["Todas"] + areas_distintas
area_sel = col3.selectbox("Filtrar por área", areas)

if mes1 == mes2:
    st.warning("Selecione dois meses diferentes.")
else:
    df = carregar_dados(mes1, mes2, area_sel)
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        st.markdown("### Gráfico Comparativo")
        fig = gerar_grafico(df, mes1, mes2)
        st.pyplot(fig)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Baixar para Excel"):
                try:
                    excel_data = exportar_para_excel_pivo()
                    if excel_data:
                        st.download_button("Download Excel Completo", excel_data, file_name="treinamentos_completo.xlsx")
                except Exception as e:
                    st.error(f"Erro ao exportar: {e}")

# === ÁREA PROTEGIDA PARA EDIÇÃO ===
with st.expander("Editar dados (restrito)", expanded=st.session_state["autenticado"]):
    if not st.session_state["autenticado"]:
        senha = st.text_input("Senha de edição", type="password")
        if senha == SENHA_EDICAO:
            st.success("Acesso liberado")
            st.session_state["autenticado"] = True
            st.rerun()
        elif senha:
            st.error("Senha incorreta")
    else:
        if st.button("Encerrar sessão de edição"):
            st.session_state["autenticado"] = False
            st.rerun()

        with col2:
                    st.markdown("### Importar Excel para atualização de dados")
                    uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=["xlsx"])

                    if uploaded_file is not None:
                        df_importado = pd.read_excel(uploaded_file)
                        df_importado.columns = df_importado.columns.str.strip().str.lower()

                        if "area" in df_importado.columns:
                            meses_base = [col.split()[0] for col in df_importado.columns if "em dia" in col or "vencido" in col]
                            meses_base = sorted(set(meses_base))

                            for _, row in df_importado.iterrows():
                                area = str(row["area"]).strip().upper()

                                for mes in meses_base:
                                    em_dia_col = f"{mes.lower()} em dia"
                                    vencido_col = f"{mes.lower()} vencido"

                                    if em_dia_col in df_importado.columns and vencido_col in df_importado.columns:
                                        em_dia = int(row[em_dia_col]) if not pd.isna(row[em_dia_col]) else 0
                                        vencido = int(row[vencido_col]) if not pd.isna(row[vencido_col]) else 0

                                        salvar_registro(area, mes, em_dia, vencido)
                            st.session_state['importacao_sucesso'] = True  # Flag temporária
                            st.session_state['autenticado'] = False
                            st.rerun()
                        else:
                            st.error(f"O arquivo Excel deve conter a coluna: 'area'.")
        # === EDIÇÃO MANUAL ===
        mes_edicao = st.selectbox("Mês para editar", meses)
        area = st.selectbox("Nome da área", areas)
        em_dia = st.number_input("Em dia", min_value=0, step=1)
        vencido = st.number_input("Vencido", min_value=0, step=1)

        if st.button("Salvar dados", key="botao_salvar"):
            salvar_registro(area, mes_edicao, em_dia, vencido)
            st.success("Dados atualizados com sucesso!")
