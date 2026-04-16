import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import os
import time
import base64


def get_logo_base64(path="logo_normatel.png"):
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

st.set_page_config(
    page_title="Dashboard Materiais | Normatel",
    page_icon="logo_normatel.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css"/>
<style>
    .main .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 100%; overflow-x: hidden; }
    [data-testid="stPlotlyChart"] { overflow: hidden; }
    .kpi-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border-radius: 12px; padding: 20px; text-align: center;
        border-left: 4px solid #4CAF50; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin-bottom: 10px;
    }
    .kpi-card h2 { color: #4CAF50; font-size: 2rem; margin: 0; font-weight: 700; }
    .kpi-card p { color: #b0b0b0; font-size: 0.85rem; margin: 5px 0 0 0; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-card-warn { border-left: 4px solid #FF9800; }
    .kpi-card-warn h2 { color: #FF9800; }
    .kpi-card-info { border-left: 4px solid #2196F3; }
    .kpi-card-info h2 { color: #2196F3; }
    .kpi-card-danger { border-left: 4px solid #f44336; }
    .kpi-card-danger h2 { color: #f44336; }
    .main-header {
        background: linear-gradient(90deg, #2e7d32 0%, #43a047 100%);
        color: white; padding: 15px 25px; border-radius: 10px; margin-bottom: 20px; text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; letter-spacing: 2px; }
    .main-header p { margin: 5px 0 0 0; font-size: 0.9rem; opacity: 0.9; }
    .section-header {
        background: #2e7d32; color: white; padding: 8px 15px; border-radius: 6px;
        font-size: 0.9rem; font-weight: 600; letter-spacing: 1px; margin: 15px 0 10px 0;
    }
    .section-header i { margin-right: 7px; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%); }
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3 { color: #4CAF50; }
    [data-testid="stSidebar"] .stMarkdown i { margin-right: 6px; }
    .dataframe { font-size: 0.8rem !important; }
    .refresh-indicator {
        background: #1a1a2e; border: 1px solid #4CAF50; border-radius: 8px;
        padding: 8px 15px; color: #4CAF50; font-size: 0.8rem; text-align: center;
    }
</style>
""", unsafe_allow_html=True)

COLORS = {
    "primary": "#4CAF50", "secondary": "#2196F3", "warning": "#FF9800",
    "danger": "#f44336", "dark": "#1a1a2e",
    "chart_palette": ["#4CAF50", "#2196F3", "#FF9800", "#f44336", "#9C27B0",
                       "#00BCD4", "#795548", "#607D8B", "#E91E63", "#CDDC39",
                       "#FF5722", "#3F51B5"],
    "sequential": px.colors.sequential.Greens,
}


def format_brl(value):
    if abs(value) >= 1_000_000: return f"R$ {value/1_000_000:,.2f}M"
    elif abs(value) >= 1_000: return f"R$ {value/1_000:,.1f}K"
    else: return f"R$ {value:,.2f}"

def format_qty(value):
    if abs(value) >= 1_000_000: return f"{value/1_000_000:,.1f}M"
    elif abs(value) >= 1_000: return f"{value/1_000:,.1f}K"
    else: return f"{value:,.0f}"

@st.cache_data(ttl=30)
def load_data(file_path):
    df = pd.read_excel(file_path, sheet_name="BASE MATERIAIS (2)", engine="openpyxl")
    df.columns = df.columns.str.strip()
    if "DATAEMISSAO" in df.columns:
        df["DATAEMISSAO"] = pd.to_datetime(df["DATAEMISSAO"], errors="coerce")
    if "RECCREATEDON" in df.columns:
        df["RECCREATEDON"] = pd.to_datetime(df["RECCREATEDON"], errors="coerce")
    for col in ["QUANTIDADE", "VALOR"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "DATAEMISSAO" in df.columns:
        df["MES"] = df["DATAEMISSAO"].dt.to_period("M").astype(str)
        df["MES_ANO"] = df["DATAEMISSAO"].dt.strftime("%m/%Y")
        df["ANO"] = df["DATAEMISSAO"].dt.year
        df["TRIMESTRE"] = df["DATAEMISSAO"].dt.to_period("Q").astype(str)
        df["SEMANA"] = df["DATAEMISSAO"].dt.isocalendar().week.fillna(0).astype(int)
        df["DIA_SEMANA"] = df["DATAEMISSAO"].dt.day_name()
    df["VALOR_UNITARIO"] = np.where(df["QUANTIDADE"] > 0, df["VALOR"] / df["QUANTIDADE"], 0)
    df["BASE"] = df["PROJETO"].apply(
        lambda x: "UTGCAB - BARRA DO FURADO E SEVERINA" if "CABIUNAS" in str(x).upper() and "UTE" not in str(x).upper()
        else "UTE-TMA, TAPERA e ÁREAS EXTERNAS" if "UTE" in str(x).upper()
        else str(x)
    )
    return df


def render_sidebar(df):
    with st.sidebar:
        if os.path.exists("logo_normatel.png"):
            st.image("logo_normatel.png", use_container_width=True)
        st.markdown("## Normatel Engenharia")
        st.markdown("---")
        st.markdown('<h3><i class="fa-solid fa-gear"></i> Configurações</h3>', unsafe_allow_html=True)
        auto_refresh = st.toggle("Auto-refresh (30s)", value=False)
        if auto_refresh:
            st.markdown('<div class="refresh-indicator"><i class="fa-solid fa-circle" style="color:#4CAF50;font-size:0.6rem;"></i> Atualizando a cada 30s</div>', unsafe_allow_html=True)
        st.markdown("---")
        st.markdown('<h3><i class="fa-solid fa-sliders"></i> Filtros</h3>', unsafe_allow_html=True)
        if "DATAEMISSAO" in df.columns:
            min_date = df["DATAEMISSAO"].min()
            max_date = df["DATAEMISSAO"].max()
            if pd.notna(min_date) and pd.notna(max_date):
                date_range = st.date_input("Período", value=(min_date.date(), max_date.date()),
                    min_value=min_date.date(), max_value=max_date.date())
            else:
                date_range = None
        else:
            date_range = None
        bases = ["Todos"] + sorted(df["BASE"].dropna().unique().tolist())
        selected_base = st.selectbox("Base / Agrupamento", bases)
        disciplinas = ["Todas"] + sorted(df["DISCIPLINA"].dropna().unique().tolist())
        selected_disciplina = st.selectbox("Disciplina", disciplinas)
        tipos = ["Todos"] + sorted(df["TIPO"].dropna().unique().tolist())
        selected_tipo = st.selectbox("Tipo de Material", tipos)
        classificacoes = ["Todas"] + sorted(df["CLASSIFICAÇÃO DO MATERIAL"].dropna().unique().tolist())
        selected_classif = st.selectbox("Classificação", classificacoes)
        fornecedores = ["Todos"] + sorted(df["NOMEFANTASIA"].dropna().unique().tolist())
        selected_fornecedor = st.selectbox("Fornecedor", fornecedores)
        st.markdown("---")
        st.markdown(f"**Última atualização:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        st.markdown(f"**Total de registros:** {len(df):,}")
        mask = pd.Series([True] * len(df), index=df.index)
        if date_range and len(date_range) == 2:
            mask &= (df["DATAEMISSAO"].dt.date >= date_range[0]) & (df["DATAEMISSAO"].dt.date <= date_range[1])
        if selected_base != "Todos": mask &= df["BASE"] == selected_base
        if selected_disciplina != "Todas": mask &= df["DISCIPLINA"] == selected_disciplina
        if selected_tipo != "Todos": mask &= df["TIPO"] == selected_tipo
        if selected_classif != "Todas": mask &= df["CLASSIFICAÇÃO DO MATERIAL"] == selected_classif
        if selected_fornecedor != "Todos": mask &= df["NOMEFANTASIA"] == selected_fornecedor
        return df[mask], auto_refresh


def render_kpis(df, df_filtered):
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    total_qty = df_filtered["QUANTIDADE"].sum()
    total_valor = df_filtered["VALOR"].sum()
    total_itens = len(df_filtered)
    n_fornecedores = df_filtered["NOMEFANTASIA"].nunique()
    n_descricoes = df_filtered["DESCRIÇÃO"].nunique()
    valor_medio = total_valor / total_itens if total_itens > 0 else 0
    with col1:
        st.markdown(f'<div class="kpi-card"><h2>{format_qty(total_qty)}</h2><p>Quantidade Total</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="kpi-card kpi-card-info"><h2>{format_brl(total_valor)}</h2><p>Valor Total</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="kpi-card kpi-card-warn"><h2>{total_itens:,}</h2><p>Total de Registros</p></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="kpi-card"><h2>{n_descricoes:,}</h2><p>Itens Únicos</p></div>', unsafe_allow_html=True)
    with col5:
        st.markdown(f'<div class="kpi-card kpi-card-info"><h2>{n_fornecedores}</h2><p>Fornecedores</p></div>', unsafe_allow_html=True)
    with col6:
        st.markdown(f'<div class="kpi-card kpi-card-danger"><h2>{format_brl(valor_medio)}</h2><p>Valor Médio / Registro</p></div>', unsafe_allow_html=True)

def render_row1(df_filtered):
    col1, col2, col3 = st.columns([4, 3, 3])
    with col1:
        st.markdown('<div class="section-header"><i class="fa-solid fa-box"></i> TIPO DE MATERIAL</div>', unsafe_allow_html=True)
        tipo_data = df_filtered.groupby("TIPO").agg(
            QTD=("QUANTIDADE", "sum"), VALOR=("VALOR", "sum"), REGISTROS=("TIPO", "count")
        ).sort_values("QTD", ascending=True).tail(12)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=tipo_data.index, x=tipo_data["QTD"], orientation="h",
            marker_color=COLORS["primary"],
            text=tipo_data["QTD"].apply(lambda x: format_qty(x)),
            textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Quantidade: %{x:,.0f}<extra></extra>"
        ))
        fig.update_layout(height=400, margin=dict(l=10, r=80, t=10, b=10),
            xaxis_title="Quantidade", plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)", font=dict(size=11),
            xaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", automargin=True))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="section-header"><i class="fa-solid fa-tag"></i> CLASSIFICAÇÃO</div>', unsafe_allow_html=True)
        classif_data = df_filtered.groupby("CLASSIFICAÇÃO DO MATERIAL")["QUANTIDADE"].sum().sort_values(ascending=False)
        fig = go.Figure(data=[go.Pie(
            labels=classif_data.index, values=classif_data.values, hole=0.5,
            marker_colors=COLORS["chart_palette"][:len(classif_data)],
            textinfo="label+value", textfont_size=11,
            hovertemplate="<b>%{label}</b><br>Quantidade: %{value:,.0f}<br>Percentual: %{percent}<extra></extra>"
        )])
        fig.update_layout(height=400, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with col3:
        st.markdown('<div class="section-header"><i class="fa-solid fa-list"></i> DISCIPLINA</div>', unsafe_allow_html=True)
        disc_data = df_filtered.groupby("DISCIPLINA")["QUANTIDADE"].sum().sort_values(ascending=True)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=disc_data.index, x=disc_data.values, orientation="h",
            marker_color=COLORS["secondary"], text=disc_data.values,
            texttemplate="%{text:,.0f}", textposition="outside",
            cliponaxis=False,
        ))
        fig.update_layout(height=400, margin=dict(l=10, r=80, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=11), xaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", automargin=True))
        st.plotly_chart(fig, use_container_width=True)


def render_row2(df_filtered):
    col1, col2 = st.columns([6, 4])
    with col1:
        st.markdown('<div class="section-header"><i class="fa-solid fa-chart-line"></i> EVOLUÇÃO MENSAL</div>', unsafe_allow_html=True)
        monthly = df_filtered.groupby("MES").agg(
            QTD=("QUANTIDADE", "sum"), VALOR=("VALOR", "sum"), REGISTROS=("QUANTIDADE", "count")
        ).sort_index()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            x=monthly.index.astype(str), y=monthly["QTD"], name="Quantidade",
            marker_color=COLORS["primary"], opacity=0.8,
            hovertemplate="<b>%{x}</b><br>Quantidade: %{y:,.0f}<extra></extra>"
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=monthly.index.astype(str), y=monthly["VALOR"], name="Valor (R$)",
            line=dict(color=COLORS["warning"], width=3), mode="lines+markers",
            hovertemplate="<b>%{x}</b><br>Valor: R$ %{y:,.2f}<extra></extra>"
        ), secondary_y=True)
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            font=dict(size=11), xaxis=dict(showgrid=False), barmode="group")
        fig.update_yaxes(title_text="Quantidade", secondary_y=False, showgrid=True, gridcolor="rgba(128,128,128,0.2)")
        fig.update_yaxes(title_text="Valor (R$)", secondary_y=True, showgrid=False)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="section-header"><i class="fa-solid fa-building"></i> BASE / AGRUPAMENTO</div>', unsafe_allow_html=True)
        base_data = df_filtered.groupby("BASE").agg(
            QTD=("QUANTIDADE", "sum"), VALOR=("VALOR", "sum"), REGISTROS=("BASE", "count")
        ).sort_values("QTD", ascending=False)
        fig = go.Figure(data=[go.Bar(
            x=base_data.index, y=base_data["QTD"],
            marker_color=[COLORS["primary"], COLORS["secondary"]][:len(base_data)],
            text=base_data["QTD"].apply(lambda x: format_qty(x)), textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>Quantidade: %{y:,.0f}<extra></extra>"
        )])
        fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=80),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=11), yaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", automargin=True),
            xaxis=dict(tickangle=-15, automargin=True))
        st.plotly_chart(fig, use_container_width=True)

def render_row3(df_filtered):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="section-header"><i class="fa-solid fa-industry"></i> TOP 10 FORNECEDORES (VALOR)</div>', unsafe_allow_html=True)
        top_forn = df_filtered.groupby("NOMEFANTASIA")["VALOR"].sum().sort_values(ascending=False).head(10)
        labels = [n[:40] + "..." if len(n) > 40 else n for n in top_forn.index]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=labels[::-1], x=top_forn.values[::-1], orientation="h",
            marker_color=COLORS["chart_palette"][:10][::-1],
            text=[format_brl(v) for v in top_forn.values[::-1]],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Valor: R$ %{x:,.2f}<extra></extra>"
        ))
        fig.update_layout(height=400, margin=dict(l=10, r=100, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=10), xaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)", automargin=True))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="section-header"><i class="fa-solid fa-sitemap"></i> TREEMAP — TIPO × CLASSIFICAÇÃO</div>', unsafe_allow_html=True)
        tree_data = df_filtered.groupby(["TIPO", "CLASSIFICAÇÃO DO MATERIAL"])["VALOR"].sum().reset_index()
        tree_data = tree_data[tree_data["VALOR"] > 0]
        if not tree_data.empty:
            fig = px.treemap(tree_data, path=["TIPO", "CLASSIFICAÇÃO DO MATERIAL"],
                values="VALOR", color="VALOR", color_continuous_scale="Greens")
            fig.update_layout(height=400, margin=dict(l=5, r=5, t=5, b=5),
                font=dict(size=11), coloraxis_showscale=False)
            fig.update_traces(hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados para exibir o Treemap.")
    with col3:
        st.markdown('<div class="section-header"><i class="fa-solid fa-table-cells"></i> HEATMAP — DISCIPLINA × MÊS</div>', unsafe_allow_html=True)
        heat_data = df_filtered.pivot_table(index="DISCIPLINA", columns="MES",
            values="QUANTIDADE", aggfunc="sum", fill_value=0)
        if not heat_data.empty:
            heat_data = heat_data[sorted(heat_data.columns)]
            fig = go.Figure(data=go.Heatmap(
                z=heat_data.values, x=[str(c) for c in heat_data.columns],
                y=heat_data.index, colorscale="Greens",
                hovertemplate="Disciplina: %{y}<br>Mês: %{x}<br>Quantidade: %{z:,.0f}<extra></extra>"
            ))
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=10, b=10),
                font=dict(size=10), xaxis=dict(tickangle=-45))
            st.plotly_chart(fig, use_container_width=True)


def render_row4(df_filtered):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-header"><i class="fa-solid fa-chart-bar"></i> ANÁLISE DE PARETO — TOP 20 ITENS (VALOR)</div>', unsafe_allow_html=True)
        pareto = df_filtered.groupby("DESCRIÇÃO")["VALOR"].sum().sort_values(ascending=False).head(20)
        cumulative = pareto.cumsum() / pareto.sum() * 100
        labels = [d[:50] + "..." if len(d) > 50 else d for d in pareto.index]
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            x=labels, y=pareto.values, name="Valor",
            marker_color=COLORS["primary"],
            hovertemplate="<b>%{x}</b><br>Valor: R$ %{y:,.2f}<extra></extra>"
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=labels, y=cumulative.values, name="% Acumulado",
            line=dict(color=COLORS["danger"], width=2), mode="lines+markers",
            hovertemplate="%{y:.1f}%<extra></extra>"
        ), secondary_y=True)
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=120),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            font=dict(size=10), xaxis=dict(tickangle=-45, showgrid=False))
        fig.update_yaxes(title_text="Valor (R$)", secondary_y=False, showgrid=True, gridcolor="rgba(128,128,128,0.2)")
        fig.update_yaxes(title_text="% Acumulado", secondary_y=True, range=[0, 105], showgrid=False)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="section-header"><i class="fa-solid fa-chart-area"></i> TENDÊNCIA DE GASTOS POR CLASSIFICAÇÃO</div>', unsafe_allow_html=True)
        trend = df_filtered.pivot_table(index="MES", columns="CLASSIFICAÇÃO DO MATERIAL",
            values="VALOR", aggfunc="sum", fill_value=0).sort_index()
        fig = go.Figure()
        for i, col in enumerate(trend.columns):
            fig.add_trace(go.Scatter(
                x=trend.index.astype(str), y=trend[col], name=col,
                mode="lines+markers",
                line=dict(width=2, color=COLORS["chart_palette"][i % len(COLORS["chart_palette"])]),
                fill="tonexty" if i > 0 else "tozeroy",
                hovertemplate=f"<b>{col}</b><br>Mês: %{{x}}<br>Valor: R$ %{{y:,.2f}}<extra></extra>"
            ))
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, font=dict(size=10)),
            font=dict(size=11), xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)"))
        st.plotly_chart(fig, use_container_width=True)

def render_row5(df_filtered):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="section-header"><i class="fa-solid fa-bullseye"></i> CONCENTRAÇÃO DE FORNECEDORES</div>', unsafe_allow_html=True)
        forn_valor = df_filtered.groupby("NOMEFANTASIA")["VALOR"].sum().sort_values(ascending=False)
        total = forn_valor.sum()
        if total > 0:
            top5_pct = forn_valor.head(5).sum() / total * 100
            top10_pct = forn_valor.head(10).sum() / total * 100
            others_pct = 100 - top10_pct
            fig = go.Figure(data=[go.Indicator(
                mode="gauge+number+delta", value=top5_pct,
                title={"text": "Top 5 Fornecedores (% do Valor Total)"},
                gauge={"axis": {"range": [0, 100]}, "bar": {"color": COLORS["primary"]},
                    "steps": [
                        {"range": [0, 50], "color": "rgba(76,175,80,0.1)"},
                        {"range": [50, 75], "color": "rgba(255,152,0,0.1)"},
                        {"range": [75, 100], "color": "rgba(244,67,54,0.1)"}],
                    "threshold": {"line": {"color": COLORS["danger"], "width": 4},
                        "thickness": 0.75, "value": 80}},
                number={"suffix": "%", "font": {"size": 40}}
            )])
            fig.update_layout(height=300, margin=dict(l=30, r=30, t=80, b=10), font=dict(size=12))
            st.plotly_chart(fig, use_container_width=True)
            st.markdown(f"""
            | Métrica | Valor |
            |---------|-------|
            | Top 5 fornecedores | **{top5_pct:.1f}%** do valor total |
            | Top 10 fornecedores | **{top10_pct:.1f}%** do valor total |
            | Demais fornecedores | **{others_pct:.1f}%** do valor total |
            | Total de fornecedores | **{len(forn_valor)}** |
            """)
    with col2:
        st.markdown('<div class="section-header"><i class="fa-solid fa-chart-column"></i> DISTRIBUIÇÃO DE VALORES POR PEDIDO</div>', unsafe_allow_html=True)
        valor_data = df_filtered[df_filtered["VALOR"] > 0]["VALOR"]
        if not valor_data.empty:
            fig = go.Figure()
            fig.add_trace(go.Histogram(
                x=valor_data, nbinsx=50, marker_color=COLORS["primary"], opacity=0.7,
                hovertemplate="Faixa: R$ %{x:,.0f}<br>Frequência: %{y}<extra></extra>"
            ))
            fig.add_vline(x=valor_data.median(), line_dash="dash", line_color=COLORS["warning"],
                         annotation_text=f"Mediana: {format_brl(valor_data.median())}")
            fig.add_vline(x=valor_data.mean(), line_dash="dot", line_color=COLORS["danger"],
                         annotation_text=f"Média: {format_brl(valor_data.mean())}")
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=10, b=10),
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                xaxis_title="Valor (R$)", yaxis_title="Frequência", font=dict(size=11),
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)"))
            st.plotly_chart(fig, use_container_width=True)
    with col3:
        st.markdown('<div class="section-header"><i class="fa-solid fa-triangle-exclamation"></i> TOP ITENS COM MAIOR VARIAÇÃO DE PREÇO</div>', unsafe_allow_html=True)
        multi_purchase = df_filtered[df_filtered["VALOR_UNITARIO"] > 0].groupby("DESCRIÇÃO").agg(
            COMPRAS=("VALOR_UNITARIO", "count"), PRECO_MIN=("VALOR_UNITARIO", "min"),
            PRECO_MAX=("VALOR_UNITARIO", "max"), PRECO_MEDIO=("VALOR_UNITARIO", "mean"),
            PRECO_STD=("VALOR_UNITARIO", "std")
        )
        multi_purchase = multi_purchase[multi_purchase["COMPRAS"] >= 3].copy()
        if not multi_purchase.empty:
            multi_purchase["VARIACAO_%"] = ((multi_purchase["PRECO_MAX"] - multi_purchase["PRECO_MIN"]) / multi_purchase["PRECO_MIN"] * 100)
            multi_purchase = multi_purchase.sort_values("VARIACAO_%", ascending=False).head(10)
            display_df = multi_purchase[["COMPRAS", "PRECO_MIN", "PRECO_MAX", "VARIACAO_%"]].copy()
            display_df.index = [d[:45] + "..." if len(d) > 45 else d for d in display_df.index]
            display_df.columns = ["Compras", "Preço Mín", "Preço Máx", "Variação %"]
            display_df["Preço Mín"] = display_df["Preço Mín"].apply(lambda x: f"R$ {x:,.2f}")
            display_df["Preço Máx"] = display_df["Preço Máx"].apply(lambda x: f"R$ {x:,.2f}")
            display_df["Variação %"] = display_df["Variação %"].apply(lambda x: f"{x:,.1f}%")
            st.dataframe(display_df, use_container_width=True, height=400)
        else:
            st.info("Dados insuficientes para análise de variação de preço.")


def render_detail_table(df_filtered):
    st.markdown('<div class="section-header"><i class="fa-solid fa-table"></i> DETALHAMENTO DE MATERIAIS</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([4, 2, 2])
    with col1:
        search = st.text_input("Buscar descrição", placeholder="Digite para filtrar...")
    with col2:
        sort_by = st.selectbox("Ordenar por", ["DATAEMISSAO", "VALOR", "QUANTIDADE", "DESCRIÇÃO"])
    with col3:
        sort_order = st.selectbox("Ordem", ["Decrescente", "Crescente"])
    display_df = df_filtered.copy()
    if search:
        display_df = display_df[display_df["DESCRIÇÃO"].str.contains(search, case=False, na=False)]
    ascending = sort_order == "Crescente"
    display_df = display_df.sort_values(sort_by, ascending=ascending)
    cols_display = ["DESCRIÇÃO", "TIPO", "DISCIPLINA", "CLASSIFICAÇÃO DO MATERIAL",
                    "QUANTIDADE", "UNIDADE", "VALOR", "NOMEFANTASIA", "DATAEMISSAO", "BASE"]
    display_df = display_df[cols_display].copy()
    display_df["VALOR"] = display_df["VALOR"].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "")
    display_df["DATAEMISSAO"] = display_df["DATAEMISSAO"].dt.strftime("%d/%m/%Y")
    display_df["QUANTIDADE"] = display_df["QUANTIDADE"].apply(lambda x: f"{x:,.0f}")
    st.dataframe(display_df, use_container_width=True, height=500,
        column_config={
            "DESCRIÇÃO": st.column_config.TextColumn("Descrição", width="large"),
            "TIPO": st.column_config.TextColumn("Tipo"),
            "DISCIPLINA": st.column_config.TextColumn("Disciplina"),
            "CLASSIFICAÇÃO DO MATERIAL": st.column_config.TextColumn("Classificação"),
            "QUANTIDADE": st.column_config.TextColumn("Qtd"),
            "UNIDADE": st.column_config.TextColumn("Un"),
            "VALOR": st.column_config.TextColumn("Valor"),
            "NOMEFANTASIA": st.column_config.TextColumn("Fornecedor", width="medium"),
            "DATAEMISSAO": st.column_config.TextColumn("Data Emissão"),
            "BASE": st.column_config.TextColumn("Base"),
        })
    st.markdown(f"**{len(display_df):,} registros exibidos**")

def render_statistics(df_filtered):
    st.markdown('<div class="section-header"><i class="fa-solid fa-square-poll-vertical"></i> ANÁLISE ESTATÍSTICA</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**Estatísticas de Valor (R$)**")
        valor_stats = df_filtered[df_filtered["VALOR"] > 0]["VALOR"].describe()
        stats_df = pd.DataFrame({
            "Métrica": ["Contagem", "Média", "Desvio Padrão", "Mínimo", "25%", "Mediana", "75%", "Máximo"],
            "Valor": [
                f"{valor_stats['count']:,.0f}", f"R$ {valor_stats['mean']:,.2f}",
                f"R$ {valor_stats['std']:,.2f}", f"R$ {valor_stats['min']:,.2f}",
                f"R$ {valor_stats['25%']:,.2f}", f"R$ {valor_stats['50%']:,.2f}",
                f"R$ {valor_stats['75%']:,.2f}", f"R$ {valor_stats['max']:,.2f}"]
        })
        st.dataframe(stats_df, use_container_width=True, hide_index=True)
    with col2:
        st.markdown("**Estatísticas de Quantidade**")
        qty_stats = df_filtered[df_filtered["QUANTIDADE"] > 0]["QUANTIDADE"].describe()
        stats_df2 = pd.DataFrame({
            "Métrica": ["Contagem", "Média", "Desvio Padrão", "Mínimo", "25%", "Mediana", "75%", "Máximo"],
            "Valor": [
                f"{qty_stats['count']:,.0f}", f"{qty_stats['mean']:,.1f}",
                f"{qty_stats['std']:,.1f}", f"{qty_stats['min']:,.0f}",
                f"{qty_stats['25%']:,.0f}", f"{qty_stats['50%']:,.0f}",
                f"{qty_stats['75%']:,.0f}", f"{qty_stats['max']:,.0f}"]
        })
        st.dataframe(stats_df2, use_container_width=True, hide_index=True)
    with col3:
        st.markdown("**Resumo por Classificação**")
        class_summary = df_filtered.groupby("CLASSIFICAÇÃO DO MATERIAL").agg(
            Registros=("CLASSIFICAÇÃO DO MATERIAL", "count"),
            Valor_Total=("VALOR", "sum"), Qtd_Total=("QUANTIDADE", "sum")
        ).sort_values("Valor_Total", ascending=False)
        class_summary["Valor_Total"] = class_summary["Valor_Total"].apply(lambda x: format_brl(x))
        class_summary["Qtd_Total"] = class_summary["Qtd_Total"].apply(lambda x: format_qty(x))
        class_summary.columns = ["Registros", "Valor Total", "Qtd Total"]
        st.dataframe(class_summary, use_container_width=True)


def main():
    logo_b64 = get_logo_base64()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:60px; margin-bottom:8px;" /><br/>' if logo_b64 else ""
    st.markdown(f"""
    <div class="main-header">
        {logo_html}
        <h1>UTILIZAÇÃO DE MATERIAIS | CONTRATO 4300682358 | RJ</h1>
        <p>Normatel Engenharia — Dashboard Analítico de Materiais e Equipamentos</p>
    </div>
    """, unsafe_allow_html=True)
    _default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "741 E 743_MATERIAIS - ABRIL.xlsx")
    FILE_PATH = os.environ.get("EXCEL_FILE", _default)
    if not os.path.exists(FILE_PATH):
        st.error(f"Arquivo não encontrado: **{FILE_PATH}**")
        st.info("""
        **Como configurar:**
        1. Coloque o arquivo Excel na mesma pasta do `app.py`
        2. Ou defina a variável de ambiente: `EXCEL_FILE=caminho/do/arquivo.xlsx`
        3. Ou altere a variável `FILE_PATH` no código
        """)
        uploaded = st.file_uploader("Fazer upload do arquivo Excel", type=["xlsx", "xls"])
        if uploaded:
            FILE_PATH = "/tmp/uploaded_excel.xlsx"
            with open(FILE_PATH, "wb") as f:
                f.write(uploaded.getbuffer())
        else:
            return
    try:
        df = load_data(FILE_PATH)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return
    df_filtered, auto_refresh = render_sidebar(df)
    if df_filtered.empty:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
        return
    render_kpis(df, df_filtered)
    render_row1(df_filtered)
    render_row2(df_filtered)
    render_row3(df_filtered)
    render_row4(df_filtered)
    render_row5(df_filtered)
    render_detail_table(df_filtered)
    render_statistics(df_filtered)
    st.markdown("---")
    st.markdown(
        f"<div style='text-align:center; color:#888; font-size:0.8rem;'>"
        f"Dashboard Analítico | Normatel Engenharia | Gerente: JOSÉ DANIEL | "
        f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>",
        unsafe_allow_html=True)
    if auto_refresh:
        if "last_refresh" not in st.session_state:
            st.session_state.last_refresh = time.time()
        elapsed = time.time() - st.session_state.last_refresh
        if elapsed >= 30:
            st.session_state.last_refresh = time.time()
            st.cache_data.clear()
            st.rerun()


if __name__ == "__main__":
    main()
