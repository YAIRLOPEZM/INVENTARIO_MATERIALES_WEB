import streamlit as st
import pandas as pd
import plotly.express as px
import psycopg2
from zoneinfo import ZoneInfo

st.set_page_config(
    page_title="Inventario de Materiales",
    page_icon="📦",
    layout="wide",
)
# ==========================================
# LOGIN
# ==========================================
import streamlit_authenticator as stauth

credentials = {
    "usernames": {
        user: dict(data)
        for user, data in st.secrets["credentials"]["usernames"].items()
    }
}

authenticator = stauth.Authenticate(
    credentials,
    st.secrets["cookie"]["name"],
    st.secrets["cookie"]["key"],
    st.secrets["cookie"]["expiry_days"],
    auto_hash=False,
)

try:
    authenticator.login()
except Exception as e:
    st.error(e)

if st.session_state.get("authentication_status") is False:
    st.error("Usuario o contraseña incorrectos")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("Por favor ingresa tu usuario y contraseña")
    st.stop()

authenticator.logout("Cerrar sesión", "sidebar")
st.sidebar.write(f'Bienvenido, {st.session_state.get("name")}')


@st.cache_data(ttl=300)
def cargar_inventario():
    database_url = st.secrets["DATABASE_URL"]
    conexion = psycopg2.connect(database_url, sslmode="require")
    df = pd.read_sql("SELECT * FROM inventario", conexion)
    conexion.close()
    return df


@st.cache_data(ttl=300)
def cargar_movimientos():
    database_url = st.secrets["DATABASE_URL"]
    conexion = psycopg2.connect(database_url, sslmode="require")
    df = pd.read_sql(
        "SELECT * FROM inv_movimiento",
        conexion,
        parse_dates=["fecha_movimiento"],
    )
    conexion.close()
    return df


@st.cache_data(ttl=300)
def obtener_ultima_actualizacion():
    database_url = st.secrets["DATABASE_URL"]
    conexion = psycopg2.connect(database_url, sslmode="require")
    cur = conexion.cursor()
    cur.execute(
        "SELECT ultima_actualizacion FROM metadata_actualizacion "
        "ORDER BY id DESC LIMIT 1;"
    )
    resultado = cur.fetchone()
    conexion.close()
    return resultado[0] if resultado else None


df_inv = cargar_inventario()
df_mov = cargar_movimientos()
ultima_actualizacion = obtener_ultima_actualizacion()

# ==========================================
# FILTROS
# ==========================================

st.sidebar.title("📦 Filtros")

bodegas = sorted(df_inv["bodega_nombre"].dropna().unique())
bodega_sel = st.sidebar.multiselect("Bodega", bodegas)

alertas = sorted(df_inv["alerta_stock"].dropna().unique())
alerta_sel = st.sidebar.multiselect("Alerta", alertas)

item_busqueda = st.sidebar.text_input("Buscar ítem (código o descripción)")

df_filtrado = df_inv.copy()
if bodega_sel:
    df_filtrado = df_filtrado[df_filtrado["bodega_nombre"].isin(bodega_sel)]
if alerta_sel:
    df_filtrado = df_filtrado[df_filtrado["alerta_stock"].isin(alerta_sel)]
if item_busqueda:
    mask = (
        df_filtrado["item_codigo"].astype(str).str.contains(item_busqueda, case=False, na=False)
        | df_filtrado["item_descripcion"].str.contains(item_busqueda, case=False, na=False)
    )
    df_filtrado = df_filtrado[mask]

# ==========================================
# ENCABEZADO Y KPIs
# ==========================================

st.title("Inventario de Materiales")
st.caption("Stock, consumo promedio, alertas de reposición y valorización por bodega.")

if ultima_actualizacion:
    # Neon guarda la hora en UTC; la convertimos a hora de Colombia
    # solo para mostrarla (el dato guardado sigue siendo UTC).
    hora_colombia = ultima_actualizacion.replace(
        tzinfo=ZoneInfo("UTC")
    ).astimezone(ZoneInfo("America/Bogota"))

    st.info(
        f"🕒 Última actualización de datos: "
        f"{hora_colombia.strftime('%d/%m/%Y %I:%M %p')}"
    )
else:
    st.warning("🕒 Aún no hay registro de la última actualización de datos.")

col1, col2, col3, col4, col5 = st.columns(5)

def formato_compacto(valor):
    """Abrevia números grandes: 2,915,123 -> 2.92M, 850,000 -> 850K."""
    if valor >= 1_000_000:
        return f"${valor / 1_000_000:,.2f}M"
    if valor >= 1_000:
        return f"${valor / 1_000:,.0f}K"
    return f"${valor:,.0f}"


col1.metric("Stock total (unidades)", f"{df_filtrado['existencias'].sum():,.0f}")
col2.metric(
    "Valor total inventario",
    formato_compacto(df_filtrado['valor_total'].sum()),
    help=f"Valor exacto: ${df_filtrado['valor_total'].sum():,.0f}",
)
col3.metric(
    "Ítems críticos",
    f"{(df_filtrado['alerta_stock'] == '🔴 CRÍTICO').sum():,}",
)
col4.metric(
    "En reposición",
    f"{(df_filtrado['alerta_stock'] == '🟡 REPOSICIÓN').sum():,}",
)
col5.metric(
    "Sin rotación",
    f"{(df_filtrado['alerta_stock'] == '⚪ SIN ROTACIÓN').sum():,}",
)

st.divider()

# ==========================================
# GRÁFICOS
# ==========================================

col_izq, col_der = st.columns([1, 1.3])

with col_izq:
    st.subheader("Distribución de alertas")
    conteo_alerta = df_filtrado["alerta_stock"].value_counts().reset_index()
    conteo_alerta.columns = ["Alerta", "Cantidad"]
    fig = px.pie(
        conteo_alerta, names="Alerta", values="Cantidad", hole=0.5,
        color="Alerta",
        color_discrete_map={
            "🔴 CRÍTICO": "#e63946",
            "🟡 REPOSICIÓN": "#f4a300",
            "🟢 OK": "#1a7a3a",
            "⚪ SIN ROTACIÓN": "#9ca3af",
            "🟣 ITEM SIN REGISTRO Y SIN MOVIMIENTO": "#8b5cf6",
        },
    )
    fig.update_traces(textinfo="percent+value")
    st.plotly_chart(fig, use_container_width=True)

with col_der:
    st.subheader("Top 15 ítems por valor en inventario")
    top_valor = df_filtrado.nlargest(15, "valor_total")[["item_descripcion", "valor_total"]]
    fig2 = px.bar(top_valor, x="valor_total", y="item_descripcion", orientation="h")
    fig2.update_layout(yaxis={"categoryorder": "total ascending"}, yaxis_title="", xaxis_title="Valor total")
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

# ==========================================
# MEJORA: días estimados hasta agotarse
# ==========================================

st.subheader("⏳ Estimado de días hasta agotar stock")
st.caption("Basado en el consumo promedio mensual. Solo ítems con consumo activo.")

con_consumo = df_filtrado[df_filtrado["consumo_promedio_mensual"] > 0].copy()
con_consumo["dias_restantes"] = (
    con_consumo["existencias"] / (con_consumo["consumo_promedio_mensual"] / 30)
).round(0)

tabla_dias = con_consumo.nsmallest(20, "dias_restantes")[
    ["item_codigo", "item_descripcion", "bodega_nombre", "existencias",
     "consumo_promedio_mensual", "dias_restantes", "alerta_stock"]
]

st.dataframe(
    tabla_dias,
    use_container_width=True,
    hide_index=True,
    column_config={
        "item_codigo": "Código",
        "item_descripcion": "Descripción",
        "bodega_nombre": "Bodega",
        "existencias": "Stock actual",
        "consumo_promedio_mensual": st.column_config.NumberColumn("Consumo/mes", format="%.1f"),
        "dias_restantes": st.column_config.NumberColumn("Días restantes", format="%.0f"),
        "alerta_stock": "Alerta",
    },
)

st.divider()

# ==========================================
# TABLA COMPLETA (equivalente a la tabla del dashboard viejo)
# ==========================================

st.subheader("📋 Detalle completo de inventario")

st.dataframe(
    df_filtrado[[
        "bodega_nombre", "item_codigo", "item_descripcion", "existencias",
        "consumo_promedio_mensual", "stock_minimos", "punto_reposicion",
        "stock_maximo", "cantidad_a_pedir", "valor_unidad", "valor_total",
        "alerta_stock",
    ]],
    use_container_width=True,
    hide_index=True,
    column_config={
        "bodega_nombre": "Bodega",
        "item_codigo": "Código",
        "item_descripcion": "Descripción",
        "existencias": "Stock actual",
        "consumo_promedio_mensual": st.column_config.NumberColumn("Consumo/mes", format="%.1f"),
        "stock_minimos": st.column_config.NumberColumn("Stock mínimo", format="%.1f"),
        "punto_reposicion": st.column_config.NumberColumn("Pto. reposición", format="%.1f"),
        "stock_maximo": st.column_config.NumberColumn("Stock máximo", format="%.1f"),
        "cantidad_a_pedir": st.column_config.NumberColumn("Cant. a pedir", format="%.1f"),
        "valor_unidad": st.column_config.NumberColumn("Valor unidad", format="$%.0f"),
        "valor_total": st.column_config.NumberColumn("Valor total", format="$%.0f"),
        "alerta_stock": "Alerta",
    },
)

st.caption(f"{len(df_filtrado):,} ítems mostrados de {len(df_inv):,} totales.")

# ==========================================
# HISTÓRICO DE MOVIMIENTOS (por ítem, opcional)
# ==========================================

st.divider()
st.subheader("📈 Histórico de movimientos por ítem")

item_hist = st.selectbox(
    "Elige un ítem para ver su historial",
    options=sorted(df_mov["item_descripcion"].dropna().unique()),
    index=None,
    placeholder="Busca un ítem...",
)

if item_hist:
    mov_item = df_mov[df_mov["item_descripcion"] == item_hist].sort_values("fecha_movimiento")
    fig3 = px.line(
        mov_item, x="fecha_movimiento", y="saldo_parcial",
        title=f"Saldo en el tiempo — {item_hist}",
        markers=True,
    )
    st.plotly_chart(fig3, use_container_width=True)

    st.dataframe(
        mov_item[["fecha_movimiento", "bodega_nombre", "tipo", "numero", "entradas_san", "salidas_san", "saldo_parcial"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "fecha_movimiento": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
            "bodega_nombre": "Bodega",
            "tipo": "Tipo",
            "numero": "N° documento",
            "entradas_san": "Entradas",
            "salidas_san": "Salidas",
            "saldo_parcial": "Saldo",
        },
    )

# ==========================================
# EXPORTAR
# ==========================================

@st.cache_data
def convertir_a_excel(dataframe):
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Inventario")
    return buffer.getvalue()

st.download_button(
    "📊 Exportar esta vista a Excel",
    data=convertir_a_excel(df_filtrado),
    file_name="inventario_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
