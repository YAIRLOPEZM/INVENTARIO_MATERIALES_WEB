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
def cargar_salidas():
    database_url = st.secrets["DATABASE_URL"]
    conexion = psycopg2.connect(database_url, sslmode="require")
    df = pd.read_sql(
        "SELECT * FROM inv_salida",
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
df_sal = cargar_salidas()
# Nombre de la bodega para mostrarlo en el historial
nombres_bodega = (
    df_inv.dropna(subset=["bodega_codigo"])
    .drop_duplicates("bodega_codigo")
    .set_index("bodega_codigo")["bodega_nombre"]
    .to_dict()
)
df_sal["bodega_nombre"] = df_sal["bodega_codigo"].map(nombres_bodega)
ultima_actualizacion = obtener_ultima_actualizacion()

tab_inv, tab_sal = st.tabs(["📦 Inventario", "📤 Salidas por tercero / equipo"])

with tab_inv:
    # ==========================================
    # FILTROS
    # ==========================================

    st.sidebar.title("📦 Filtros de inventario")
    st.sidebar.caption("Aplican a la pestaña Inventario.")

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
    # HISTÓRICO DE SALIDAS (por ítem)
    # ==========================================

    st.divider()
    st.subheader("📈 Histórico de salidas por ítem")
    st.caption("Fuente: consulta dinámica «Salidas no contabilizadas» de Novasoft.")

    item_hist = st.selectbox(
        "Elige un ítem para ver su historial",
        options=sorted(df_sal["item_descripcion"].dropna().unique()),
        index=None,
        placeholder="Busca un ítem...",
    )

    if item_hist:
        sal_item = df_sal[df_sal["item_descripcion"] == item_hist].sort_values("fecha_movimiento")

        por_mes = (
            sal_item.assign(mes=sal_item["fecha_movimiento"].dt.to_period("M").dt.to_timestamp())
            .groupby("mes", as_index=False)["cantidad"].sum()
        )
        fig3 = px.bar(
            por_mes, x="mes", y="cantidad",
            title=f"Unidades entregadas por mes — {item_hist}",
        )
        fig3.update_layout(xaxis_title="", yaxis_title="Unidades")
        st.plotly_chart(fig3, use_container_width=True)

        st.dataframe(
            sal_item[[
                "fecha_movimiento", "bodega_nombre", "numero", "cantidad",
                "costo_unitario", "costo_total", "tercero", "rig", "proyecto",
            ]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "fecha_movimiento": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "bodega_nombre": "Bodega",
                "numero": "N° documento",
                "cantidad": "Cantidad",
                "costo_unitario": st.column_config.NumberColumn("Costo unitario", format="$%.0f"),
                "costo_total": st.column_config.NumberColumn("Costo total", format="$%.0f"),
                "tercero": "Tercero",
                "rig": "Equipo (Rig)",
                "proyecto": "Proyecto",
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


# ==========================================
# PESTAÑA: SALIDAS POR TERCERO / EQUIPO (RIG)
# ==========================================

with tab_sal:

    st.title("Salidas por tercero y equipo (Rig)")
    st.caption(
        "Fuente: consulta dinámica «Salidas no contabilizadas» de Novasoft. "
        "Cada fila es una línea de salida con su tercero, equipo, proyecto y costo."
    )

    if df_sal.empty:
        st.warning("Todavía no hay salidas cargadas.")
        st.stop()

    d = df_sal.copy()
    d["anio"] = d["fecha_movimiento"].dt.year
    d["mes"] = d["fecha_movimiento"].dt.month
    for c in ("tercero", "rig", "proyecto", "centro_costo", "bodega_nombre"):
        d[c] = d[c].fillna("SIN DATO")

    MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

    st.subheader("🔎 Filtros")

    f1, f2, f3 = st.columns(3)
    anios = sorted(d["anio"].unique(), reverse=True)
    anio_sel = f1.multiselect("Año", anios, default=anios[:1], key="sal_anio")
    mes_sel = f2.multiselect(
        "Mes", list(MESES), format_func=lambda m: MESES[m], key="sal_mes"
    )
    bodega_sal = f3.multiselect(
        "Bodega", sorted(d["bodega_nombre"].unique()), key="sal_bodega"
    )

    f4, f5, f6 = st.columns(3)
    rig_sel = f4.multiselect(
        "Equipo (Rig)", sorted(d["rig"].unique()), key="sal_rig"
    )
    tercero_sel = f5.multiselect(
        "Tercero (quien recibe)", sorted(d["tercero"].unique()), key="sal_tercero"
    )
    proyecto_sel = f6.multiselect(
        "Proyecto", sorted(d["proyecto"].unique()), key="sal_proyecto"
    )

    busqueda_sal = st.text_input(
        "Buscar ítem o documento (código, descripción o número SAN)",
        key="sal_busqueda",
    )

    ds = d
    if anio_sel:
        ds = ds[ds["anio"].isin(anio_sel)]
    if mes_sel:
        ds = ds[ds["mes"].isin(mes_sel)]
    if bodega_sal:
        ds = ds[ds["bodega_nombre"].isin(bodega_sal)]
    if rig_sel:
        ds = ds[ds["rig"].isin(rig_sel)]
    if tercero_sel:
        ds = ds[ds["tercero"].isin(tercero_sel)]
    if proyecto_sel:
        ds = ds[ds["proyecto"].isin(proyecto_sel)]
    if busqueda_sal:
        m = (
            ds["item_codigo"].astype(str).str.contains(busqueda_sal, case=False, na=False)
            | ds["item_descripcion"].astype(str).str.contains(busqueda_sal, case=False, na=False)
            | ds["numero"].astype(str).str.contains(busqueda_sal, case=False, na=False)
        )
        ds = ds[m]

    if ds.empty:
        st.info("No hay salidas con esos filtros.")
        st.stop()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Costo total", formato_compacto(ds["costo_total"].sum()),
              help=f"Valor exacto: ${ds['costo_total'].sum():,.0f}")
    k2.metric("Unidades", f"{ds['cantidad'].sum():,.0f}")
    k3.metric("Documentos", f"{ds['numero'].nunique():,}")
    k4.metric("Ítems distintos", f"{ds['item_codigo'].nunique():,}")

    st.divider()

    g1, g2 = st.columns(2)

    with g1:
        st.subheader("Costo por mes")
        por_mes = (
            ds.assign(periodo=ds["fecha_movimiento"].dt.to_period("M").dt.to_timestamp())
            .groupby("periodo", as_index=False)["costo_total"].sum()
        )
        fig_mes = px.bar(por_mes, x="periodo", y="costo_total")
        fig_mes.update_layout(xaxis_title="", yaxis_title="Costo total")
        st.plotly_chart(fig_mes, use_container_width=True)

    with g2:
        st.subheader("Costo por equipo (Rig)")
        por_rig = (
            ds.groupby("rig", as_index=False)["costo_total"].sum()
            .sort_values("costo_total", ascending=False).head(15)
        )
        fig_rig = px.bar(por_rig, x="costo_total", y="rig", orientation="h")
        fig_rig.update_layout(yaxis={"categoryorder": "total ascending"},
                              yaxis_title="", xaxis_title="Costo total")
        st.plotly_chart(fig_rig, use_container_width=True)

    st.subheader("Costo por tercero (top 15)")
    por_ter = (
        ds.groupby("tercero", as_index=False)["costo_total"].sum()
        .sort_values("costo_total", ascending=False).head(15)
    )
    fig_ter = px.bar(por_ter, x="costo_total", y="tercero", orientation="h")
    fig_ter.update_layout(yaxis={"categoryorder": "total ascending"},
                          yaxis_title="", xaxis_title="Costo total")
    st.plotly_chart(fig_ter, use_container_width=True)

    st.divider()

    r1, r2 = st.columns(2)

    with r1:
        st.subheader("Resumen por equipo (Rig)")
        res_rig = (
            ds.groupby("rig", as_index=False)
            .agg(lineas=("item_codigo", "size"), unidades=("cantidad", "sum"),
                 costo_total=("costo_total", "sum"))
            .sort_values("costo_total", ascending=False)
        )
        st.dataframe(
            res_rig, use_container_width=True, hide_index=True,
            column_config={
                "rig": "Equipo (Rig)", "lineas": "Líneas",
                "unidades": st.column_config.NumberColumn("Unidades", format="%.0f"),
                "costo_total": st.column_config.NumberColumn("Costo total", format="$%.0f"),
            },
        )

    with r2:
        st.subheader("Resumen por tercero")
        res_ter = (
            ds.groupby("tercero", as_index=False)
            .agg(lineas=("item_codigo", "size"), unidades=("cantidad", "sum"),
                 costo_total=("costo_total", "sum"))
            .sort_values("costo_total", ascending=False)
        )
        st.dataframe(
            res_ter, use_container_width=True, hide_index=True,
            column_config={
                "tercero": "Tercero", "lineas": "Líneas",
                "unidades": st.column_config.NumberColumn("Unidades", format="%.0f"),
                "costo_total": st.column_config.NumberColumn("Costo total", format="$%.0f"),
            },
        )

    st.divider()

    st.subheader("📋 Detalle de salidas")

    detalle = ds.sort_values("fecha_movimiento", ascending=False)[[
        "fecha_movimiento", "bodega_nombre", "numero", "item_codigo",
        "item_descripcion", "cantidad", "costo_unitario", "costo_total",
        "tercero", "rig", "proyecto", "centro_costo",
    ]]

    st.dataframe(
        detalle, use_container_width=True, hide_index=True,
        column_config={
            "fecha_movimiento": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
            "bodega_nombre": "Bodega",
            "numero": "N° documento",
            "item_codigo": "Código",
            "item_descripcion": "Descripción",
            "cantidad": st.column_config.NumberColumn("Cantidad", format="%.2f"),
            "costo_unitario": st.column_config.NumberColumn("Costo unitario", format="$%.0f"),
            "costo_total": st.column_config.NumberColumn("Costo total", format="$%.0f"),
            "tercero": "Tercero",
            "rig": "Equipo (Rig)",
            "proyecto": "Proyecto",
            "centro_costo": "Centro de costo",
        },
    )

    st.caption(f"{len(detalle):,} líneas mostradas de {len(d):,} totales.")

    st.download_button(
        "📊 Exportar estas salidas a Excel",
        data=convertir_a_excel(detalle),
        file_name="salidas_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_salidas",
    )
