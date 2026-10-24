"""
Script para actualizar los datos de INVENTARIO_MATERIALES en Neon,
leyendo DIRECTAMENTE los archivos descargados de Novasoft:

  - Consulta dinámica "SALIDAS NO CONTABILIZADAS": una fila por cada
    línea de salida (fecha, documento, ítem, cantidad, costo, tercero,
    centro de costo, proyecto, equipo/Rig). Reemplaza al kárdex INV0206.
  - INV0230: Inventario valorizado por bodegas (stock actual).

Puedes pasar VARIOS archivos juntos, en cualquier orden — el script
detecta solo si cada uno es la consulta de salidas o un INV0230.

CÓMO USARLO:

   python actualizar_datos_inventario.py "ruta\\Resultado.xlsx" "ruta\\INV0230_bodega1.xlsx" ...

REQUISITO: la variable de entorno DATABASE_URL_INVENTARIO debe
apuntar a tu base de datos de Neon del proyecto INVENTARIO_MATERIALES
(nombre distinto a DATABASE_URL a propósito, para no chocar con
el de SEGUIMIENTO_RQ cuando ambos scripts corren en el mismo proceso).

============================================================
REGLAS DE NEGOCIO REPRODUCIDAS (del modelo original en Power BI):

- Consumo_Promedio_Mensual = Salidas del año actual / meses con consumo > 0
- Punto_Reposicion  = Consumo_Promedio_Mensual * 3
- Stock_Minimos     = Consumo_Promedio_Mensual * 1.5
- Stock_Maximo      = Consumo_Promedio_Mensual * 3.5
- Cantidad_a_Pedir  = Stock_Maximo - Stock_Actual (si Stock_Actual < Stock_Maximo, si no 0)

- Alertas_Stock (en este orden de prioridad):
    Consumo=0 y Stock_Actual=0        -> "🟣 ITEM SIN REGISTRO Y SIN MOVIMIENTO"
    Consumo=0 y Stock_Actual>0        -> "⚪ SIN ROTACIÓN"
    Stock_Actual <= Stock_Minimos     -> "🔴 CRÍTICO"
    Stock_Actual <= Punto_Reposicion  -> "🟡 REPOSICIÓN"
    (cualquier otro caso)              -> "🟢 OK"
============================================================
"""

import sys
import os
import re
import datetime

import pandas as pd
import psycopg2
import psycopg2.extras


DATABASE_URL_INVENTARIO = os.environ.get("DATABASE_URL_INVENTARIO", "")


BODEGAS_VALIDAS = ["02", "03", "04", "15"]

COLUMNAS_SALIDAS = {
    "Fecha": "fecha_movimiento",
    "Centro de Costo": "centro_costo",
    "ID CC": "id_cc",
    "ID Proyecto": "id_proyecto",
    "Proyecto": "proyecto",
    "Rig": "rig",
    "Tipo Documento": "tipo_documento",
    "Documento": "numero",
    "Bodega": "bodega_codigo",
    "Ítem": "item_codigo",
    "Descripción": "item_descripcion",
    "Cantidad": "cantidad",
    "Costo Unitario": "costo_unitario",
    "Costo Total": "costo_total",
    "ID Tercero": "id_tercero",
    "Nombre Tercero": "tercero",
    "Generador": "generador",
}


def _leer_tabla(ruta, **kwargs):
    """
    Lee el Excel descargado. Si Novasoft entrega un .xls que en realidad
    es una tabla HTML, lo intenta leer como HTML.
    """
    try:
        return pd.read_excel(ruta, **kwargs)
    except Exception:
        tablas = pd.read_html(ruta, header=kwargs.get("header", 0), decimal=",", thousands=".")
        return tablas[0]


def detectar_tipo_reporte(ruta):
    df = _leer_tabla(ruta, header=None, nrows=15)

    # Consulta dinámica de salidas: la primera fila trae los encabezados.
    primera = [str(v).strip() for v in df.iloc[0].tolist()]
    if "Documento" in primera and "Costo Unitario" in primera:
        return "SALIDAS"

    # Reportes clásicos de Novasoft (INV0230): fila con "Reporte:".
    mask = df.apply(
        lambda fila: fila.astype(str).str.contains("Reporte:", na=False).any(),
        axis=1,
    )
    if not mask.any():
        return None
    fila_reporte = df[mask].iloc[0].dropna().tolist()
    if len(fila_reporte) < 2:
        return None
    return str(fila_reporte[1]).strip()


def parsear_salidas(ruta):
    """
    Lee la consulta "SALIDAS NO CONTABILIZADAS" y la deja con nombres de
    columna simples. Se queda solo con las bodegas del inventario.
    """
    df = _leer_tabla(ruta, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    faltan = [c for c in COLUMNAS_SALIDAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Al archivo de salidas le faltan columnas: {faltan}")

    df = df[list(COLUMNAS_SALIDAS)].rename(columns=COLUMNAS_SALIDAS)

    for c in df.columns:
        df[c] = df[c].astype(str).str.strip().replace({"nan": None, "None": None, "": None})

    df["fecha_movimiento"] = pd.to_datetime(df["fecha_movimiento"], errors="coerce")
    for c in ("cantidad", "costo_unitario", "costo_total"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=["fecha_movimiento", "item_codigo"])

    # Bodega: "03 " -> "03" (y "3" -> "03")
    df["bodega_codigo"] = df["bodega_codigo"].fillna("").str.strip().str.zfill(2)
    df = df[df["bodega_codigo"].isin(BODEGAS_VALIDAS)].reset_index(drop=True)

    return df


def parsear_valorizado(ruta):
    df = _leer_tabla(ruta, header=None)

    bodega_codigo = None
    bodega_nombre = None

    filas_finales = []

    for _, fila in df.iterrows():
        col0 = fila[0] if len(fila) > 0 else None

        if isinstance(col0, str) and col0.strip().startswith("BODEGA:"):
            m = re.match(r"BODEGA:\s*(\S+)\s+(.*)", col0.strip())
            if m:
                bodega_codigo, bodega_nombre = m.group(1), m.group(2).strip()
            continue

        if isinstance(col0, str) and len(fila) > 8 and pd.notna(fila[8]):
            texto = col0.strip()
            if texto.isdigit() or (len(texto) > 5 and texto[0].isdigit()):
                filas_finales.append({
                    "bodega_codigo": bodega_codigo,
                    "bodega_nombre": bodega_nombre,
                    "item_codigo": texto,
                    "item_descripcion": str(fila[1]).strip() if pd.notna(fila[1]) else None,
                    "existencias": float(fila[8]),
                    "valor_unidad": float(fila[13]) if pd.notna(fila[13]) else None,
                    "valor_total": float(fila[26]) if pd.notna(fila[26]) else None,
                })

    return pd.DataFrame(filas_finales)


def cargar_archivos(rutas):
    salidas_por_archivo = []
    valorizado_por_archivo = []

    for ruta in rutas:
        tipo = detectar_tipo_reporte(ruta)
        if tipo == "SALIDAS":
            print(f"  {ruta} -> salidas no contabilizadas (consulta dinámica)")
            salidas_por_archivo.append(parsear_salidas(ruta))
        elif tipo == "INV0230":
            print(f"  {ruta} -> valorizado (INV0230)")
            valorizado_por_archivo.append(parsear_valorizado(ruta))
        else:
            print(f"  ⚠️  {ruta}: no se reconoce el tipo de archivo (se omite)")

    df_salidas = (
        pd.concat(salidas_por_archivo, ignore_index=True)
        if salidas_por_archivo else pd.DataFrame()
    )
    df_valorizado = (
        pd.concat(valorizado_por_archivo, ignore_index=True)
        if valorizado_por_archivo else pd.DataFrame()
    )

    return df_salidas, df_valorizado


def calcular_medidas(df_salidas, df_valorizado):

    anio_actual = datetime.date.today().year

    mov_anio = df_salidas[
        pd.to_datetime(df_salidas["fecha_movimiento"]).dt.year == anio_actual
    ].copy()
    mov_anio["anio_mes"] = pd.to_datetime(mov_anio["fecha_movimiento"]).dt.strftime("%Y-%m")

    consumo = (
        mov_anio[mov_anio["cantidad"] > 0]
        .groupby(["bodega_codigo", "item_codigo"])
        .agg(
            salidas_totales=("cantidad", "sum"),
            meses_con_consumo=("anio_mes", "nunique"),
        )
        .reset_index()
    )
    consumo["consumo_promedio_mensual"] = (
        consumo["salidas_totales"] / consumo["meses_con_consumo"]
    )

    resultado = df_valorizado.merge(
        consumo[["bodega_codigo", "item_codigo", "consumo_promedio_mensual"]],
        on=["bodega_codigo", "item_codigo"],
        how="left",
    )
    resultado["consumo_promedio_mensual"] = resultado["consumo_promedio_mensual"].fillna(0)

    resultado["punto_reposicion"] = resultado["consumo_promedio_mensual"] * 3
    resultado["stock_minimos"] = resultado["consumo_promedio_mensual"] * 1.5
    resultado["stock_maximo"] = resultado["consumo_promedio_mensual"] * 3.5

    resultado["cantidad_a_pedir"] = (
        resultado["stock_maximo"] - resultado["existencias"]
    ).clip(lower=0)
    resultado.loc[
        resultado["existencias"] >= resultado["stock_maximo"], "cantidad_a_pedir"
    ] = 0

    def calcular_alerta(fila):
        consumo_val = fila["consumo_promedio_mensual"]
        stock_val = fila["existencias"]
        if consumo_val == 0 and stock_val == 0:
            return "🟣 ITEM SIN REGISTRO Y SIN MOVIMIENTO"
        if consumo_val == 0 and stock_val > 0:
            return "⚪ SIN ROTACIÓN"
        if stock_val <= fila["stock_minimos"]:
            return "🔴 CRÍTICO"
        if stock_val <= fila["punto_reposicion"]:
            return "🟡 REPOSICIÓN"
        return "🟢 OK"

    resultado["alerta_stock"] = resultado.apply(calcular_alerta, axis=1)

    return resultado


def subir_a_neon(df_salidas, df_resultado):

    if not DATABASE_URL_INVENTARIO:
        print("ERROR: falta configurar la variable de entorno DATABASE_URL_INVENTARIO.")
        sys.exit(1)

    conexion = psycopg2.connect(DATABASE_URL_INVENTARIO, sslmode="require")
    cursor = conexion.cursor()

    # --- INV_SALIDA: una fila por cada línea de salida (consulta dinámica).
    #     La tabla vieja inv_movimiento (kárdex) ya no se actualiza.
    cursor.execute("DROP TABLE IF EXISTS inv_salida")
    cursor.execute("""
        CREATE TABLE inv_salida (
            fecha_movimiento DATE,
            centro_costo TEXT,
            id_cc TEXT,
            id_proyecto TEXT,
            proyecto TEXT,
            rig TEXT,
            tipo_documento TEXT,
            numero TEXT,
            bodega_codigo TEXT,
            item_codigo TEXT,
            item_descripcion TEXT,
            cantidad NUMERIC,
            costo_unitario NUMERIC,
            costo_total NUMERIC,
            id_tercero TEXT,
            tercero TEXT,
            generador TEXT
        )
    """)
    columnas = list(df_salidas.columns)
    columnas_sql = ", ".join(columnas)
    filas = [
        tuple(None if pd.isna(v) else v for v in fila)
        for fila in df_salidas.itertuples(index=False, name=None)
    ]
    if filas:
        psycopg2.extras.execute_values(
            cursor,
            f"INSERT INTO inv_salida ({columnas_sql}) VALUES %s",
            filas,
            page_size=1000,
        )

    cursor.execute("DROP TABLE IF EXISTS inventario")
    cursor.execute("""
        CREATE TABLE inventario (
            bodega_codigo TEXT,
            bodega_nombre TEXT,
            item_codigo TEXT,
            item_descripcion TEXT,
            existencias NUMERIC,
            valor_unidad NUMERIC,
            valor_total NUMERIC,
            consumo_promedio_mensual NUMERIC,
            punto_reposicion NUMERIC,
            stock_minimos NUMERIC,
            stock_maximo NUMERIC,
            cantidad_a_pedir NUMERIC,
            alerta_stock TEXT
        )
    """)
    columnas2 = list(df_resultado.columns)
    columnas2_sql = ", ".join(columnas2)
    filas2 = [
        tuple(None if pd.isna(v) else v for v in fila)
        for fila in df_resultado.itertuples(index=False, name=None)
    ]
    if filas2:
        psycopg2.extras.execute_values(
            cursor,
            f"INSERT INTO inventario ({columnas2_sql}) VALUES %s",
            filas2,
            page_size=1000,
        )

    # --- Registrar la hora de esta actualización, para que el dashboard
    #     la muestre.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metadata_actualizacion (
            id SERIAL PRIMARY KEY,
            ultima_actualizacion TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)
    cursor.execute(
        "INSERT INTO metadata_actualizacion (ultima_actualizacion) VALUES (NOW());"
    )

    conexion.commit()
    cursor.close()
    conexion.close()

    print(f"Listo: {len(df_salidas)} salidas y {len(df_resultado)} ítems de inventario subidos a Neon.")
    print("Listo: hora de actualización registrada para el dashboard.")


if __name__ == "__main__":

    if len(sys.argv) < 2:
        print('Uso: python actualizar_datos_inventario.py "ruta\\Resultado.xlsx" "ruta\\INV0230_bodegaX.xlsx" ...')
        sys.exit(1)

    rutas = sys.argv[1:]

    print("Leyendo y detectando archivos...")
    df_salidas, df_valorizado = cargar_archivos(rutas)

    print(f"Salidas: {len(df_salidas)} filas | Valorizado: {len(df_valorizado)} filas")

    print("Calculando medidas (consumo, stock min/max, alertas)...")
    df_resultado = calcular_medidas(df_salidas, df_valorizado)

    print("Subiendo a Neon...")
    subir_a_neon(df_salidas, df_resultado)
