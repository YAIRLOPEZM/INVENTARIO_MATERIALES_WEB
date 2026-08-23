"""
Script para actualizar los datos de INVENTARIO_MATERIALES en Neon,
leyendo DIRECTAMENTE los reportes crudos de Novasoft:

  - INV0206: Kárdex de inventario por bodega/producto (movimientos)
  - INV0230: Inventario valorizado por bodegas (stock actual)

Puedes pasar VARIOS archivos de distintas bodegas juntos, en
cualquier orden — el script detecta solo si cada uno es un
INV0206 o un INV0230, leyendo el encabezado del reporte.

CÓMO USARLO:

   python actualizar_datos.py "ruta\\INV0206_bodega1.xlsx" "ruta\\INV0230_bodega1.xlsx" "ruta\\INV0206_bodega2.xlsx" "ruta\\INV0230_bodega2.xlsx"

   (agrega tantos archivos como bodegas tengas, mezclados sin problema)

REQUISITO: la variable de entorno DATABASE_URL debe apuntar a tu
base de datos de Neon del proyecto INVENTARIO_MATERIALES.

============================================================
REGLAS DE NEGOCIO REPRODUCIDAS (del modelo original en Power BI):

- Consumo_Promedio_Mensual = Salidas del año actual / meses con consumo > 0
- Punto_Reposicion  = Consumo_Promedio_Mensual * 3
- Stock_Minimos     = Consumo_Promedio_Mensual * 1 (lead time) + 50% seguridad
                     = Consumo_Promedio_Mensual * 1.5
- Stock_Maximo      = Consumo_Promedio_Mensual * 3 (lead time + revisión) + 50% seguridad
                     = Consumo_Promedio_Mensual * 3.5
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


DATABASE_URL = os.environ.get("DATABASE_URL", "")


# ============================================================
# DETECCIÓN AUTOMÁTICA DEL TIPO DE REPORTE
# ============================================================

def detectar_tipo_reporte(ruta):
    df = pd.read_excel(ruta, header=None, nrows=15)
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


# ============================================================
# PARSER: INV0206 (kárdex de movimientos)
# ============================================================

def parsear_kardex(ruta):
    df = pd.read_excel(ruta, header=None)

    bodega_codigo = None
    bodega_nombre = None
    item_codigo = None
    item_descripcion = None

    movimientos = []

    for _, fila in df.iterrows():
        col2 = fila[2] if len(fila) > 2 else None
        col9 = fila[9] if len(fila) > 9 else None

        if isinstance(col2, str) and col2.strip().startswith("BODEGA:"):
            m = re.match(r"BODEGA:\s*(\S+)\s+(.*)", col2.strip())
            if m:
                bodega_codigo, bodega_nombre = m.group(1), m.group(2).strip()
            continue

        if isinstance(col2, str) and col2.strip().startswith("ITEM:"):
            texto = col2.strip().replace("ITEM:", "", 1).strip()
            if " - " in texto:
                item_codigo, item_descripcion = texto.split(" - ", 1)
                item_codigo = item_codigo.strip()
                item_descripcion = item_descripcion.strip()
            continue

        if isinstance(col9, str) and "Saldo" in col9:
            continue  # "Saldo Anterior:" / "Saldo Nuevo:" -> no son movimientos

        if isinstance(col2, (pd.Timestamp, datetime.datetime)):
            movimientos.append({
                "bodega_codigo": bodega_codigo,
                "bodega_nombre": bodega_nombre,
                "item_codigo": item_codigo,
                "item_descripcion": item_descripcion,
                "fecha_movimiento": col2,
                "tipo": str(fila[4]).strip() if pd.notna(fila[4]) else None,
                "numero": str(fila[5]).strip() if pd.notna(fila[5]) else None,
                "ped_orc": str(fila[7]).strip() if pd.notna(fila[7]) else None,
                "entradas_san": float(fila[9]) if pd.notna(fila[9]) else 0.0,
                "salidas_san": float(fila[12]) if pd.notna(fila[12]) else 0.0,
                "saldo_parcial": float(fila[17]) if pd.notna(fila[17]) else None,
            })

    return pd.DataFrame(movimientos)


# ============================================================
# PARSER: INV0230 (inventario valorizado)
# ============================================================

def parsear_valorizado(ruta):
    df = pd.read_excel(ruta, header=None)

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


# ============================================================
# COMBINAR TODOS LOS ARCHIVOS (varias bodegas, mezclados)
# ============================================================

def cargar_archivos(rutas):
    movimientos_por_archivo = []
    valorizado_por_archivo = []

    for ruta in rutas:
        tipo = detectar_tipo_reporte(ruta)
        if tipo == "INV0206":
            print(f"  {ruta} -> kárdex (INV0206)")
            movimientos_por_archivo.append(parsear_kardex(ruta))
        elif tipo == "INV0230":
            print(f"  {ruta} -> valorizado (INV0230)")
            valorizado_por_archivo.append(parsear_valorizado(ruta))
        else:
            print(f"  ⚠️  {ruta}: no se reconoce el tipo de reporte (se omite)")

    df_movimientos = (
        pd.concat(movimientos_por_archivo, ignore_index=True)
        if movimientos_por_archivo else pd.DataFrame()
    )
    df_valorizado = (
        pd.concat(valorizado_por_archivo, ignore_index=True)
        if valorizado_por_archivo else pd.DataFrame()
    )

    return df_movimientos, df_valorizado


# ============================================================
# CALCULAR LA TABLA FINAL (equivalente a las medidas DAX)
# ============================================================

def calcular_medidas(df_movimientos, df_valorizado):

    anio_actual = datetime.date.today().year

    mov_anio = df_movimientos[
        pd.to_datetime(df_movimientos["fecha_movimiento"]).dt.year == anio_actual
    ].copy()
    mov_anio["anio_mes"] = pd.to_datetime(mov_anio["fecha_movimiento"]).dt.strftime("%Y-%m")

    # Consumo_Promedio_Mensual: salidas del año actual / meses con consumo > 0
    consumo = (
        mov_anio[mov_anio["salidas_san"] > 0]
        .groupby(["bodega_codigo", "item_codigo"])
        .agg(
            salidas_totales=("salidas_san", "sum"),
            meses_con_consumo=("anio_mes", "nunique"),
        )
        .reset_index()
    )
    consumo["consumo_promedio_mensual"] = (
        consumo["salidas_totales"] / consumo["meses_con_consumo"]
    )

    # Partimos del inventario valorizado (una fila por ítem/bodega)
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


# ============================================================
# SUBIR A NEON
# ============================================================

def subir_a_neon(df_movimientos, df_resultado):

    if not DATABASE_URL:
        print("ERROR: falta configurar la variable de entorno DATABASE_URL.")
        sys.exit(1)

    conexion = psycopg2.connect(DATABASE_URL, sslmode="require")
    cursor = conexion.cursor()

    # ---- INV_MOVIMIENTO ----
    cursor.execute("DROP TABLE IF EXISTS inv_movimiento")
    cursor.execute("""
        CREATE TABLE inv_movimiento (
            bodega_codigo TEXT,
            bodega_nombre TEXT,
            item_codigo TEXT,
            item_descripcion TEXT,
            fecha_movimiento DATE,
            tipo TEXT,
            numero TEXT,
            ped_orc TEXT,
            entradas_san NUMERIC,
            salidas_san NUMERIC,
            saldo_parcial NUMERIC
        )
    """)
    columnas = list(df_movimientos.columns)
    columnas_sql = ", ".join(columnas)
    filas = [
        tuple(None if pd.isna(v) else v for v in fila)
        for fila in df_movimientos.itertuples(index=False, name=None)
    ]
    if filas:
        psycopg2.extras.execute_values(
            cursor,
            f"INSERT INTO inv_movimiento ({columnas_sql}) VALUES %s",
            filas,
            page_size=1000,
        )

    # ---- INVENTARIO (Dim_Valores + medidas calculadas, todo en 1 tabla) ----
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

    conexion.commit()
    cursor.close()
    conexion.close()

    print(f"Listo: {len(df_movimientos)} movimientos y {len(df_resultado)} ítems de inventario subidos a Neon.")


if __name__ == "__main__":

    if len(sys.argv) < 2:
        print('Uso: python actualizar_datos.py "ruta\\INV0206_bodegaX.xlsx" "ruta\\INV0230_bodegaX.xlsx" ...')
        sys.exit(1)

    rutas = sys.argv[1:]

    print("Leyendo y detectando archivos...")
    df_movimientos, df_valorizado = cargar_archivos(rutas)

    print(f"Movimientos: {len(df_movimientos)} filas | Valorizado: {len(df_valorizado)} filas")

    print("Calculando medidas (consumo, stock min/max, alertas)...")
    df_resultado = calcular_medidas(df_movimientos, df_valorizado)

    print("Subiendo a Neon...")
    subir_a_neon(df_movimientos, df_resultado)
