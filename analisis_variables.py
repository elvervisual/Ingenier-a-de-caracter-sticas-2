#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para cargar datos e identificar:
1. Número total de variables (columnas).
2. Tipo de datos técnico (Pandas) y estadístico (Numéricas vs Categóricas).
3. Clasificación y detalle de cada variable.
"""

import os
import sys
from pathlib import Path
import pandas as pd


def buscar_archivo_datos():
    """
    Busca de manera flexible el archivo de datos:
    1. Si se pasa como argumento de línea de comandos.
    2. En la subcarpeta 'datos/'.
    3. En el directorio actual (.).
    """
    # 1. Argumento por línea de comandos
    if len(sys.argv) > 1:
        ruta = Path(sys.argv[1])
        if ruta.exists():
            return ruta

    directorio_actual = Path(__file__).resolve().parent

    # 2. Buscar en carpeta 'datos'
    carpeta_datos = directorio_actual / "datos"
    if carpeta_datos.exists() and carpeta_datos.is_dir():
        archivos = list(carpeta_datos.glob("*.xlsx")) + list(carpeta_datos.glob("*.csv")) + list(carpeta_datos.glob("*.xls"))
        if archivos:
            return archivos[0]

    # 3. Archivo específico conocido o cualquier excel/csv en la carpeta actual
    archivo_especifico = directorio_actual / "2-1-ingenieria-caracteristicas.xlsx"
    if archivo_especifico.exists():
        return archivo_especifico

    candidatos = list(directorio_actual.glob("*.xlsx")) + list(directorio_actual.glob("*.csv")) + list(directorio_actual.glob("*.xls"))
    # Excluir archivos temporales que empiezan por ~$
    candidatos = [c for c in candidatos if not c.name.startswith("~$")]
    if candidatos:
        return candidatos[0]

    return None


def cargar_dataset(ruta_archivo):
    """
    Carga el dataset y el diccionario si está disponible (en archivos Excel).
    """
    diccionario_df = None
    nombre_hoja = None

    if ruta_archivo.suffix.lower() in [".xlsx", ".xls"]:
        excel = pd.ExcelFile(ruta_archivo)
        hojas = excel.sheet_names

        # Si existe hoja 'hogares', usamos esa como principal
        if "hogares" in hojas:
            nombre_hoja = "hogares"
        else:
            nombre_hoja = hojas[0]

        df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja)

        # Si hay hoja 'diccionario', la cargamos para enriquecer el análisis
        if "diccionario" in hojas:
            diccionario_df = pd.read_excel(ruta_archivo, sheet_name="diccionario")
    else:
        df = pd.read_csv(ruta_archivo)

    return df, diccionario_df, nombre_hoja


def analizar_variables(df, diccionario_df=None):
    """
    Analiza y clasifica las variables en Numéricas, Categóricas y Especiales (Fechas/IDs).
    """
    total_filas, total_columnas = df.shape

    categoricas = []
    numericas = []
    fechas = []
    identificadores = []

    resumen_filas = []

    # Mapa del diccionario si existe
    tipo_diccionario = {}
    if diccionario_df is not None and "variable" in diccionario_df.columns and "tipo" in diccionario_df.columns:
        tipo_diccionario = dict(zip(diccionario_df["variable"], diccionario_df["tipo"]))

    for col in df.columns:
        dtype = df[col].dtype
        n_unicos = df[col].nunique(dropna=True)
        n_nulos = df[col].isnull().sum()
        pct_nulos = (n_nulos / total_filas) * 100

        # Clasificación
        clasificacion = ""
        # 1. Fechas
        if pd.api.types.is_datetime64_any_dtype(dtype) or "fecha" in col.lower() or "date" in col.lower():
            fechas.append(col)
            clasificacion = "Fecha / Temporal"
        # 2. Identificadores
        elif col.lower().startswith("id_") or col.lower() == "id":
            identificadores.append(col)
            clasificacion = "Identificador (ID)"
        # 3. Categóricas (texto / object / category / booleano o discreto con significado cualitativo)
        elif pd.api.types.is_object_dtype(dtype) or pd.api.types.is_categorical_dtype(dtype) or pd.api.types.is_bool_dtype(dtype):
            categoricas.append(col)
            clasificacion = "Categórica"
        # 4. Numéricas (enteros o flotantes)
        elif pd.api.types.is_numeric_dtype(dtype):
            # Caso especial: variables ordinales o binarias que pueden estar codificadas como números
            # Si el diccionario lo define como categórica/ordinal/binaria, lo indicamos
            tipo_teorico = tipo_diccionario.get(col, "")
            if "ordinal" in tipo_teorico.lower() or "categórica" in tipo_teorico.lower():
                categoricas.append(col)
                clasificacion = "Categórica (Ordinal/Numérica)"
            elif col.lower() in ["estrato"]:  # convención usual
                categoricas.append(col)
                clasificacion = "Categórica (Ordinal/Numérica)"
            elif col.lower() in ["pobreza_multidimensional"] or n_unicos == 2 and set(df[col].dropna().unique()).issubset({0, 1}):
                categoricas.append(col)
                clasificacion = "Categórica (Binaria 0/1)"
            else:
                numericas.append(col)
                if pd.api.types.is_integer_dtype(dtype):
                    clasificacion = "Numérica Discreta"
                else:
                    clasificacion = "Numérica Continua"
        else:
            categoricas.append(col)
            clasificacion = "Categórica / Otro"

        # Ejemplo o rango de valores
        if pd.api.types.is_numeric_dtype(dtype) and "Categórica" not in clasificacion:
            val_min = df[col].min()
            val_max = df[col].max()
            ejemplo = f"Mín: {val_min} | Máx: {val_max}"
        else:
            ejemplos_lista = df[col].dropna().unique()[:3]
            ejemplo = ", ".join(str(x) for x in ejemplos_lista)
            if len(df[col].dropna().unique()) > 3:
                ejemplo += ", ..."

        tipo_teorico_str = tipo_diccionario.get(col, "—")

        resumen_filas.append({
            "Variable": col,
            "Tipo Pandas": str(dtype),
            "Clasificación": clasificacion,
            "Tipo en Diccionario": tipo_teorico_str,
            "Valores Únicos": n_unicos,
            "Nulos (Faltantes)": f"{n_nulos} ({pct_nulos:.1f}%)",
            "Muestra / Rango": ejemplo
        })

    return {
        "total_filas": total_filas,
        "total_columnas": total_columnas,
        "categoricas": categoricas,
        "numericas": numericas,
        "fechas": fechas,
        "identificadores": identificadores,
        "tabla_resumen": pd.DataFrame(resumen_filas)
    }


def imprimir_reporte(ruta_archivo, nombre_hoja, resultados):
    ancho = 90
    print("=" * ancho)
    print(" ANÁLISIS Y CLASIFICACIÓN DE VARIABLES DEL DATASET ".center(ancho, "="))
    print("=" * ancho)
    print(f"📁 Archivo cargado : {ruta_archivo.name}")
    print(f"📍 Ruta completa   : {ruta_archivo}")
    if nombre_hoja:
        print(f"📄 Hoja analizada  : {nombre_hoja}")
    print(f"📊 Dimensiones     : {resultados['total_filas']} registros (filas) y {resultados['total_columnas']} variables (columnas)")
    print("-" * ancho)

    print("\n" + " RESUMEN GENERAL DE VARIABLES ".center(ancho, "-"))
    print(f"• Total de variables analizadas : {resultados['total_columnas']}")
    print(f"• Variables Numéricas           : {len(resultados['numericas'])}")
    print(f"• Variables Categóricas         : {len(resultados['categoricas'])}")
    print(f"• Variables Temporales / Fecha  : {len(resultados['fechas'])}")
    print(f"• Variables Identificadoras (ID): {len(resultados['identificadores'])}")

    print("\n" + " DETALLE: VARIABLES NUMÉRICAS ".center(ancho, "-"))
    if resultados['numericas']:
        for i, var in enumerate(resultados['numericas'], 1):
            print(f"  {i:2d}. {var}")
    else:
        print("  (No se encontraron variables numéricas)")

    print("\n" + " DETALLE: VARIABLES CATEGÓRICAS ".center(ancho, "-"))
    if resultados['categoricas']:
        for i, var in enumerate(resultados['categoricas'], 1):
            print(f"  {i:2d}. {var}")
    else:
        print("  (No se encontraron variables categóricas)")

    if resultados['fechas']:
        print("\n" + " DETALLE: VARIABLES TEMPORALES (FECHAS) ".center(ancho, "-"))
        for i, var in enumerate(resultados['fechas'], 1):
            print(f"  {i:2d}. {var}")

    if resultados['identificadores']:
        print("\n" + " DETALLE: IDENTIFICADORES (IDs) ".center(ancho, "-"))
        for i, var in enumerate(resultados['identificadores'], 1):
            print(f"  {i:2d}. {var}")

    print("\n" + " TABLA DETALLADA DE CADA VARIABLE ".center(ancho, "-"))
    tabla = resultados["tabla_resumen"]
    
    # Formatear la tabla con columnas seleccionadas para visualización limpia en consola
    columnas_mostrar = ["Variable", "Tipo Pandas", "Clasificación", "Valores Únicos", "Nulos (Faltantes)"]
    print(tabla[columnas_mostrar].to_string(index=False))
    print("=" * ancho)


def main():
    ruta_archivo = buscar_archivo_datos()

    if ruta_archivo is None:
        print("❌ Error: No se encontró ningún archivo de datos (.xlsx, .xls o .csv).")
        print("Verifica que el archivo se encuentre en este directorio o dentro de una carpeta llamada 'datos/'.")
        sys.exit(1)

    try:
        df, diccionario_df, nombre_hoja = cargar_dataset(ruta_archivo)
        resultados = analizar_variables(df, diccionario_df)
        imprimir_reporte(ruta_archivo, nombre_hoja, resultados)

    except Exception as e:
        print(f"❌ Ocurrió un error al procesar el archivo: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
