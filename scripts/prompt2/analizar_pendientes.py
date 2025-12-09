#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de análisis de transacciones pendientes
Responde las 3 preguntas sobre códigos faltantes, transacciones y documentos afectados
"""

import pandas as pd
from pathlib import Path
import glob

# Archivos base
ARCHIVO_MP_KEY = "MP KEY.xlsx"
ARCHIVO_VENTAS = "Ventas JUL-AGO-SEP-OCT.xlsx"

def normalizar_clave(clave: str) -> str:
    """Normaliza una clave para comparación"""
    if pd.isna(clave):
        return ''
    return str(clave).upper().strip().replace(' ', '')

def detectar_columna_clave_p(df: pd.DataFrame, posibles_nombres: list = None) -> str:
    """Detecta automáticamente la columna CLAVE_P"""
    if posibles_nombres is None:
        posibles_nombres = ['CLAVE_P', 'KEY_MS', 'KEY', 'CLAVE PRODUCTO']
    
    df_cols_upper = [str(col).upper().strip() for col in df.columns]
    
    # Priorizar KEY_MS sobre otras columnas
    if 'KEY_MS' in df.columns:
        return 'KEY_MS'
    
    for nombre in posibles_nombres:
        nombre_upper = nombre.upper().strip()
        for idx, col_upper in enumerate(df_cols_upper):
            if nombre_upper in col_upper or col_upper in nombre_upper:
                # Evitar KEY_ECLOUD si hay otras opciones
                if 'ECLOUD' not in col_upper or len(posibles_nombres) == 1:
                    return df.columns[idx]
    
    # Búsqueda más flexible (evitar KEY_ECLOUD)
    for col in df.columns:
        col_str = str(col).upper().strip()
        if any(keyword in col_str for keyword in ['CLAVE', 'KEY']):
            if 'ECLOUD' not in col_str:
                if 'P' in col_str or 'PRODUCTO' in col_str or 'MS' in col_str:
                    return col
    
    return None

def detectar_columna_numero(df: pd.DataFrame) -> str:
    """Detecta automáticamente la columna Numero"""
    posibles = ['NUMERO', 'NÚMERO', 'NUM', 'NRO', 'DOCUMENTO', 'FACTURA']
    df_cols_upper = [str(col).upper().strip() for col in df.columns]
    
    for nombre in posibles:
        nombre_upper = nombre.upper().strip()
        for idx, col_upper in enumerate(df_cols_upper):
            if nombre_upper == col_upper or nombre_upper in col_upper:
                return df.columns[idx]
    
    return None

def obtener_archivo_pendientes_mas_reciente():
    """Obtiene el archivo TX_Pendientes más reciente"""
    archivos = glob.glob("TX_Pendientes_*.xlsx")
    if not archivos:
        archivos = glob.glob("transacciones pendientes*.xlsx")
    
    if not archivos:
        raise FileNotFoundError("No se encontró archivo de transacciones pendientes")
    
    # Ordenar por fecha de modificación
    archivos.sort(key=lambda x: Path(x).stat().st_mtime, reverse=True)
    return archivos[0]

def analizar_transacciones_pendientes():
    """Analiza las transacciones pendientes y responde las 3 preguntas"""
    
    print("=" * 70)
    print("📊 ANÁLISIS DE TRANSACCIONES PENDIENTES")
    print("=" * 70)
    
    # 1. Cargar archivo de ventas original (para totales)
    print("\n📂 Cargando archivo de ventas original...")
    try:
        # Intentar diferentes configuraciones de lectura
        df_ventas = None
        col_clave_p_ventas = None
        col_numero_ventas = None
        
        # Buscar la fila de encabezado que contenga las columnas necesarias
        for i in range(15):
            try:
                df_test = pd.read_excel(ARCHIVO_VENTAS, header=i, nrows=10)
                if len(df_test.columns) > 3:  # Archivo tiene suficientes columnas
                    # Intentar detectar columnas en este encabezado
                    temp_clave_p = detectar_columna_clave_p(df_test, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                    temp_numero = detectar_columna_numero(df_test)
                    
                    if temp_clave_p is not None and temp_numero is not None:
                        # Encontramos el encabezado correcto
                        df_ventas = pd.read_excel(ARCHIVO_VENTAS, header=i)
                        col_clave_p_ventas = temp_clave_p
                        col_numero_ventas = temp_numero
                        break
            except Exception:
                continue
        
        if df_ventas is None:
            # Último intento: leer sin encabezado y buscar manualmente
            df_raw = pd.read_excel(ARCHIVO_VENTAS, header=None)
            # Buscar fila que contenga 'KEY_MS' o 'Numero'
            for i in range(min(15, len(df_raw))):
                row_values = [str(val).upper() for val in df_raw.iloc[i].values if pd.notna(val)]
                if 'KEY_MS' in row_values or 'NUMERO' in row_values:
                    df_ventas = pd.read_excel(ARCHIVO_VENTAS, header=i)
                    col_clave_p_ventas = detectar_columna_clave_p(df_ventas, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                    col_numero_ventas = detectar_columna_numero(df_ventas)
                    break
            
            if df_ventas is None:
                df_ventas = pd.read_excel(ARCHIVO_VENTAS)
                col_clave_p_ventas = detectar_columna_clave_p(df_ventas, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                col_numero_ventas = detectar_columna_numero(df_ventas)
        
        if col_clave_p_ventas is None:
            raise ValueError(f"❌ No se encontró la columna CLAVE_P en {ARCHIVO_VENTAS}. Columnas disponibles: {df_ventas.columns.tolist()}")
        
        if col_numero_ventas is None:
            raise ValueError(f"❌ No se encontró la columna Numero en {ARCHIVO_VENTAS}. Columnas disponibles: {df_ventas.columns.tolist()}")
        
        df_ventas['CLAVE_P'] = df_ventas[col_clave_p_ventas].astype(str).str.strip()
        df_ventas['Numero'] = df_ventas[col_numero_ventas].astype(str).str.strip()
        df_ventas = df_ventas.dropna(subset=['CLAVE_P', 'Numero'])
        df_ventas = df_ventas[(df_ventas['CLAVE_P'] != '') & (df_ventas['Numero'] != '')]
        
        # Totales del archivo original
        total_transacciones = len(df_ventas)
        total_codigos_unicos_ventas = df_ventas['CLAVE_P'].nunique()
        total_documentos_unicos = df_ventas['Numero'].nunique()
        
        print(f"✅ Ventas cargado: {total_transacciones} transacciones")
        print(f"   • Códigos únicos: {total_codigos_unicos_ventas}")
        print(f"   • Documentos únicos: {total_documentos_unicos}")
        
    except Exception as e:
        print(f"❌ Error al cargar ventas: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # 2. Cargar archivo MP KEY (para total de códigos válidos)
    print("\n📂 Cargando archivo MP KEY...")
    try:
        # Intentar diferentes configuraciones de lectura
        df_mp_key = None
        col_clave_p_mp = None
        
        # Buscar la fila de encabezado
        for i in range(10):
            try:
                df_test = pd.read_excel(ARCHIVO_MP_KEY, header=i, nrows=5)
                if len(df_test.columns) > 3:  # Archivo tiene suficientes columnas
                    temp_clave_p = detectar_columna_clave_p(df_test, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
                    if temp_clave_p is not None:
                        df_mp_key = pd.read_excel(ARCHIVO_MP_KEY, header=i)
                        col_clave_p_mp = temp_clave_p
                        break
            except Exception:
                continue
        
        if df_mp_key is None:
            df_mp_key = pd.read_excel(ARCHIVO_MP_KEY)
            col_clave_p_mp = detectar_columna_clave_p(df_mp_key, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
        
        if col_clave_p_mp is None:
            raise ValueError(f"❌ No se encontró la columna CLAVE_P en {ARCHIVO_MP_KEY}. Columnas disponibles: {df_mp_key.columns.tolist()}")
        
        df_mp_key['CLAVE_P'] = df_mp_key[col_clave_p_mp].astype(str).str.strip()
        df_mp_key = df_mp_key.dropna(subset=['CLAVE_P'])
        df_mp_key = df_mp_key[df_mp_key['CLAVE_P'] != '']
        
        # Normalizar para comparación
        df_mp_key['CLAVE_P_NORM'] = df_mp_key['CLAVE_P'].apply(normalizar_clave)
        total_codigos_validos = df_mp_key['CLAVE_P_NORM'].nunique()
        
        print(f"✅ MP KEY cargado: {total_codigos_validos} códigos válidos únicos")
        
    except Exception as e:
        print(f"⚠️  Error al cargar MP KEY: {e}")
        import traceback
        traceback.print_exc()
        total_codigos_validos = None
    
    # 3. Cargar archivo de pendientes más reciente
    print("\n📂 Cargando archivo de transacciones pendientes...")
    # Inicializar variables
    transacciones_afectadas = 0
    documentos_afectados = 0
    lineas_sin_codigo = 0
    lineas_con_codigo = 0
    codigos_faltantes = 0
    
    try:
        archivo_pendientes = obtener_archivo_pendientes_mas_reciente()
        print(f"   Archivo: {archivo_pendientes}")
        
        # Hoja 1: Transacciones pendientes
        df_pendientes = pd.read_excel(archivo_pendientes, sheet_name='Sheet1')
        col_clave_p_pend = detectar_columna_clave_p(df_pendientes, ['CLAVE_P', 'KEY', 'KEY_MS', 'CLAVE PRODUCTO'])
        col_numero_pend = detectar_columna_numero(df_pendientes)
        
        if col_clave_p_pend is None:
            raise ValueError(f"❌ No se encontró columna CLAVE_P en {archivo_pendientes}")
        if col_numero_pend is None:
            raise ValueError(f"❌ No se encontró columna Numero en {archivo_pendientes}")
        
        df_pendientes['CLAVE_P'] = df_pendientes[col_clave_p_pend].astype(str).str.strip()
        df_pendientes['Numero'] = df_pendientes[col_numero_pend].astype(str).str.strip()
        df_pendientes = df_pendientes.dropna(subset=['CLAVE_P', 'Numero'])
        df_pendientes = df_pendientes[(df_pendientes['CLAVE_P'] != '') & (df_pendientes['Numero'] != '')]
        
        transacciones_afectadas = len(df_pendientes)
        documentos_afectados = df_pendientes['Numero'].nunique()
        
        # Contar líneas sin código (ENC=0) y con código (ENC=1)
        col_enc = None
        for col in df_pendientes.columns:
            if str(col).upper().strip() == 'ENC':
                col_enc = col
                break
        
        if col_enc and col_enc in df_pendientes.columns:
            lineas_sin_codigo = (df_pendientes[col_enc] == 0).sum()
            lineas_con_codigo = (df_pendientes[col_enc] == 1).sum()
        else:
            # Si no hay columna ENC, calcular basándose en si el código está en MP KEY
            print("   ⚠️  No se encontró columna ENC, calculando manualmente...")
            if 'df_mp_key' in locals() and df_mp_key is not None and 'CLAVE_P_NORM' in df_mp_key.columns:
                claves_validas = set(df_mp_key['CLAVE_P_NORM'].unique())
                df_pendientes['CLAVE_P_NORM'] = df_pendientes['CLAVE_P'].apply(normalizar_clave)
                df_pendientes['TIENE_CODIGO'] = df_pendientes['CLAVE_P_NORM'].isin(claves_validas)
                lineas_sin_codigo = (~df_pendientes['TIENE_CODIGO']).sum()
                lineas_con_codigo = df_pendientes['TIENE_CODIGO'].sum()
            else:
                print("   ⚠️  No se puede calcular sin MP KEY o columna ENC")
                lineas_sin_codigo = 0
                lineas_con_codigo = 0
        
        print(f"✅ Pendientes cargado: {transacciones_afectadas} transacciones")
        print(f"   • Documentos afectados: {documentos_afectados}")
        print(f"   • Líneas sin código válido: {lineas_sin_codigo:,}")
        print(f"   • Líneas con código válido (en docs afectados): {lineas_con_codigo:,}")
        
        # Hoja 2: Códigos no encontrados
        try:
            df_codigos_faltantes = pd.read_excel(archivo_pendientes, sheet_name='Sheet2')
            if 'CLAVE_P' in df_codigos_faltantes.columns:
                df_codigos_faltantes['CLAVE_P'] = df_codigos_faltantes['CLAVE_P'].astype(str).str.strip()
                df_codigos_faltantes = df_codigos_faltantes.dropna(subset=['CLAVE_P'])
                df_codigos_faltantes = df_codigos_faltantes[df_codigos_faltantes['CLAVE_P'] != '']
                codigos_faltantes = df_codigos_faltantes['CLAVE_P'].nunique()
            else:
                codigos_faltantes = 0
                print(f"⚠️  Hoja 2 no tiene columna CLAVE_P")
        except Exception as e:
            print(f"⚠️  No se pudo leer Hoja 2: {e}")
            codigos_faltantes = 0
        
        print(f"   • Códigos faltantes: {codigos_faltantes}")
        
    except Exception as e:
        print(f"❌ Error al cargar pendientes: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # 4. Calcular y mostrar resultados
    print("\n" + "=" * 70)
    print("📈 RESULTADOS DEL ANÁLISIS")
    print("=" * 70)
    
    # Pregunta 1: Códigos faltantes
    print("\n1️⃣  CÓDIGOS FALTANTES")
    print("-" * 70)
    if total_codigos_validos is not None:
        print(f"   Total de códigos válidos en MP KEY: {total_codigos_validos:,}")
    print(f"   Total de códigos únicos en ventas: {total_codigos_unicos_ventas:,}")
    print(f"   Códigos faltantes (no encontrados): {codigos_faltantes:,}")
    if total_codigos_unicos_ventas > 0:
        porcentaje = (codigos_faltantes / total_codigos_unicos_ventas) * 100
        print(f"   Porcentaje faltante: {porcentaje:.2f}%")
    
    # Pregunta 2: Transacciones afectadas
    print("\n2️⃣  TRANSACCIONES AFECTADAS")
    print("-" * 70)
    print(f"   Total de transacciones: {total_transacciones:,}")
    print("")
    print(f"   📊 LÍNEAS SIN CÓDIGO VÁLIDO:")
    print(f"      • Líneas sin código: {lineas_sin_codigo:,}")
    porcentaje_sin_codigo = (lineas_sin_codigo / total_transacciones) * 100
    print(f"      • Porcentaje del total: {porcentaje_sin_codigo:.2f}%")
    print("")
    print(f"   📊 LÍNEAS AFECTADAS (Todas las líneas de documentos con código faltante):")
    print(f"      • Líneas afectadas: {transacciones_afectadas:,}")
    porcentaje_trans = (transacciones_afectadas / total_transacciones) * 100
    print(f"      • Porcentaje del total: {porcentaje_trans:.2f}%")
    print(f"      • Incluye {lineas_con_codigo:,} líneas con código válido")
    print(f"        (que están en documentos con al menos 1 línea sin código)")
    
    # Pregunta 3: Documentos afectados
    print("\n3️⃣  DOCUMENTOS/FACTURAS AFECTADAS")
    print("-" * 70)
    print(f"   Total de documentos únicos: {total_documentos_unicos:,}")
    print(f"   Documentos afectados: {documentos_afectados:,}")
    porcentaje_doc = (documentos_afectados / total_documentos_unicos) * 100
    print(f"   Porcentaje afectado: {porcentaje_doc:.2f}%")
    
    print("\n" + "=" * 70)
    print("✅ Análisis completado")
    print("=" * 70)

if __name__ == "__main__":
    analizar_transacciones_pendientes()

