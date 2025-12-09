"""
Script para enriquecer BASE_TX.xlsx con SAP_ID desde MP KEY.xlsx

Proceso:
1. Lee MP KEY.xlsx y extrae KEY y Número de artículo (SAP_ID)
2. Quita todos los espacios de KEY para crear CLAVE_P normalizada
3. Crea tabla temporal de mapeo CLAVE_P -> SAP_ID
4. Enriquece BASE_TX.xlsx con SAP_ID mediante match con CLAVE_P
5. Valida que todas las transacciones tengan SAP_ID (no debería haber errores)
6. Genera BASE_TX_Enriquecida.xlsx
"""

import pandas as pd
import sys
from datetime import datetime
import os

def main():
    print("=" * 80)
    print("PROCESO DE ENRIQUECIMIENTO: TX_Carga + SAP_ID")
    print("=" * 80)
    
    # Rutas de archivos
    archivo_base = "TX_Carga.xlsx"
    archivo_key = "MP KEY.xlsx"
    archivo_salida = f"TX_Carga_Enriquecida_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # ============================================================
    # PASO 1: Cargar MP KEY.xlsx
    # ============================================================
    print("\n[PASO 1] Cargando MP KEY.xlsx...")
    try:
        df_key = pd.read_excel(archivo_key)
        print(f"   ✓ Archivo cargado: {df_key.shape[0]} filas x {df_key.shape[1]} columnas")
    except Exception as e:
        print(f"   ✗ ERROR: No se pudo cargar {archivo_key}")
        print(f"   Detalle: {e}")
        sys.exit(1)
    
    # Verificar columnas requeridas
    if "KEY" not in df_key.columns:
        print("   ✗ ERROR: Columna 'KEY' no encontrada en MP KEY.xlsx")
        print(f"   Columnas disponibles: {list(df_key.columns)}")
        sys.exit(1)
    
    if "Número de artículo" not in df_key.columns:
        print("   ✗ ERROR: Columna 'Número de artículo' no encontrada en MP KEY.xlsx")
        print(f"   Columnas disponibles: {list(df_key.columns)}")
        sys.exit(1)
    
    # ============================================================
    # PASO 2: Normalizar CLAVE_P (quitar todos los espacios de KEY)
    # ============================================================
    print("\n[PASO 2] Normalizando CLAVE_P (quitando espacios de KEY)...")
    df_key['CLAVE_P_NORMALIZADA'] = df_key['KEY'].astype(str).str.replace(' ', '')
    print(f"   ✓ CLAVE_P normalizada creada")
    print(f"   Ejemplos: '{df_key['KEY'].iloc[0]}' -> '{df_key['CLAVE_P_NORMALIZADA'].iloc[0]}'")
    
    # ============================================================
    # PASO 3: Crear tabla temporal de mapeo CLAVE_P -> SAP_ID
    # ============================================================
    print("\n[PASO 3] Creando tabla temporal de mapeo CLAVE_P -> SAP_ID...")
    df_mapeo = df_key[['CLAVE_P_NORMALIZADA', 'Número de artículo']].copy()
    df_mapeo.columns = ['CLAVE_P', 'SAP_ID']
    
    # Eliminar duplicados (mantener el primero si hay duplicados)
    duplicados_antes = df_mapeo.shape[0]
    df_mapeo = df_mapeo.drop_duplicates(subset=['CLAVE_P'], keep='first')
    duplicados_eliminados = duplicados_antes - df_mapeo.shape[0]
    
    if duplicados_eliminados > 0:
        print(f"   ⚠️  ADVERTENCIA: Se encontraron {duplicados_eliminados} CLAVE_P duplicados, se mantuvo el primero")
    
    print(f"   ✓ Tabla de mapeo creada: {df_mapeo.shape[0]} registros únicos")
    print(f"   Ejemplos de mapeo:")
    for i in range(min(3, len(df_mapeo))):
        print(f"      {df_mapeo.iloc[i]['CLAVE_P']} -> {df_mapeo.iloc[i]['SAP_ID']}")
    
    # ============================================================
    # PASO 4: Cargar TX_Carga.xlsx
    # ============================================================
    print("\n[PASO 4] Cargando TX_Carga.xlsx...")
    try:
        df_base = pd.read_excel(archivo_base)
        print(f"   ✓ Archivo cargado: {df_base.shape[0]} filas x {df_base.shape[1]} columnas")
    except Exception as e:
        print(f"   ✗ ERROR: No se pudo cargar {archivo_base}")
        print(f"   Detalle: {e}")
        sys.exit(1)
    
    # Verificar columna CLAVE_P
    if "CLAVE_P" not in df_base.columns:
        print("   ✗ ERROR: Columna 'CLAVE_P' no encontrada en TX_Carga.xlsx")
        print(f"   Columnas disponibles: {list(df_base.columns)}")
        sys.exit(1)
    
    # ============================================================
    # PASO 5: Enriquecer TX_Carga con SAP_ID
    # ============================================================
    print("\n[PASO 5] Enriqueciendo TX_Carga con SAP_ID...")
    
    # Verificar si ya existe columna SAP_ID
    tiene_sap_id = "SAP_ID" in df_base.columns
    if tiene_sap_id:
        filas_con_sap_id_existente = df_base['SAP_ID'].notna().sum()
        filas_sin_sap_id_existente = df_base['SAP_ID'].isna().sum()
        print(f"   ⚠️  TX_Carga ya tiene columna SAP_ID")
        print(f"      Filas con SAP_ID: {filas_con_sap_id_existente}")
        print(f"      Filas sin SAP_ID: {filas_sin_sap_id_existente}")
        
        if filas_sin_sap_id_existente == 0:
            print(f"   ✓ SAP_ID ya está completo, actualizando desde MP KEY...")
            # Eliminar columna SAP_ID existente para reemplazarla
            df_base = df_base.drop(columns=['SAP_ID'])
        else:
            print(f"   ⚠️  SAP_ID incompleto, completando desde MP KEY...")
            # Eliminar solo los valores nulos para reemplazarlos
            df_base = df_base[df_base['SAP_ID'].notna() | df_base['SAP_ID'].isna()].copy()
            df_base = df_base.drop(columns=['SAP_ID'])
    
    # Convertir CLAVE_P a string para el merge
    df_base['CLAVE_P'] = df_base['CLAVE_P'].astype(str)
    df_mapeo['CLAVE_P'] = df_mapeo['CLAVE_P'].astype(str)
    
    # Realizar merge (left join para mantener todas las filas de TX_Carga)
    filas_antes = df_base.shape[0]
    df_enriquecido = df_base.merge(
        df_mapeo[['CLAVE_P', 'SAP_ID']],
        on='CLAVE_P',
        how='left'
    )
    
    print(f"   ✓ Merge completado: {df_enriquecido.shape[0]} filas (esperado: {filas_antes})")
    
    # ============================================================
    # PASO 6: VALIDACIÓN CRÍTICA - Verificar que todas tengan SAP_ID
    # ============================================================
    print("\n[PASO 6] VALIDACIÓN: Verificando que todas las transacciones tengan SAP_ID...")
    
    filas_sin_sap_id = df_enriquecido['SAP_ID'].isna().sum()
    filas_con_sap_id = df_enriquecido['SAP_ID'].notna().sum()
    
    print(f"   Filas con SAP_ID: {filas_con_sap_id}")
    print(f"   Filas sin SAP_ID: {filas_sin_sap_id}")
    
    if filas_sin_sap_id > 0:
        print("\n" + "=" * 80)
        print("⚠️  ALERTA DE INCONSISTENCIA ⚠️")
        print("=" * 80)
        print(f"Se encontraron {filas_sin_sap_id} transacciones SIN SAP_ID.")
        print("Esto NO debería ocurrir ya que las transacciones fueron pre-filtradas.")
        
        # Guardar reporte de inconsistencias en Excel
        archivo_reporte = f"REPORTE_INCONSISTENCIAS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df_sin_sap_completo = df_enriquecido[df_enriquecido['SAP_ID'].isna()].copy()
        
        try:
            # Guardar todas las transacciones sin SAP_ID
            df_sin_sap_completo.to_excel(archivo_reporte, index=False)
            print(f"\n✓ REPORTE GUARDADO: {archivo_reporte}")
            print(f"   Ubicación: {os.path.abspath(archivo_reporte)}")
            print(f"   Total de filas sin SAP_ID: {len(df_sin_sap_completo):,}")
        except Exception as e:
            print(f"\n⚠️  No se pudo guardar el reporte: {e}")
        
        # Mostrar resumen en consola
        print("\nDetalles de las transacciones sin SAP_ID (primeras 20):")
        df_sin_sap = df_sin_sap_completo[['CLAVE_P', 'Fecha', 'Numero', 'Articulo']].head(20)
        print(df_sin_sap.to_string())
        
        if filas_sin_sap_id > 20:
            print(f"\n... y {filas_sin_sap_id - 20} filas más (ver archivo {archivo_reporte})")
        
        print("\nCLAVE_P únicos sin match:")
        clave_p_sin_match = df_sin_sap_completo['CLAVE_P'].unique()
        print(f"   Total de CLAVE_P únicos sin match: {len(clave_p_sin_match)}")
        for clave in clave_p_sin_match[:10]:
            print(f"   - {clave}")
        if len(clave_p_sin_match) > 10:
            print(f"   ... y {len(clave_p_sin_match) - 10} más (ver archivo {archivo_reporte})")
        
        print("\n" + "=" * 80)
        print("PROCESO DETENIDO POR INCONSISTENCIA")
        print(f"Revisa el archivo: {archivo_reporte}")
        print("=" * 80)
        sys.exit(1)
    
    print("   ✓ VALIDACIÓN EXITOSA: Todas las transacciones tienen SAP_ID")
    
    # ============================================================
    # PASO 7: Guardar archivo enriquecido
    # ============================================================
    print(f"\n[PASO 7] Guardando archivo enriquecido: {archivo_salida}...")
    
    try:
        df_enriquecido.to_excel(archivo_salida, index=False)
        print(f"   ✓ Archivo guardado exitosamente")
        print(f"   Ubicación: {os.path.abspath(archivo_salida)}")
    except Exception as e:
        print(f"   ✗ ERROR: No se pudo guardar el archivo")
        print(f"   Detalle: {e}")
        sys.exit(1)
    
    # ============================================================
    # RESUMEN FINAL
    # ============================================================
    print("\n" + "=" * 80)
    print("RESUMEN DEL PROCESO")
    print("=" * 80)
    print(f"Archivo origen: {archivo_base}")
    print(f"Archivo clave: {archivo_key}")
    print(f"Archivo salida: {archivo_salida}")
    print(f"\nEstadísticas:")
    print(f"  - Filas procesadas: {df_enriquecido.shape[0]:,}")
    print(f"  - Columnas en salida: {df_enriquecido.shape[1]}")
    print(f"  - Registros en tabla de mapeo: {df_mapeo.shape[0]:,}")
    print(f"  - Transacciones con SAP_ID: {filas_con_sap_id:,} (100%)")
    print(f"\n✓ PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 80)

if __name__ == "__main__":
    main()

