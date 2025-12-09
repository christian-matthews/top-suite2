#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para explicar la lógica de cálculo de transacciones afectadas
"""

import pandas as pd

print("=" * 70)
print("📊 EXPLICACIÓN: CÓMO SE CALCULAN LAS TRANSACCIONES AFECTADAS")
print("=" * 70)

print("""
La lógica funciona en 2 pasos:

1️⃣  PRIMER PASO: Se evalúa CADA LÍNEA individualmente
   - Cada línea se marca como 'OK' si su CLAVE_P está en MP KEY
   - O como 'NO_ENCONTRADO' si su CLAVE_P NO está en MP KEY

2️⃣  SEGUNDO PASO: Se agrupa por DOCUMENTO/FACTURA (Numero)
   - Si TODAS las líneas de un documento tienen 'OK' → TODO el documento va a CARGAR
   - Si ALGUNA línea tiene 'NO_ENCONTRADO' → TODO el documento va a PENDIENTES
   (Esto incluye TODAS las líneas del documento, incluso las que tienen código válido)

EJEMPLO:
""")

print("Documento 001:")
print("  Línea 1: CLAVE_P='ABC' → ✅ Encontrado → STATUS='OK'")
print("  Línea 2: CLAVE_P='XYZ' → ❌ NO encontrado → STATUS='NO_ENCONTRADO'")
print("  Línea 3: CLAVE_P='DEF' → ✅ Encontrado → STATUS='OK'")
print("")
print("  Resultado: Documento 001 → PENDIENTE (porque tiene al menos 1 línea con código faltante)")
print("  Transacciones afectadas: 3 líneas (TODAS las líneas del documento)")
print("")
print("=" * 70)
print("")

# Cargar archivo de pendientes para mostrar estadísticas reales
try:
    archivo_pendientes = "TX_Pendientes_20251119_160825.xlsx"
    df_pendientes = pd.read_excel(archivo_pendientes, sheet_name='Sheet1')
    
    # Detectar columnas
    def detectar_columna(df, posibles):
        for col in df.columns:
            col_str = str(col).upper().strip()
            for posible in posibles:
                if posible.upper() in col_str:
                    return col
        return None
    
    col_enc = detectar_columna(df_pendientes, ['ENC'])
    
    if col_enc is not None:
        # Contar líneas con código encontrado vs no encontrado
        total_lineas = len(df_pendientes)
        lineas_con_codigo = (df_pendientes[col_enc] == 1).sum()
        lineas_sin_codigo = (df_pendientes[col_enc] == 0).sum()
        
        print("📈 ESTADÍSTICAS REALES DEL ARCHIVO PENDIENTES:")
        print("-" * 70)
        print(f"Total de líneas en pendientes: {total_lineas:,}")
        print(f"  • Líneas CON código válido (ENC=1): {lineas_con_codigo:,}")
        print(f"  • Líneas SIN código válido (ENC=0): {lineas_sin_codigo:,}")
        print("")
        print("💡 CONCLUSIÓN:")
        print(f"Las {total_lineas:,} transacciones afectadas incluyen:")
        print(f"  - {lineas_sin_codigo:,} líneas que NO tienen código válido")
        print(f"  - {lineas_con_codigo:,} líneas que SÍ tienen código válido")
        print("    (pero están en documentos que tienen al menos 1 línea sin código)")
        print("")
        print(f"Por eso el total es {total_lineas:,} y no solo {lineas_sin_codigo:,}")
        
except Exception as e:
    print(f"⚠️  No se pudo cargar el archivo para mostrar estadísticas: {e}")

print("=" * 70)


