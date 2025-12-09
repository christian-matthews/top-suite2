#!/usr/bin/env python3
"""
Script para contar solo los registros con datos en la columna A
"""

import pandas as pd

def contar_registros_con_datos_columna_a():
    """
    Cuenta solo los registros que tienen datos (no vacíos) en la columna A
    """
    archivo_excel = "/Users/christianmatthews/Library/Mobile Documents/com~apple~CloudDocs/CURSOR/TOP/MAESTRO PRODUCTO/Test1.xlsx"
    
    try:
        print("📊 Leyendo hoja 'BASE GS1 (2)'...")
        df = pd.read_excel(archivo_excel, sheet_name="BASE GS1 (2)")
        
        # Obtener la columna A
        columna_a = df.iloc[:, 0]  # Primera columna
        nombre_columna = df.columns[0]
        
        # Contar solo registros con datos (no NaN, no vacíos)
        registros_con_datos = columna_a.notna().sum()
        
        print(f"\n📋 Análisis de la columna A ('{nombre_columna}'):")
        print(f"   📊 Total de filas en la hoja: {len(df):,}")
        print(f"   ✅ Registros CON datos: {registros_con_datos:,}")
        print(f"   ❌ Registros SIN datos: {(len(df) - registros_con_datos):,}")
        print(f"   📍 Porcentaje con datos: {(registros_con_datos/len(df)*100):.2f}%")
        
        return registros_con_datos
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

if __name__ == "__main__":
    print("=" * 50)
    print("🔍 CONTADOR DE REGISTROS CON DATOS - COLUMNA A")
    print("=" * 50)
    
    cantidad = contar_registros_con_datos_columna_a()
    
    if cantidad:
        print(f"\n🎯 RESULTADO: {cantidad:,} registros con datos en la columna A")
    else:
        print("❌ No se pudo obtener el conteo")


