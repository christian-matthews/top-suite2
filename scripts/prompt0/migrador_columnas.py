#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Programa para migrar datos de formato Antiguo a Final
usando la lógica de mapeo de Training
"""

import pandas as pd
import numpy as np
from pathlib import Path
import sys
from difflib import SequenceMatcher
from datetime import datetime
import shutil

def similaridad(a, b):
    """Calcula la similaridad entre dos strings"""
    if pd.isna(a) or pd.isna(b):
        return 0.0
    return SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def analizar_estructura(df):
    """Analiza la estructura de un DataFrame"""
    info = {
        'num_filas': len(df),
        'num_columnas': len(df.columns),
        'columnas': list(df.columns),
        'tipos': df.dtypes.to_dict(),
        'valores_nulos': df.isnull().sum().to_dict(),
        'ejemplos': {}
    }
    
    # Obtener ejemplos de valores no nulos para cada columna
    for col in df.columns:
        valores_no_nulos = df[col].dropna()
        if len(valores_no_nulos) > 0:
            info['ejemplos'][col] = valores_no_nulos.head(3).tolist()
    
    return info

def estimar_mapeo_columnas(df_antiguo, df_training, df_final):
    """
    Estima la relación entre columnas de Antiguo y Final
    usando la lógica de Training
    """
    print("\n" + "="*80)
    print("FASE 1: ANÁLISIS Y ESTIMACIÓN DE MAPEO DE COLUMNAS")
    print("="*80)
    
    # Analizar estructuras
    print("\n[1] Analizando estructura de las hojas...")
    info_antiguo = analizar_estructura(df_antiguo)
    info_training = analizar_estructura(df_training)
    info_final = analizar_estructura(df_final)
    
    print(f"\n  Hoja 'ANTIGUO': {info_antiguo['num_filas']} filas, {info_antiguo['num_columnas']} columnas")
    print(f"  Hoja 'TRAINING': {info_training['num_filas']} filas, {info_training['num_columnas']} columnas")
    print(f"  Hoja 'FINAL': {info_final['num_filas']} filas, {info_final['num_columnas']} columnas")
    
    # Mapeo basado en nombres de columnas
    print("\n[2] Analizando similitud de nombres de columnas...")
    mapeo_por_nombre = {}
    
    for col_final in df_final.columns:
        mejor_match = None
        mejor_score = 0.0
        
        for col_antiguo in df_antiguo.columns:
            score = similaridad(col_final, col_antiguo)
            if score > mejor_score:
                mejor_score = score
                mejor_match = col_antiguo
        
        if mejor_score > 0.3:  # Umbral mínimo de similitud
            mapeo_por_nombre[col_final] = {
                'columna_antigua': mejor_match,
                'score_nombre': mejor_score
            }
    
    # Mapeo basado en tipos de datos y valores usando TRAINING como referencia
    print("\n[3] Analizando similitud de tipos de datos y valores usando TRAINING...")
    mapeo_por_contenido = {}
    
    for col_final in df_final.columns:
        mejor_match = None
        mejor_score = 0.0
        
        tipo_final = df_final[col_final].dtype
        
        for col_antiguo in df_antiguo.columns:
            tipo_antiguo = df_antiguo[col_antiguo].dtype
            
            # Comparar tipos
            score_tipo = 1.0 if tipo_final == tipo_antiguo else 0.5
            
            # Comparar valores usando TRAINING como referencia
            score_valores = 0.0
            if col_final in df_training.columns and col_antiguo in df_antiguo.columns:
                # Comparar distribuciones de valores únicos entre TRAINING y ANTIGUO
                valores_training = set(df_training[col_final].dropna().astype(str))
                valores_antiguo = set(df_antiguo[col_antiguo].dropna().astype(str))
                
                if len(valores_training) > 0 and len(valores_antiguo) > 0:
                    interseccion = valores_training.intersection(valores_antiguo)
                    union = valores_training.union(valores_antiguo)
                    score_valores = len(interseccion) / len(union) if len(union) > 0 else 0.0
                    
                    # Bonus si hay muchos valores en común (mayor confianza)
                    if len(interseccion) > 10:
                        score_valores = min(1.0, score_valores * 1.2)
            
            score_total = (score_tipo * 0.2) + (score_valores * 0.8)
            
            if score_total > mejor_score:
                mejor_score = score_total
                mejor_match = col_antiguo
        
        if mejor_score > 0.15:  # Umbral más alto para contenido
            mapeo_por_contenido[col_final] = {
                'columna_antigua': mejor_match,
                'score_contenido': mejor_score
            }
    
    # Combinar mapeos
    print("\n[4] Combinando mapeos...")
    mapeo_final = {}
    
    for col_final in df_final.columns:
        mapeo = {
            'columna_final': col_final,
            'columna_antigua': None,
            'confianza': 0.0,
            'metodo': None,
            'score_nombre': 0.0,
            'score_contenido': 0.0,
            'alternativas_nombre': [],
            'alternativas_contenido': []
        }
        
        # Priorizar mapeo por contenido si existe y tiene buena confianza
        if col_final in mapeo_por_contenido:
            score_contenido = mapeo_por_contenido[col_final]['score_contenido']
            mapeo['score_contenido'] = score_contenido
            # Si el score de contenido es alto (>0.3), priorizarlo sobre nombre
            if score_contenido > 0.3:
                mapeo['columna_antigua'] = mapeo_por_contenido[col_final]['columna_antigua']
                mapeo['confianza'] = score_contenido
                mapeo['metodo'] = 'contenido'
        
        # Si no hay mapeo por contenido o es de baja confianza, usar nombre
        if not mapeo['columna_antigua'] and col_final in mapeo_por_nombre:
            mapeo['columna_antigua'] = mapeo_por_nombre[col_final]['columna_antigua']
            mapeo['confianza'] = mapeo_por_nombre[col_final]['score_nombre']
            mapeo['score_nombre'] = mapeo_por_nombre[col_final]['score_nombre']
            mapeo['metodo'] = 'nombre'
            
            # Guardar alternativas por nombre
            for col_ant in df_antiguo.columns:
                score = similaridad(col_final, col_ant)
                if score > 0.2 and col_ant != mapeo['columna_antigua']:
                    mapeo['alternativas_nombre'].append({
                        'columna': col_ant,
                        'score': score
                    })
            mapeo['alternativas_nombre'].sort(key=lambda x: x['score'], reverse=True)
        
        mapeo_final[col_final] = mapeo
    
    # Aplicar reglas especiales después de construir todos los mapeos
    print("\n[5] Aplicando reglas especiales de mapeo...")
    
    # Primero, identificar qué columnas de ANTIGUO están mapeadas a qué columnas de FINAL
    mapeo_inverso = {}  # columna_antigua -> lista de columnas_final
    for col_final, info in mapeo_final.items():
        if info['columna_antigua']:
            col_ant = info['columna_antigua']
            if col_ant not in mapeo_inverso:
                mapeo_inverso[col_ant] = []
            mapeo_inverso[col_ant].append(col_final)
    
    # Regla 1: EAN ANTIGUO debe mapearse desde EAN13 de ANTIGUO
    for col_final, info in mapeo_final.items():
        col_final_clean = col_final.strip().upper()
        if 'EAN ANTIGUO' in col_final_clean and 'EAN13' in df_antiguo.columns:
            # Si EAN13 está mapeado a EAN13 (columna final), cambiar EAN ANTIGUO para usar EAN13
            if 'EAN13' in mapeo_inverso.get('EAN13', []):
                # EAN13 ya está mapeado a EAN13, entonces EAN ANTIGUO también debe usar EAN13
                info['columna_antigua'] = 'EAN13'
                info['confianza'] = 0.9
                info['metodo'] = 'regla_especial'
            elif info['columna_antigua'] != 'EAN13' or info['confianza'] < 0.5:
                # Forzar mapeo a EAN13
                info['columna_antigua'] = 'EAN13'
                info['confianza'] = 0.9
                info['metodo'] = 'regla_especial'
    
    # Regla 2: EAN NUEVO no debe tener mapeo (debe quedar vacío) a menos que sea muy seguro
    for col_final, info in mapeo_final.items():
        col_final_clean = col_final.strip().upper()
        if 'EAN NUEVO' in col_final_clean:
            # Solo mantener mapeo si hay coincidencia muy fuerte (>0.85)
            if info['confianza'] < 0.85:
                info['columna_antigua'] = None
                info['confianza'] = 0.0
                info['metodo'] = None
    
    # Regla 3: Verificador debe quedar vacío (no mapear valores)
    for col_final, info in mapeo_final.items():
        col_final_clean = col_final.strip().upper()
        if 'VERIFICADOR' in col_final_clean:
            # Siempre dejar sin mapeo (vacío)
            info['columna_antigua'] = None
            info['confianza'] = 0.0
            info['metodo'] = None
    
    # Regla 4: EAN13 en FINAL no debe mapearse desde EAN13 de ANTIGUO
    # porque EAN13 de ANTIGUO va a EAN ANTIGUO
    for col_final, info in mapeo_final.items():
        col_final_clean = col_final.strip().upper()
        if col_final_clean == 'EAN13' and info['columna_antigua'] == 'EAN13':
            # Verificar si EAN13 está siendo usado para EAN ANTIGUO
            ean13_usado_en_antiguo = False
            for otro_col, otro_info in mapeo_final.items():
                if otro_col != col_final and 'EAN ANTIGUO' in otro_col.strip().upper():
                    if otro_info['columna_antigua'] == 'EAN13':
                        ean13_usado_en_antiguo = True
                        break
            
            # Si EAN13 está siendo usado para EAN ANTIGUO, entonces EAN13 en FINAL debe quedar vacío
            if ean13_usado_en_antiguo:
                info['columna_antigua'] = None
                info['confianza'] = 0.0
                info['metodo'] = None
    
    return mapeo_final, info_antiguo, info_training, info_final, mapeo_por_nombre, mapeo_por_contenido

def mostrar_mapeo(mapeo_final):
    """Muestra el mapeo propuesto de forma clara"""
    print("\n" + "="*80)
    print("MAPEO PROPUESTO: Antiguo → Final")
    print("="*80)
    print(f"\n{'Columna Final':<30} {'Columna Antigua':<30} {'Confianza':<12} {'Método':<15}")
    print("-" * 90)
    
    mapeos_validos = []
    mapeos_no_encontrados = []
    
    for col_final, info in mapeo_final.items():
        if info['columna_antigua']:
            confianza_pct = f"{info['confianza']*100:.1f}%"
            print(f"{col_final:<30} {info['columna_antigua']:<30} {confianza_pct:<12} {info['metodo']:<15}")
            mapeos_validos.append(info)
        else:
            print(f"{col_final:<30} {'NO ENCONTRADA':<30} {'0.0%':<12} {'-':<15}")
            mapeos_no_encontrados.append(col_final)
    
    print("\n" + "-" * 90)
    print(f"Total columnas en Final: {len(mapeo_final)}")
    print(f"Columnas mapeadas: {len(mapeos_validos)}")
    print(f"Columnas sin mapeo: {len(mapeos_no_encontrados)}")
    
    if mapeos_no_encontrados:
        print(f"\n⚠️  Columnas sin mapeo encontrado:")
        for col in mapeos_no_encontrados:
            print(f"   - {col}")
    
    return mapeos_validos, mapeos_no_encontrados

def guardar_reporte_mapeo(mapeo_final, archivo_salida, df_antiguo, df_training, carpeta_ejecucion):
    """Guarda un reporte detallado del mapeo en Excel"""
    print("\n[5] Generando reporte detallado del mapeo...")
    
    # Crear DataFrame con el reporte
    datos_reporte = []
    
    for col_final, info in mapeo_final.items():
        fila = {
            'Columna Final': col_final,
            'Columna Antigua': info['columna_antigua'] if info['columna_antigua'] else 'NO ENCONTRADA',
            'Confianza (%)': f"{info['confianza']*100:.1f}%",
            'Método': info['metodo'] if info['metodo'] else '-',
            'Score Nombre': f"{info['score_nombre']*100:.1f}%" if info['score_nombre'] > 0 else '-',
            'Score Contenido': f"{info['score_contenido']*100:.1f}%" if info['score_contenido'] > 0 else '-',
            'Racional': ''
        }
        
        # Construir el racional (solo si hay información relevante)
        racional_parts = []
        
        # Reglas especiales para racional
        col_final_clean = col_final.strip().upper()
        
        if 'VERIFICADOR' in col_final_clean and not info['columna_antigua']:
            # Verificador debe quedar vacío
            fila['Racional'] = 'Dejar en blanco valor inutil'
        elif info['columna_antigua']:
            if info['metodo'] == 'nombre':
                racional_parts.append(f"Coincidencia por nombre: {info['score_nombre']*100:.1f}%")
                if info['alternativas_nombre']:
                    alt = info['alternativas_nombre'][0]
                    racional_parts.append(f"Alternativa más cercana: '{alt['columna']}' ({alt['score']*100:.1f}%)")
            elif info['metodo'] == 'contenido':
                racional_parts.append(f"Coincidencia por contenido: {info['score_contenido']*100:.1f}%")
                if info['score_nombre'] > 0:
                    racional_parts.append(f"Score nombre: {info['score_nombre']*100:.1f}% (no usado)")
            elif info['metodo'] == 'regla_especial':
                racional_parts.append("Mapeo por regla especial")
            
            # Agregar información sobre tipos de datos
            if info['columna_antigua'] in df_antiguo.columns:
                tipo_ant = str(df_antiguo[info['columna_antigua']].dtype)
                if col_final in df_training.columns:
                    tipo_final = str(df_training[col_final].dtype)
                    if tipo_ant == tipo_final:
                        racional_parts.append(f"Tipos coinciden: {tipo_ant}")
                    else:
                        racional_parts.append(f"Tipos diferentes: Antiguo={tipo_ant}, Final={tipo_final}")
            
            # Solo asignar racional si hay información
            if racional_parts:
                fila['Racional'] = " | ".join(racional_parts)
            else:
                fila['Racional'] = ''  # Dejar en blanco si no hay información
        else:
            fila['Racional'] = ''  # Dejar en blanco si no hay mapeo
        datos_reporte.append(fila)
    
    df_reporte = pd.DataFrame(datos_reporte)
    
    # Guardar en Excel en la carpeta de ejecución
    archivo_reporte = carpeta_ejecucion / "REPORTE_MAPEO.xlsx"
    with pd.ExcelWriter(archivo_reporte, engine='openpyxl') as writer:
        df_reporte.to_excel(writer, sheet_name='Mapeo Detallado', index=False)
        
        # Ajustar ancho de columnas
        worksheet = writer.sheets['Mapeo Detallado']
        worksheet.column_dimensions['A'].width = 30  # Columna Final
        worksheet.column_dimensions['B'].width = 30  # Columna Antigua
        worksheet.column_dimensions['C'].width = 15  # Confianza
        worksheet.column_dimensions['D'].width = 15  # Método
        worksheet.column_dimensions['E'].width = 15  # Score Nombre
        worksheet.column_dimensions['F'].width = 15  # Score Contenido
        worksheet.column_dimensions['G'].width = 80  # Racional
    
    print(f"  ✓ Reporte guardado en: {archivo_reporte.name}")
    return archivo_reporte

def migrar_datos(df_antiguo, mapeo_final, df_training):
    """
    Migra los datos de Antiguo a Final usando el mapeo estimado
    """
    print("\n" + "="*80)
    print("FASE 2: MIGRACIÓN DE DATOS")
    print("="*80)
    
    # Crear DataFrame con estructura de Final
    df_resultado = pd.DataFrame()
    
    print("\n[1] Creando estructura de columnas Final...")
    for col_final, info in mapeo_final.items():
        if info['columna_antigua']:
            col_antigua = info['columna_antigua']
            
            # Copiar datos de la columna antigua
            if col_antigua in df_antiguo.columns:
                df_resultado[col_final] = df_antiguo[col_antigua].copy()
                print(f"  ✓ {col_final} ← {col_antigua}")
            else:
                # Si no existe, crear columna vacía con el tipo correcto
                tipo_esperado = df_training[col_final].dtype if col_final in df_training.columns else 'object'
                df_resultado[col_final] = pd.Series(dtype=tipo_esperado)
                print(f"  ⚠ {col_final} ← (columna no encontrada, se crea vacía)")
        else:
            # Columna sin mapeo, crear vacía con tipo de training
            tipo_esperado = df_training[col_final].dtype if col_final in df_training.columns else 'object'
            df_resultado[col_final] = pd.Series(dtype=tipo_esperado)
            print(f"  ⚠ {col_final} ← (sin mapeo, se crea vacía)")
    
    # Asegurar que el orden de columnas sea el mismo que en Final/training
    if len(df_training.columns) > 0:
        columnas_orden = df_training.columns.tolist()
        # Agregar columnas que estén en resultado pero no en training
        for col in df_resultado.columns:
            if col not in columnas_orden:
                columnas_orden.append(col)
        df_resultado = df_resultado.reindex(columns=columnas_orden)
    
    print(f"\n[2] Migración completada: {len(df_resultado)} filas, {len(df_resultado.columns)} columnas")
    
    return df_resultado

def main(auto_confirm=False):
    """Función principal"""
    # Buscar CORE.xlsx en el directorio de trabajo actual (para uso desde API)
    # o en el directorio del script (para uso directo)
    archivo_excel = Path.cwd() / "CORE.xlsx"
    if not archivo_excel.exists():
        archivo_excel = Path(__file__).parent / "CORE.xlsx"
    
    if not archivo_excel.exists():
        print(f"❌ Error: No se encontró el archivo {archivo_excel}")
        sys.exit(1)
    
    # Crear carpeta con timestamp para esta ejecución
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta_ejecucion = Path(__file__).parent / f"ejecucion_{timestamp}"
    carpeta_ejecucion.mkdir(exist_ok=True)
    
    print("="*80)
    print("MIGRADOR DE COLUMNAS: Antiguo → Final")
    print("="*80)
    print(f"\n📂 Abriendo archivo: {archivo_excel.name}")
    print(f"📁 Carpeta de ejecución: {carpeta_ejecucion.name}")
    
    # Copiar archivo de entrada a la carpeta de ejecución
    archivo_entrada_copia = carpeta_ejecucion / "CORE.xlsx"
    shutil.copy2(archivo_excel, archivo_entrada_copia)
    print(f"  ✓ Archivo de entrada copiado a: {archivo_entrada_copia.name}")
    
    try:
        # Leer las hojas
        print("\n[0] Leyendo hojas del Excel...")
        df_antiguo = pd.read_excel(archivo_excel, sheet_name='ANTIGUO')
        df_training = pd.read_excel(archivo_excel, sheet_name='TRAINING')
        df_final = pd.read_excel(archivo_excel, sheet_name='FINAL')
        
        print("  ✓ Hoja 'ANTIGUO' leída")
        print("  ✓ Hoja 'TRAINING' leída")
        print("  ✓ Hoja 'FINAL' leída")
        
        # FASE 1: Estimar mapeo
        mapeo_final, info_antiguo, info_training, info_final, mapeo_por_nombre, mapeo_por_contenido = estimar_mapeo_columnas(
            df_antiguo, df_training, df_final
        )
        
        # Mostrar mapeo
        mapeos_validos, mapeos_no_encontrados = mostrar_mapeo(mapeo_final)
        
        # Guardar reporte detallado
        archivo_reporte = guardar_reporte_mapeo(mapeo_final, archivo_excel, df_antiguo, df_training, carpeta_ejecucion)
        
        # Preguntar confirmación antes de migrar
        if not auto_confirm:
            print("\n" + "="*80)
            respuesta = input("\n¿Desea proceder con la migración? (s/n): ").strip().lower()
            
            if respuesta != 's':
                print("\n❌ Migración cancelada por el usuario")
                return
        else:
            print("\n" + "="*80)
            print("\n✓ Procediendo automáticamente con la migración...")
        
        # FASE 2: Migrar datos
        df_resultado = migrar_datos(df_antiguo, mapeo_final, df_training)
        
        # Guardar resultado en la carpeta de ejecución
        print("\n[3] Guardando resultado...")
        archivo_salida = carpeta_ejecucion / "CORE_MIGRADO.xlsx"
        
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Leer el archivo original para mantener otras hojas
            try:
                libro_original = pd.ExcelFile(archivo_excel)
                for sheet in libro_original.sheet_names:
                    if sheet not in ['ANTIGUO', 'TRAINING', 'FINAL']:
                        df_temp = pd.read_excel(archivo_excel, sheet_name=sheet)
                        df_temp.to_excel(writer, sheet_name=sheet, index=False)
            except:
                pass
            
            # Escribir la hoja migrada
            df_resultado.to_excel(writer, sheet_name='FINAL', index=False)
            print(f"  ✓ Resultado guardado en: {archivo_salida.name}")
            print(f"  ✓ Hoja 'FINAL' actualizada con datos migrados")
        
        print("\n" + "="*80)
        print("✅ PROCESO COMPLETADO")
        print("="*80)
        print(f"\n📊 Resumen:")
        print(f"   - Filas migradas: {len(df_resultado)}")
        print(f"   - Columnas: {len(df_resultado.columns)}")
        print(f"   - Mapeos exitosos: {len(mapeos_validos)}")
        print(f"   - Columnas sin mapeo: {len(mapeos_no_encontrados)}")
        print(f"\n📁 Carpeta de ejecución: {carpeta_ejecucion}")
        print(f"\n📄 Archivos en la carpeta:")
        print(f"   - Entrada: CORE.xlsx")
        print(f"   - Salida: CORE_MIGRADO.xlsx")
        print(f"   - Reporte: REPORTE_MAPEO.xlsx")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    import sys
    # Si se pasa --auto como argumento, ejecuta sin confirmación
    auto_confirm = '--auto' in sys.argv
    main(auto_confirm=auto_confirm)

