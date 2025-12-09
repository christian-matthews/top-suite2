"""
TOP Suite 2 - Aplicación Unificada
Todo en un solo servidor FastAPI (sin frontend separado)
"""

from fastapi import FastAPI, File, UploadFile, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import shutil
import sys
from datetime import datetime
from typing import Optional
import traceback
import io
import contextlib
import os

# Agregar el directorio scripts al path
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "scripts"))

app = FastAPI(title="TOP Suite 2", version="2.0.0")


# Manejador global de excepciones - siempre devuelve JSON
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={
            "status": "error",
            "message": f"Error interno: {str(exc)}",
            "files": [],
            "log": [f"Error: {str(exc)}", traceback.format_exc()]
        }
    )


# Directorio para archivos temporales y resultados
TMP_DIR = BASE_DIR / "tmp"
TMP_DIR.mkdir(exist_ok=True)

# Montar archivos estáticos
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


def create_job_dir() -> Path:
    """Crea un directorio único para cada job"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    job_dir = TMP_DIR / f"job_{timestamp}"
    job_dir.mkdir(exist_ok=True)
    return job_dir


@contextlib.contextmanager
def capture_output():
    """Captura stdout y stderr"""
    stdout_capture = io.StringIO()
    stderr_capture = io.StringIO()
    old_stdout, old_stderr = sys.stdout, sys.stderr
    try:
        sys.stdout = stdout_capture
        sys.stderr = stderr_capture
        yield stdout_capture, stderr_capture
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr


# ============================================================
# PÁGINA PRINCIPAL - Interfaz HTML
# ============================================================

@app.get("/", response_class=HTMLResponse)
async def home():
    """Página principal con la interfaz de usuario"""
    html_file = BASE_DIR / "static" / "index.html"
    if html_file.exists():
        return HTMLResponse(content=html_file.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>Error: index.html no encontrado</h1>", status_code=500)


# ============================================================
# ENDPOINT: Descargar archivos generados
# ============================================================

@app.get("/download/{job_id}/{filename}")
async def download_file(job_id: str, filename: str):
    """Descarga un archivo generado"""
    file_path = TMP_DIR / job_id / filename
    if file_path.exists():
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return JSONResponse({"error": "Archivo no encontrado"}, status_code=404)


# ============================================================
# API: Procesar PROMPT0 - Migrador de columnas
# ============================================================

@app.post("/api/prompt0")
async def process_prompt0(file: UploadFile = File(...)):
    """Migra datos de formato antiguo a nuevo"""
    job_dir = None
    log = []
    
    try:
        job_dir = create_job_dir()
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando PROMPT0 - Migrador...")
        
        # Guardar archivo como CORE.xlsx
        core_file = job_dir / "CORE.xlsx"
        content = await file.read()
        with open(core_file, "wb") as f:
            f.write(content)
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Archivo recibido: {file.filename} ({len(content)} bytes)")
        
        # Ejecutar script
        from prompt0.migrador_columnas import main as run_migrador
        
        original_cwd = os.getcwd()
        os.chdir(str(job_dir))
        
        try:
            with capture_output() as (stdout, stderr):
                try:
                    run_migrador(auto_confirm=True)
                    success = True
                except SystemExit as e:
                    # El script usa sys.exit() para indicar error
                    success = e.code == 0
                    if not success:
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Script terminó con código: {e.code}")
                except Exception as e:
                    success = False
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error script: {str(e)}")
                    log.append(traceback.format_exc())
            
            # Capturar output
            if stdout.getvalue():
                for line in stdout.getvalue().strip().split('\n'):
                    if line.strip():
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {line}")
            
            # Buscar archivos generados (en job_dir y en subcarpetas ejecucion_*)
            files = []
            # Archivos en job_dir
            for f in job_dir.glob("*.xlsx"):
                if f.name != "CORE.xlsx":
                    files.append({"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"})
            # Archivos en subcarpetas ejecucion_* (el script crea estas carpetas)
            for ejecucion_dir in job_dir.glob("ejecucion_*"):
                for f in ejecucion_dir.glob("*.xlsx"):
                    # Copiar al job_dir para facilitar descarga
                    dest = job_dir / f.name
                    if not dest.exists():
                        shutil.copy2(f, dest)
                        files.append({"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"})
            # También buscar en el directorio del script (por si el script los creó ahí)
            script_dir = BASE_DIR / "scripts" / "prompt0"
            for ejecucion_dir in script_dir.glob("ejecucion_*"):
                for f in ejecucion_dir.glob("*.xlsx"):
                    dest = job_dir / f.name
                    if not dest.exists():
                        shutil.copy2(f, dest)
                        files.append({"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"})
                # Limpiar la carpeta de ejecución del script después de copiar
                try:
                    shutil.rmtree(ejecucion_dir)
                except:
                    pass
            
            return {
                "status": "ok" if success or files else "error",
                "message": "Migración completada" if files else "Error en migración",
                "files": files,
                "log": log,
                "job_id": job_dir.name
            }
        finally:
            os.chdir(original_cwd)
            
    except Exception as e:
        error_msg = str(e)
        error_trace = traceback.format_exc()
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error general: {error_msg}")
        log.append(error_trace)
        return JSONResponse(content={
            "status": "error", 
            "message": error_msg, 
            "files": [], 
            "log": log
        })


# ============================================================
# API: Procesar PROMPT1 - Comparar MP
# ============================================================

@app.post("/api/prompt1")
async def process_prompt1(
    base_file: UploadFile = File(...),
    final_file: UploadFile = File(...)
):
    """Compara BASE vs FINAL"""
    job_dir = create_job_dir()
    log = []
    
    try:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando PROMPT1 - Comparador...")
        
        # Guardar archivos
        with open(job_dir / "BASE.xlsx", "wb") as f:
            shutil.copyfileobj(base_file.file, f)
        with open(job_dir / "FINAL.xlsx", "wb") as f:
            shutil.copyfileobj(final_file.file, f)
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Archivos recibidos")
        
        from prompt1.main import main as run_comparador
        
        original_cwd = os.getcwd()
        os.chdir(str(job_dir))
        
        try:
            with capture_output() as (stdout, stderr):
                try:
                    run_comparador()
                    success = True
                except SystemExit as e:
                    success = e.code == 0 if e.code is not None else False
                    exit_code = e.code if e.code is not None else "unknown"
                    if not success:
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Script terminó con código: {exit_code}")
                        # Capturar stderr si hay errores
                        if stderr.getvalue():
                            for line in stderr.getvalue().strip().split('\n'):
                                if line.strip():
                                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {line}")
                except Exception as e:
                    success = False
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
                    log.append(traceback.format_exc())
            
            # Capturar stdout
            if stdout.getvalue():
                for line in stdout.getvalue().strip().split('\n'):
                    if line.strip():
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {line}")
            
            # Capturar stderr (errores adicionales)
            if stderr.getvalue():
                for line in stderr.getvalue().strip().split('\n'):
                    if line.strip() and "ERROR:" not in line:  # Evitar duplicados
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {line}")
            
            files = []
            for f in job_dir.glob("*.xlsx"):
                if f.name not in ["BASE.xlsx", "FINAL.xlsx"]:
                    files.append({"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"})
            for f in job_dir.glob("*.txt"):
                files.append({"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"})
            
            return {
                "status": "ok" if success else "error",
                "message": "Comparación completada" if success else "Error",
                "files": files,
                "log": log,
                "job_id": job_dir.name
            }
        finally:
            os.chdir(original_cwd)
            
    except Exception as e:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
        return {"status": "error", "message": str(e), "files": [], "log": log}


# ============================================================
# API: Procesar PROMPT2 - Procesar Ventas
# ============================================================

@app.post("/api/prompt2")
async def process_prompt2(
    mp_key_file: UploadFile = File(...),
    ventas_file: UploadFile = File(...)
):
    """Procesa ventas contra MP KEY"""
    job_dir = create_job_dir()
    log = []
    
    try:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando PROMPT2 - Ventas...")
        
        with open(job_dir / "MP KEY.xlsx", "wb") as f:
            shutil.copyfileobj(mp_key_file.file, f)
        with open(job_dir / "Ventas JUL-AGO-SEP-OCT.xlsx", "wb") as f:
            shutil.copyfileobj(ventas_file.file, f)
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Archivos recibidos")
        
        from prompt2.procesar_mp_ventas import main as run_ventas
        
        original_cwd = os.getcwd()
        os.chdir(str(job_dir))
        
        try:
            with capture_output() as (stdout, stderr):
                try:
                    run_ventas()
                    success = True
                except SystemExit as e:
                    success = e.code == 0
                    if not success:
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Script terminó con código: {e.code}")
                except Exception as e:
                    success = False
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
                    log.append(traceback.format_exc())
            
            if stdout.getvalue():
                for line in stdout.getvalue().strip().split('\n'):
                    if line.strip():
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {line}")
            
            files = []
            for proceso_dir in job_dir.glob("Proceso_*"):
                for f in proceso_dir.glob("*.xlsx"):
                    files.append({"name": f.name, "url": f"/download/{job_dir.name}/{proceso_dir.name}/{f.name}"})
            
            return {
                "status": "ok" if success else "error",
                "message": "Procesamiento completado" if success else "Error",
                "files": files,
                "log": log,
                "job_id": job_dir.name
            }
        finally:
            os.chdir(original_cwd)
            
    except Exception as e:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
        return {"status": "error", "message": str(e), "files": [], "log": log}


# ============================================================
# API: Procesar PROMPT3 - Enriquecer Transacciones
# ============================================================

@app.post("/api/prompt3")
async def process_prompt3(
    tx_carga_file: UploadFile = File(...),
    mp_key_file: UploadFile = File(...)
):
    """Enriquece transacciones con SAP_ID"""
    job_dir = create_job_dir()
    log = []
    
    try:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando PROMPT3 - Enriquecimiento...")
        
        with open(job_dir / "TX_Carga.xlsx", "wb") as f:
            shutil.copyfileobj(tx_carga_file.file, f)
        with open(job_dir / "MP KEY.xlsx", "wb") as f:
            shutil.copyfileobj(mp_key_file.file, f)
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Archivos recibidos")
        
        from prompt3.enriquecer_base_tx import main as run_enriquecer
        
        original_cwd = os.getcwd()
        os.chdir(str(job_dir))
        
        try:
            with capture_output() as (stdout, stderr):
                try:
                    run_enriquecer()
                    success = True
                except SystemExit:
                    success = False
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Proceso detenido por inconsistencias")
                except Exception as e:
                    success = False
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
            
            if stdout.getvalue():
                for line in stdout.getvalue().strip().split('\n'):
                    if line.strip():
                        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {line}")
            
            files = [{"name": f.name, "url": f"/download/{job_dir.name}/{f.name}"} 
                     for f in job_dir.glob("*.xlsx") if f.name not in ["TX_Carga.xlsx", "MP KEY.xlsx"]]
            
            return {
                "status": "ok" if success or files else "error",
                "message": "Enriquecimiento completado" if success else "Completado con advertencias",
                "files": files,
                "log": log,
                "job_id": job_dir.name
            }
        finally:
            os.chdir(original_cwd)
            
    except Exception as e:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
        return {"status": "error", "message": str(e), "files": [], "log": log}


# ============================================================
# API: Procesar Maestro Producto
# ============================================================

@app.post("/api/maestro-producto")
async def process_maestro_producto(file: UploadFile = File(...)):
    """Procesa Maestro de Productos a formato SAP"""
    job_dir = create_job_dir()
    log = []
    
    try:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando Maestro Producto...")
        
        input_file = job_dir / file.filename
        with open(input_file, "wb") as f:
            shutil.copyfileobj(file.file, f)
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Archivo recibido: {file.filename}")
        
        from maestro_producto.procesador_excel import generar_tabla_tgt
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = job_dir / f"resultado_{input_file.stem}_{timestamp}.xlsx"
        
        with capture_output() as (stdout, stderr):
            try:
                success = generar_tabla_tgt(str(input_file), str(output_file))
            except Exception as e:
                success = False
                log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
        
        if stdout.getvalue():
            for line in stdout.getvalue().strip().split('\n'):
                if line.strip():
                    log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {line}")
        
        files = []
        if success and output_file.exists():
            files.append({"name": output_file.name, "url": f"/download/{job_dir.name}/{output_file.name}"})
        
        return {
            "status": "ok" if success else "error",
            "message": "Procesamiento completado" if success else "Error",
            "files": files,
            "log": log,
            "job_id": job_dir.name
        }
        
    except Exception as e:
        log.append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {str(e)}")
        return {"status": "error", "message": str(e), "files": [], "log": log}


# ============================================================
# Health Check
# ============================================================

@app.get("/health")
async def health():
    """Verifica el estado del servidor"""
    return {"status": "ok", "message": "TOP Suite 2 funcionando"}


@app.post("/api/test-upload")
async def test_upload(file: UploadFile = File(...)):
    """Endpoint de prueba para verificar uploads"""
    try:
        content = await file.read()
        return JSONResponse(content={
            "status": "ok",
            "filename": file.filename,
            "size": len(content),
            "content_type": file.content_type
        })
    except Exception as e:
        return JSONResponse(content={
            "status": "error",
            "message": str(e),
            "traceback": traceback.format_exc()
        })


if __name__ == "__main__":
    import uvicorn
    # Obtener puerto de variable de entorno (Render) o usar 8000 por defecto
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "127.0.0.1")
    print("=" * 50)
    print("  TOP Suite 2 - Iniciando servidor...")
    print(f"  Abre http://{host}:{port} en tu navegador")
    print("=" * 50)
    uvicorn.run(app, host=host, port=port)

