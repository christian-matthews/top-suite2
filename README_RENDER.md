# Desplegar TOP Suite 2 en Render

## Pasos para subir a Render

### 1. Sube tu código a GitHub
```bash
cd top_suite2
git init
git add .
git commit -m "Initial commit"
git remote add origin TU_REPO_GITHUB
git push -u origin main
```

### 2. En Render.com

1. **Crea una cuenta** en https://render.com (gratis)

2. **Nuevo Web Service:**
   - Click en "New" → "Web Service"
   - Conecta tu repositorio de GitHub
   - Selecciona el repositorio con `top_suite2`

3. **Configuración:**
   - **Name:** `top-suite2` (o el nombre que prefieras)
   - **Environment:** `Python 3`
   - **Python Version:** `3.10.13` (IMPORTANTE: Configurar manualmente en Settings)
   - **Build Command:** `bash build.sh` (o `pip install --upgrade pip setuptools wheel && pip install -r requirements.txt`)
   - **Start Command:** `uvicorn app:app --host 0.0.0.0 --port $PORT`
   - **Plan:** Free (gratis)

4. **Variables de entorno (opcional):**
   - No necesitas configurar ninguna por ahora

5. **Click "Create Web Service"**

### 3. Acceder desde Internet

Una vez desplegado, Render te dará una URL como:
```
https://top-suite2.onrender.com
```

**¡Esa es tu URL pública!** Compártela con quien necesites.

---

## Notas importantes

- **Primera carga:** Puede tardar 30-60 segundos (Render "duerme" servicios gratuitos después de 15 min de inactividad)
- **Límites del plan gratuito:**
  - 750 horas/mes gratis
  - Se "duerme" después de 15 min sin uso
  - Se "despierta" automáticamente al recibir una petición

---

## Solución de problemas

### Si el servicio no inicia:
- Verifica los logs en Render Dashboard
- Asegúrate de que `requirements.txt` tenga todas las dependencias
- Verifica que el `Procfile` esté correcto

### Si hay errores de importación:
- Verifica que todos los archivos `__init__.py` estén presentes
- Asegúrate de que la estructura de carpetas sea correcta

