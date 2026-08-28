# Dashboard Libro Maestro

## Despliegue recomendado: GitHub Pages

El repositorio incluye una versión estática del dashboard en `index.html`.
Esta versión funciona directamente en el navegador y conserva los filtros,
gráficos y la descarga CSV sin ejecutar Python.

1. En GitHub abra **Settings > Pages**.
2. En **Build and deployment**, seleccione **GitHub Actions** como origen.
3. Haga push a la rama `main` o ejecute manualmente el workflow
   **Deploy static dashboard to GitHub Pages**.
4. Espere a que termine el workflow y abra la URL publicada en el entorno
   `github-pages`.

Cada `push` a `main` vuelve a publicar automáticamente la versión estática.

## ¿Qué se publica?

GitHub Pages publica `index.html`, que carga `data.json` y Plotly desde el
navegador. Por eso no necesita instalar `requirements.txt` ni ejecutar
`app.py`. Los cambios realizados directamente en el archivo Excel deben
convertirse nuevamente a `data.json` para reflejarse en la página.

## Alternativa: Render

Si se requiere ejecutar la versión Python completa de `app.py`, conecte el
repositorio en [Render](https://render.com). Render leerá `render.yaml` y
ejecutará `gunicorn app:server`.

## Ejecución local

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Abra `http://127.0.0.1:8050`.