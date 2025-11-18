Calculador de Balance General (Proyecto de la Facultad)

Proyecto para generar un balance general a partir de archivos Excel (libro diario y libro mayor).

Integrantes:
- Mauro Maiese 28.315.101
- Jeckson Torres 29.797.183
- Santiago Velázquez 28.333.409

Objetivo: aplicación de escritorio para cargar archivos Excel (.xlsx) con libro diario y libro mayor, generar y visualizar un balance general y exportarlo en Excel.

Estado: scaffold inicial con GUI mínima, núcleo de procesamiento y tests.

Setup rápido (Windows / PowerShell):

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m src.main
```

Notas importantes:
- Este repositorio incluye un fichero `LICENSE-INSTITUTIONAL.txt` como marcador: debe reemplazarse por la licencia institucional entregada por la Facultad.
- El código inicial usa `pandas` y `openpyxl` para leer/escribir Excel.

Estructura principal:
- `src/` : código fuente (GUI + processor)
- `tests/` : pruebas unitarias
- `examples/` : (vacío) lugar para subir libros de ejemplo

Web front-end (opcional):

También hay una versión web basada en JavaScript/HTML/CSS en `web/` que procesa los archivos Excel en el navegador y puede guardar resultados en Firebase Firestore.

Cómo ejecutar la versión web localmente:

1. Abra la carpeta `web` con un servidor estático (ej. `Live Server` de VSCode o `http-server`).

Con `npx http-server` en PowerShell desde la raíz del repo:
```powershell
npx http-server ./web -p 8080
# luego abra http://localhost:8080 en el navegador
```

2. Para activar guardado en Firebase:
 - Copie `web/firebase.sample.js` a `web/firebase.config.js` y reemplace los valores por su configuración de Firebase (proyecto, apiKey, etc.).
 - En la consola de Firebase, habilite Firestore en modo de prueba (o configure reglas adecuadas) y ajuste el código si desean autenticación.

Notas:
- La web usa SheetJS (`xlsx`) para leer/escribir Excel en el navegador.
- El guardado en Firebase usa la librería `firebase-compat` para simplificar la inicialización desde el navegador.

Flujo de desarrollo recomendado para los 3 programadores:
- Usar ramas por funcionalidad: `feature/<nombre>`
- Revisar por PRs y asignar revisores (uno de los integrantes)
- Agregar issues y milestones para pruebas y validación con datos reales

Si quieren, puedo crear un repositorio remoto en GitHub/GitLab y subir esto (necesitaré acceso/token o que ustedes lo creen y me den permiso).
