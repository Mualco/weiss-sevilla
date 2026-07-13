# Weiss Sevilla — Estadísticas de torneos

App de estadísticas para la comunidad de Weiss Schwarz de Sevilla. Convierte los
Excel de estadísticas de temporada en una SQLite limpia y la sirve como una app
Streamlit **pública** (sin login) con una página de administración protegida
por contraseña para importar/corregir datos.

## Estructura

- `ingesta.py` — lee uno o varios `.xlsx` (uno por temporada) y construye
  `weiss.db`. Sólo importa las hojas **fuente**; las hojas derivadas (rankings,
  `Set_Stats`, `Deck Stats`, etc.) se recalculan siempre desde los datos fuente,
  nunca se leen del Excel. `build_database()` acepta tanto rutas locales como
  archivos subidos desde la web (usado por la página Administrar).
- `app.py` — app Streamlit con dos páginas:
  - **Estadísticas** (pública): Rankings, Resumen, Partidas y Resultados
    individuales (ficha por jugador con gráficos, historial de partidas y
    desgloses de winrate, al estilo de la hoja `Resultados_Individuales` del
    Excel).
  - **Administrar** (con contraseña): importar Excel desde la web y corregir
    datos directamente si algo se importó mal.
  La lógica de agregación (`wilson_lower`, `deck_rankings`,
  `set_representation`, `player_leaderboard`, `text_filter`) son funciones puras
  al principio del archivo, sin dependencias de Streamlit.
- `weiss.db` — base de datos generada, versionada en el repo (ver
  [Despliegue](#despliegue)).
- `.streamlit/secrets.toml` — contraseña de administrador (no se sube a git;
  usa `secrets.toml.example` como plantilla).

## Uso local

```powershell
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt

# generar/actualizar weiss.db a partir de los Excel de temporada
.venv\Scripts\python ingesta.py WeissSevilla-26.xlsx --db weiss.db

# copiar la plantilla de contraseña de administrador y editarla
copy secrets.toml.example .streamlit\secrets.toml

.venv\Scripts\streamlit run app.py
```

## Acceso público vs. administración

Hay dos enlaces distintos a la misma app:

- **Enlace público** — `https://tu-app.streamlit.app/` — solo ve la página
  **Estadísticas**; el menú lateral ni siquiera muestra la opción
  "Administrar" (es este el que se comparte con la comunidad).
- **Enlace de administración** — `https://tu-app.streamlit.app/?admin=1` —
  además muestra "Administrar" en el menú lateral. Esto solo **revela** el
  enlace; entrar sigue pidiendo la contraseña de `st.secrets["admin_password"]`
  igual que antes. Una vez cargada la app con `?admin=1` una vez, el enlace se
  mantiene visible el resto de la sesión aunque cambies de pestaña.

La página **Administrar** permite:
  1. **Importar Excel** — subir uno o varios `.xlsx` desde el navegador; se
     reconstruye `weiss.db` al instante y muestra los mismos avisos de
     validación que la CLI.
  2. **Corregir datos** — editar/añadir/borrar filas directamente en cualquier
     tabla (`matches`, `entries`, `attendance`, `player_refs`, `set_refs`,
     `deck_info`) con una tabla editable, para arreglar algo mal importado sin
     tocar el Excel.

> **Persistencia en Streamlit Community Cloud:** el sistema de archivos del
> contenedor es efímero. Los cambios hechos desde "Administrar" (importar o
> corregir) se ven al instante para todo el mundo mientras la app siga viva,
> pero **se pierden en el siguiente reinicio/redeploy** si no se guardan en
> git. Usa el botón "Descargar weiss.db actual" y súbela al repositorio para
> dejarlas fijas. En local (`streamlit run app.py` en tu máquina) los cambios
> sí persisten en el `weiss.db` de disco sin pasos extra.

## Ingesta y multi-temporada

Cada `.xlsx` corresponde a una temporada (p.ej. `WeissSevilla-26.xlsx` →
temporada 2026). La temporada de cada fila se calcula a partir del **año de su
fecha**, no del nombre de archivo, así que basta con pasar todos los Excel
juntos:

```powershell
.venv\Scripts\python ingesta.py WeissSevilla-26.xlsx WeissSevilla-27.xlsx --db weiss.db
```

`ingesta.py` **no es incremental**: cada ejecución reconstruye `weiss.db` entera
desde cero a partir de los `.xlsx` que le pases (por CLI o subidos desde
Administrar). Para añadir una temporada nueva, pasa el archivo nuevo *más*
todos los anteriores.

### Validaciones incluidas

- Comprueba que las hojas y columnas esperadas existan; si el Excel cambia de
  formato, falla con un mensaje claro indicando qué falta (no falla en
  silencio ni importa datos a medias).
- Descarta y avisa de: filas sin jugador o sin resultado (rondas
  incompletas/byes), fechas no interpretables, resultados que no son `W`/`L`,
  wins/losses no numéricos o negativos, y filas duplicadas.
- Avisa (sin bloquear) de jugadores sin foto en `Player Refs` o códigos de set
  sin referencia en `Set Refs`.

### Hojas fuente vs. derivadas

El Excel tiene hojas *fuente* (`RAW RESULTS`, `Raw Data`, `Asistencia`,
`Player Refs`, `Set Refs`, `Deck Info`) y hojas *derivadas* que son cálculos
manuales (`Ranking Temporada`, `Ranking mensual`, `Torneos Recientes`,
`TorneosMayo`, `Set_Stats`, `Deck Stats`, `Hoja 3`, `Resultados_Individuales`).
Sólo se importan las primeras; la app recalcula el resto. Está verificado que
el recálculo reproduce el Excel al decimal (Wilson score de "OSK Door Pants" =
0.5014068556 con 31 partidas, idéntico en ambos).

## Autenticación

Solo la página **Administrar** requiere contraseña, guardada en
`st.secrets["admin_password"]`:

- **Local**: `.streamlit/secrets.toml` (gitignored) con
  `admin_password = "..."`.
- **Streamlit Community Cloud**: pestaña *Settings → Secrets* de la app, mismo
  formato TOML.

> El archivo `secrets.toml` debe guardarse **sin BOM UTF-8** — un editor que
> añada BOM rompe el parseo TOML y Streamlit actúa como si no hubiera
> contraseña configurada.

## Despliegue (Streamlit Community Cloud)

Se versiona `weiss.db` directamente en el repo (pesa ~180 KB): Streamlit Cloud
no ejecuta pasos de build personalizados, sólo `pip install -r requirements.txt`
y `streamlit run app.py`, así que la app debe poder arrancar leyendo la DB tal
cual está en el repo.

Flujo para actualizar datos (dos opciones equivalentes):

- **Desde la web**: entra en Administrar → Importar Excel → sube los `.xlsx` →
  descarga la `weiss.db` resultante → `git add weiss.db` → commit/push.
- **En local**: `python ingesta.py <archivos.xlsx...> --db weiss.db` → revisar
  avisos → `git add weiss.db` (+ el `.xlsx` si es nuevo) → commit/push.

Streamlit Community Cloud redespliega automáticamente al detectar el push.

Pasos de alta en Streamlit Community Cloud:

1. Subir este repo a GitHub.
2. En share.streamlit.io, "New app" → seleccionar el repo, rama y `app.py`.
3. Configurar el secreto `admin_password` en Settings → Secrets.

## Pendiente / ideas de mejora

- Cachear en local las imágenes de `Player Refs`/`Set Refs`/`Deck Info` cuando
  la URL externa falla (algunas imágenes de ws-tcg.com bloquean hotlinking
  fuera de su dominio).
- Mostrar imágenes de cartas de deck en la pestaña Rankings.
