"""App Streamlit de estadisticas Weiss Sevilla.

Lee weiss.db (generada por ingesta.py) y muestra dos paginas:

- **Estadisticas** (publica, sin login): Rankings, Resumen, Partidas, Jugadores.
- **Administrar** (con contrasena): importar Excel desde la web y corregir
  datos directamente si algo se importo mal.

La logica de agregacion vive en funciones puras al inicio del archivo,
separadas de la UI, para que sean testeables sin Streamlit.
"""

from __future__ import annotations

import html as html_lib
import math
import sqlite3
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from ingesta import IngestaError, build_database

DB_PATH = Path(__file__).parent / "weiss.db"

# Paleta derivada de los colores reales del Excel (cabeceras cian, titulos
# naranja), reajustada de luminosidad/croma para pasar el validador de
# contraste y CVD de la skill dataviz (ver scripts/validate_palette.js).
COLOR_ORANGE = "#C97700"
COLOR_BLUE = "#3D85C6"
COLOR_TEAL = "#00899F"
COLOR_RED = "#C9302C"

# paleta suave (mismas familias de color de la skill dataviz, pero mezcladas
# hacia blanco) para los graficos de barras/tarta por set: menos saturada,
# mas facil de leer con muchas categorias a la vez.
SET_PALETTE = [
    "#D9A066",  # naranja suave
    "#6DBFAE",  # teal suave
    "#7FA8D9",  # azul suave
    "#D98C82",  # rojo suave
    "#A79BD9",  # violeta suave
    "#D992B0",  # magenta suave
    "#8FBF7F",  # verde suave
    "#D9B25C",  # dorado suave
]

# bandas alternas por torneo en la tabla de pairings individuales
TORNEO_BANDS = ["#FDF3E7", "#E7F1F7"]


def build_set_color_map(*set_series: pd.Series) -> dict[str, str]:
    """Asigna a cada codigo de set un color fijo de SET_PALETTE, en orden
    alfabetico, para que el mismo set tenga siempre el mismo color en
    cualquier grafico de la app (no solo dentro de un grafico concreto)."""
    codes: set[str] = set()
    for s in set_series:
        codes |= set(s.dropna().unique().tolist())
    return {code: SET_PALETTE[i % len(SET_PALETTE)] for i, code in enumerate(sorted(codes))}

XL_CSS = """
<style>
html, body, [class*="css"] { font-family: Arial, "Helvetica Neue", sans-serif; }

.xl-title {
  color: #C97700;
  font-weight: 700;
  border-bottom: 2px solid #E3B673;
  padding-bottom: 6px;
  margin: 4px 0 18px 0;
  letter-spacing: 0.2px;
}

[data-testid="stMetricLabel"] { color: #007685; font-weight: 600; }
[data-testid="stMetricValue"] { color: #2B2A28; }

.xl-table-wrap {
  overflow: auto;
  max-height: 480px;
  border: 1px solid #DCDAD3;
  border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  margin-bottom: 1.2rem;
}
table.xl-table { border-collapse: collapse; width: 100%; font-size: 0.9rem; }
table.xl-table thead th {
  position: sticky;
  top: 0;
  background: #00899F;
  color: #FFFFFF;
  border: 1px solid #05808F;
  padding: 7px 12px;
  text-align: left;
  font-weight: 700;
  white-space: nowrap;
}
table.xl-table tbody td {
  border: 1px solid #D2CFC5;
  padding: 6px 12px;
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
  color: #14130F;
  font-weight: 500;
}
table.xl-table tbody tr:nth-child(even) { background: #F2F9FA; }
table.xl-table tbody tr:hover { background: #FDF0DC; }
table.xl-table.xl-table--bold tbody td { font-weight: 700; }

.xl-avatar {
  height: 26px;
  width: 26px;
  border-radius: 50%;
  object-fit: cover;
  vertical-align: middle;
  margin-left: 8px;
  border: 1px solid #DCDAD3;
  background: #EEEDE8;
}

.xl-card-thumb {
  height: 118px;
  width: 84px;
  object-fit: cover;
  border-radius: 4px;
  border: 1px solid #C7C4B9;
  margin-right: 5px;
  background: #EEEDE8;
  vertical-align: middle;
}
.xl-set-thumb {
  height: 118px;
  width: 84px;
  object-fit: cover;
  border-radius: 4px;
  border: 1px solid #C7C4B9;
  background: #EEEDE8;
  vertical-align: middle;
}
.xl-banner {
  background: #C97700;
  color: #FFFFFF;
  font-weight: 700;
  text-align: center;
  padding: 9px 10px;
  border-radius: 8px 8px 0 0;
  font-size: 1rem;
}
.xl-banner-sub {
  background: #FBE7CB;
  color: #6B4A00;
  text-align: center;
  font-size: 0.78rem;
  padding: 4px 8px;
}
.xl-panel .xl-table-wrap {
  border-radius: 0 0 8px 8px;
  border-top: none;
  margin-top: 0;
}

.xl-admin-warning {
  background: #FDF0DC;
  border: 1px solid #E3B673;
  border-radius: 8px;
  padding: 10px 14px;
  margin-bottom: 1rem;
  color: #6B4A00;
  font-size: 0.92rem;
}
</style>
"""

# ------------------------------------------------------------------
# Funciones puras (agregacion, sin dependencias de Streamlit)
# ------------------------------------------------------------------


def wilson_lower(wins: int, losses: int, z: float = 1.96) -> float:
    """Cota inferior del intervalo de Wilson al 95% para una tasa de victorias."""
    n = wins + losses
    if n == 0:
        return 0.0
    p = wins / n
    denom = 1 + z**2 / n
    center = p + z**2 / (2 * n)
    margin = z * math.sqrt((p * (1 - p) + z**2 / (4 * n)) / n)
    return (center - margin) / denom


def deck_rankings(entries: pd.DataFrame, min_partidas: int = 10) -> pd.DataFrame:
    """Ranking de decks (set+arquetipo) por cota de Wilson, con umbral minimo de partidas."""
    cols = ["set", "arquetipo", "wins", "losses", "partidas", "winrate", "wilson"]
    if entries.empty:
        return pd.DataFrame(columns=cols)
    grp = entries.groupby(["set", "arquetipo"], as_index=False).agg(wins=("wins", "sum"), losses=("losses", "sum"))
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp = grp[grp["partidas"] >= min_partidas].copy()
    if grp.empty:
        return pd.DataFrame(columns=cols)
    grp["winrate"] = grp["wins"] / grp["partidas"]
    grp["wilson"] = grp.apply(lambda r: wilson_lower(int(r["wins"]), int(r["losses"])), axis=1)
    return grp.sort_values("wilson", ascending=False).reset_index(drop=True)[cols]


def set_representation(entries: pd.DataFrame) -> pd.DataFrame:
    """Partidas (wins+losses, no filas de inscripcion) y jugadores distintos por set."""
    cols = ["set", "partidas", "jugadores"]
    if entries.empty:
        return pd.DataFrame(columns=cols)
    grp = entries.groupby("set", as_index=False).agg(
        wins=("wins", "sum"), losses=("losses", "sum"), jugadores=("jugador", "nunique")
    )
    grp["partidas"] = grp["wins"] + grp["losses"]
    return grp.sort_values("partidas", ascending=False).reset_index(drop=True)[cols]


def player_leaderboard(entries: pd.DataFrame, min_partidas: int = 1) -> pd.DataFrame:
    """Ranking de jugadores por winrate, con umbral minimo de partidas."""
    cols = ["jugador", "wins", "losses", "partidas", "winrate", "sets_distintos"]
    if entries.empty:
        return pd.DataFrame(columns=cols)
    grp = entries.groupby("jugador", as_index=False).agg(
        wins=("wins", "sum"), losses=("losses", "sum"), sets_distintos=("set", "nunique")
    )
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp = grp[grp["partidas"] >= min_partidas].copy()
    if grp.empty:
        return pd.DataFrame(columns=cols)
    grp["winrate"] = grp["wins"] / grp["partidas"]
    return grp.sort_values("winrate", ascending=False).reset_index(drop=True)[cols]


def player_activity_leaderboard(entries: pd.DataFrame, min_partidas: int = 1) -> pd.DataFrame:
    """Ranking de jugadores por partidas totales jugadas (actividad), no por winrate."""
    cols = ["jugador", "partidas", "sets_distintos"]
    if entries.empty:
        return pd.DataFrame(columns=cols)
    grp = entries.groupby("jugador", as_index=False).agg(
        wins=("wins", "sum"), losses=("losses", "sum"), sets_distintos=("set", "nunique")
    )
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp = grp[grp["partidas"] >= min_partidas].copy()
    return grp.sort_values("partidas", ascending=False).reset_index(drop=True)[cols]


def player_summary(entries: pd.DataFrame, jugador: str) -> dict:
    """Totales de temporada de un jugador: jugadas, wins, losses, winrate."""
    pe = entries[entries["jugador"] == jugador]
    wins = int(pe["wins"].sum())
    losses = int(pe["losses"].sum())
    partidas = wins + losses
    return {
        "jugadas": partidas,
        "wins": wins,
        "losses": losses,
        "winrate": wins / partidas if partidas else 0.0,
    }


def player_set_breakdown(entries: pd.DataFrame, jugador: str) -> pd.DataFrame:
    """Wins/losses/winrate de un jugador, agrupado por su propio set."""
    cols = ["set", "wins", "losses", "partidas", "winrate"]
    pe = entries[entries["jugador"] == jugador]
    if pe.empty:
        return pd.DataFrame(columns=cols)
    grp = pe.groupby("set", as_index=False).agg(wins=("wins", "sum"), losses=("losses", "sum"))
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp["winrate"] = grp["wins"] / grp["partidas"]
    return grp.sort_values("partidas", ascending=False).reset_index(drop=True)[cols]


def player_deck_breakdown(entries: pd.DataFrame, jugador: str) -> pd.DataFrame:
    """Wins/losses/winrate de un jugador, agrupado por su propio set+arquetipo."""
    cols = ["set", "arquetipo", "wins", "losses", "partidas", "winrate"]
    pe = entries[entries["jugador"] == jugador]
    if pe.empty:
        return pd.DataFrame(columns=cols)
    grp = pe.groupby(["set", "arquetipo"], as_index=False).agg(wins=("wins", "sum"), losses=("losses", "sum"))
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp["winrate"] = grp["wins"] / grp["partidas"]
    return grp.sort_values("partidas", ascending=False).reset_index(drop=True)[cols]


def player_rival_breakdown(matches: pd.DataFrame, jugador: str) -> pd.DataFrame:
    """Wins/losses/winrate de un jugador, agrupado por rival (partidas individuales)."""
    cols = ["rival", "wins", "losses", "winrate"]
    pm = matches[matches["jugador"] == jugador]
    if pm.empty:
        return pd.DataFrame(columns=cols)
    grp = pm.groupby("rival", as_index=False).agg(
        wins=("resultado", lambda s: int((s == "W").sum())),
        losses=("resultado", lambda s: int((s == "L").sum())),
    )
    grp["partidas"] = grp["wins"] + grp["losses"]
    grp = grp[grp["partidas"] > 0].copy()
    grp["winrate"] = grp["wins"] / grp["partidas"]
    return grp.sort_values("partidas", ascending=False).drop(columns="partidas").reset_index(drop=True)[cols]


def text_filter(df: pd.DataFrame, query: str, columns: list[str]) -> pd.DataFrame:
    """Filtro de texto libre: coincide si `query` aparece en cualquiera de `columns`."""
    if not query or not query.strip():
        return df
    q = query.strip().lower()
    mask = pd.Series(False, index=df.index)
    for col in columns:
        if col in df.columns:
            mask |= df[col].astype(str).str.lower().str.contains(q, na=False, regex=False)
    return df[mask]


# ------------------------------------------------------------------
# Autenticacion de administrador (el resto de la app es publica)
# ------------------------------------------------------------------


def check_admin_password() -> bool:
    if st.session_state.get("admin_authenticated"):
        return True

    if "admin_password" not in st.secrets:
        st.error(
            "No hay contrasena de administrador configurada. Crea "
            "`.streamlit/secrets.toml` con `admin_password = \"...\"` (ver "
            "secrets.toml.example) o configura el secreto en Streamlit Community Cloud."
        )
        return False

    def on_submit() -> None:
        if st.session_state.get("admin_password_input") == st.secrets["admin_password"]:
            st.session_state["admin_authenticated"] = True
        else:
            st.session_state["admin_authenticated"] = False

    st.subheader("Acceso de administrador")
    st.text_input("Contrasena", type="password", key="admin_password_input", on_change=on_submit)
    if st.session_state.get("admin_authenticated") is False:
        st.error("Contrasena incorrecta.")
    return False


# ------------------------------------------------------------------
# Carga de datos
# ------------------------------------------------------------------


@st.cache_data
def load_data() -> tuple[pd.DataFrame, ...]:
    if not DB_PATH.exists():
        st.error(f"No se encuentra {DB_PATH.name}. Ve a la pagina Administrar para importar un Excel.")
        st.stop()
    conn = sqlite3.connect(DB_PATH)
    try:
        matches = pd.read_sql("SELECT * FROM matches", conn, parse_dates=["fecha"])
        entries = pd.read_sql("SELECT * FROM entries", conn, parse_dates=["fecha"])
        attendance = pd.read_sql("SELECT * FROM attendance", conn, parse_dates=["fecha"])
        player_refs = pd.read_sql("SELECT * FROM player_refs", conn)
        set_refs = pd.read_sql("SELECT * FROM set_refs", conn)
        deck_info = pd.read_sql("SELECT * FROM deck_info", conn)
    finally:
        conn.close()
    return matches, entries, attendance, player_refs, set_refs, deck_info


def style_fig(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=40, b=10),
        font=dict(color="#898781"),
    )
    fig.update_xaxes(gridcolor="#e1e0d9", zeroline=False)
    fig.update_yaxes(gridcolor="#e1e0d9", zeroline=False)
    return fig


def render_table(
    df: pd.DataFrame,
    formatters: dict[str, str] | None = None,
    html_columns: list[str] | None = None,
    bold: bool = False,
) -> None:
    """Tabla HTML con cabecera fija en cian y cuadricula, imitando el aspecto
    de una hoja de calculo. `formatters` es un dict columna -> format-string.
    `html_columns` marca columnas cuyo contenido ya es HTML de confianza
    (p.ej. un avatar) y no debe escaparse. `bold` pone el texto en negrita."""
    if df.empty:
        st.caption("Sin datos.")
        return
    view = df.copy()
    html_cols = set(html_columns or [])
    for col in view.columns:
        if pd.api.types.is_datetime64_any_dtype(view[col]):
            view[col] = view[col].dt.strftime("%Y-%m-%d")
    if formatters:
        for col, fmt in formatters.items():
            if col in view.columns:
                view[col] = view[col].map(lambda v: fmt.format(v) if pd.notna(v) else "")
    for col in view.columns:
        if col not in html_cols:
            view[col] = view[col].map(lambda v: html_lib.escape(str(v)) if pd.notna(v) else "")
    classes = "xl-table xl-table--bold" if bold else "xl-table"
    table_html = view.to_html(index=False, classes=classes, border=0, na_rep="", escape=False)
    st.markdown(f'<div class="xl-table-wrap">{table_html}</div>', unsafe_allow_html=True)


def render_panel_table(
    title: str,
    df: pd.DataFrame,
    subtitle: str | None = None,
    formatters: dict[str, str] | None = None,
    html_columns: list[str] | None = None,
    headers: dict[str, str] | None = None,
    bold: bool = True,
) -> None:
    """Panel con banner naranja de titulo + tabla de cabecera cian, como las
    tablas 'Deck con mayor Win%', 'Mayor asistencia' y 'Representacion de
    sets' de la hoja Ranking Temporada del Excel. Todo en un unico bloque
    HTML para que el banner y la tabla queden visualmente unidos."""
    sub_html = f'<div class="xl-banner-sub">{html_lib.escape(subtitle)}</div>' if subtitle else ""
    banner = f'<div class="xl-banner">{html_lib.escape(title)}</div>{sub_html}'
    if df.empty:
        st.markdown(f'<div class="xl-panel">{banner}</div>', unsafe_allow_html=True)
        st.caption("Sin datos.")
        return
    view = df.copy()
    html_cols = set(html_columns or [])
    headers = headers or {}
    for col in view.columns:
        if pd.api.types.is_datetime64_any_dtype(view[col]):
            view[col] = view[col].dt.strftime("%Y-%m-%d")
    if formatters:
        for col, fmt in formatters.items():
            if col in view.columns:
                view[col] = view[col].map(lambda v: fmt.format(v) if pd.notna(v) else "")
    for col in view.columns:
        if col not in html_cols:
            view[col] = view[col].map(lambda v: html_lib.escape(str(v)) if pd.notna(v) else "")
    view = view.rename(columns={c: headers.get(c, c) for c in view.columns})
    classes = "xl-table xl-table--bold" if bold else "xl-table"
    table_html = view.to_html(index=False, classes=classes, border=0, na_rep="", escape=False)
    full_html = f'<div class="xl-panel">{banner}<div class="xl-table-wrap" style="max-height:520px">{table_html}</div></div>'
    st.markdown(full_html, unsafe_allow_html=True)


def player_name_with_avatar(jugador: str, player_refs: pd.DataFrame) -> str:
    """Nombre de jugador + su imagen (si existe en player_refs) a continuacion, como HTML."""
    name_html = html_lib.escape(str(jugador))
    ref = player_refs[player_refs["jugador"] == jugador]
    if not ref.empty and pd.notna(ref.iloc[0]["imagen"]):
        url = html_lib.escape(str(ref.iloc[0]["imagen"]), quote=True)
        return f'{name_html}<img src="{url}" class="xl-avatar" loading="lazy">'
    return name_html


def winrate_bg(pct: float) -> str:
    """Color de fondo suave: rojo -> gris neutro -> verde, con el 50% en el medio."""
    pct = max(0.0, min(1.0, pct))
    red, neutral, green = (0xF7, 0xD9, 0xD6), (0xF1, 0xF0, 0xEA), (0xD9, 0xEE, 0xD5)
    if pct <= 0.5:
        t = pct / 0.5
        c = [round(red[i] + (neutral[i] - red[i]) * t) for i in range(3)]
    else:
        t = (pct - 0.5) / 0.5
        c = [round(neutral[i] + (green[i] - neutral[i]) * t) for i in range(3)]
    return "#{:02X}{:02X}{:02X}".format(*c)


def img_tag(url, css_class: str = "xl-card-thumb") -> str:
    """Miniatura de imagen (carta/set/jugador) o cadena vacia si no hay URL."""
    if pd.isna(url) or not str(url).strip():
        return ""
    safe_url = html_lib.escape(str(url), quote=True)
    return f'<img src="{safe_url}" class="{css_class}" loading="lazy">'


def render_stat_table(
    df: pd.DataFrame,
    winrate_col: str = "winrate",
    formatters: dict[str, str] | None = None,
    headers: dict[str, str] | None = None,
) -> None:
    """Tabla HTML pequena con la columna de winrate coloreada en rojo/verde
    segun su valor (heat scale), usada en la ficha de Resultados individuales."""
    if df.empty:
        st.caption("Sin datos.")
        return
    headers = headers or {}
    head_html = "".join(f"<th>{html_lib.escape(headers.get(c, c))}</th>" for c in df.columns)
    rows_html = []
    # itertuples (no iterrows): iterrows fuerza un dtype comun para toda la
    # fila, y una fila con columnas int + float (p.ej. wins y winrate) se
    # vuelve toda float ("22" -> "22.0"); itertuples respeta el dtype de cada columna.
    for row in df.itertuples(index=False):
        tds = []
        for col in df.columns:
            val = getattr(row, col)
            if formatters and col in formatters and pd.notna(val):
                val_str = formatters[col].format(val)
            else:
                val_str = "" if pd.isna(val) else str(val)
            val_str = html_lib.escape(val_str)
            if col == winrate_col and pd.notna(val):
                tds.append(f'<td style="background:{winrate_bg(float(val))}">{val_str}</td>')
            else:
                tds.append(f"<td>{val_str}</td>")
        rows_html.append(f"<tr>{''.join(tds)}</tr>")
    table_html = f'<table class="xl-table"><thead><tr>{head_html}</tr></thead><tbody>{"".join(rows_html)}</tbody></table>'
    st.markdown(f'<div class="xl-table-wrap">{table_html}</div>', unsafe_allow_html=True)


def render_pairings_table(matches: pd.DataFrame) -> None:
    """Historial de partidas de un jugador: bandas alternas por torneo y
    resultado W/L coloreado, como en la hoja 'Resultados_Individuales' del Excel."""
    if matches.empty:
        st.caption("Sin datos.")
        return
    view = matches.sort_values(["fecha", "ronda"], ascending=[False, False])
    torneo_order = list(dict.fromkeys(view["id_torneo"].tolist()))
    band = {t: TORNEO_BANDS[i % 2] for i, t in enumerate(torneo_order)}

    headers = ["ID Torneo", "Fecha", "Ronda", "Set", "Arquetipo", "Rival", "Set rival", "Arquetipo rival", "Resultado"]
    head_html = "".join(f"<th>{h}</th>" for h in headers)

    rows_html = []
    for r in view.itertuples(index=False):
        bg = band.get(r.id_torneo, "#FFFFFF")
        fecha_str = r.fecha.strftime("%Y-%m-%d") if pd.notna(r.fecha) else ""
        vals = [r.id_torneo, fecha_str, r.ronda, r.set1, r.arquetipo1, r.rival, r.set2, r.arquetipo2]
        tds = "".join(
            f'<td style="background:{bg}">{html_lib.escape(str(v) if pd.notna(v) else "")}</td>' for v in vals
        )
        if r.resultado == "W":
            res_td = '<td style="background:#DCF3DC;color:#0B5E0B;font-weight:600;">W</td>'
        else:
            res_td = '<td style="background:#FBDCDC;color:#8A1F1F;font-weight:600;">L</td>'
        rows_html.append(f"<tr>{tds}{res_td}</tr>")

    table_html = f'<table class="xl-table"><thead><tr>{head_html}</tr></thead><tbody>{"".join(rows_html)}</tbody></table>'
    st.markdown(f'<div class="xl-table-wrap">{table_html}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------
# Pagina publica: Estadisticas
# ------------------------------------------------------------------


DECK_RANKING_MIN_PARTIDAS = 10
PLAYER_RANKING_MIN_PARTIDAS = 5


def render_rankings(
    entries: pd.DataFrame, player_refs: pd.DataFrame, set_refs: pd.DataFrame, deck_info: pd.DataFrame
) -> None:
    st.markdown('<div class="xl-title"><h2>Rankings</h2></div>', unsafe_allow_html=True)

    # el panel de decks necesita mucho ancho (3 imagenes de carta + 5 columnas
    # de texto), asi que ocupa toda la fila para no necesitar scroll horizontal
    deck_top = deck_rankings(entries, min_partidas=DECK_RANKING_MIN_PARTIDAS)
    deck_top = deck_top.merge(deck_info, on=["set", "arquetipo"], how="left")
    for i in (1, 2, 3):
        deck_top[f"Carta {i}"] = deck_top[f"img{i}"].map(img_tag)
    deck_view = deck_top[["set", "arquetipo", "partidas", "wilson", "winrate", "Carta 1", "Carta 2", "Carta 3"]]
    render_panel_table(
        "Deck con mayor Win% de la temporada",
        deck_view,
        subtitle=f"El minimo de partidas para que un deck entre en la lista es {DECK_RANKING_MIN_PARTIDAS}.",
        formatters={"wilson": "{:.2f}", "winrate": "{:.0%}"},
        html_columns=["Carta 1", "Carta 2", "Carta 3"],
        headers={
            "set": "Set", "arquetipo": "Deck", "partidas": "Partidas Jugadas",
            "wilson": "Wilson Score", "winrate": "Win %",
        },
    )

    # estos dos solo llevan una columna de imagen cada uno: caben bien a media anchura
    col_activity, col_sets = st.columns(2)

    with col_activity:
        activity = player_activity_leaderboard(entries).head(10)
        activity = activity.merge(player_refs, on="jugador", how="left")
        activity["Signature"] = activity["imagen"].map(img_tag)
        activity_view = activity[["jugador", "partidas", "sets_distintos", "Signature"]]
        render_panel_table(
            "Mayor asistencia mensual",
            activity_view,
            html_columns=["Signature"],
            headers={"jugador": "Jugador", "partidas": "Partidas", "sets_distintos": "Sets diferentes"},
        )

    with col_sets:
        rep = set_representation(entries).head(10)
        rep = rep.merge(set_refs.rename(columns={"codigo": "set"}), on="set", how="left")
        rep["Imagen"] = rep["imagen"].map(lambda u: img_tag(u, "xl-set-thumb"))
        rep_view = rep[["set", "partidas", "jugadores", "Imagen"]]
        render_panel_table(
            "Representacion de sets",
            rep_view,
            html_columns=["Imagen"],
            headers={"set": "Set", "partidas": "Partidas", "jugadores": "Jugadores"},
        )

    st.divider()
    st.markdown("#### Explorar ranking completo")
    tab_decks, tab_jugadores = st.tabs(["Decks", "Jugadores"])

    with tab_decks:
        ranking = deck_rankings(entries, min_partidas=DECK_RANKING_MIN_PARTIDAS)
        render_table(ranking, formatters={"winrate": "{:.1%}", "wilson": "{:.4f}"}, bold=True)

    with tab_jugadores:
        leaderboard = player_leaderboard(entries, min_partidas=PLAYER_RANKING_MIN_PARTIDAS)
        if not leaderboard.empty:
            leaderboard = leaderboard.copy()
            leaderboard["jugador"] = leaderboard["jugador"].map(lambda j: player_name_with_avatar(j, player_refs))
        render_table(leaderboard, formatters={"winrate": "{:.1%}"}, html_columns=["jugador"], bold=True)


def render_resumen(matches: pd.DataFrame, entries: pd.DataFrame, attendance: pd.DataFrame, set_color_map: dict[str, str]) -> None:
    st.markdown('<div class="xl-title"><h2>Resumen de la temporada</h2></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Partidas registradas", len(matches))
    c2.metric("Jugadores distintos", entries["jugador"].nunique() if not entries.empty else 0)
    c3.metric("Torneos", attendance["id_torneo"].nunique() if not attendance.empty else 0)
    top_deck = deck_rankings(entries, min_partidas=10)
    c4.metric("Mejor deck (Wilson)", f"{top_deck.iloc[0]['arquetipo']}" if not top_deck.empty else "—")

    col1, col2 = st.columns(2)
    with col1:
        if not attendance.empty:
            att = attendance.sort_values("fecha")
            fig = px.line(att, x="fecha", y="jugadores", title="Asistencia por torneo", markers=True, text="jugadores")
            fig.update_traces(line_color=COLOR_BLUE, marker=dict(size=8, color=COLOR_BLUE), textposition="top center")
            fig.update_xaxes(tickformat="%d/%m/%Y", showgrid=False)
            fig.update_yaxes(rangemode="tozero")
            st.plotly_chart(style_fig(fig), use_container_width=True)
    with col2:
        rep = set_representation(entries).head(10)
        if not rep.empty:
            fig = px.bar(
                rep, x="set", y="partidas", title="Representacion de sets (top 10)",
                color="set", color_discrete_map=set_color_map,
            )
            fig.update_layout(showlegend=False)
            st.plotly_chart(style_fig(fig), use_container_width=True)

    top10 = deck_rankings(entries, min_partidas=10).head(10)
    if not top10.empty:
        top10["deck"] = top10["set"] + " · " + top10["arquetipo"]
        top10 = top10.sort_values("wilson")
        fig = px.bar(
            top10,
            x="wilson",
            y="deck",
            orientation="h",
            title="Top 10 decks por Wilson score (min. 10 partidas)",
            color="set",
            color_discrete_map=set_color_map,
        )
        fig.update_layout(showlegend=False)
        st.plotly_chart(style_fig(fig), use_container_width=True)
    else:
        st.info("No hay decks con al menos 10 partidas todavia.")


def render_partidas(matches: pd.DataFrame) -> None:
    st.markdown('<div class="xl-title"><h2>Partidas</h2></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    jugadores = sorted(matches["jugador"].dropna().unique())
    jugador_sel = col1.multiselect("Jugador", jugadores)

    # el desplegable de Set solo ofrece los sets que han jugado los jugadores elegidos
    sets_pool = matches[matches["jugador"].isin(jugador_sel)] if jugador_sel else matches
    sets = sorted(sets_pool["set1"].dropna().unique())
    set_sel = col2.multiselect("Set", sets)

    resultado_sel = col3.multiselect("Resultado", ["W", "L"])
    query = st.text_input("Busqueda libre (jugador, rival, set, arquetipo)")

    df = matches.copy()
    if jugador_sel:
        df = df[df["jugador"].isin(jugador_sel)]
    if set_sel:
        df = df[df["set1"].isin(set_sel)]
    if resultado_sel:
        df = df[df["resultado"].isin(resultado_sel)]
    df = text_filter(df, query, ["jugador", "rival", "set1", "set2", "arquetipo1", "arquetipo2"])

    render_table(df.sort_values("fecha", ascending=False))
    st.download_button(
        "Descargar CSV",
        df.to_csv(index=False).encode("utf-8"),
        file_name="partidas_filtradas.csv",
        mime="text/csv",
    )


def render_resultados_individuales(
    entries: pd.DataFrame, matches: pd.DataFrame, player_refs: pd.DataFrame, set_color_map: dict[str, str]
) -> None:
    st.markdown('<div class="xl-title"><h2>Resultados individuales</h2></div>', unsafe_allow_html=True)
    jugadores = sorted(entries["jugador"].dropna().unique())
    if not jugadores:
        st.info("No hay jugadores todavia.")
        return

    col_sel, col_img, col_pie, col_bar = st.columns([1, 1.2, 1.6, 1.6])
    with col_sel:
        st.markdown("**Jugador**")
        jugador = st.selectbox("Jugador", jugadores, label_visibility="collapsed")

    ref = player_refs[player_refs["jugador"] == jugador]
    with col_img:
        if not ref.empty and pd.notna(ref.iloc[0]["imagen"]):
            st.image(ref.iloc[0]["imagen"], width=200)
        else:
            st.caption("(sin imagen)")

    set_stats = player_set_breakdown(entries, jugador)
    with col_pie:
        if not set_stats.empty:
            fig = px.pie(
                set_stats, values="partidas", names="set", title="Sets mas jugados", hole=0.35,
                color="set", color_discrete_map=set_color_map,
            )
            fig.update_traces(textinfo="percent+label")
            st.plotly_chart(style_fig(fig), use_container_width=True)
    with col_bar:
        if not set_stats.empty:
            fig = px.bar(
                set_stats.sort_values("winrate", ascending=False), x="set", y="winrate", title="Set con mayor Win%",
                color="set", color_discrete_map=set_color_map,
            )
            fig.update_layout(showlegend=False)
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(style_fig(fig), use_container_width=True)

    st.markdown("**Pairings individuales**")
    player_matches = matches[matches["jugador"] == jugador]
    render_pairings_table(player_matches)

    summary_df = pd.DataFrame([player_summary(entries, jugador)])
    rival_stats = player_rival_breakdown(matches, jugador)
    deck_stats = player_deck_breakdown(entries, jugador)

    row1c1, row1c2 = st.columns(2)
    with row1c1:
        st.markdown("**Winrate temporada**")
        render_stat_table(
            summary_df,
            formatters={"winrate": "{:.2%}"},
            headers={"jugadas": "Jugadas", "wins": "Win Totales", "losses": "Losses Totales", "winrate": "Win% Temporada"},
        )
    with row1c2:
        st.markdown("**Winrate x Set**")
        render_stat_table(
            set_stats,
            formatters={"winrate": "{:.2%}"},
            headers={"set": "Set", "wins": "Wins", "losses": "Losses", "partidas": "Partidas", "winrate": "Win% x Set"},
        )

    row2c1, row2c2 = st.columns(2)
    with row2c1:
        st.markdown("**Winrate x Rival**")
        render_stat_table(
            rival_stats,
            formatters={"winrate": "{:.2%}"},
            headers={"rival": "Rival", "wins": "Wins", "losses": "Losses", "winrate": "Win% contra Rival"},
        )
    with row2c2:
        st.markdown("**Winrate x (Set+Deck)**")
        render_stat_table(
            deck_stats,
            formatters={"winrate": "{:.2%}"},
            headers={
                "set": "Set", "arquetipo": "Deck", "wins": "Wins", "losses": "Losses",
                "partidas": "Partidas", "winrate": "Win% x Set",
            },
        )


def page_estadisticas() -> None:
    matches, entries, attendance, player_refs, set_refs, deck_info = load_data()

    # color por set calculado sobre el universo completo (sin filtrar por
    # temporada) para que un set siempre tenga el mismo color en toda la app
    set_color_map = build_set_color_map(entries["set"], matches["set1"], matches["set2"])

    st.sidebar.title("Weiss Sevilla")
    temporadas = sorted(matches["temporada"].dropna().unique().tolist(), reverse=True)
    temporada_sel = st.sidebar.selectbox("Temporada", ["Todas"] + [str(t) for t in temporadas])
    if temporada_sel != "Todas":
        year = int(temporada_sel)
        matches = matches[matches["temporada"] == year]
        entries = entries[entries["temporada"] == year]
        attendance = attendance[attendance["temporada"] == year]

    tab_rankings, tab_resumen, tab_partidas, tab_individuales = st.tabs(
        ["Rankings", "Resumen", "Partidas", "Resultados individuales"]
    )
    with tab_rankings:
        render_rankings(entries, player_refs, set_refs, deck_info)
    with tab_resumen:
        render_resumen(matches, entries, attendance, set_color_map)
    with tab_partidas:
        render_partidas(matches)
    with tab_individuales:
        render_resultados_individuales(entries, matches, player_refs, set_color_map)


# ------------------------------------------------------------------
# Pagina de administracion (con contrasena)
# ------------------------------------------------------------------


ADMIN_TABLES = ["matches", "entries", "attendance", "player_refs", "set_refs", "deck_info"]


def render_admin_import() -> None:
    st.subheader("Importar Excel")
    st.markdown(
        '<div class="xl-admin-warning">Sube <b>todos</b> los .xlsx de todas las temporadas cada vez '
        "(la base se reconstruye entera desde cero, no de forma incremental). Si la app esta "
        "desplegada en Streamlit Community Cloud, esto actualiza los datos al instante para "
        "quien la este viendo, pero <b>no sobrevive a un reinicio o redeploy</b>: descarga la "
        "weiss.db resultante (abajo) y subela al repositorio de git para dejarla fija.</div>",
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader("Excel de temporada(s)", type="xlsx", accept_multiple_files=True)
    if st.button("Importar y reconstruir weiss.db", disabled=not uploaded):
        try:
            report = build_database(uploaded, DB_PATH)
        except IngestaError as exc:
            st.error(str(exc))
        else:
            st.success("weiss.db reconstruida correctamente.")
            if report.warnings:
                st.warning("\n\n".join(f"- {w}" for w in report.warnings))
            load_data.clear()

    if DB_PATH.exists():
        st.download_button(
            "Descargar weiss.db actual",
            DB_PATH.read_bytes(),
            file_name="weiss.db",
            mime="application/octet-stream",
        )


def render_admin_editor() -> None:
    st.subheader("Corregir datos")
    st.markdown(
        '<div class="xl-admin-warning">Estos cambios se guardan directamente en weiss.db. '
        "Si mas adelante vuelves a importar un Excel, esta tabla se recalcula desde el Excel y "
        "estas correcciones manuales se pierden: para una correccion permanente, arreglala en el "
        "Excel de origen cuando puedas.</div>",
        unsafe_allow_html=True,
    )
    if not DB_PATH.exists():
        st.info("Todavia no hay weiss.db. Importa un Excel primero.")
        return

    table = st.selectbox("Tabla", ADMIN_TABLES)
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(f"SELECT * FROM {table}", conn)
    finally:
        conn.close()

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"editor_{table}")
    if st.button(f"Guardar cambios en '{table}'"):
        with sqlite3.connect(DB_PATH) as conn:
            edited.to_sql(table, conn, if_exists="replace", index=False)
        st.success("Guardado.")
        load_data.clear()


def page_admin() -> None:
    st.markdown('<div class="xl-title"><h2>Administrar</h2></div>', unsafe_allow_html=True)

    if not check_admin_password():
        st.stop()

    if st.sidebar.button("Cerrar sesion de administrador"):
        st.session_state["admin_authenticated"] = False
        st.rerun()

    tab_import, tab_edit = st.tabs(["Importar Excel", "Corregir datos"])
    with tab_import:
        render_admin_import()
    with tab_edit:
        render_admin_editor()


def main() -> None:
    st.set_page_config(page_title="Weiss Sevilla", page_icon="🃏", layout="wide")
    st.markdown(XL_CSS, unsafe_allow_html=True)

    # el enlace publico (sin ?admin=1) no muestra "Administrar" en el menu;
    # el enlace de administracion (con ?admin=1) si lo muestra, pero la
    # contrasena sigue haciendo falta para entrar - esto solo oculta el enlace.
    # Se guarda en session_state porque al navegar entre paginas Streamlit no
    # conserva el query param en la URL, y sin esto el enlace desaparecería
    # justo al hacer click en "Administrar".
    if st.query_params.get("admin") == "1":
        st.session_state["admin_link_used"] = True

    pages = [st.Page(page_estadisticas, title="Estadisticas", icon="📊", default=True)]
    if st.session_state.get("admin_link_used"):
        pages.append(st.Page(page_admin, title="Administrar", icon="🔒"))

    pg = st.navigation(pages)
    pg.run()


if __name__ == "__main__":
    main()
