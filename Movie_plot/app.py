"""
CINEVECTOR — Movie Intelligence Engine
A Netflix-styled RAG / vector-search dashboard for the SE488 project.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Deploy on Streamlit Community Cloud the same way you deployed your
market-basket dashboard: push this folder (app.py, requirements.txt,
movie_plots_premium_filled_partial.xlsx) to a GitHub repo and point
Streamlit Cloud at app.py.
"""

import os
import re
import time
import sqlite3
import string
import warnings
import base64
import io
import textwrap

import numpy as np
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image, ImageDraw, ImageFont

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────
# PAGE CONFIG + NETFLIX THEME
# ─────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="CineVector | Movie Intelligence Engine",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="collapsed",
)

NETFLIX_RED = "#E50914"
NETFLIX_BLACK = "#141414"
NETFLIX_DARK = "#181818"
NETFLIX_CARD = "#232323"
NETFLIX_GRAY = "#808080"
NETFLIX_WHITE = "#F5F5F1"

CUSTOM_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background-color: {NETFLIX_BLACK} !important;
    color: {NETFLIX_WHITE};
}}

.stApp {{
    background: linear-gradient(180deg, #0b0b0b 0%, {NETFLIX_BLACK} 8%, {NETFLIX_BLACK} 100%);
}}

/* Hide default streamlit chrome */
#MainMenu, footer, header {{visibility: hidden;}}

/* ── Hero ───────────────────────────────────────────────────────── */
.cv-hero {{
    padding: 3.5rem 0 2rem 0;
    text-align: center;
    border-bottom: 1px solid #2a2a2a;
    margin-bottom: 2rem;
}}
.cv-logo {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 4.2rem;
    letter-spacing: 0.08em;
    color: {NETFLIX_RED};
    margin-bottom: 0;
    line-height: 1;
    text-shadow: 0 0 40px rgba(229,9,20,0.35);
}}
.cv-tagline {{
    font-size: 0.95rem;
    letter-spacing: 0.35em;
    text-transform: uppercase;
    color: {NETFLIX_GRAY};
    margin-top: 0.4rem;
    font-weight: 600;
}}
.cv-sub {{
    color: #b3b3b3;
    font-size: 1.05rem;
    margin-top: 1.2rem;
    max-width: 640px;
    margin-left: auto;
    margin-right: auto;
}}

/* ── Stat chips ─────────────────────────────────────────────────── */
.cv-chip-row {{
    display: flex;
    justify-content: center;
    gap: 2.5rem;
    margin-top: 2rem;
    flex-wrap: wrap;
}}
.cv-chip {{
    text-align: center;
}}
.cv-chip-value {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.1rem;
    color: {NETFLIX_WHITE};
}}
.cv-chip-label {{
    font-size: 0.72rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: {NETFLIX_GRAY};
    margin-top: -0.3rem;
}}

/* ── Section labels ─────────────────────────────────────────────── */
.cv-section-tag {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 0.85rem;
    letter-spacing: 0.3em;
    color: {NETFLIX_RED};
    text-transform: uppercase;
    margin-bottom: -0.4rem;
}}
.cv-section-title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.1rem;
    letter-spacing: 0.03em;
    color: {NETFLIX_WHITE};
    margin-bottom: 0.6rem;
}}

/* ── Movie / result card ────────────────────────────────────────── */
.cv-card {{
    background: {NETFLIX_CARD};
    border: 1px solid #2f2f2f;
    border-radius: 6px;
    padding: 1.1rem 1.3rem;
    margin-bottom: 0.9rem;
    transition: border-color 0.15s ease;
}}
.cv-card:hover {{
    border-color: {NETFLIX_RED};
}}
.cv-card-title {{
    font-size: 1.15rem;
    font-weight: 700;
    color: {NETFLIX_WHITE};
}}
.cv-card-meta {{
    font-size: 0.78rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: {NETFLIX_GRAY};
    margin-bottom: 0.5rem;
}}
.cv-card-plot {{
    font-size: 0.88rem;
    color: #cfcfcf;
    line-height: 1.5;
}}
.cv-badge {{
    display: inline-block;
    background: rgba(229,9,20,0.15);
    border: 1px solid {NETFLIX_RED};
    color: #ff6b6f;
    font-size: 0.72rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 0.15rem 0.55rem;
    border-radius: 3px;
    margin-top: 0.4rem;
}}

/* ── Index lab card ─────────────────────────────────────────────── */
.cv-index-card {{
    background: {NETFLIX_DARK};
    border: 1px solid #2f2f2f;
    border-radius: 8px;
    padding: 1.4rem;
    text-align: center;
    height: 100%;
}}
.cv-index-name {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.5rem;
    letter-spacing: 0.05em;
    color: {NETFLIX_WHITE};
}}
.cv-index-tag {{
    font-size: 0.7rem;
    letter-spacing: 0.15em;
    color: {NETFLIX_RED};
    text-transform: uppercase;
    margin-bottom: 0.6rem;
}}

/* ── misc ───────────────────────────────────────────────────────── */
.cv-divider {{
    border-top: 1px solid #2a2a2a;
    margin: 2.5rem 0 2rem 0;
}}
.cv-mono-label {{
    font-size: 0.7rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: {NETFLIX_GRAY};
}}

/* ── Channel switcher (module navigation, one view at a time) ──────── */
.cv-channel {{
    text-align: center;
    padding: 0.3rem 0 0.1rem 0;
}}
.cv-channel-num {{
    font-size: 0.68rem;
    letter-spacing: 0.3em;
    color: {NETFLIX_RED};
    font-weight: 700;
}}
.cv-channel-name {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.1rem;
    letter-spacing: 0.05em;
    color: {NETFLIX_WHITE};
    line-height: 1.15;
    text-shadow: 0 0 24px rgba(229,9,20,0.25);
}}
.cv-channel-desc {{
    font-size: 0.78rem;
    color: {NETFLIX_GRAY};
    margin-top: -0.1rem;
}}
div[data-testid="column"]:has(button[kind="secondary"]) button {{
    background: {NETFLIX_DARK} !important;
    border: 1px solid #2a2a2a !important;
    color: {NETFLIX_WHITE} !important;
    font-size: 1.1rem !important;
    border-radius: 50% !important;
    aspect-ratio: 1 / 1 !important;
    padding: 0 !important;
}}
div[data-testid="column"]:has(button[kind="secondary"]) button:hover {{
    border-color: {NETFLIX_RED} !important;
    color: {NETFLIX_RED} !important;
}}
div[data-baseweb="select"] {{
    max-width: 320px;
    margin: 0.6rem auto 0 auto;
}}

.stButton>button {{
    background-color: {NETFLIX_RED};
    color: white;
    border: none;
    border-radius: 4px;
    font-weight: 700;
    letter-spacing: 0.05em;
    padding: 0.5rem 1.6rem;
}}
.stButton>button:hover {{
    background-color: #f6121d;
    color: white;
}}

input, textarea {{
    background-color: {NETFLIX_DARK} !important;
    color: {NETFLIX_WHITE} !important;
}}

/* ── Sliding poster row (marquee) ───────────────────────────────── */
.cv-marquee-label {{
    font-size: 0.72rem;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: {NETFLIX_GRAY};
    margin: 0 0 0.6rem 0.1rem;
}}
.cv-marquee-wrap {{
    overflow: hidden;
    width: 100%;
    -webkit-mask-image: linear-gradient(90deg, transparent, black 4%, black 96%, transparent);
    mask-image: linear-gradient(90deg, transparent, black 4%, black 96%, transparent);
    margin-bottom: 1rem;
}}
.cv-marquee-track {{
    display: flex;
    gap: 14px;
    width: max-content;
    animation: cv-slide 45s linear infinite;
}}
.cv-marquee-track:hover {{
    animation-play-state: paused;
}}
@keyframes cv-slide {{
    from {{ transform: translateX(0); }}
    to {{ transform: translateX(-50%); }}
}}
.cv-poster-tile {{
    position: relative;
    width: 118px;
    height: 176px;
    border-radius: 6px;
    overflow: hidden;
    flex: 0 0 auto;
    border: 1px solid rgba(229, 9, 20, 0.28);
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.55);
    transition: transform 0.2s ease, border-color 0.2s ease;
}}
.cv-poster-tile:hover {{
    transform: scale(1.045);
    border-color: {NETFLIX_RED};
    z-index: 2;
}}
.cv-poster-tile img {{
    width: 100%;
    height: 100%;
    object-fit: cover;
    display: block;
    transition: transform 0.5s ease;
}}
.cv-poster-tile:hover img {{
    transform: scale(1.08);
}}
.cv-poster-tile .cv-poster-caption {{
    position: absolute;
    left: 0; right: 0; bottom: 0;
    background: linear-gradient(180deg, transparent, rgba(0,0,0,0.9) 65%);
    padding: 24px 7px 7px;
    font-size: 0.66rem;
    color: {NETFLIX_WHITE};
    font-weight: 800;
    letter-spacing: 0.01em;
    line-height: 1.25;
}}
.cv-poster-badge {{
    position: absolute;
    top: 6px; left: 6px;
    background: {NETFLIX_RED};
    color: {NETFLIX_WHITE};
    font-size: 0.6rem;
    font-weight: 800;
    letter-spacing: 0.03em;
    padding: 2px 6px;
    border-radius: 3px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.5);
    z-index: 3;
}}

/* ── Poster-thumb row card (used in search/evidence results) ──────── */
.cv-row-card {{
    display: flex;
    gap: 1rem;
    background: {NETFLIX_CARD};
    border: 1px solid #2f2f2f;
    border-radius: 6px;
    padding: 0.7rem;
    margin-bottom: 0.8rem;
    align-items: flex-start;
}}
.cv-row-card:hover {{ border-color: {NETFLIX_RED}; }}
.cv-row-card img {{
    width: 64px;
    height: 96px;
    object-fit: cover;
    border-radius: 4px;
    border: 1px solid rgba(229, 9, 20, 0.28);
    flex: 0 0 auto;
}}

/* ── Filmstrip timeline (signature Explore visual) ─────────────────── */
.cv-filmstrip {{
    background: {NETFLIX_DARK};
    border: 1px solid #2a2a2a;
    border-radius: 4px;
    padding: 0.9rem 0 0.4rem 0;
    margin-top: 0.6rem;
    position: relative;
}}
.cv-sprockets {{
    display: flex;
    justify-content: space-evenly;
    padding: 0 0.6rem;
}}
.cv-sprockets span {{
    width: 9px;
    height: 9px;
    border-radius: 2px;
    background: {NETFLIX_BLACK};
    border: 1px solid #333;
    display: inline-block;
}}

/* ── Ranked leaderboard rows (genre breakdown) ─────────────────────── */
.cv-rank-row {{
    display: flex;
    align-items: center;
    gap: 0.8rem;
    padding: 0.42rem 0;
    border-bottom: 1px solid #232323;
}}
.cv-rank-num {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1rem;
    color: #4a4a4a;
    width: 1.6rem;
    flex: 0 0 auto;
}}
.cv-rank-label {{
    font-size: 0.82rem;
    color: {NETFLIX_WHITE};
    width: 13rem;
    flex: 0 0 auto;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.cv-rank-track {{
    flex: 1 1 auto;
    background: #1c1c1c;
    border-radius: 3px;
    height: 9px;
    overflow: hidden;
}}
.cv-rank-fill {{
    height: 100%;
    border-radius: 3px;
    background: linear-gradient(90deg, #5c0509, {NETFLIX_RED});
}}
.cv-rank-count {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1rem;
    color: {NETFLIX_GRAY};
    width: 1.8rem;
    flex: 0 0 auto;
    text-align: right;
}}

/* ── Collection breakdown: one panel, leaderboard + inline stat rail ─ */
.cv-breakdown-panel {{
    display: flex;
    align-items: stretch;
    background: {NETFLIX_DARK};
    border: 1px solid #2a2a2a;
    border-radius: 4px;
    overflow: hidden;
}}
.cv-breakdown-main {{
    flex: 1 1 auto;
    padding: 0.9rem 1.3rem;
    min-width: 0;
}}
.cv-breakdown-rail {{
    flex: 0 0 200px;
    padding: 1.1rem 1.3rem;
    border-left: 1px solid #2a2a2a;
    background: repeating-linear-gradient(
        135deg, rgba(229,9,20,0.03) 0px, rgba(229,9,20,0.03) 10px,
        transparent 10px, transparent 20px
    );
}}
.cv-wc-big {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.4rem;
    color: {NETFLIX_WHITE};
    line-height: 1;
}}
.cv-wc-range {{
    font-size: 0.68rem;
    letter-spacing: 0.1em;
    color: {NETFLIX_GRAY};
    text-transform: uppercase;
    margin-top: 0.2rem;
}}
.cv-wc-bars {{
    display: flex;
    align-items: flex-end;
    gap: 3px;
    height: 40px;
    margin-top: 0.9rem;
}}
.cv-wc-bars div {{
    flex: 1 1 auto;
    background: {NETFLIX_RED};
    opacity: 0.85;
    border-radius: 2px 2px 0 0;
    min-height: 3px;
}}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────
# DATA LOADING — session-state backed, so it can be edited/replaced live
# ─────────────────────────────────────────────────────────────────────────

DEFAULT_DATA_PATH = os.path.join(
    os.path.dirname(__file__), "movie_plots_premium_filled_partial.xlsx"
)

REQUIRED_COLUMNS = [
    "movie_id", "movie_title", "genre", "year",
    "main_characters", "themes", "keywords", "plot",
]


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize any uploaded/edited dataframe into the shape the app expects:
    all required columns present, clean IDs, numeric year, and a rebuilt
    'text' column used for embeddings/TF-IDF."""
    df = df.copy()
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[REQUIRED_COLUMNS].fillna("")
    df = df.reset_index(drop=True)

    # Assign clean, unique movie_id values — fill in blanks/duplicates.
    seen, ids, next_id = set(), [], 1
    for v in df["movie_id"]:
        v = str(v).strip()
        if v and v not in seen:
            ids.append(v)
        else:
            while str(next_id) in seen:
                next_id += 1
            v = str(next_id)
            ids.append(v)
        seen.add(v)
    df["movie_id"] = ids

    df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(0).astype(int)

    for col in ["movie_title", "genre", "main_characters", "themes", "keywords", "plot"]:
        df[col] = df[col].astype(str)

    df["text"] = (
        df["movie_title"] + ". " + df["genre"] + ". " + df["main_characters"] + ". " +
        df["themes"] + ". " + df["keywords"] + ". " + df["plot"]
    )
    return df


def load_dataframe_from_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raw = pd.read_excel(uploaded_file)
    return prepare_dataframe(raw)


if "movies" not in st.session_state:
    if os.path.exists(DEFAULT_DATA_PATH):
        st.session_state.movies = prepare_dataframe(pd.read_excel(DEFAULT_DATA_PATH))
    else:
        st.session_state.movies = prepare_dataframe(pd.DataFrame(columns=REQUIRED_COLUMNS))

movies = st.session_state.movies
HAS_MOVIES = len(movies) > 0


def dataset_fingerprint(df: pd.DataFrame) -> int:
    """A cache key that changes on ANY edit — added/deleted/edited rows,
    any column — so cached embeddings/indexes/SQL views never go stale."""
    if df.empty:
        return 0
    return int(pd.util.hash_pandas_object(df, index=False).sum())


# ─────────────────────────────────────────────────────────────────────────
# POSTERS
# ─────────────────────────────────────────────────────────────────────────
# Real poster art requires a licensed source. We never scrape/hotlink
# posters ourselves — instead: (1) if the user supplies their own free
# TMDB API key, we fetch real posters through TMDB's official API, which
# is exactly what that API is for; (2) otherwise we generate an original,
# genre-tinted placeholder "poster" locally with Pillow, so the app still
# looks like a real poster grid with zero setup and zero copyright risk.

GENRE_PALETTES = {
    "Superhero": [(30, 10, 12), (120, 15, 20)],
    "Superhero,Fantasy": [(25, 8, 40), (90, 20, 110)],
    "Superhero,Science Fiction": [(6, 20, 40), (10, 70, 110)],
    "Animation, Family": [(60, 30, 5), (170, 100, 10)],
    "Animation, Adventure": [(4, 35, 45), (10, 110, 130)],
    "Animation, Fantasy": [(35, 8, 45), (110, 30, 130)],
    "Science Fiction, Adventure": [(3, 25, 20), (10, 90, 70)],
}
DEFAULT_PALETTE = [(30, 30, 30), (70, 70, 70)]


def _palette_for(genre):
    return GENRE_PALETTES.get(genre, DEFAULT_PALETTE)


@st.cache_data(show_spinner=False)
def generate_placeholder_poster(title, year, genre, size=(300, 450)):
    w, h = size
    c1, c2 = _palette_for(genre)
    img = Image.new("RGB", (w, h), c1)
    draw = ImageDraw.Draw(img)
    for y in range(h):
        t = y / h
        r = int(c1[0] + (c2[0] - c1[0]) * t)
        g = int(c1[1] + (c2[1] - c1[1]) * t)
        b = int(c1[2] + (c2[2] - c1[2]) * t)
        draw.line([(0, y), (w, y)], fill=(r, g, b))

    overlay = Image.new("L", (w, h), 0)
    odraw = ImageDraw.Draw(overlay)
    odraw.ellipse([-w * 0.3, -h * 0.3, w * 1.3, h * 1.3], fill=60)
    img = Image.composite(Image.new("RGB", (w, h), (0, 0, 0)), img, overlay.point(lambda p: 255 - p))

    # thin outer border, like the CSS accent border real posters get
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, w - 1, h - 1], outline=(229, 9, 20, 90), width=2)

    font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    try:
        tag_font = ImageFont.truetype(font_path, int(size[0] * 0.038))
        title_font = ImageFont.truetype(font_path, int(size[0] * 0.088))
        meta_font = ImageFont.truetype(font_path, int(size[0] * 0.042))
        badge_font = ImageFont.truetype(font_path, int(size[0] * 0.14))
    except Exception:
        tag_font = title_font = meta_font = badge_font = ImageFont.load_default()

    # top tag row — "CINEVECTOR" left, year right, like the proflix badge card
    pad = int(w * 0.07)
    draw.text((pad, pad * 0.6), "CINEVECTOR", font=tag_font, fill=(229, 9, 20))
    year_txt = str(year)
    ybbox = draw.textbbox((0, 0), year_txt, font=tag_font)
    draw.text((w - pad - (ybbox[2] - ybbox[0]), pad * 0.6), year_txt, font=tag_font, fill=(210, 210, 210))

    # centered circular initial badge
    initial = str(title).strip()[:1].upper() or "?"
    cx, cy, cr = w / 2, h * 0.36, w * 0.11
    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], outline=(229, 9, 20), width=3)
    ibbox = draw.textbbox((0, 0), initial, font=badge_font)
    draw.text((cx - (ibbox[2] - ibbox[0]) / 2, cy - (ibbox[3] - ibbox[1]) / 2 - ibbox[1]), initial, font=badge_font, fill=(245, 245, 241))

    # title, wrapped and centered below the badge
    wrapped = textwrap.wrap(str(title), width=15)
    line_h = int(size[0] * 0.105)
    y0 = h * 0.52
    for i, line in enumerate(wrapped[:3]):
        bbox = draw.textbbox((0, 0), line, font=title_font)
        lw = bbox[2] - bbox[0]
        draw.text(((w - lw) / 2, y0 + i * line_h), line, font=title_font, fill=(245, 245, 241))

    # category / genre line beneath the title
    if genre:
        cat_txt = str(genre).upper()
        cbbox = draw.textbbox((0, 0), cat_txt, font=meta_font)
        cw = cbbox[2] - cbbox[0]
        cat_y = y0 + len(wrapped[:3]) * line_h + int(size[0] * 0.02)
        draw.text(((w - cw) / 2, cat_y), cat_txt, font=meta_font, fill=(200, 200, 200))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def _tmdb_key():
    key = st.session_state.get("tmdb_api_key", "").strip()
    if key:
        return key
    try:
        return st.secrets.get("TMDB_API_KEY", "")
    except Exception:
        return ""


@st.cache_data(show_spinner=False)
def fetch_tmdb_poster_url(title, year, api_key):
    if not api_key:
        return None
    try:
        clean_title = re.sub(r"\s*\(\d{4}\)$", "", str(title))
        resp = requests.get(
            "https://api.themoviedb.org/3/search/movie",
            params={"api_key": api_key, "query": clean_title, "year": year},
            timeout=6,
        )
        data = resp.json()
        results = data.get("results", [])
        if not results and year:
            # retry without year in case TMDB's year match is too strict
            resp = requests.get(
                "https://api.themoviedb.org/3/search/movie",
                params={"api_key": api_key, "query": clean_title},
                timeout=6,
            )
            results = resp.json().get("results", [])
        if results:
            poster_path = results[0].get("poster_path")
            if poster_path:
                return f"https://image.tmdb.org/t/p/w342{poster_path}"
    except Exception:
        pass
    return None


def _omdb_key():
    key = st.session_state.get("omdb_api_key", "").strip()
    if key:
        return key
    try:
        return st.secrets.get("OMDB_API_KEY", "")
    except Exception:
        return ""


@st.cache_data(show_spinner=False)
def fetch_omdb_poster_url(title, year, api_key):
    if not api_key:
        return None
    try:
        clean_title = re.sub(r"\s*\(\d{4}\)$", "", str(title))
        resp = requests.get(
            "https://www.omdbapi.com/",
            params={"apikey": api_key, "t": clean_title, "y": year, "type": "movie"},
            timeout=6,
        )
        data = resp.json()
        if data.get("Response") == "False" and year:
            # retry without the year in case OMDb's year match is too strict
            resp = requests.get(
                "https://www.omdbapi.com/",
                params={"apikey": api_key, "t": clean_title, "type": "movie"},
                timeout=6,
            )
            data = resp.json()
        poster = data.get("Poster")
        if poster and poster != "N/A":
            return poster
    except Exception:
        pass
    return None


def _poster_seed(title, year):
    """Deterministic seed so the same movie always gets the same stock photo."""
    raw = f"{title}-{year}"
    return re.sub(r"[^a-zA-Z0-9]+", "-", raw).strip("-").lower() or "movie"


def get_real_photo_url(title, year, size=(300, 450)):
    """A real, licensed-for-reuse photo with no API key or signup required
    (Lorem Picsum, backed by Unsplash-sourced CC0 photography). Not an
    actual movie poster — just a real photograph so the card doesn't look
    like generated art — seeded per movie so it stays consistent."""
    seed = _poster_seed(title, year)
    w, h = size
    return f"https://picsum.photos/seed/{seed}/{w}/{h}"


def get_poster(title, year, genre):
    """Primary image source, in priority order: TMDB real poster → OMDb real
    poster (either, if a key is set) → free no-signup generic stock photo →
    nothing (caller should use get_poster_fallback for a guaranteed-available
    generated card)."""
    tmdb = _tmdb_key()
    if tmdb:
        url = fetch_tmdb_poster_url(title, year, tmdb)
        if url:
            return url
    omdb = _omdb_key()
    if omdb:
        url = fetch_omdb_poster_url(title, year, omdb)
        if url:
            return url
    return get_real_photo_url(title, year)


def get_poster_fallback(title, year, genre):
    """Guaranteed-available generated card, used as the onerror fallback in
    case a real photo fails to load (e.g. no internet egress)."""
    return generate_placeholder_poster(title, year, genre)


def poster_img_tag(title, year, genre, css_class="", style="", loading="lazy"):
    """A single <img> tag that tries a real photo first and swaps itself to
    the generated card if that photo fails to load — use this everywhere
    instead of building <img src="..."> by hand."""
    primary = get_poster(title, year, genre)
    fallback = get_poster_fallback(title, year, genre)
    class_attr = f' class="{css_class}"' if css_class else ""
    style_attr = f' style="{style}"' if style else ""
    return (
        f'<img{class_attr}{style_attr} loading="{loading}" src="{primary}" '
        f"onerror=\"this.onerror=null;this.src='{fallback}';\" />"
    )


# ─────────────────────────────────────────────────────────────────────────
# MODELS + INDEXES (cached — built once per session)
# ─────────────────────────────────────────────────────────────────────────


@st.cache_data(show_spinner=False)
def generate_backdrop(title, genre, size=(1280, 480)):
    """Wide cinematic backdrop for the hero banner — dark on the left for
    text legibility, genre-tinted color bleeding in from the right, same
    spirit as a real streaming-service hero image."""
    w, h = size
    c1, c2 = _palette_for(genre)
    img = Image.new("RGB", (w, h), c1)
    draw = ImageDraw.Draw(img)
    for x in range(w):
        t = x / w
        r = int(c1[0] + (c2[0] - c1[0]) * t)
        g = int(c1[1] + (c2[1] - c1[1]) * t)
        b = int(c1[2] + (c2[2] - c1[2]) * t)
        draw.line([(x, 0), (x, h)], fill=(r, g, b))
    overlay = Image.new("L", (w, h), 0)
    odraw = ImageDraw.Draw(overlay)
    for x in range(w):
        alpha = max(0, 235 - int(235 * (x / (w * 0.65))))
        odraw.line([(x, 0), (x, h)], fill=alpha)
    black = Image.new("RGB", (w, h), (8, 8, 8))
    img = Image.composite(black, img, overlay)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def get_backdrop(title, year, genre):
    return generate_backdrop(title, genre)


@st.cache_resource(show_spinner="Loading embedding model...")
def load_embedder():
    from sentence_transformers import SentenceTransformer
    return SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")


@st.cache_resource(show_spinner="Building embeddings...")
def build_embeddings(movies_text_tuple):
    embedder = load_embedder()
    texts = list(movies_text_tuple)
    emb = embedder.encode(texts, convert_to_numpy=True, normalize_embeddings=True)
    return emb


@st.cache_resource(show_spinner="Building FAISS indexes...")
def build_indexes(_embeddings, cache_key):
    import faiss
    dimension = _embeddings.shape[1]
    n = _embeddings.shape[0]

    flat_index = faiss.IndexFlatL2(dimension)
    flat_index.add(_embeddings)

    nlist = max(1, min(5, n))
    quantizer = faiss.IndexFlatL2(dimension)
    ivf_index = faiss.IndexIVFFlat(quantizer, dimension, nlist)
    ivf_index.train(_embeddings)
    ivf_index.add(_embeddings)
    ivf_index.nprobe = min(3, nlist)

    m, bits = 8, 4
    pq_index = faiss.IndexPQ(dimension, m, bits)
    pq_index.train(_embeddings)
    pq_index.add(_embeddings)

    quantizer2 = faiss.IndexFlatL2(dimension)
    ivfpq_index = faiss.IndexIVFPQ(quantizer2, dimension, nlist, m, bits)
    ivfpq_index.train(_embeddings)
    ivfpq_index.add(_embeddings)
    ivfpq_index.nprobe = min(3, nlist)

    hnsw_index = faiss.IndexHNSWFlat(dimension, 32)
    hnsw_index.add(_embeddings)

    return {
        "IndexFlatL2": flat_index,
        "IndexIVFFlat": ivf_index,
        "IndexPQ": pq_index,
        "IndexIVFPQ": ivfpq_index,
        "IndexHNSWFlat": hnsw_index,
    }


@st.cache_resource(show_spinner="Loading FLAN-T5...")
def load_generator():
    from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
    tok = AutoTokenizer.from_pretrained("google/flan-t5-base", truncation_side="left")
    model = AutoModelForSeq2SeqLM.from_pretrained("google/flan-t5-base")
    return tok, model


@st.cache_resource(show_spinner=False)
def build_tfidf(movie_documents_tuple):
    from sklearn.feature_extraction.text import TfidfVectorizer
    docs = list(movie_documents_tuple)
    vec = TfidfVectorizer(lowercase=True, stop_words="english", ngram_range=(1, 2), sublinear_tf=True)
    matrix = vec.fit_transform(docs)
    return vec, matrix


if HAS_MOVIES:
    text_tuple = tuple(movies["text"].tolist())
    embeddings = build_embeddings(text_tuple)
    dimension = embeddings.shape[1]
    indexes = build_indexes(embeddings, text_tuple)
else:
    dimension = 384
    embeddings = np.zeros((0, dimension), dtype="float32")
    indexes = {}

STOPWORDS = set("""
a an the and or but if then than is are was were be been being
of to in on for from with by about into over after before during
which who what where when why how movie film films does do did
this that these those a lot tell me please involving involves
features feature focuses focus based around as at it its they them
he she his her their your you i we our about story follows
""".split())


def normalize_text(text):
    text = str(text).lower()
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"[^a-z0-9\s-]", " ", text)


def tokens(text):
    words = re.findall(r"[a-z0-9]+", normalize_text(text))
    return [w for w in words if w not in STOPWORDS and len(w) > 1]


def token_set(text):
    return set(tokens(text))


movie_documents = [
    f"{row['movie_title']} {row.get('genre','')} {row.get('main_characters','')} "
    f"{row.get('themes','')} {row.get('keywords','')} {row.get('plot','')}"
    for _, row in movies.iterrows()
]
if HAS_MOVIES:
    tfidf_vectorizer, tfidf_matrix = build_tfidf(tuple(movie_documents))
else:
    tfidf_vectorizer, tfidf_matrix = None, None


def cleaned_title(title):
    return re.sub(r"\s*\(\d{4}\)$", "", str(title).lower()).strip()


def find_direct_entities(query):
    q = normalize_text(query)
    title_hits, character_hits, keyword_hits = [], [], []
    for idx, row in movies.iterrows():
        title = cleaned_title(row["movie_title"])
        chars = [c.strip().lower() for c in str(row.get("main_characters", "")).split(",") if c.strip()]
        keywords = [k.strip().lower() for k in str(row.get("keywords", "")).split(",") if k.strip()]

        if title and title in q:
            title_hits.append(idx)
            continue
        if any(len(c) >= 5 and c in q for c in chars):
            character_hits.append(idx)
            continue
        if any(len(k) >= 5 and k in q for k in keywords):
            keyword_hits.append(idx)

    if title_hits:
        return title_hits
    if character_hits:
        return character_hits
    if keyword_hits:
        return keyword_hits
    return []


def structured_match_score(query, row):
    q_tokens = set(tokens(query))
    if not q_tokens:
        return 0.0
    fields = [
        (str(row.get("movie_title", "")), 5.0),
        (str(row.get("main_characters", "")), 4.0),
        (str(row.get("keywords", "")), 4.0),
        (str(row.get("themes", "")), 2.5),
        (str(row.get("plot", "")), 1.0),
    ]
    matched_weight, total_weight = 0.0, sum(w for _, w in fields)
    for text, weight in fields:
        field_tokens = token_set(text)
        overlap = len(q_tokens & field_tokens)
        if overlap:
            matched_weight += weight * min(1.0, overlap / max(1, min(3, len(q_tokens))))
    return min(1.0, matched_weight / total_weight)


CONCEPTS = {
    "music": {"music", "musician", "guitar", "song", "songs"},
    "boy": {"boy", "young", "child", "miguel"},
    "girl": {"girl", "young", "child", "riley"},
    "love": {"love", "loves", "loving", "family", "relationship", "dream", "dreams", "passion"},
    "mysterio": {"mysterio", "quentin", "beck", "illusion", "drones"},
    "multiverse": {"multiverse", "dimension", "dimensions", "universe", "reality"},
    "nature": {"nature", "environment", "ecosystem", "biodiversity", "conservation"},
    "time travel": {"time", "travel", "time-travel", "endgame"},
    "magic": {"magic", "magical", "sorcerer", "mystic", "spells", "powers"},
    "wakanda": {"wakanda", "tchalla", "killmonger", "leadership"},
}


def expanded_concept_score(query, row):
    q = normalize_text(query)
    query_concepts = set()
    for phrase, words in CONCEPTS.items():
        if phrase in q or any(w in q.split() for w in words):
            query_concepts.update(words)
    if not query_concepts:
        return 0.0
    doc_tokens = token_set(
        f"{row.get('movie_title','')} {row.get('main_characters','')} "
        f"{row.get('themes','')} {row.get('keywords','')} {row.get('plot','')}"
    )
    overlap = query_concepts & doc_tokens
    return len(overlap) / max(1, min(6, len(query_concepts)))


def retrieve_movies(query, k=3, min_combined_score=0.18):
    if not HAS_MOVIES:
        return pd.DataFrame()

    direct = find_direct_entities(query)
    if direct:
        results = movies.iloc[direct[:k]].copy()
        for col in ["semantic_score", "tfidf_score", "structured_score", "concept_score", "combined_score"]:
            results[col] = 1.0
        return results

    embedder = load_embedder()
    query_vec = embedder.encode([query], convert_to_numpy=True, normalize_embeddings=True)
    semantic_scores = np.clip(np.dot(embeddings, query_vec[0]), 0.0, 1.0)

    from sklearn.metrics.pairwise import cosine_similarity
    q_tfidf = tfidf_vectorizer.transform([query])
    tfidf_scores = np.clip(cosine_similarity(q_tfidf, tfidf_matrix)[0], 0.0, 1.0)

    records = []
    for i, row in movies.iterrows():
        structured = structured_match_score(query, row)
        concept = expanded_concept_score(query, row)
        combined = (
            0.45 * float(semantic_scores[i]) +
            0.30 * float(tfidf_scores[i]) +
            0.15 * float(structured) +
            0.10 * float(concept)
        )
        records.append({
            "_idx": i, "semantic_score": float(semantic_scores[i]),
            "tfidf_score": float(tfidf_scores[i]), "structured_score": float(structured),
            "concept_score": float(concept), "combined_score": float(combined),
        })

    score_df = pd.DataFrame(records).sort_values("combined_score", ascending=False)
    score_df = score_df[score_df["combined_score"] >= min_combined_score]
    if score_df.empty:
        return pd.DataFrame()

    top = score_df.head(k)
    results = movies.iloc[top["_idx"].astype(int).tolist()].copy().reset_index(drop=True)
    for col in ["semantic_score", "tfidf_score", "structured_score", "concept_score", "combined_score"]:
        results[col] = top[col].to_numpy()
    return results


def build_context(results, max_plot_words=180):
    context = ""
    for _, row in results.iterrows():
        plot = str(row.get("plot", ""))
        words = plot.split()
        short_plot = " ".join(words[:max_plot_words])
        if len(words) > max_plot_words:
            short_plot += "..."
        context += f"""
MOVIE: {row['movie_title']} ({row.get('year', 'N/A')})
GENRE: {row.get('genre', '')}
MAIN CHARACTERS: {row.get('main_characters', '')}
THEMES: {row.get('themes', '')}
KEYWORDS: {row.get('keywords', '')}
PLOT: {short_plot}

"""
    return context


def detect_question_type(question):
    q = normalize_text(question)
    if re.search(r"\bwho\s+appears\b|\bwho\s+are\s+the\s+(main\s+)?characters\b|\bcharacters\s+in\b", q):
        return "character_list"
    if re.search(r"\bwho\s+is\b|\bwho\s+was\b|\btell\s+me\s+about\s+.+character", q):
        return "character"
    if re.search(r"\bwhich\s+movie\b|\bwhat\s+movie\b|\bwhich\s+film\b", q):
        return "movie_identification"
    if re.search(r"\btheme|themes\b", q):
        return "themes"
    if re.search(r"\bplot\b|\bwhat\s+happens\b|\bwhat\s+is\s+.+about\b|\bstory\b", q):
        return "plot"
    return "general"


def build_prompt(question, context):
    qtype = detect_question_type(question)
    if qtype == "character":
        task = """
Identify the character from the grounded context and give a detailed character explanation.
Start with the character's name and identity. Then explain their background, role in the movie,
important transformation or journey, abilities or powers ONLY when supported by the context,
and the main lesson/theme connected to them. Do not invent facts that are absent from the context.
"""
    elif qtype == "character_list":
        task = """
List the main characters from the requested movie. Give each character's name and a short role or
trait supported by the context. Do not add characters that are not listed or described in the context.
"""
    elif qtype == "movie_identification":
        task = """
Identify the best matching movie clearly in the first sentence. Then explain why it matches the
question using the movie's characters, keywords, themes, and plot. Give a useful plot/background
summary rather than only the title. Do not choose a different movie merely because it is semantically similar.
"""
    elif qtype == "themes":
        task = "Give the movie's main themes and explain each theme briefly using evidence from the plot."
    elif qtype == "plot":
        task = """
Give a clear, detailed plot summary based only on the retrieved context. Explain the main character,
central conflict, major development, and outcome or lesson when the context supports it.
"""
    else:
        task = """
Answer the question comprehensively using the retrieved movie context. Include the movie title/year,
relevant characters, background, plot, themes, and important facts when supported. Prefer a detailed
answer over a one-line answer, but never invent information that is not in the context.
"""
    return f"""You are a grounded movie question-answering assistant.

IMPORTANT RULES:
1. Use ONLY the movie information provided in the Context below.
2. Do NOT use outside movie knowledge to fill missing details.
3. Do NOT guess.
4. Answer in complete, natural sentences.
5. The user wants a useful detailed answer, not a one-word or one-name answer.
6. If the context does not support a requested detail, explicitly say that the dataset does not provide that detail.

QUESTION TYPE: {qtype}

TASK:
{task}

CONTEXT:
{context}

USER QUESTION:
{question}

DETAILED ANSWER:
"""


def grounded_fallback(question, docs):
    if docs.empty:
        return "I don't have a sufficiently relevant movie in this dataset to answer that question reliably."
    row = docs.iloc[0]
    title, year = row["movie_title"], row.get("year", "N/A")
    chars = str(row.get("main_characters", ""))
    themes = str(row.get("themes", ""))
    keywords = str(row.get("keywords", ""))
    plot = str(row.get("plot", ""))
    qtype = detect_question_type(question)

    if qtype == "character":
        q_tokens = token_set(question)
        char = None
        for c in [x.strip() for x in chars.split(",")]:
            if c and (c.lower() in normalize_text(question) or any(p in q_tokens for p in tokens(c))):
                char = c
                break
        char = char or chars.split(",")[0].strip()
        return f"{char} is a main character in {title} ({year}). {plot} The movie's relevant themes are {themes}."
    if qtype == "character_list":
        return f"The main characters listed for {title} ({year}) are {chars}. {plot}"
    if qtype == "movie_identification":
        return f"The best matching movie is {title} ({year}). It matches through these keywords: {keywords}. Main characters: {chars}. Plot: {plot}"
    if qtype == "themes":
        return f"The main themes of {title} ({year}) are {themes}. {plot}"
    return f"{title} ({year}) — {plot} The main themes are {themes}, and the key characters are {chars}."


def generate_answer(prompt, max_new_tokens=220):
    tok, model = load_generator()
    inputs = tok(prompt, return_tensors="pt", truncation=True, max_length=512)
    output_ids = model.generate(
        **inputs, max_new_tokens=max_new_tokens, num_beams=4, do_sample=False,
        no_repeat_ngram_size=3, repetition_penalty=1.12, length_penalty=1.05, early_stopping=True,
    )
    return tok.decode(output_ids[0], skip_special_tokens=True).strip()


def answer_movie_question(question):
    docs = retrieve_movies(question, k=3)
    if docs.empty:
        return "I don't have a sufficiently relevant movie in this dataset to answer that question reliably.", docs
    qtype = detect_question_type(question)
    context_docs = docs.head(1) if qtype in {"character", "character_list", "movie_identification"} else docs.head(2)
    context = build_context(context_docs)
    prompt = build_prompt(question, context)
    response = generate_answer(prompt, max_new_tokens=220).strip()
    weak = (
        len(response.split()) < 35
        or response.lower() in {"stephen strange", "anna, elsa, kristoff, olaf"}
        or response.lower().startswith("i don't know")
    )
    if weak:
        response = grounded_fallback(question, context_docs)
    return response, docs


# ─────────────────────────────────────────────────────────────────────────
# HERO BANNER — featured movie, backdrop image, real synopsis
# ─────────────────────────────────────────────────────────────────────────

n_genres = movies["genre"].nunique()
year_min, year_max = int(movies["year"].min()), int(movies["year"].max())

st.markdown(f"""
<div style="text-align:center; padding: 1.6rem 0 0.4rem;">
    <div class="cv-logo" style="font-size:2.6rem;">CINEVECTOR</div>
    <div class="cv-tagline" style="font-size:0.72rem;">Movie Intelligence Engine</div>
    <div class="cv-chip-row" style="margin-top:1rem; gap:1.8rem;">
        <div class="cv-chip"><div class="cv-chip-value" style="font-size:1.4rem;">{len(movies)}</div><div class="cv-chip-label">Movies</div></div>
        <div class="cv-chip"><div class="cv-chip-value" style="font-size:1.4rem;">{n_genres}</div><div class="cv-chip-label">Genres</div></div>
        <div class="cv-chip"><div class="cv-chip-value" style="font-size:1.4rem;">{year_min}–{year_max}</div><div class="cv-chip-label">Years</div></div>
        <div class="cv-chip"><div class="cv-chip-value" style="font-size:1.4rem;">{dimension}D</div><div class="cv-chip-label">Embeddings</div></div>
        <div class="cv-chip"><div class="cv-chip-value" style="font-size:1.4rem;">5</div><div class="cv-chip-label">FAISS Indexes</div></div>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("Use official posters (optional)"):
    st.caption(
        "By default, cards show a real stock photo (no signup needed, but not movie-related) with a "
        "generated title card as a safety net if it can't load. To show each movie's actual official "
        "poster instead, add a free key from **either** service below (TMDB is tried first if both are set)."
    )
    oc1, oc2 = st.columns(2)
    with oc1:
        st.markdown("**OMDb** — easiest: just an email address")
        st.caption("Free key at [omdbapi.com/apikey.aspx](https://www.omdbapi.com/apikey.aspx) — emailed to you, no address/billing form.")
        st.text_input("OMDb API key", key="omdb_api_key", type="password", label_visibility="collapsed")
    with oc2:
        st.markdown("**TMDB** — larger catalog, more detail")
        st.caption("Free key at [themoviedb.org/settings/api](https://www.themoviedb.org/settings/api) — requires filling out a short application form.")
        st.text_input("TMDB API key", key="tmdb_api_key", type="password", label_visibility="collapsed")
    st.caption(
        "To make a key permanent (no retyping, works for every visitor), don't use the boxes above on a "
        "deployed app — go to **share.streamlit.io → your app → ⋮ → Settings → Secrets** and paste:\n\n"
        "```\nOMDB_API_KEY = \"your-actual-key-here\"\n```\n\n"
        "then **Reboot app**. The boxes above are only for quick local testing."
    )

    st.markdown('<div class="cv-divider" style="margin:1rem 0;"></div>', unsafe_allow_html=True)
    st.markdown("**Diagnose why posters aren't showing**")
    diag_title = st.text_input("Test title", value="Iron Man", key="diag_title")
    if st.button("Run diagnostic", key="diag_run"):
        omdb_key = _omdb_key()
        tmdb_key = _tmdb_key()
        st.write(f"OMDb key detected: `{'yes — ' + omdb_key[:3] + '…' + omdb_key[-2:] if omdb_key else 'NO KEY FOUND'}`")
        st.write(f"TMDB key detected: `{'yes — ' + tmdb_key[:3] + '…' + tmdb_key[-2:] if tmdb_key else 'NO KEY FOUND'}`")
        if not omdb_key and not tmdb_key:
            st.error(
                "Neither key was found in session_state OR st.secrets. That means the secret either "
                "isn't saved on Streamlit Cloud, is misspelled, or the app hasn't rebooted since you saved it."
            )
        if omdb_key:
            try:
                raw = requests.get(
                    "https://www.omdbapi.com/",
                    params={"apikey": omdb_key, "t": diag_title, "type": "movie"},
                    timeout=6,
                )
                st.write(f"OMDb HTTP status: `{raw.status_code}`")
                st.json(raw.json())
            except Exception as e:
                st.error(f"OMDb request raised an exception: {e}")
        if tmdb_key:
            try:
                raw = requests.get(
                    "https://api.themoviedb.org/3/search/movie",
                    params={"api_key": tmdb_key, "query": diag_title},
                    timeout=6,
                )
                st.write(f"TMDB HTTP status: `{raw.status_code}`")
                st.json(raw.json())
            except Exception as e:
                st.error(f"TMDB request raised an exception: {e}")


@st.dialog("Movie Details", width="large")
def show_movie_modal(row):
    c1, c2 = st.columns([1, 2])
    with c1:
        st.markdown(
            poster_img_tag(row["movie_title"], row["year"], row["genre"], style="width:100%;border-radius:6px;"),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(f"### {row['movie_title']} ({row['year']})")
        st.caption(row.get("genre", ""))
        st.markdown(f"**Characters:** {row.get('main_characters','')}")
        st.markdown(f"**Themes:** {row.get('themes','')}")
        st.markdown(f"**Keywords:** {row.get('keywords','')}")
        st.markdown("**Synopsis**")
        st.write(row.get("plot", ""))


if HAS_MOVIES:
    if "hero_movie_id" not in st.session_state:
        st.session_state.hero_movie_id = movies.iloc[0]["movie_id"]
    hero_row = movies[movies["movie_id"] == st.session_state.hero_movie_id]
    hero_row = hero_row.iloc[0] if not hero_row.empty else movies.iloc[0]

    backdrop = get_backdrop(hero_row["movie_title"], hero_row["year"], hero_row["genre"])
    synopsis_preview = str(hero_row.get("plot", ""))
    if len(synopsis_preview) > 260:
        synopsis_preview = synopsis_preview[:260].rsplit(" ", 1)[0] + "..."

    st.markdown(f"""
    <div style="position:relative; border-radius:10px; overflow:hidden; height:340px;
                background-image:url('{backdrop}'); background-size:cover; background-position:center;
                margin: 1rem 0 1.4rem; display:flex; align-items:flex-end;">
        <div style="padding:2rem; max-width:600px;">
            <div style="font-size:0.7rem; letter-spacing:0.2em; color:#ff6b6f; text-transform:uppercase; margin-bottom:0.5rem;">
                Featured &bull; {hero_row.get('genre','')}
            </div>
            <div style="font-family:'Bebas Neue',sans-serif; font-size:2.6rem; color:#F5F5F1; line-height:1;">
                {hero_row['movie_title']}
            </div>
            <div style="font-size:0.78rem; color:#b3b3b3; margin:0.6rem 0 1rem;">{hero_row['year']} &bull; {hero_row.get('main_characters','')}</div>
            <div style="font-size:0.92rem; color:#e8e8e8; line-height:1.5;">{synopsis_preview}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    hc1, hc2, _ = st.columns([1, 1, 3])
    with hc1:
        if st.button("ℹ️ More Info", key="hero_more_info"):
            show_movie_modal(hero_row)
    with hc2:
        if st.button("🔀 Shuffle Featured", key="hero_shuffle"):
            candidates = movies[movies["movie_id"] != hero_row["movie_id"]]
            pick = candidates.sample(1).iloc[0] if not candidates.empty else hero_row
            st.session_state.hero_movie_id = pick["movie_id"]
            st.rerun()

    # ─────────────────────────────────────────────────────────────────
    # REAL INTERACTIVE CAROUSEL — hover-reveal arrows, autoplay toggle,
    # smooth native scroll. Genuinely interactive (unlike a CSS marquee):
    # buttons call scrollBy() directly, autoplay is a real setInterval
    # loop that resets at the end, same behavior as a native Netflix row.
    # ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="cv-marquee-label">TRENDING IN THIS COLLECTION</div>', unsafe_allow_html=True)

    _carousel_items = "".join(
        f'''<div class="tile">
            <span class="badge">CV</span>
            {poster_img_tag(row['movie_title'], row['year'], row['genre'])}
            <div class="cap">{row['movie_title']}<br/><span>{row['year']}</span></div>
        </div>'''
        for _, row in movies.iterrows()
    )

    carousel_html = f"""
    <style>
      body {{ margin:0; background:transparent; font-family:'Inter',sans-serif; }}
      .wrap {{ position:relative; }}
      .track {{
        display:flex; gap:14px; overflow-x:auto; scroll-behavior:smooth;
        padding-bottom:6px; scrollbar-width:none;
      }}
      .track::-webkit-scrollbar {{ display:none; }}
      .tile {{
        position:relative; width:150px; height:225px; flex:0 0 auto;
        border-radius:6px; overflow:hidden; border:1px solid rgba(229,9,20,0.28);
        box-shadow: 0 8px 20px rgba(0,0,0,0.55);
        transition: transform 0.2s ease, border-color 0.2s ease;
      }}
      .tile:hover {{ transform: scale(1.06); border-color:#E50914; z-index:2; }}
      .tile img {{ width:100%; height:100%; object-fit:cover; display:block; transition: transform 0.5s ease; }}
      .tile:hover img {{ transform: scale(1.08); }}
      .badge {{
        position:absolute; top:6px; left:6px; z-index:3;
        background:#E50914; color:#F5F5F1; font-size:0.6rem; font-weight:800;
        letter-spacing:0.03em; padding:2px 6px; border-radius:3px;
        box-shadow:0 2px 6px rgba(0,0,0,0.5);
      }}
      .cap {{
        position:absolute; left:0; right:0; bottom:0;
        background:linear-gradient(180deg, transparent, rgba(0,0,0,0.9) 65%);
        padding:28px 8px 8px; font-size:0.78rem; color:#F5F5F1; font-weight:800; line-height:1.25;
      }}
      .cap span {{ opacity:0.6; font-weight:400; }}
      .arrow {{
        position:absolute; top:0; bottom:0; width:44px; display:flex; align-items:center; justify-content:center;
        background:linear-gradient(90deg, #141414, transparent); color:#fff; cursor:pointer;
        opacity:0; transition:opacity 0.2s ease; font-size:22px; z-index:3; border:none;
      }}
      .arrow.right {{ right:0; background:linear-gradient(270deg, #141414, transparent); }}
      .wrap:hover .arrow {{ opacity:1; }}
      .autoplay-btn {{
        margin-top:8px; background:#232323; border:1px solid #2f2f2f; color:#F5F5F1;
        font-size:0.72rem; letter-spacing:0.08em; text-transform:uppercase; padding:5px 12px;
        border-radius:4px; cursor:pointer;
      }}
      .autoplay-btn:hover {{ border-color:#E50914; }}
    </style>
    <div class="wrap">
      <div class="track" id="track">{_carousel_items}</div>
      <button class="arrow left" onclick="document.getElementById('track').scrollBy({{left:-320,behavior:'smooth'}})">&#10094;</button>
      <button class="arrow right" onclick="document.getElementById('track').scrollBy({{left:320,behavior:'smooth'}})">&#10095;</button>
    </div>
    <button class="autoplay-btn" id="autoplayBtn" onclick="toggleAutoplay()">▶ Autoplay</button>
    <script>
      let autoplayInterval = null;
      function toggleAutoplay() {{
        const btn = document.getElementById('autoplayBtn');
        const track = document.getElementById('track');
        if (autoplayInterval) {{
          clearInterval(autoplayInterval);
          autoplayInterval = null;
          btn.textContent = '▶ Autoplay';
        }} else {{
          btn.textContent = '⏸ Pause';
          autoplayInterval = setInterval(() => {{
            if (track.scrollLeft + track.clientWidth >= track.scrollWidth - 20) {{
              track.scrollTo({{left:0, behavior:'smooth'}});
            }} else {{
              track.scrollBy({{left:320, behavior:'smooth'}});
            }}
          }}, 3000);
        }}
      }}
    </script>
    """
    components.html(carousel_html, height=290)

    with st.expander("More info on a movie from the row above"):
        pick_title = st.selectbox("Movie", movies["movie_title"].tolist(), key="carousel_pick")
        if st.button("ℹ️ More Info", key="carousel_more_info"):
            show_movie_modal(movies[movies["movie_title"] == pick_title].iloc[0])

_MODULE_META = [
    ("EXPLORE", "The dataset, at a glance"),
    ("VECTOR OBSERVATORY", "The embedding space, visualized"),
    ("INDEX LAB", "FAISS index types, compared"),
    ("DATA STUDIO", "SQL, semantic, and hybrid queries"),
    ("EVIDENCE ENGINE", "Ask a question, get a grounded answer"),
    ("MANAGE DATA", "Upload, edit, and export the catalog"),
]
_jump_labels = [f"{i+1:02d} · {m[0]}" for i, m in enumerate(_MODULE_META)]

if "cv_module_idx" not in st.session_state:
    st.session_state.cv_module_idx = 0
if "cv_jump" not in st.session_state:
    st.session_state.cv_jump = _jump_labels[0]

def _cv_go(delta):
    new_idx = (st.session_state.cv_module_idx + delta) % len(_MODULE_META)
    st.session_state.cv_module_idx = new_idx
    st.session_state.cv_jump = _jump_labels[new_idx]  # keep the dropdown in sync

def _cv_jump_changed():
    st.session_state.cv_module_idx = _jump_labels.index(st.session_state.cv_jump)

nav_l, nav_c, nav_r = st.columns([1, 6, 1])
with nav_l:
    st.button("◀", key="cv_prev", use_container_width=True, on_click=_cv_go, args=(-1,))
with nav_c:
    _idx = st.session_state.cv_module_idx
    _name, _desc = _MODULE_META[_idx]
    st.markdown(f"""
    <div class="cv-channel">
        <div class="cv-channel-num">{_idx+1:02d} / {len(_MODULE_META):02d}</div>
        <div class="cv-channel-name">{_name}</div>
        <div class="cv-channel-desc">{_desc}</div>
    </div>
    """, unsafe_allow_html=True)
with nav_r:
    st.button("▶", key="cv_next", use_container_width=True, on_click=_cv_go, args=(1,))

st.selectbox(
    "Jump to a module", _jump_labels,
    key="cv_jump", label_visibility="collapsed", on_change=_cv_jump_changed,
)

_view = st.session_state.cv_module_idx

# ─────────────────────────────────────────────────────────────────────────
# TAB 1 — EXPLORE
# ─────────────────────────────────────────────────────────────────────────

if _view == 0:
    st.markdown('<div class="cv-section-tag">DATASET</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Explore the Collection</div>', unsafe_allow_html=True)

    # ── Signature visual: release timeline, one marker per movie ───────
    st.markdown('<div class="cv-mono-label">RELEASE TIMELINE</div>', unsafe_allow_html=True)
    movies_sorted = movies.sort_values("year", kind="stable").reset_index(drop=True)
    n_movies = len(movies_sorted)
    # graded red per movie — darkest (Genre Distribution's tallest-bar shade) for the
    # oldest release, fading to the lightest salmon for the newest, same Reds ramp
    red_shades = px.colors.sample_colorscale(
        px.colors.sequential.Reds[::-1], [i / max(n_movies - 1, 1) for i in range(n_movies)]
    )
    st.markdown('<div class="cv-filmstrip"><div class="cv-sprockets">'
                + "".join(["<span></span>"] * 26) + "</div>", unsafe_allow_html=True)
    timeline = go.Figure()
    for i, row in movies_sorted.iterrows():
        timeline.add_trace(go.Scatter(
            x=[row["year"]], y=[i],
            mode="markers+text",
            marker=dict(size=13, color=red_shades[i], line=dict(width=1, color=NETFLIX_DARK)),
            text=[f"  {row['movie_title']}"], textposition="middle right",
            textfont=dict(color=NETFLIX_WHITE, size=11),
            hovertemplate=f"{row['movie_title']} — {row['year']}<extra></extra>",
            showlegend=False,
        ))
        timeline.add_shape(
            type="line", x0=movies_sorted["year"].min(), x1=row["year"], y0=i, y1=i,
            line=dict(color=red_shades[i], width=1.5),
        )
    timeline.update_layout(
        plot_bgcolor=NETFLIX_DARK, paper_bgcolor=NETFLIX_DARK, font_color=NETFLIX_WHITE,
        height=max(220, n_movies * 26), margin=dict(t=15, b=30, l=10, r=160),
        xaxis=dict(showgrid=False, tickfont=dict(size=10, color=NETFLIX_GRAY), title="Release Year"),
        yaxis=dict(visible=False, autorange="reversed"),
    )
    st.plotly_chart(timeline, use_container_width=True, config={"displayModeBar": False})
    st.markdown('<div class="cv-sprockets">' + "".join(["<span></span>"] * 26)
                + "</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height:1.4rem'></div>", unsafe_allow_html=True)

    st.markdown('<div class="cv-mono-label">COLLECTION BREAKDOWN</div>', unsafe_allow_html=True)
    genre_counts = movies["genre"].value_counts()
    max_count = int(genre_counts.max())
    n_genre_rows = len(genre_counts)
    genre_shades = px.colors.sample_colorscale(
        px.colors.sequential.Reds[::-1], [i / max(n_genre_rows - 1, 1) for i in range(n_genre_rows)]
    )
    rows_html = ""
    for i, (genre, count) in enumerate(genre_counts.items()):
        pct = (count / max_count) * 100
        rows_html += (
            f'<div class="cv-rank-row">'
            f'<div class="cv-rank-num">{i+1:02d}</div>'
            f'<div class="cv-rank-label">{genre}</div>'
            f'<div class="cv-rank-track"><div class="cv-rank-fill" '
            f'style="width:{pct}%; background:{genre_shades[i]};"></div></div>'
            f'<div class="cv-rank-count">{count}</div>'
            f'</div>'
        )

    plot_lengths = movies["plot"].str.split().apply(len)
    counts, edges = np.histogram(plot_lengths, bins=10)
    max_c = max(counts.max(), 1)
    bars_html = "".join(
        f'<div style="height:{max(6, int(c / max_c * 100))}%;" title="{c} movie(s)"></div>'
        for c in counts
    )

    breakdown_html = (
        f'<div class="cv-breakdown-panel">'
        f'<div class="cv-breakdown-main">{rows_html}</div>'
        f'<div class="cv-breakdown-rail">'
        f'<div class="cv-wc-big">{int(plot_lengths.mean())}</div>'
        f'<div class="cv-wc-range">AVG WORDS / SYNOPSIS</div>'
        f'<div class="cv-wc-bars">{bars_html}</div>'
        f'<div class="cv-wc-range" style="margin-top:0.4rem;">'
        f'RANGE {plot_lengths.min()}–{plot_lengths.max()}</div>'
        f'</div>'
        f'</div>'
    )
    st.markdown(breakdown_html, unsafe_allow_html=True)

    st.markdown('<div class="cv-mono-label" style="margin-top:1.5rem;">BROWSE THE CATALOG</div>', unsafe_allow_html=True)
    poster_cols = st.columns(6)
    for i, (_, row) in enumerate(movies.iterrows()):
        with poster_cols[i % 6]:
            poster_tag = poster_img_tag(row['movie_title'], row['year'], row['genre'])
            st.markdown(f"""
            <div class="cv-poster-tile" style="width:100%; height:150px; margin-bottom:0.4rem;">
                <span class="cv-poster-badge">CV</span>
                {poster_tag}
                <div class="cv-poster-caption">{row['movie_title']}</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Info", key=f"catalog_info_{row['movie_id']}", use_container_width=True):
                show_movie_modal(row)

    st.markdown('<div class="cv-mono-label" style="margin-top:1rem;">DATASET PREVIEW</div>', unsafe_allow_html=True)
    st.dataframe(
        movies[["movie_title", "year", "genre", "main_characters", "themes"]],
        use_container_width=True, height=320,
    )

# ─────────────────────────────────────────────────────────────────────────
# TAB 2 — VECTOR OBSERVATORY
# ─────────────────────────────────────────────────────────────────────────

if _view == 1:
    st.markdown('<div class="cv-section-tag">EMBEDDING SPACE</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Vector Observatory</div>', unsafe_allow_html=True)

    from sklearn.decomposition import PCA
    pca = PCA(n_components=2)
    coords = pca.fit_transform(embeddings)
    plot_df = movies.copy()
    plot_df["pca_x"] = coords[:, 0]
    plot_df["pca_y"] = coords[:, 1]

    fig4 = px.scatter(
        plot_df, x="pca_x", y="pca_y", color="genre", hover_name="movie_title",
        hover_data={"year": True, "pca_x": False, "pca_y": False},
        title="Movie Collection in 2D Embedding Space (PCA)",
        color_discrete_sequence=px.colors.sequential.Reds_r,
    )
    fig4.update_traces(marker=dict(size=14, line=dict(width=1, color=NETFLIX_BLACK)))
    fig4.update_layout(
        plot_bgcolor=NETFLIX_DARK, paper_bgcolor=NETFLIX_BLACK, font_color=NETFLIX_WHITE,
        height=560, title_font_size=16,
        xaxis_title="PCA Component 1", yaxis_title="PCA Component 2",
    )
    st.plotly_chart(fig4, use_container_width=True)

    norms = np.linalg.norm(embeddings, axis=1)
    fig5 = px.histogram(
        x=norms, nbins=10, title="Embedding Norm Distribution",
        color_discrete_sequence=[NETFLIX_RED],
    )
    fig5.update_layout(
        plot_bgcolor=NETFLIX_BLACK, paper_bgcolor=NETFLIX_BLACK, font_color=NETFLIX_WHITE,
        height=260, title_font_size=15,
    )
    st.plotly_chart(fig5, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────
# TAB 3 — INDEX LAB
# ─────────────────────────────────────────────────────────────────────────

if _view == 2:
    st.markdown('<div class="cv-section-tag">FAISS EXPERIMENTS</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Index Lab</div>', unsafe_allow_html=True)

    index_info = {
        "IndexFlatL2": ("EXACT", "Brute-force ground truth — scans every vector."),
        "IndexIVFFlat": ("CLUSTERED", "Partitions vectors into clusters, searches nearby ones only."),
        "IndexPQ": ("COMPRESSED", "Compresses vectors into compact codes for smaller memory."),
        "IndexIVFPQ": ("HYBRID", "Combines clustering with compression."),
        "IndexHNSWFlat": ("GRAPH", "Navigates a multi-layer proximity graph — no training needed."),
    }

    cols = st.columns(5)
    for col, (name, (tag, desc)) in zip(cols, index_info.items()):
        with col:
            st.markdown(f"""
            <div class="cv-index-card">
                <div class="cv-index-tag">{tag}</div>
                <div class="cv-index-name">{name.replace('Index','').replace('Flat','')}</div>
                <div style="font-size:0.78rem;color:#b3b3b3;margin-top:0.6rem;">{desc}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown('<div class="cv-divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-tag">RETRIEVAL ARENA</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Compare All Five Indexes on One Query</div>', unsafe_allow_html=True)

    arena_query = st.text_input("Query for the arena", value="Spider-Man fights Green Goblin")
    run_arena = st.button("Run Arena", key="arena_btn")

    if run_arena and arena_query.strip():
        embedder = load_embedder()
        qvec = embedder.encode([arena_query], convert_to_numpy=True, normalize_embeddings=True)

        flat_ids = indexes["IndexFlatL2"].search(qvec, 5)[1][0]
        ground_truth = set(flat_ids.tolist())

        rows = []
        for name, idx in indexes.items():
            start = time.time()
            distances, ids = idx.search(qvec, 5)
            latency_ms = (time.time() - start) * 1000
            retrieved = set(ids[0].tolist())
            recall = len(retrieved & ground_truth) / max(1, len(ground_truth))
            rows.append({"Index": name, "Latency (ms)": round(latency_ms, 3), "Recall@5 vs Flat": round(recall * 100, 1)})

        arena_df = pd.DataFrame(rows)

        cA, cB = st.columns(2)
        with cA:
            figA = px.bar(
                arena_df, x="Index", y="Recall@5 vs Flat", title="Recall@5 (vs exact IndexFlatL2)",
                color_discrete_sequence=[NETFLIX_RED],
            )
            figA.update_layout(plot_bgcolor=NETFLIX_BLACK, paper_bgcolor=NETFLIX_BLACK, font_color=NETFLIX_WHITE, title_font_size=14)
            st.plotly_chart(figA, use_container_width=True)
        with cB:
            figB = px.bar(
                arena_df, x="Index", y="Latency (ms)", title="Search Latency",
                color_discrete_sequence=["#b3b3b3"],
            )
            figB.update_layout(plot_bgcolor=NETFLIX_BLACK, paper_bgcolor=NETFLIX_BLACK, font_color=NETFLIX_WHITE, title_font_size=14)
            st.plotly_chart(figB, use_container_width=True)

        st.dataframe(arena_df, use_container_width=True)
        st.caption(
            "With only 20 movies, latency differences are in the millisecond/microsecond range and "
            "will vary run to run — this is a small-dataset teaching experiment, not a production benchmark."
        )

# ─────────────────────────────────────────────────────────────────────────
# TAB 4 — DATA STUDIO (SQL + Semantic + Hybrid)
# ─────────────────────────────────────────────────────────────────────────

if _view == 3:
    st.markdown('<div class="cv-section-tag">STRUCTURED + SEMANTIC</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Data Studio</div>', unsafe_allow_html=True)

    @st.cache_resource(show_spinner=False)
    def get_sqlite_conn(_movies_df, cache_key):
        conn = sqlite3.connect(":memory:", check_same_thread=False)
        _movies_df.drop(columns=["text"], errors="ignore").to_sql("movies", conn, index=False, if_exists="replace")
        return conn

    conn = get_sqlite_conn(movies, dataset_fingerprint(movies))

    mode = st.radio("Query mode", ["Structured (SQL)", "Semantic (Natural Language)", "Hybrid"], horizontal=True)

    if mode == "Structured (SQL)":
        default_sql = "SELECT movie_title, year, genre FROM movies WHERE year >= 2010 ORDER BY year;"
        sql = st.text_area("SQL query against the `movies` table", value=default_sql, height=90)
        if st.button("Run SQL"):
            try:
                result = pd.read_sql_query(sql, conn)
                st.dataframe(result, use_container_width=True)
            except Exception as e:
                st.error(f"SQL error: {e}")

    elif mode == "Semantic (Natural Language)":
        nl_query = st.text_input("Describe what you're looking for", value="a hero learning responsibility")
        if st.button("Search"):
            docs = retrieve_movies(nl_query, k=5)
            if docs.empty:
                st.warning("Nothing in the dataset scored above the relevance threshold for this query.")
            else:
                for _, row in docs.iterrows():
                    poster_tag = poster_img_tag(row['movie_title'], row.get('year',''), row.get('genre',''))
                    st.markdown(f"""
                    <div class="cv-row-card">
                        {poster_tag}
                        <div>
                            <div class="cv-card-title">{row['movie_title']}</div>
                            <div class="cv-card-meta">{row.get('year','')} • {row.get('genre','')}</div>
                            <div class="cv-card-plot">{str(row.get('plot',''))[:220]}...</div>
                            <div class="cv-badge">RELEVANCE {row['combined_score']*100:.1f}%</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

    else:  # Hybrid
        c1, c2 = st.columns(2)
        with c1:
            genre_filter = st.selectbox("Genre filter (SQL)", ["Any"] + sorted(movies["genre"].unique().tolist()))
        with c2:
            year_filter = st.slider("Minimum year (SQL)", year_min, year_max, year_min)
        hybrid_query = st.text_input("Semantic query (FAISS)", value="responsibility and sacrifice")

        if st.button("Run Hybrid Query"):
            sql = "SELECT movie_title FROM movies WHERE year >= ?"
            params = [year_filter]
            if genre_filter != "Any":
                sql += " AND genre = ?"
                params.append(genre_filter)
            sql_matches = set(pd.read_sql_query(sql, conn, params=params)["movie_title"].tolist())

            sem_docs = retrieve_movies(hybrid_query, k=len(movies))
            sem_docs = sem_docs[sem_docs["movie_title"].isin(sql_matches)]

            if sem_docs.empty:
                st.warning("No movies satisfy both the SQL filter and the semantic query.")
            else:
                st.caption(f"SQL filtered candidates: {len(sql_matches)}  →  after semantic ranking: {len(sem_docs)}")
                for _, row in sem_docs.head(5).iterrows():
                    poster_tag = poster_img_tag(row['movie_title'], row.get('year',''), row.get('genre',''))
                    st.markdown(f"""
                    <div class="cv-row-card">
                        {poster_tag}
                        <div>
                            <div class="cv-card-title">{row['movie_title']}</div>
                            <div class="cv-card-meta">{row.get('year','')} • {row.get('genre','')}</div>
                            <div class="cv-card-plot">{str(row.get('plot',''))[:220]}...</div>
                            <div class="cv-badge">RELEVANCE {row['combined_score']*100:.1f}%</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────
# TAB 5 — EVIDENCE ENGINE (RAG)
# ─────────────────────────────────────────────────────────────────────────

if _view == 4:
    st.markdown('<div class="cv-section-tag">RETRIEVAL-AUGMENTED GENERATION</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Evidence Engine</div>', unsafe_allow_html=True)

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    question = st.text_input("Ask a movie question", placeholder="Who is Doctor Strange?", key="rag_q")
    ask = st.button("Ask")

    if ask and question.strip():
        with st.spinner("Retrieving evidence and generating an answer..."):
            answer, docs = answer_movie_question(question)
        st.session_state.chat_history.insert(0, (question, answer, docs))

    for q, a, docs in st.session_state.chat_history:
        st.markdown(f"""
        <div class="cv-card">
            <div class="cv-card-meta">QUESTION</div>
            <div class="cv-card-title">{q}</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="cv-card" style="border-left:3px solid {NETFLIX_RED};">
            <div class="cv-card-meta">GENERATED ANSWER</div>
            <div class="cv-card-plot" style="font-size:0.95rem;color:{NETFLIX_WHITE};">{a}</div>
        </div>
        """, unsafe_allow_html=True)

        if not docs.empty:
            with st.expander(f"GROUNDED IN {len(docs)} RETRIEVED RECORD(S) — view evidence"):
                for _, row in docs.iterrows():
                    score = row.get("combined_score", None)
                    score_txt = f"{score*100:.1f}%" if score is not None else "n/a"
                    poster_tag = poster_img_tag(row['movie_title'], row.get('year',''), row.get('genre',''))
                    st.markdown(f"""
                    <div class="cv-row-card">
                        {poster_tag}
                        <div>
                            <div class="cv-card-title">{row['movie_title']} ({row.get('year','')})</div>
                            <div class="cv-card-meta">RELEVANCE {score_txt}</div>
                            <div class="cv-card-plot">{str(row.get('plot',''))[:300]}...</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        st.markdown('<div class="cv-divider"></div>', unsafe_allow_html=True)

    if not st.session_state.chat_history:
        st.caption(
            "Example questions: \"Who is Doctor Strange?\" · \"Which movie is about a boy who loves music?\" "
            "· \"Who appears in Frozen?\" · \"What are the themes of Avengers: Infinity War?\""
        )

# ─────────────────────────────────────────────────────────────────────────
# TAB 6 — MANAGE DATA (upload, inline add/edit/delete, export)
# ─────────────────────────────────────────────────────────────────────────

if _view == 5:
    st.markdown('<div class="cv-section-tag">DATASET CONTROL</div>', unsafe_allow_html=True)
    st.markdown('<div class="cv-section-title">Manage Data</div>', unsafe_allow_html=True)
    st.caption(
        "Changes here rebuild embeddings, FAISS indexes, and TF-IDF the next time any "
        "tab needs them — that can take a moment on a large dataset, instant on this one."
    )

    st.markdown('<div class="cv-mono-label" style="margin-top:0.5rem;">UPLOAD A DIFFERENT DATASET</div>', unsafe_allow_html=True)
    upload_mode = st.radio(
        "Upload mode", ["Replace current dataset", "Append to current dataset"],
        horizontal=True, label_visibility="collapsed",
    )
    uploaded_file = st.file_uploader(
        "Upload .xlsx or .csv", type=["xlsx", "csv"],
        help=f"Expected columns: {', '.join(REQUIRED_COLUMNS)}. Missing columns are added empty.",
    )
    if uploaded_file is not None and st.button("Apply upload", key="apply_upload"):
        try:
            new_df = load_dataframe_from_upload(uploaded_file)
            if upload_mode.startswith("Replace"):
                st.session_state.movies = new_df
            else:
                combined = pd.concat(
                    [st.session_state.movies.drop(columns=["text"]), new_df.drop(columns=["text"])],
                    ignore_index=True,
                )
                st.session_state.movies = prepare_dataframe(combined)
            st.success(f"Dataset updated — now {len(st.session_state.movies)} movies.")
            st.rerun()
        except Exception as e:
            st.error(f"Couldn't read that file: {e}")

    st.markdown('<div class="cv-divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="cv-mono-label">EDIT INLINE — ADD, EDIT, OR DELETE ROWS</div>', unsafe_allow_html=True)
    st.caption("Use the trash icon on a row to delete it, or the blank row at the bottom to add a new movie. Click Save when done.")

    editable_view = movies.drop(columns=["text"]) if HAS_MOVIES else pd.DataFrame(columns=REQUIRED_COLUMNS)
    edited_df = st.data_editor(
        editable_view,
        num_rows="dynamic",
        use_container_width=True,
        height=420,
        key="movies_data_editor",
        column_config={
            "movie_id": st.column_config.TextColumn("ID", width="small"),
            "movie_title": st.column_config.TextColumn("Title", width="medium"),
            "year": st.column_config.NumberColumn("Year", width="small", format="%d"),
        },
    )

    if st.button("Save table changes", key="save_table"):
        st.session_state.movies = prepare_dataframe(edited_df)
        st.success(f"Saved — {len(st.session_state.movies)} movies in the dataset now.")
        st.rerun()

    st.markdown('<div class="cv-divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="cv-mono-label">EXPORT CURRENT DATASET</div>', unsafe_allow_html=True)
    export_df = movies.drop(columns=["text"]) if HAS_MOVIES else pd.DataFrame(columns=REQUIRED_COLUMNS)
    ce1, ce2 = st.columns(2)
    with ce1:
        st.download_button(
            "Download as CSV",
            export_df.to_csv(index=False).encode("utf-8"),
            file_name="movies_export.csv",
            mime="text/csv",
        )
    with ce2:
        xbuf = io.BytesIO()
        export_df.to_excel(xbuf, index=False)
        st.download_button(
            "Download as Excel",
            xbuf.getvalue(),
            file_name="movies_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
