# -*- coding: utf-8 -*-
"""
Punjab Water Quality Dashboard — Version 4.0 (Single-file)
BLV Prasad

Dashboard mode: 1 parameter
Comparison mode: up to 4 parameters

Revisions (this version):
1) Default pins: size=3 and dark blue color (later changes by traffic light)
2) Company logo top-left: 'Resilience actions logo.png' in /data
3) Village boundary outline only, 90% transparent (very faint)
4) Remove sidebar text: "Loaded boundary: ... Field: ..."
5) Fix heading going above page: increase top padding + safer header layout
"""

#%% 01 — IMPORTS + PAGE CONFIG + RESPONSIVE CSS + HEADER (LOGO + TITLE)

import streamlit as st
import pandas as pd
import numpy as np
import pydeck as pdk
import matplotlib.pyplot as plt
import io
import re
import sys
import base64
from pathlib import Path

# Optional boundary overlay dependency
try:
    import geopandas as gpd
    _HAS_GPD = True
except Exception:
    gpd = None
    _HAS_GPD = False

st.set_page_config(page_title="Punjab WQ Dashboard", layout="wide")

# Responsive layout + fixed header spacing + heading font sizes + logo styling
st.markdown(
    """
<style>
/* Give enough space so title never hides behind Streamlit top bar */
.block-container { padding-top: 3.2rem !important; padding-bottom: 1rem; }

/* Heading sizes (adjust here) */
h1 { font-size: 30px !important; line-height: 1.15 !important; margin-top: 0rem !important; padding-top: 0rem !important; }
h2 { font-size: 22px !important; line-height: 1.20 !important; }
h3 { font-size: 18px !important; line-height: 1.25 !important; }

/* Auto stack columns on smaller widths */
@media (max-width: 1100px) {
  .block-container { padding-left: 0.85rem; padding-right: 0.85rem; }
  div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; }
  div[data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; }
}

/* Logo box (square, white, no rounding) */
.resl-logo-box{
  background:#fff;
  padding:8px;
  border-radius:0 !important;
  display:inline-block;
  line-height:0;
}

/* Bigger logo */
.resl-logo{
  width:140px;               /* <-- make bigger/smaller here (try 140–160) */
  height:auto;
  object-fit:contain;
  border-radius:0 !important;
  display:block;
}
</style>
""",
    unsafe_allow_html=True
)

# ---------------- PATHS (works in normal Python + PyInstaller) ----------------
def _base_dir() -> Path:
    # For bundled (PyInstaller): read bundled data from _MEIPASS
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    # For normal runs: folder containing this script
    return Path(__file__).resolve().parent

BASE_DIR = _base_dir()
DATA_DIR = BASE_DIR / "data"

def runtime_dir() -> Path:
    # For writing outputs/logs: alongside the exe (NOT inside _MEIPASS)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return BASE_DIR

RUNTIME_DIR = runtime_dir()
# -----------------------------------------------------------------------------


def draw_header():
    logo_path = DATA_DIR / "Resilience actions logo.png"

    # Give logo its own space and keep title aligned
    c1, c2 = st.columns([2.5, 9.5], vertical_alignment="center")

    with c1:
        if logo_path.exists():
            b64 = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
            st.markdown(
                f"""
                <div class="resl-logo-box">
                  <img class="resl-logo" src="data:image/png;base64,{b64}">
                </div>
                """,
                unsafe_allow_html=True
            )

    with c2:
        st.title("SRI MUKTSAR SAHIB District - Water Quality Dashboard")

draw_header()

#%% 02 — LOAD EXCEL

FILE_PATH = DATA_DIR / "villagedata2.xlsx"

if not FILE_PATH.exists():
    st.error(
        "Excel file not found.\n\n"
        f"Expected here:\n{FILE_PATH}\n\n"
        "Fix:\n"
        "1) Create a folder named 'data' next to the EXE (or app.py)\n"
        "2) Put 'villagedata2.xlsx' inside it."
    )
    st.stop()

df = pd.read_excel(FILE_PATH, sheet_name="Sheet1")
limits = pd.read_excel(FILE_PATH, sheet_name="Sheet2")

df.columns = df.columns.astype(str).str.strip()
limits.columns = limits.columns.astype(str).str.strip()

def parse_measurement_to_float(x):
    """Convert mixed lab strings to numeric for logic/plots."""
    if pd.isna(x):
        return np.nan

    # already numeric
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return np.nan

    sl = s.lower()

    # qualitative bacteriological
    if sl in ("absent", "not detected", "nd", "nil", "negative"):
        return 0.0
    if sl in ("present", "detected", "pos", "positive"):
        return 1.0

    # values like <0.05 or >1.2
    if s.startswith("<") or s.startswith(">"):
        try:
            return float(s[1:].strip())
        except Exception:
            return np.nan

    # plain numeric in string
    try:
        return float(s)
    except Exception:
        return np.nan


def display_value(raw, num):
    """Prefer raw text (<0.05 / Absent / Present). Else show formatted numeric."""
    if pd.isna(raw):
        return smart_fmt(num)
    s = str(raw).strip()
    if s == "" or s.lower() == "nan":
        return smart_fmt(num)
    # keep qualitative / censoring marks
    if any(ch.isalpha() for ch in s) or ("<" in s) or (">" in s):
        return s
    return smart_fmt(num)


def display_limit(raw, num):
    """Show original limit text if it exists (e.g., Absent / No relaxation)."""
    if pd.isna(raw):
        return smart_fmt(num)
    s = str(raw).strip()
    if s == "" or s.lower() == "nan":
        return smart_fmt(num)
    if any(ch.isalpha() for ch in s) or ("<" in s) or (">" in s):
        return s
    return smart_fmt(num)

df["Latitude"]  = pd.to_numeric(df.get("Latitude"), errors="coerce")
df["Longitude"] = pd.to_numeric(df.get("Longitude"), errors="coerce")
df = df.dropna(subset=["Latitude", "Longitude"]).copy()

limits["Test Parameter"] = (
    limits["Test Parameter"].astype(str).str.strip().str.replace("\n", " ", regex=False)
)

# Keep RAW limits for display (Absent / No relaxation / etc.)
limits["Acceptable Limit Raw"]  = limits["Acceptable Limit"]
limits["Permissible Limit Raw"] = limits["Permissible Limit"]

# Numeric limits for logic (Absent->0, Present->1, <0.05->0.05, No relaxation->NaN)
limits["Acceptable Limit"]  = limits["Acceptable Limit"].apply(parse_measurement_to_float)
limits["Permissible Limit"] = limits["Permissible Limit"].apply(parse_measurement_to_float)


valid_params = limits["Test Parameter"].dropna().unique().tolist()
param_cols = [c for c in df.columns if c in valid_params]

# Keep RAW values for display + numeric for logic/plots
for col in param_cols:
    df[f"{col}_raw"] = df[col]                 # preserve original text (<0.05 / Absent / Present)
    df[col] = df[col].apply(parse_measurement_to_float)   # numeric working values


limit_map = limits.set_index("Test Parameter")[[
    "Acceptable Limit", "Permissible Limit",
    "Acceptable Limit Raw", "Permissible Limit Raw"
]]

#%% 03 — SIDEBAR OPTIONS + SAFE WRAPPERS (NO 13-INCH TOGGLE + FIX TABLE HEIGHT)

# st.sidebar.markdown("### Display")
# SHOW_LABELS = st.sidebar.toggle("Show village labels on map", value=True)
# No sidebar control. Default behaviour:
SHOW_LABELS = True   # set False if you want labels always OFF

def pydeck_chart_safe(deck, height):
    try:
        st.pydeck_chart(deck, height=height, width="stretch")
    except TypeError:
        st.pydeck_chart(deck, height=height, use_container_width=True)

def dataframe_safe(data, height="stretch"):
    """
    Streamlit height must be int or 'stretch'. Never pass None.
    """
    try:
        st.dataframe(data, height=height, width="stretch")
    except TypeError:
        st.dataframe(data, height=height, use_container_width=True)

def download_button_safe(label, data, file_name, mime):
    try:
        st.download_button(label, data=data, file_name=file_name, mime=mime, width="stretch")
    except TypeError:
        st.download_button(label, data=data, file_name=file_name, mime=mime, use_container_width=True)


#%% 04 — VILLAGE BOUNDARY (LOCAL FOLDER FIRST, ZIP FALLBACK) + NO "LOADED" TEXT

BOUNDARY_FOLDER = DATA_DIR / "villages_boundary"     # preferred
BOUNDARY_ZIP    = DATA_DIR / "villages_boundary.zip" # fallback
BOUNDARY_VILLAGE_FIELD_OVERRIDE = None               # optional exact field name

st.sidebar.markdown("### Village boundary overlay")
SHOW_BOUNDARY = st.sidebar.toggle("Show village boundary", value=True)

def _guess_village_field(cols):
    if not cols:
        return None
    for c in cols:
        cl = str(c).lower()
        if "village" in cl or cl in ("name", "vill_name", "village_na", "v_name"):
            return c
    return cols[0]

@st.cache_data(show_spinner=False)
def load_boundary_from_folder(folder: Path):
    if not _HAS_GPD:
        return None
    if not folder.exists() or not folder.is_dir():
        return None
    shp_files = sorted(folder.glob("*.shp"))
    if not shp_files:
        return None
    gdf = gpd.read_file(shp_files[0])
    if gdf.crs is None:
        gdf = gdf.set_crs("EPSG:4326")
    else:
        gdf = gdf.to_crs("EPSG:4326")
    gdf = gdf[gdf.geometry.notnull()].copy()
    gdf = gdf[gdf.is_valid].copy()
    return gdf

@st.cache_data(show_spinner=False)
def load_boundary_from_zip(zip_path: Path):
    if not _HAS_GPD:
        return None
    if not zip_path.exists():
        return None
    import zipfile, tempfile
    with tempfile.TemporaryDirectory() as td:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(td)
        td_path = Path(td)
        shp_files = sorted(td_path.glob("*.shp"))
        if not shp_files:
            return None
        gdf = gpd.read_file(shp_files[0])
        if gdf.crs is None:
            gdf = gdf.set_crs("EPSG:4326")
        else:
            gdf = gdf.to_crs("EPSG:4326")
        gdf = gdf[gdf.geometry.notnull()].copy()
        gdf = gdf[gdf.is_valid].copy()
        return gdf

BOUNDARY_GDF = None
BOUNDARY_FIELD = None

if SHOW_BOUNDARY and _HAS_GPD:
    BOUNDARY_GDF = load_boundary_from_folder(BOUNDARY_FOLDER)
    if BOUNDARY_GDF is None:
        BOUNDARY_GDF = load_boundary_from_zip(BOUNDARY_ZIP)

    if BOUNDARY_GDF is not None:
        cols = list(BOUNDARY_GDF.columns)
        if BOUNDARY_VILLAGE_FIELD_OVERRIDE and BOUNDARY_VILLAGE_FIELD_OVERRIDE in cols:
            BOUNDARY_FIELD = BOUNDARY_VILLAGE_FIELD_OVERRIDE
        else:
            BOUNDARY_FIELD = _guess_village_field(cols)


#%% 05 — SETTINGS + UTILS (PINS: SIZE=3, DEFAULT DARK BLUE)

CARTO_STYLE = "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json"

# Requested default pin size
PIN_SIZE = 4

def clean_filename(text: str) -> str:
    if not text:
        return "Unknown"
    return re.sub(r"[^A-Za-z0-9_]+", "_", str(text)).strip("_") or "Unknown"

def smart_fmt(v):
    if pd.isna(v):
        return ""
    try:
        f = float(v)
    except Exception:
        return str(v)
    if np.isfinite(f) and float(f).is_integer():
        return str(int(f))
    s = str(v)
    if "." in s:
        dec = s.split(".")[-1]
        if len(dec) > 4:
            return f"{f:.4f}".rstrip("0").rstrip(".")
        return s.rstrip("0").rstrip(".")
    return str(v)

def get_status(v, acc, perm):
    if pd.isna(v):
        return "blue"
    if pd.isna(perm):
        return "green" if v <= acc else "red"
    if v <= acc:
        return "green"
    elif v <= perm:
        return "yellow"
    else:
        return "red"

def calc_zoom(sel_cluster, sel_village, n_points):
    if n_points <= 1:
        return 14
    if sel_village != "All":
        return 12
    if sel_cluster != "All":
        return 11
    return 10

def _svg_pin_data_uri(hex_color: str) -> str:
    svg = f"""
    <svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 64 64">
      <path fill="{hex_color}" d="M32 2C20.4 2 11 11.4 11 23c0 15.7 21 39 21 39s21-23.3 21-39C53 11.4 43.6 2 32 2z"/>
      <circle cx="32" cy="23" r="8" fill="#ffffff" opacity="0.85"/>
    </svg>
    """
    b64 = base64.b64encode(svg.encode("utf-8")).decode("utf-8")
    return f"data:image/svg+xml;base64,{b64}"

# Dark blue default requested
PIN_BLUE   = _svg_pin_data_uri("#0B3D91")  # dark blue
PIN_GREEN  = _svg_pin_data_uri("#00c800")
PIN_YELLOW = _svg_pin_data_uri("#ffc800")
PIN_RED    = _svg_pin_data_uri("#ff0000")

ICON_SPEC  = {"width": 128, "height": 128, "anchorY": 128}

def _icon_for_status(stt: str):
    stt = str(stt).lower()
    if stt == "green":
        url = PIN_GREEN
    elif stt == "yellow":
        url = PIN_YELLOW
    elif stt == "red":
        url = PIN_RED
    else:
        url = PIN_BLUE
    d = {"url": url}
    d.update(ICON_SPEC)
    return d


#%% 06 — BOUNDARY FILTER + MAP BUILDER (BOUNDARY OUTLINE 90% TRANSPARENT)

def _bounds_to_view(bounds):
    minx, miny, maxx, maxy = bounds
    lon = (minx + maxx) / 2.0
    lat = (miny + maxy) / 2.0
    span = max(maxx - minx, maxy - miny)
    span = max(span, 1e-9)
    zoom = float(np.clip(8 - np.log2(span), 5, 16))
    return lat, lon, zoom

def boundary_for_selection(sel_village):
    if (not SHOW_BOUNDARY) or (BOUNDARY_GDF is None) or (BOUNDARY_FIELD is None):
        return None, None

    g = BOUNDARY_GDF.copy()

    if sel_village != "All":
        s = g[BOUNDARY_FIELD].astype(str).str.strip().str.lower()
        target = str(sel_village).strip().lower()
        mask = (s == target)
        if mask.sum() == 0:
            mask = s.str.contains(target, na=False)
        g2 = g[mask].copy()
        if not g2.empty:
            g = g2

    if g.empty:
        return None, None

    return g.__geo_interface__, g.total_bounds

def build_map(df_map, zoom_level, boundary_geojson=None, boundary_bounds=None, show_labels=True):
    df_map = df_map.copy()

    # ensure tooltip columns exist (avoids tooltip breaking)
    for col in [
        "Village Name", "Cluster", "Source", "Location within Village",
        "tooltip_param_name", "tooltip_param_value",
        "tooltip_acc", "tooltip_perm", "tooltip_status"
    ]:
        if col not in df_map.columns:
            df_map[col] = ""
        df_map[col] = df_map[col].fillna("")

    # Make sure coords are numeric (prevents weird centering/zoom bugs)
    df_map["Latitude"]  = pd.to_numeric(df_map.get("Latitude"), errors="coerce")
    df_map["Longitude"] = pd.to_numeric(df_map.get("Longitude"), errors="coerce")

    # Normalize status strings (prevents icon mapping misfires)
    df_map["status"] = df_map.get("status", "blue")
    df_map["status"] = df_map["status"].fillna("blue").astype(str).str.strip().str.lower()

    # ---------- PRIORITIZE RED (worst status) FOR SAME LOCATION ----------
    # severity ranking: blue < green < yellow < red
    sev_rank = {"blue": 0, "green": 1, "yellow": 2, "red": 3}
    df_map["_sev"] = df_map["status"].map(sev_rank).fillna(0).astype(int)

    # numeric value for tie-break (pick higher value if same severity)
    df_map["_valnum"] = pd.to_numeric(df_map["tooltip_param_value"], errors="coerce")
    df_map["_valnum"] = df_map["_valnum"].fillna(-1e18)

    # Count how many records share same coordinates (optional info)
    df_map["samples_here"] = df_map.groupby(["Longitude", "Latitude"])["status"].transform("size")

    # Pick the WORST row per coordinate:
    df_map = df_map.sort_values(["Longitude", "Latitude", "_sev", "_valnum"])
    df_map = df_map.groupby(["Longitude", "Latitude"], as_index=False).tail(1).copy()
    # -------------------------------------------------------------------

    # Pin icons
    df_map["icon_data"] = df_map["status"].apply(_icon_for_status)

    layers = []

    # Boundary outline (90% transparent, outline only)
    if boundary_geojson is not None:
        layers.append(
            pdk.Layer(
                "GeoJsonLayer",
                data=boundary_geojson,
                stroked=True,
                filled=False,
                get_line_color=[0, 0, 0, 25],   # 90% transparent
                line_width_min_pixels=2,        # thicker
                pickable=False,
            )
        )

    # Pins
    layers.append(
        pdk.Layer(
            "IconLayer",
            data=df_map,
            get_icon="icon_data",
            get_position=["Longitude", "Latitude"],
            get_size=PIN_SIZE,
            size_scale=10,
            pickable=True,
            auto_highlight=True,
        )
    )

    # Labels (controlled by show_labels)
    if show_labels:
        label_colors = {
            "green":  [0, 120, 0, 255],
            "yellow": [180, 140, 0, 255],
            "red":    [180, 0, 0, 255],
            "blue":   [11, 61, 145, 255],
        }
        df_map["label_color"] = df_map["status"].map(label_colors)
        df_map["label_color"] = df_map["label_color"].apply(
            lambda x: x if isinstance(x, list) else [11, 61, 145, 255]
        )

        layers.append(
            pdk.Layer(
                "TextLayer",
                data=df_map,
                get_position=["Longitude", "Latitude"],
                get_text="Village Name",
                get_color="label_color",
                get_size=16,
                pickable=False,
            )
        )

    # ---------------- VIEW (FIXED) ----------------
    # Prefer fitting to the *filtered points* bounds (cluster/village),
    # so boundary overlay never overrides zoom.
    pts = df_map[["Longitude", "Latitude"]].dropna()

    if len(pts) >= 2:
        min_lon, max_lon = float(pts["Longitude"].min()), float(pts["Longitude"].max())
        min_lat, max_lat = float(pts["Latitude"].min()),  float(pts["Latitude"].max())

        # small padding so pins near edges aren't clipped
        pad_lon = (max_lon - min_lon) * 0.08
        pad_lat = (max_lat - min_lat) * 0.08
        if pad_lon == 0: pad_lon = 0.01
        if pad_lat == 0: pad_lat = 0.01

        bounds = [min_lon - pad_lon, min_lat - pad_lat, max_lon + pad_lon, max_lat + pad_lat]
        lat, lon, zoom = _bounds_to_view(bounds)
        view_state = pdk.ViewState(latitude=lat, longitude=lon, zoom=zoom, pitch=0, bearing=0)

    elif len(pts) == 1:
        center_lon = float(pts["Longitude"].iloc[0])
        center_lat = float(pts["Latitude"].iloc[0])
        view_state = pdk.ViewState(latitude=center_lat, longitude=center_lon, zoom=zoom_level, pitch=0, bearing=0)

    else:
        # no valid points -> fall back to boundary bounds (district view) if available
        if boundary_bounds is not None:
            lat, lon, zoom = _bounds_to_view(boundary_bounds)
            view_state = pdk.ViewState(latitude=lat, longitude=lon, zoom=zoom, pitch=0, bearing=0)
        else:
            view_state = pdk.ViewState(latitude=30.90, longitude=75.85, zoom=zoom_level, pitch=0, bearing=0)
    # ---------------------------------------------

    tooltip = {
        "html": """
        <div style="font-family: Arial; line-height: 1.25;">
          <div><b>Village:</b> {Village Name}</div>
          <div><b>Cluster:</b> {Cluster}</div>
          <div><b>Source:</b> {Source}</div>
          <div><b>Location:</b> {Location within Village}</div>
          <div style="color:#666; font-size:11px;">Samples at this point: {samples_here}</div>
          <hr style="margin:6px 0;"/>
          <div><b>{tooltip_param_name}:</b> {tooltip_param_value}</div>
          <div><b>Acceptable:</b> {tooltip_acc}</div>
          <div><b>Permissible:</b> {tooltip_perm}</div>
          <div><b>Status:</b> {tooltip_status}</div>
        </div>
        """,
        "style": {
            "backgroundColor": "rgba(255,255,255,0.95)",
            "color": "#111",
            "fontSize": "12px",
            "padding": "8px",
            "borderRadius": "10px"
        }
    }

    return pdk.Deck(
        layers=layers,
        initial_view_state=view_state,
        map_provider="carto",
        map_style=CARTO_STYLE,
        tooltip=tooltip
    )
#%% 07 — GRAPH BUILDER (full block, bold village | italic location)

def build_graph(df_sorted, param, label_hint=""):
    import matplotlib as mpl

    # Force MathText (so bold/italic works without LaTeX)
    mpl.rcParams["text.usetex"] = False
    mpl.rcParams["text.parse_math"] = True

    # ---- helpers for mixed-style tick labels (bold | italic) ----
    def _mt_escape(s: str) -> str:
        # escape characters that MathText treats specially
        return (str(s)
                .replace("\\", r"\\")
                .replace("_", r"\_")
                .replace("{", r"\{")
                .replace("}", r"\}")
                .replace("%", r"\%")
                .replace("#", r"\#")
                .replace("&", r"\&")
                .replace("$", r"\$"))

    def _format_tick(left: str, right: str) -> str:
        # bold left | italic right (pipe stays plain)
        return rf"$\mathbf{{{_mt_escape(left)}}}$ | $\mathit{{{_mt_escape(right)}}}$"
    # ------------------------------------------------------------

    acc = limit_map.loc[param, "Acceptable Limit"]
    perm = limit_map.loc[param, "Permissible Limit"]

    dfp = df_sorted.dropna(subset=[param]).copy()
    if dfp.empty:
        fig, ax = plt.subplots(figsize=(11, 4))
        ax.text(0.5, 0.5, "No numeric values to plot for selected filters.", ha="center", va="center")
        ax.axis("off")
        return fig

    dfp = dfp.sort_values(param, ascending=False)
    values = dfp[param].values

    # Always: Village (bold) | Location within Village (italic)
    v = dfp["Village Name"].astype(str).fillna("")
    loc = dfp["Location within Village"].astype(str).fillna("").replace("nan", "")
    src = dfp["Source"].astype(str).fillna("").replace("nan", "")

    # if location is blank, fall back to Source (so right side isn't empty)
    right = loc.where(loc.str.strip().ne(""), src)

    x_labels = [_format_tick(v.iloc[i], right.iloc[i]) for i in range(len(dfp))]

    color_map = {"green": "green", "yellow": "gold", "red": "red", "blue": "gray"}
    colors = dfp["status"].map(color_map).fillna("gray").tolist()

    fig, ax = plt.subplots(figsize=(11, 4))

    n = len(values)
    step = 1.5 if n <= 3 else 1.0
    x_pos = np.arange(n) * step
    bar_width = 0.25

    ax.bar(x_pos, values, width=bar_width, color=colors, edgecolor="black", zorder=3)

    ax.axhline(acc, color="yellow", linestyle="--", linewidth=2, zorder=10, label="Acceptable Limit")
    if not pd.isna(perm):
        ax.axhline(perm, color="red", linestyle="--", linewidth=2, zorder=10, label="Permissible Limit")

    ax.set_xticks(x_pos)
    ax.set_xticklabels(x_labels, rotation=60, ha="right", fontsize=7)  # try 8/9/10


    ax.set_ylabel(param)
    ax.set_title(param)

    ymax = max(values.max(), acc, (0 if pd.isna(perm) else perm))
    ax.set_ylim(0, ymax * 1.12)

    ax.legend(loc="upper right", bbox_to_anchor=(0.985, 0.985), frameon=True, borderaxespad=0.2)
    ax.set_xlim(x_pos.min() - 0.5, x_pos.max() + 1.0)

    fig.tight_layout()
    return fig



#%% 08 — TABLE STYLING

def styled_table(df_input, param):
    acc = limit_map.loc[param, "Acceptable Limit"]
    perm = limit_map.loc[param, "Permissible Limit"]

    dft = df_input.copy()
    dft["status"] = dft[param].apply(lambda v: get_status(v, acc, perm))
    dft = dft.sort_values(["Village Name", "Source", "Location within Village"], ascending=True)

    df_show = dft[["Village Name", "Source", "Location within Village", param]].copy()

    # RAW column (for display)
    raw_col = f"{param}_raw" if f"{param}_raw" in dft.columns else param

    # Show RAW values (Absent / Present / <0.05) if available, else formatted numeric
    df_show[param] = [
        display_value(r, n) for r, n in zip(dft[raw_col], dft[param])
    ]

    # Show limits as RAW text if present (Absent / No relaxation), else numeric
    df_show["Acceptable Limit"] = display_limit(limit_map.loc[param, "Acceptable Limit Raw"], acc)
    df_show["Permissible Limit"] = display_limit(limit_map.loc[param, "Permissible Limit Raw"], perm)

    status_list = dft["status"].tolist()

    df_show = df_show.reset_index(drop=True)
    df_show.index = df_show.index + 1
    idx_to_status = dict(zip(df_show.index, status_list))

    def row_style(row):
        stt = idx_to_status.get(row.name, "blue")
        styles = {col: "" for col in df_show.columns}
        if stt == "red":
            for col in ["Village Name", "Source", "Location within Village", param]:
                styles[col] = "color:red; font-weight:bold;"
        return pd.Series(styles, index=df_show.columns)

    styled = df_show.style.apply(row_style, axis=1)
    export_df = df_show.copy()
    return styled, export_df


def name_tag(sel_cluster, sel_village):
    if sel_village != "All":
        return sel_village
    if sel_cluster != "All":
        return sel_cluster
    return "All"


#%% 09 — DASHBOARD MODE

def dashboard_mode():

    clusters = ["All"] + sorted(df["Cluster"].dropna().unique().tolist())
    sel_cluster = st.sidebar.selectbox("Cluster", clusters, index=0)

    df1 = df if sel_cluster == "All" else df[df["Cluster"] == sel_cluster]
    villages = ["All"] + sorted(df1["Village Name"].dropna().unique().tolist())
    sel_village = st.sidebar.selectbox("Village", villages, index=0)

    df2 = df1 if sel_village == "All" else df1[df1["Village Name"] == sel_village]

    st.sidebar.markdown("### Parameters")
    param_cols_sorted = sorted(param_cols, key=lambda x: str(x).lower())
    sel_param = st.sidebar.selectbox("Select ONE parameter", ["-- Select parameter --"] + param_cols_sorted, index=0)
    if sel_param == "-- Select parameter --":
        sel_param = ""

    sources = ["All"] + sorted(df2["Source"].dropna().unique().tolist())
    sel_source = st.sidebar.selectbox("Source", sources, index=0)
    df3 = df2 if sel_source == "All" else df2[df2["Source"] == sel_source]

    locs = ["All"] + sorted(df3["Location within Village"].dropna().unique().tolist())
    sel_loc = st.sidebar.selectbox("Location within Village", locs, index=0)

    df_view = df3 if sel_loc == "All" else df3[df3["Location within Village"] == sel_loc]
    df_view = df_view.copy()

    if df_view.empty:
        st.warning("No data for selected filters.")
        return

    zoom_level = calc_zoom(sel_cluster, sel_village, len(df_view))
    boundary_geojson, boundary_bounds = boundary_for_selection(sel_village)

    st.markdown("## 🗺 Map (Hover the mouse on any pin to see details)")

    map_col, legend_col = st.columns([6, 1])

    if not sel_param:
        st.info("Select a Cluster, Village and parameter to view detailed analysis.")

        df_default = df_view.copy()
        df_default["status"] = "blue"

        df_default["tooltip_param_name"]  = "Parameter"
        df_default["tooltip_param_value"] = "Select a parameter"
        df_default["tooltip_acc"]         = "-"
        df_default["tooltip_perm"]        = "-"
        df_default["tooltip_status"]      = "-"

        deck = build_map(
            df_default,
            zoom_level=zoom_level,
            boundary_geojson=boundary_geojson,
            boundary_bounds=boundary_bounds,
            show_labels=SHOW_LABELS
        )

        with map_col:
            st.markdown("<div style='border:2px solid #ccc; border-radius:10px; padding:6px;'>", unsafe_allow_html=True)
            pydeck_chart_safe(deck, height=560)
            st.markdown("</div>", unsafe_allow_html=True)

        with legend_col:
            st.markdown(
                f"""
                <div style="border:1px solid #e5e5e5; border-radius:12px; padding:12px; background:#fff;">
                  <div style="font-weight:700; margin-bottom:10px;">Legend</div>
                  <div style="display:flex; align-items:center; gap:8px; margin:6px 0;">
                    <img src="{PIN_BLUE}" style="width:16px; height:16px;"/>
                    <span>Points (select parameter)</span>
                  </div>
                 </div>
                """,
                unsafe_allow_html=True
            )
        return

    # --- UPDATED: use RAW limits + RAW values for tooltip display ---
    acc = limit_map.loc[sel_param, "Acceptable Limit"]
    perm = limit_map.loc[sel_param, "Permissible Limit"]
    acc_raw = limit_map.loc[sel_param, "Acceptable Limit Raw"]
    perm_raw = limit_map.loc[sel_param, "Permissible Limit Raw"]

    df_view["status"] = df_view[sel_param].apply(lambda v: get_status(v, acc, perm))

    raw_col = f"{sel_param}_raw" if f"{sel_param}_raw" in df_view.columns else sel_param

    df_view["tooltip_param_name"]  = str(sel_param)
    df_view["tooltip_param_value"] = [
        display_value(r, n) for r, n in zip(df_view[raw_col], df_view[sel_param])
    ]
    df_view["tooltip_acc"]   = display_limit(acc_raw, acc)
    df_view["tooltip_perm"]  = display_limit(perm_raw, perm)
    df_view["tooltip_status"] = df_view["status"].astype(str)

    deck = build_map(
        df_view,
        zoom_level=zoom_level + 1,
        boundary_geojson=boundary_geojson,
        boundary_bounds=boundary_bounds,
        show_labels=SHOW_LABELS
    )

    with map_col:
        st.markdown("<div style='border:2px solid #ccc; border-radius:10px; padding:6px;'>", unsafe_allow_html=True)
        pydeck_chart_safe(deck, height=560)
        st.markdown("</div>", unsafe_allow_html=True)

    with legend_col:
        st.markdown(
            f"""
            <div style="border:1px solid #e5e5e5; border-radius:12px; padding:12px; background:#fff;">
              <div style="font-weight:700; margin-bottom:10px;">Legend</div>
              <div style="display:flex; align-items:center; gap:8px; margin:6px 0;">
                <img src="{PIN_GREEN}" style="width:16px; height:16px;"/>
                <span>Within Acceptable</span>
              </div>
              <div style="display:flex; align-items:center; gap:8px; margin:6px 0;">
                <img src="{PIN_YELLOW}" style="width:16px; height:16px;"/>
                <span>Between limits</span>
              </div>
              <div style="display:flex; align-items:center; gap:8px; margin:6px 0;">
                <img src="{PIN_RED}" style="width:16px; height:16px;"/>
                <span>Above permissible</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("## 📊 Drinking Water Standards (BIS)")
    label_hint = "single_village" if sel_village != "All" else ""
    fig = build_graph(df_view, sel_param, label_hint=label_hint)
    st.pyplot(fig, use_container_width=True)

    tag = clean_filename(name_tag(sel_cluster, sel_village))
    base = clean_filename(f"{sel_param}_{tag}")

    buf = io.BytesIO()
    fig.savefig(buf, format="jpg", dpi=200, bbox_inches="tight")
    download_button_safe(
        f"Download {sel_param} Graph (JPG)",
        data=buf.getvalue(),
        file_name=f"{base}.jpg",
        mime="image/jpeg"
    )
    plt.close(fig)

    st.markdown("## 🧾 Detailed Table")
    styled_tbl, export_df = styled_table(df_view, sel_param)
    dataframe_safe(styled_tbl, height=320)   # <-- avoid height=None error

    csv_bytes = export_df.to_csv(index=False).encode("utf-8")
    download_button_safe(
        f"Download {sel_param} Table (CSV)",
        data=csv_bytes,
        file_name=f"{base}.csv",
        mime="text/csv"
    )


#%% 10 — COMPARISON MODE (SIDE-BY-SIDE)

def comparison_mode():

    st.markdown("## 🔍 Comparison Mode")

    st.sidebar.markdown("### Select up to 4 parameters")
    param_cols_sorted = sorted(param_cols, key=lambda x: str(x).lower())
    multi_params = st.sidebar.multiselect("Parameters", param_cols_sorted, default=[])

    if len(multi_params) < 2:
        st.info("Select at least Two parameters to compare.")
        return
    if len(multi_params) > 4:
        st.error("Please select a maximum of 4 parameters.")
        return

    clusters = ["All"] + sorted(df["Cluster"].dropna().unique().tolist())
    sel_cluster = st.sidebar.selectbox("Cluster", clusters, index=0)

    df1 = df if sel_cluster == "All" else df[df["Cluster"] == sel_cluster]
    villages = ["All"] + sorted(df1["Village Name"].dropna().unique().tolist())
    sel_village = st.sidebar.selectbox("Village", villages, index=0)

    df2 = df1 if sel_village == "All" else df1[df1["Village Name"] == sel_village]

    sources = ["All"] + sorted(df2["Source"].dropna().unique().tolist())
    sel_source = st.sidebar.selectbox("Source", sources, index=0)

    df3 = df2 if sel_source == "All" else df2[df2["Source"] == sel_source]

    locs = ["All"] + sorted(df3["Location within Village"].dropna().unique().tolist())
    sel_loc = st.sidebar.selectbox("Location within Village", locs, index=0)

    df_view = df3 if sel_loc == "All" else df3[df3["Location within Village"] == sel_loc]
    df_view = df_view.copy()

    if df_view.empty:
        st.warning("No data for selected filters.")
        return

    zoom_level = calc_zoom(sel_cluster, sel_village, len(df_view))
    tag = clean_filename(name_tag(sel_cluster, sel_village))
    boundary_geojson, boundary_bounds = boundary_for_selection(sel_village)

    def make_grid(n):
        if n == 1:
            return [st.container()]
        if n == 2:
            c1, c2 = st.columns(2)
            return [c1, c2]
        if n == 3:
            c1, c2 = st.columns(2)
            c3 = st.container()
            return [c1, c2, c3]
        c1, c2 = st.columns(2)
        c3, c4 = st.columns(2)
        return [c1, c2, c3, c4]

    # -------------------- MAPS --------------------
    st.markdown("## 🗺 Maps")
    map_cells = make_grid(len(multi_params))

    for i, param in enumerate(multi_params):
        acc = limit_map.loc[param, "Acceptable Limit"]
        perm = limit_map.loc[param, "Permissible Limit"]
        acc_raw = limit_map.loc[param, "Acceptable Limit Raw"]
        perm_raw = limit_map.loc[param, "Permissible Limit Raw"]

        dfp = df_view.copy()
        dfp["status"] = dfp[param].apply(lambda v: get_status(v, acc, perm))

        raw_col = f"{param}_raw" if f"{param}_raw" in dfp.columns else param

        dfp["tooltip_param_name"]  = str(param)
        dfp["tooltip_param_value"] = [
            display_value(r, n) for r, n in zip(dfp[raw_col], dfp[param])
        ]
        dfp["tooltip_acc"]    = display_limit(acc_raw, acc)
        dfp["tooltip_perm"]   = display_limit(perm_raw, perm)
        dfp["tooltip_status"] = dfp["status"].astype(str)

        with map_cells[i]:
            st.markdown(f"### {param}")
            deck = build_map(
                dfp,
                zoom_level=zoom_level + 1,
                boundary_geojson=boundary_geojson,
                boundary_bounds=boundary_bounds,
                show_labels=SHOW_LABELS
            )
            st.markdown("<div style='border:2px solid #ccc; border-radius:10px; padding:6px;'>", unsafe_allow_html=True)
            pydeck_chart_safe(deck, height=520)
            st.markdown("</div>", unsafe_allow_html=True)

    # -------------------- GRAPHS --------------------
    st.markdown("## 📈 Graphs")
    graph_cells = make_grid(len(multi_params))

    for i, param in enumerate(multi_params):
        acc = limit_map.loc[param, "Acceptable Limit"]
        perm = limit_map.loc[param, "Permissible Limit"]

        dfp = df_view.copy()
        dfp["status"] = dfp[param].apply(lambda v: get_status(v, acc, perm))

        label_hint = "single_village" if sel_village != "All" else ""
        fig = build_graph(dfp, param, label_hint=label_hint)

        with graph_cells[i]:
            card = st.container(border=True)
            with card:
                st.markdown(f"### {param}")
                st.pyplot(fig, use_container_width=True)

                base = clean_filename(f"{param}_{tag}")
                dl_buf = io.BytesIO()
                fig.savefig(dl_buf, format="jpg", dpi=200, bbox_inches="tight")
                download_button_safe(
                    f"Download {param} Graph (JPG)",
                    data=dl_buf.getvalue(),
                    file_name=f"{base}.jpg",
                    mime="image/jpeg"
                )
        plt.close(fig)

    # -------------------- TABLES --------------------
    st.markdown("## 🧾 Tables")
    table_cells = make_grid(len(multi_params))

    for i, param in enumerate(multi_params):
        styled_tbl, export_df = styled_table(df_view, param)

        with table_cells[i]:
            card = st.container(border=True)
            with card:
                st.markdown(f"### {param}")
                dataframe_safe(styled_tbl, height=240)  # <-- avoid height=None error

                base = clean_filename(f"{param}_{tag}")
                csv_bytes = export_df.to_csv(index=False).encode("utf-8")
                download_button_safe(
                    f"Download {param} Table (CSV)",
                    data=csv_bytes,
                    file_name=f"{base}.csv",
                    mime="text/csv"
                )


#%% 11 — MODE SWITCH

mode = st.sidebar.radio("Select Mode", ["Dashboard", "Comparison"], index=0)

if mode == "Dashboard":
    dashboard_mode()
else:
    comparison_mode()
