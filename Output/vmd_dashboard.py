import streamlit as st
import pandas as pd
import re
import io
from pathlib import Path
from datetime import datetime
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go


# -----------------------------------------------------------------------------
# SELF-CONTAINED DATA PATHS
# -----------------------------------------------------------------------------
# The dashboard is designed to run from GitHub/Streamlit Cloud as well as
# locally. It never depends on the user's Windows E:\ or C:\ paths.
#
# Supported repository layouts:
#   1) Excel-DashBoard/Output/app.py + Output/maps + Output/data
#   2) Excel-DashBoard/app.py + Output/maps + Output/data
#
# In both cases the dashboard automatically resolves the correct Output folder.
BASE_DIR = Path(__file__).resolve().parent

if BASE_DIR.name.lower() == "output":
    OUTPUT_FOLDER = BASE_DIR
elif (BASE_DIR / "Output").exists():
    OUTPUT_FOLDER = BASE_DIR / "Output"
elif (BASE_DIR / "output").exists():
    OUTPUT_FOLDER = BASE_DIR / "output"
else:
    # Fallback: keep the app self-contained if Output is not present.
    OUTPUT_FOLDER = BASE_DIR

OUTPUT_DIR = OUTPUT_FOLDER
MAPS_FOLDER = OUTPUT_FOLDER
DATA_FOLDER = OUTPUT_FOLDER

st.set_page_config(
    page_title="VMD Agricultural Vulnerability | Thesis Dashboard",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Source+Serif+4:wght@500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: #f5f7f9;
}

.block-container {
    max-width: 1580px;
    padding-top: 1.0rem;
    padding-bottom: 3rem;
}

[data-testid="stSidebar"] {
    background: #102a43;
}

[data-testid="stSidebar"] * {
    color: #f3f7fa !important;
}

[data-testid="stSidebar"] .stRadio label {
    padding: 4px 0;
}

.hero {
    background: linear-gradient(135deg, #12344d 0%, #1d5d70 58%, #2e7d67 100%);
    border-radius: 18px;
    padding: 28px 32px;
    color: white;
    margin-bottom: 22px;
    box-shadow: 0 8px 25px rgba(16,42,67,.13);
}

.hero h1 {
    font-family: 'Source Serif 4', serif;
    font-size: 34px;
    margin: 0 0 8px 0;
    color: white;
    font-weight: 700;
}

.hero p {
    margin: 0;
    color: #dcecf2;
    font-size: 14px;
    line-height: 1.65;
}

.section-title {
    font-family: 'Source Serif 4', serif;
    font-size: 27px;
    color: #17324d;
    font-weight: 700;
    margin: 18px 0 13px;
    padding-bottom: 8px;
    border-bottom: 2px solid #dce5eb;
}

.section-subtitle {
    font-size: 18px;
    color: #234b61;
    font-weight: 700;
    margin: 20px 0 10px;
}

.card {
    background: white;
    border: 1px solid #dfe7ed;
    border-radius: 13px;
    padding: 17px 18px;
    min-height: 116px;
    box-shadow: 0 3px 12px rgba(18,52,77,.055);
}

.card .label {
    color: #71818c;
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: .5px;
}

.card .value {
    color: #17324d;
    font-size: 27px;
    line-height: 1.2;
    font-weight: 800;
    margin-top: 7px;
}

.card .detail {
    color: #778893;
    font-size: 11px;
    margin-top: 5px;
}

.finding {
    background: #ffffff;
    border-left: 4px solid #2e7d67;
    border-radius: 8px;
    padding: 13px 16px;
    margin: 8px 0;
    color: #344e5e;
    line-height: 1.6;
    box-shadow: 0 2px 8px rgba(18,52,77,.04);
}

.note {
    background: #edf5f7;
    border: 1px solid #d8e7eb;
    border-radius: 10px;
    padding: 13px 16px;
    color: #385668;
    line-height: 1.6;
}

.warning-note {
    background: #fff8e8;
    border: 1px solid #f1dfad;
    border-radius: 10px;
    padding: 13px 16px;
    color: #665126;
    line-height: 1.6;
}

.map-frame {
    background: white;
    border: 1px solid #dce5eb;
    border-radius: 12px;
    padding: 10px;
    box-shadow: 0 3px 12px rgba(18,52,77,.05);
}

.caption {
    text-align: center;
    color: #647582;
    font-size: 12px;
    font-style: italic;
    margin: 6px 0 13px;
}

.small-muted {
    color: #73828d;
    font-size: 12px;
}

div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #dfe7ed;
    border-radius: 12px;
    padding: 10px 14px;
    box-shadow: 0 2px 8px rgba(18,52,77,.04);
}

[data-testid="stDataFrame"] {
    border-radius: 10px;
}

button[kind="secondary"] {
    border-radius: 8px;
}

.footer {
    text-align: center;
    color: #7a8993;
    font-size: 11px;
    padding-top: 12px;
}
</style>
""", unsafe_allow_html=True)


def norm(value):
    return re.sub(r"[^a-z0-9]+", "_", str(value).lower()).strip("_")


def fmt(value, decimals=3):
    if value is None or pd.isna(value):
        return "N/A"
    return f"{float(value):,.{decimals}f}"


def safe_numeric(df, column):
    if df is None or column not in df.columns:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[column], errors="coerce")


def province_col(df):
    if df is None:
        return None
    candidates = [
        "NAME_1", "Province", "province", "Province_Name",
        "province_name", "NAME", "PROVINCE"
    ]
    for c in candidates:
        if c in df.columns:
            return c
    # Flexible fallback
    for c in df.columns:
        if "province" in norm(c) or norm(c) in {"name_1", "name"}:
            return c
    return None


def period_col(df):
    if df is None:
        return None
    for c in ["Period", "period", "PERIOD", "Year", "year"]:
        if c in df.columns:
            return c
    return None


def mean_col(df):
    if df is None:
        return None
    for c in ["mean", "Mean", "MEAN", "mean_value"]:
        if c in df.columns:
            return c
    return None


def area_col(df):
    if df is None:
        return None
    for c in ["sum", "area", "Area", "SUM", "Area_Sum"]:
        if c in df.columns:
            return c
    return None


@st.cache_data(show_spinner=False)
def discover_files(folder):
    folder = Path(folder)
    if not folder.exists():
        return [], []
    image_exts = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}
    images = sorted(
        [p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() in image_exts],
        key=lambda x: str(x).lower()
    )
    csvs = sorted(
        [p for p in folder.rglob("*.csv") if p.is_file()],
        key=lambda x: str(x).lower()
    )
    return images, csvs


@st.cache_data(show_spinner=False)
def read_csv_file(path):
    if path is None:
        return None
    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        try:
            return pd.read_csv(path, encoding="latin1")
        except Exception:
            return None
    except Exception:
        return None


def find_csv(csvs, wanted):
    target = norm(wanted.replace(".csv", ""))
    for f in csvs:
        if f.name.lower() == wanted.lower():
            return f
    for f in csvs:
        if target in norm(f.stem):
            return f
    return None


def find_map(images, keywords):
    """Find a map using normalized filename aliases.

    Exact/strong matches are preferred so that similarly named thesis figures
    do not accidentally replace the intended map.
    """
    ranked = []
    for f in images:
        stem = norm(f.stem)
        for idx, key in enumerate(keywords):
            k = norm(key)
            if not k:
                continue
            if stem == k:
                ranked.append((0, idx, len(stem), f))
            elif stem.startswith(k):
                ranked.append((1, idx, len(stem), f))
            elif k in stem:
                ranked.append((2, idx, len(stem), f))
    if ranked:
        ranked.sort(key=lambda x: (x[0], x[1], x[2]))
        return ranked[0][3]
    return None


def period_summary(df):
    pc = period_col(df)
    mc = mean_col(df)
    if df is None or pc is None or mc is None:
        return None
    x = df.copy()
    x["_mean"] = pd.to_numeric(x[mc], errors="coerce")
    x["_period"] = x[pc].astype(str)
    return (
        x.dropna(subset=["_mean"])
         .groupby("_period", sort=False)["_mean"]
         .mean()
         .reset_index(name="Regional Mean")
    )


def get_periods(df):
    pc = period_col(df)
    if df is None or pc is None:
        return []
    return df[pc].dropna().astype(str).drop_duplicates().tolist()


def chart_layout(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=45, b=10),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Inter, sans-serif", color="#304b5d"),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#e8eef2")
    return fig


def line_chart(df, title, y_title):
    if df is None or df.empty:
        return None
    fig = px.line(df, x="_period", y="Regional Mean", markers=True, title=title)
    fig.update_traces(line_width=3, marker_size=8)
    fig.update_yaxes(title=y_title)
    fig.update_xaxes(title="")
    return chart_layout(fig)


def map_bytes(path):
    try:
        img = Image.open(path).convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def show_map(path, caption, height_hint=None):
    if not path:
        st.warning("Map file was not found in the bundled Output folder.")
        return
    try:
        img = Image.open(path)
        st.markdown('<div class="map-frame">', unsafe_allow_html=True)
        st.image(img, use_container_width=True)
        st.markdown(f'<div class="caption">{caption}</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Could not open map: {e}")


def download_map(path, label="Download map"):
    if path:
        b = map_bytes(path)
        if b:
            st.download_button(
                label,
                data=b,
                file_name=f"{path.stem}.png",
                mime="image/png",
                key=f"download_{path.stem}",
            )


def dataframe_download(df, filename, label="Download CSV"):
    if df is None or df.empty:
        return
    st.download_button(
        label,
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=filename,
        mime="text/csv",
    )


def metric_card(label, value, detail=""):
    st.markdown(
        f"""
        <div class="card">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
            <div class="detail">{detail}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

MAPS = {
    "Combined Agricultural Vulnerability": [
        "VMD_Combined_Agricultural_Vulnerability_250m",
        "Combined_Agricultural_Vulnerability",
        "agricultural_vulnerability",
    ],
    "Salinity Hotspots": [
        "VMD_Salinity_Hotspots_250m",
        "Salinity_Hotspots",
        "salinity_hotspot",
    ],
    "Salinity–Drought Overlap": [
        "VMD_Salinity_Drought_Overlap_250m",
        "Salinity_Drought_Overlap",
        "salinity_drought",
    ],
    "SPI-3 Drought 2023–2024": [
        "VMD_SPI3_2023_2024",
        "SPI3_2023_2024",
        "spi3",
    ],
    "NDVI 2023–2024": [
        "VMD_NDVI_2023_2024_Proper_Map",
        "NDVI_2023_2024",
    ],
    "NDVI Change 2000–2023": [
        "VMD_NDVI_Change_2000_2023_map",
        "NDVI_Change_2000_2023",
        "ndvi_change",
    ],
    "Mann–Kendall S: NDVI & Rainfall": [
        "VMD_Mann_Kendall_S_250m",
        "Mann_Kendall",
        "MK_S",
    ],
    "Sen's Slope": [
        "VMD_Sen_Slope_250m",
        "Sen_Slope",
    ],
    "VSSI 2000": ["VMD_VSSI_2000", "VSSI_2000"],
    "VSSI 2010–2011": ["VMD_VSSI_2010_2011", "VSSI_2010_2011"],
    "VSSI 2015–2016": ["VMD_VSSI_2015_2016", "VSSI_2015_2016"],
    "VSSI 2019–2020": ["VMD_VSSI_2019_2020", "VSSI_2019_2020"],
    "VSSI 2023–2024": ["VMD_VSSI_2023_2024", "VSSI_2023_2024"],
    "VSSI Change 2000–2023": [
        "VMD_VSSI_Change_2000_2023",
        "VSSI_Change_2000_2023",
        "vssi_change",
    ],
}

# Additional aliases for the actual thesis figure filenames.
# These allow the dashboard to work even when exported PNGs have descriptive
# generated names rather than the canonical VMD names above.
MAPS["Combined Agricultural Vulnerability"] += ["wide_clean_cartographic_infographic_map_a_detail"]
MAPS["Salinity Hotspots"] += ["a_detailed_cartographic_map_image_a_clean_public", "map_sal_hotspots"]
MAPS["Salinity–Drought Overlap"] += ["a_detailed_thematic_map_infographic_a_clean_carto", "map_sal_drought"]
MAPS["SPI-3 Drought 2023–2024"] += ["a_clean_detailed_infographic_style_map_figure_in", "map_spi"]
MAPS["NDVI 2023–2024"] += ["map_ndvi_current"]
MAPS["NDVI Change 2000–2023"] += ["map_ndvi_change"]
MAPS["Mann–Kendall S: NDVI & Rainfall"] += ["a_wide_infographic_style_scientific_map_figure_cl", "map_mk"]
MAPS["Sen's Slope"] += ["a_detailed_map_infographic_a_landscape_oriented_s", "map_sen"]
MAPS["VSSI 2000"] += ["VMD_VSSI_2000_panel"]
MAPS["VSSI 2010–2011"] += ["VMD_VSSI_2010_2011_panel"]
MAPS["VSSI 2015–2016"] += ["VMD_VSSI_2015_2016_panel"]
MAPS["VSSI 2019–2020"] += ["VMD_VSSI_2019_2020_panel"]
MAPS["VSSI 2023–2024"] += ["VMD_VSSI_2023_2024_panel"]
MAPS["VSSI Change 2000–2023"] += ["a_detailed_infographic_map_figure_a_high_resoluti", "map_vssi_change"]

CSV_NAMES = {
    "NDVI": "VMD_Province_NDVI_Statistics.csv",
    "SPI3": "VMD_Province_SPI3_Statistics.csv",
    "VSSI": "VMD_Province_VSSI_Statistics.csv",
    "Salinity": "VMD_Salinity_Hotspot_Area.csv",
    "Vulnerability": "VMD_Agricultural_Vulnerability_Area.csv",
}

images, csvs = discover_files(MAPS_FOLDER)
data = {k: read_csv_file(find_csv(csvs, v)) for k, v in CSV_NAMES.items()}

st.sidebar.markdown("## 🌾 VMD Thesis Dashboard")
st.sidebar.caption("Professional GIS Research Interface")

pages = [
    "Executive Dashboard",
    "Thematic Map Gallery",
    "NDVI | Vegetation",
    "SPI-3 | Drought",
    "Salinity & Vulnerability",
    "VSSI | Salinity Stress",
    "Province Explorer",
    "Research Findings",
    "Data & Reproducibility",
]
page = st.sidebar.radio("Navigation", pages)

st.sidebar.markdown("---")
st.sidebar.markdown("### Data Connection")

if OUTPUT_FOLDER.exists():
    st.sidebar.success("Bundled Output connected")
else:
    st.sidebar.error("Output folder not found")

st.sidebar.caption(str(OUTPUT_FOLDER))
st.sidebar.write(f"Maps detected: **{len(images)}**")
st.sidebar.write(f"CSV files detected: **{len(csvs)}**")
st.sidebar.write(
    f"Core datasets: **{sum(v is not None for v in data.values())}/{len(data)}**"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### Thesis Indicators")
for item in ["NDVI", "SPI-3", "Salinity", "Agricultural Vulnerability", "VSSI"]:
    st.sidebar.markdown(f"• {item}")

st.sidebar.markdown("---")
st.sidebar.caption("Data and figures are bundled with the dashboard deployment.")
st.sidebar.caption(f"Dashboard refreshed: {datetime.now().strftime('%d %b %Y, %H:%M')}")


st.markdown(
    """
    <div class="hero">
        <h1>Vietnamese Mekong Delta Agricultural Vulnerability Assessment</h1>
        <p>
        Thesis-oriented GIS dashboard integrating vegetation condition, drought,
        salinity hotspots, agricultural vulnerability and salinity-stress indicators.
        The interface is designed for research interpretation, figure review,
        province-level comparison and reproducible reporting.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if page == "Executive Dashboard":

    st.markdown('<div class="section-title">Executive Research Dashboard</div>', unsafe_allow_html=True)

    ndvi = data["NDVI"]
    spi = data["SPI3"]
    sal = data["Salinity"]
    vul = data["Vulnerability"]
    vssi = data["VSSI"]

    ndvi_s = period_summary(ndvi)
    spi_s = period_summary(spi)
    vssi_s = period_summary(vssi)

    ndvi_change = None
    spi_change = None
    vssi_change = None

    if ndvi_s is not None and len(ndvi_s) >= 2:
        ndvi_change = ndvi_s.iloc[-1]["Regional Mean"] - ndvi_s.iloc[0]["Regional Mean"]
    if spi_s is not None and len(spi_s) >= 2:
        spi_change = spi_s.iloc[-1]["Regional Mean"] - spi_s.iloc[0]["Regional Mean"]
    if vssi_s is not None and len(vssi_s) >= 2:
        vssi_change = vssi_s.iloc[-1]["Regional Mean"] - vssi_s.iloc[0]["Regional Mean"]

    sal_total = None
    sal_top = "N/A"
    vul_total = None
    vul_top = "N/A"

    if sal is not None and area_col(sal):
        s = sal.copy()
        ac = area_col(s)
        pc = province_col(s)
        s["_area"] = safe_numeric(s, ac)
        sal_total = s["_area"].sum()
        if pc and not s["_area"].dropna().empty:
            sal_top = str(s.loc[s["_area"].idxmax(), pc])

    if vul is not None and area_col(vul):
        v = vul.copy()
        ac = area_col(v)
        pc = province_col(v)
        v["_area"] = safe_numeric(v, ac)
        vul_total = v["_area"].sum()
        if pc and not v["_area"].dropna().empty:
            vul_top = str(v.loc[v["_area"].idxmax(), pc])

    cols = st.columns(5)
    with cols[0]:
        metric_card("Study Provinces", "13", "Mekong Delta provincial assessment")
    with cols[1]:
        metric_card("NDVI Periods", "5", "2000 to 2023–2024")
    with cols[2]:
        metric_card("NDVI Change", fmt(ndvi_change, 3), "First to latest supplied period")
    with cols[3]:
        metric_card("SPI-3 Change", fmt(spi_change, 3), "First to latest supplied period")
    with cols[4]:
        metric_card("VSSI Change", fmt(vssi_change, 3), "Matched regional period summary")

    st.markdown('<div class="section-subtitle">Research Evidence at a Glance</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f"""
            <div class="finding">
            <b>Salinity:</b> The supplied salinity-area statistics have a reported
            total of <b>{fmt(sal_total, 0)}</b>, with <b>{sal_top}</b> having the
            largest provincial reported area.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="finding">
            <b>Agricultural vulnerability:</b> The supplied vulnerability statistics
            total <b>{fmt(vul_total, 0)}</b> in reported area sum, led by
            <b>{vul_top}</b>.
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        if ndvi_s is not None:
            st.markdown(
                f"""
                <div class="finding">
                <b>Vegetation:</b> The regional mean NDVI changes from
                <b>{fmt(ndvi_s.iloc[0]['Regional Mean'])}</b> in the first supplied
                period to <b>{fmt(ndvi_s.iloc[-1]['Regional Mean'])}</b> in the
                latest period.
                </div>
                """,
                unsafe_allow_html=True,
            )
        if spi_s is not None:
            st.markdown(
                f"""
                <div class="finding">
                <b>Drought:</b> The regional mean SPI-3 changes from
                <b>{fmt(spi_s.iloc[0]['Regional Mean'])}</b> to
                <b>{fmt(spi_s.iloc[-1]['Regional Mean'])}</b>; more negative SPI
                values indicate stronger precipitation deficit.
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown('<div class="section-subtitle">Integrated Indicator Framework</div>', unsafe_allow_html=True)
    framework = [
        ("01", "NDVI", "Vegetation condition"),
        ("02", "SPI-3", "Short-term drought stress"),
        ("03", "Salinity", "Salinity hotspot distribution"),
        ("04", "Vulnerability", "Integrated agricultural exposure"),
        ("05", "VSSI", "Salinity-related stress"),
    ]
    cols = st.columns(5)
    for col, (n, title, desc) in zip(cols, framework):
        with col:
            metric_card(f"{n} · {title}", title, desc)

    st.markdown('<div class="section-subtitle">Key Thesis Figures</div>', unsafe_allow_html=True)
    key_maps = [
        "Combined Agricultural Vulnerability",
        "NDVI 2023–2024",
        "SPI-3 Drought 2023–2024",
        "Salinity Hotspots",
    ]
    cols = st.columns(2)
    for col, label in zip(cols, key_maps):
        with col:
            st.markdown(f"**{label}**")
            show_map(find_map(images, MAPS[label]), f"Figure: {label}")

    if ndvi_s is not None:
        st.markdown('<div class="section-subtitle">Regional Temporal Context</div>', unsafe_allow_html=True)
        fig = line_chart(ndvi_s, "Regional Mean NDVI Through Study Periods", "Mean NDVI")
        if fig:
            st.plotly_chart(fig, use_container_width=True)


elif page == "Thematic Map Gallery":

    st.markdown('<div class="section-title">Thematic Map Gallery</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note">All figures are discovered automatically from the bundled Output folder. '
        'This gallery preserves the supplied cartographic outputs and provides thesis-ready viewing and download.</div>',
        unsafe_allow_html=True,
    )

    label = st.selectbox("Select thematic figure", list(MAPS.keys()))
    path = find_map(images, MAPS[label])

    if path:
        show_map(path, f"Figure: {label}")
        download_map(path, "Download Selected Figure")

        st.markdown(
            f'<div class="small-muted">Source file: {path.name}</div>',
            unsafe_allow_html=True,
        )
    else:
        st.error("Selected map was not detected.")

    st.markdown('<div class="section-subtitle">Available Map Inventory</div>', unsafe_allow_html=True)
    inventory = []
    for label2, keys in MAPS.items():
        p = find_map(images, keys)
        inventory.append({
            "Thematic Figure": label2,
            "Status": "Available" if p else "Missing",
            "File": p.name if p else "—",
        })
    inv_df = pd.DataFrame(inventory)
    st.dataframe(inv_df, use_container_width=True, hide_index=True)
    dataframe_download(inv_df, "VMD_Map_Inventory.csv")


elif page == "NDVI | Vegetation":

    st.markdown('<div class="section-title">NDVI | Vegetation Condition</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note">NDVI is interpreted as a vegetation-condition indicator. '
        'The dashboard distinguishes current spatial condition from long-term change and province-level variation.</div>',
        unsafe_allow_html=True,
    )

    a, b = st.columns(2)
    with a:
        show_map(find_map(images, MAPS["NDVI 2023–2024"]), "Figure: VMD NDVI 2023–2024")
    with b:
        show_map(find_map(images, MAPS["NDVI Change 2000–2023"]), "Figure: VMD NDVI Change 2000–2023")

    df = data["NDVI"]
    if df is not None:
        s = period_summary(df)

        if s is not None:
            st.markdown('<div class="section-subtitle">Regional NDVI Trend</div>', unsafe_allow_html=True)
            fig = line_chart(s, "Regional Mean NDVI Across Study Periods", "Mean NDVI")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("First supplied period", fmt(s.iloc[0]["Regional Mean"]))
            c2.metric("Latest supplied period", fmt(s.iloc[-1]["Regional Mean"]))
            c3.metric("Overall change", f"{s.iloc[-1]['Regional Mean'] - s.iloc[0]['Regional Mean']:+.3f}")

        st.markdown('<div class="section-subtitle">Province Comparison</div>', unsafe_allow_html=True)
        periods = get_periods(df)
        if periods:
            selected = st.selectbox("Select NDVI period", periods, key="ndvi_period")
            pc = province_col(df)
            mc = mean_col(df)
            sub = df[df[period_col(df)].astype(str) == str(selected)].copy()

            if pc and mc:
                sub["_mean"] = safe_numeric(sub, mc)
                sub = sub.sort_values("_mean", ascending=False)
                plot_df = sub[[pc, "_mean"]].dropna()

                if not plot_df.empty:
                    fig = px.bar(
                        plot_df,
                        x="_mean",
                        y=pc,
                        orientation="h",
                        title=f"Province-Level Mean NDVI — {selected}",
                        labels={"_mean": "Mean NDVI", pc: "Province"},
                    )
                    fig.update_layout(yaxis={"categoryorder": "total ascending"})
                    st.plotly_chart(chart_layout(fig, 500), use_container_width=True)

                display_cols = [c for c in [pc, mc, "min", "max", "stdDev"] if c in sub.columns]
                st.dataframe(sub[display_cols], use_container_width=True, hide_index=True)
                dataframe_download(sub[display_cols], f"VMD_NDVI_{norm(selected)}.csv")

    st.markdown('<div class="section-subtitle">Trend Statistics Maps</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        show_map(find_map(images, MAPS["Mann–Kendall S: NDVI & Rainfall"]),
                 "Figure: Mann–Kendall S statistic for NDVI and rainfall")
    with c2:
        show_map(find_map(images, MAPS["Sen's Slope"]),
                 "Figure: Sen's slope analysis")


elif page == "SPI-3 | Drought":

    st.markdown('<div class="section-title">SPI-3 | Drought Assessment</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note">SPI-3 represents precipitation/moisture stress over a three-month accumulation period. '
        'More negative values indicate stronger precipitation deficits relative to the reference distribution.</div>',
        unsafe_allow_html=True,
    )

    show_map(find_map(images, MAPS["SPI-3 Drought 2023–2024"]),
             "Figure: VMD SPI-3 Drought 2023–2024")

    df = data["SPI3"]
    if df is not None:
        s = period_summary(df)
        if s is not None:
            st.markdown('<div class="section-subtitle">Regional SPI-3 Temporal Pattern</div>', unsafe_allow_html=True)
            fig = line_chart(s, "Regional Mean SPI-3 Through Study Periods", "Mean SPI-3")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

            delta = s.iloc[-1]["Regional Mean"] - s.iloc[0]["Regional Mean"]
            c1, c2, c3 = st.columns(3)
            c1.metric("First supplied period", fmt(s.iloc[0]["Regional Mean"]))
            c2.metric("Latest supplied period", fmt(s.iloc[-1]["Regional Mean"]))
            c3.metric("Change", f"{delta:+.3f}")

        periods = get_periods(df)
        if periods:
            selected = st.selectbox("Select SPI-3 period", periods, key="spi_period")
            pc = province_col(df)
            mc = mean_col(df)
            sub = df[df[period_col(df)].astype(str) == str(selected)].copy()

            if pc and mc:
                sub["_mean"] = safe_numeric(sub, mc)
                sub = sub.sort_values("_mean", ascending=True)
                plot_df = sub[[pc, "_mean"]].dropna()

                if not plot_df.empty:
                    fig = px.bar(
                        plot_df,
                        x="_mean",
                        y=pc,
                        orientation="h",
                        title=f"Province-Level Mean SPI-3 — {selected}",
                        labels={"_mean": "Mean SPI-3", pc: "Province"},
                    )
                    st.plotly_chart(chart_layout(fig, 500), use_container_width=True)

                display_cols = [c for c in [pc, mc, "min", "max", "stdDev"] if c in sub.columns]
                st.dataframe(sub[display_cols], use_container_width=True, hide_index=True)
                dataframe_download(sub[display_cols], f"VMD_SPI3_{norm(selected)}.csv")

        st.markdown(
            '<div class="warning-note"><b>Interpretation note:</b> Any drought class labels used in the thesis '
            'should follow the specific SPI classification scheme documented in the methodology. '
            'The dashboard does not invent class thresholds.</div>',
            unsafe_allow_html=True,
        )

elif page == "Salinity & Vulnerability":

    st.markdown('<div class="section-title">Salinity, Drought Interaction & Agricultural Vulnerability</div>',
                unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        show_map(find_map(images, MAPS["Salinity Hotspots"]),
                 "Figure: VMD Salinity Hotspots 250 m")
    with c2:
        show_map(find_map(images, MAPS["Salinity–Drought Overlap"]),
                 "Figure: VMD Salinity–Drought Overlap 250 m")

    show_map(find_map(images, MAPS["Combined Agricultural Vulnerability"]),
             "Figure: VMD Combined Agricultural Vulnerability 250 m")

    sal = data["Salinity"]
    if sal is not None and area_col(sal):
        st.markdown('<div class="section-subtitle">Salinity Hotspot Area by Province</div>', unsafe_allow_html=True)
        s = sal.copy()
        pc = province_col(s)
        ac = area_col(s)
        s["_area"] = safe_numeric(s, ac)

        if pc:
            s = s.sort_values("_area", ascending=False)
            fig = px.bar(
                s.dropna(subset=["_area"]),
                x="_area",
                y=pc,
                orientation="h",
                title="Reported Salinity Hotspot Area by Province",
                labels={"_area": "Reported area sum", pc: "Province"},
            )
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(chart_layout(fig, 500), use_container_width=True)

            display = s[[pc, ac]].copy()
            st.dataframe(display, use_container_width=True, hide_index=True)
            dataframe_download(display, "VMD_Salinity_Hotspot_Area_Province.csv")

            total = s["_area"].sum()
            top3 = s.head(3)["_area"].sum() / total * 100 if total else None
            st.markdown(
                f'<div class="finding">The three largest provincial salinity totals account for approximately '
                f'<b>{top3:.1f}%</b> of the supplied reported total.</div>',
                unsafe_allow_html=True,
            )

    vul = data["Vulnerability"]
    if vul is not None and area_col(vul):
        st.markdown('<div class="section-subtitle">Agricultural Vulnerability Area by Province</div>',
                    unsafe_allow_html=True)
        v = vul.copy()
        pc = province_col(v)
        ac = area_col(v)
        v["_area"] = safe_numeric(v, ac)

        if pc:
            v = v.sort_values("_area", ascending=False)
            fig = px.bar(
                v.dropna(subset=["_area"]),
                x="_area",
                y=pc,
                orientation="h",
                title="Reported Agricultural Vulnerability Area by Province",
                labels={"_area": "Reported area sum", pc: "Province"},
            )
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(chart_layout(fig, 500), use_container_width=True)

            display = v[[pc, ac]].copy()
            st.dataframe(display, use_container_width=True, hide_index=True)
            dataframe_download(display, "VMD_Agricultural_Vulnerability_Area_Province.csv")

    st.markdown(
        '<div class="warning-note"><b>Area-unit caution:</b> The CSV field is treated as a reported '
        'area sum. The dashboard does not label it as hectares or another unit unless that unit is explicitly '
        'contained in the source data or methodology.</div>',
        unsafe_allow_html=True,
    )


elif page == "VSSI | Salinity Stress":

    st.markdown('<div class="section-title">VSSI | Vietnamese Salinity Stress Index</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note">The dashboard preserves the numerical scale contained in the supplied VSSI CSV. '
        'No unsupported 0–1 normalization is applied.</div>',
        unsafe_allow_html=True,
    )

    period_map = {
        "VSSI 2000": "2000",
        "VSSI 2010–2011": "2010_2011",
        "VSSI 2015–2016": "2015_2016",
        "VSSI 2019–2020": "2019_2020",
        "VSSI 2023–2024": "2023_2024",
    }

    selected = st.selectbox("Select VSSI period", list(period_map.keys()), key="vssi_period")
    show_map(find_map(images, MAPS[selected]), f"Figure: {selected}")

    df = data["VSSI"]
    if df is not None:
        pc = province_col(df)
        mc = mean_col(df)
        pcol = period_col(df)

        if pc and mc and pcol:
            desired = period_map[selected]
            sub = df[df[pcol].astype(str).map(norm) == norm(desired)].copy()
            if sub.empty:
                # fallback to contains
                sub = df[df[pcol].astype(str).str.replace("-", "_").str.lower() == desired.lower()].copy()

            if not sub.empty:
                sub["_mean"] = safe_numeric(sub, mc)
                sub = sub.sort_values("_mean", ascending=False)

                fig = px.bar(
                    sub.dropna(subset=["_mean"]),
                    x="_mean",
                    y=pc,
                    orientation="h",
                    title=f"Province-Level VSSI — {selected}",
                    labels={"_mean": "Mean VSSI", pc: "Province"},
                )
                st.plotly_chart(chart_layout(fig, 500), use_container_width=True)

                display_cols = [c for c in [pc, mc, "min", "max", "stdDev"] if c in sub.columns]
                st.dataframe(sub[display_cols], use_container_width=True, hide_index=True)
                dataframe_download(sub[display_cols], f"VMD_VSSI_{norm(desired)}.csv")

        s = period_summary(df)
        if s is not None:
            st.markdown('<div class="section-subtitle">Regional VSSI Temporal Pattern</div>', unsafe_allow_html=True)
            fig = line_chart(s, "Regional Mean VSSI Through Study Periods", "Mean VSSI")
            if fig:
                st.plotly_chart(fig, use_container_width=True)

            delta = s.iloc[-1]["Regional Mean"] - s.iloc[0]["Regional Mean"]
            c1, c2, c3 = st.columns(3)
            c1.metric("First supplied period", fmt(s.iloc[0]["Regional Mean"]))
            c2.metric("Latest supplied period", fmt(s.iloc[-1]["Regional Mean"]))
            c3.metric("Change", f"{delta:+.3f}")

    st.markdown('<div class="section-subtitle">VSSI Long-Term Change</div>', unsafe_allow_html=True)
    show_map(find_map(images, MAPS["VSSI Change 2000–2023"]),
             "Figure: VMD VSSI Change 2000–2023")

elif page == "Province Explorer":

    st.markdown('<div class="section-title">Province-Level Statistical Explorer</div>', unsafe_allow_html=True)

    indicator = st.selectbox(
        "Select indicator",
        ["NDVI", "SPI3", "VSSI", "Salinity", "Vulnerability"],
    )
    df = data[indicator]

    if df is None:
        st.error(f"{indicator} CSV was not detected in the bundled Output folder.")
    else:
        pc = province_col(df)
        if pc is None:
            st.error("Province field could not be identified.")
        else:
            provinces = sorted(df[pc].dropna().astype(str).unique().tolist())
            province = st.selectbox("Select province", provinces)
            sub = df[df[pc].astype(str) == province].copy()

            st.markdown(f'<div class="section-subtitle">{province}</div>', unsafe_allow_html=True)

            if indicator in ["NDVI", "SPI3", "VSSI"]:
                mc = mean_col(sub)
                pcol = period_col(sub)
                if mc and pcol:
                    sub["_mean"] = safe_numeric(sub, mc)

                    if indicator == "NDVI":
                        title = f"{province} — NDVI Temporal Profile"
                    elif indicator == "SPI3":
                        title = f"{province} — SPI-3 Temporal Profile"
                    else:
                        title = f"{province} — VSSI Temporal Profile"

                    fig = px.line(
                        sub,
                        x=pcol,
                        y="_mean",
                        markers=True,
                        title=title,
                        labels={"_mean": f"Mean {indicator}", pcol: "Period"},
                    )
                    st.plotly_chart(chart_layout(fig), use_container_width=True)

                    display_cols = [c for c in [pc, pcol, mc, "min", "max", "stdDev"] if c in sub.columns]
                    st.dataframe(sub[display_cols], use_container_width=True, hide_index=True)
                    dataframe_download(sub[display_cols], f"VMD_{indicator}_{norm(province)}.csv")

            else:
                ac = area_col(sub)
                if ac:
                    value = safe_numeric(sub, ac)
                    st.metric("Reported area sum", fmt(value.iloc[0] if len(value) else None, 3))
                    display_cols = [c for c in [pc, ac] if c in sub.columns]
                    st.dataframe(sub[display_cols], use_container_width=True, hide_index=True)
                    dataframe_download(sub[display_cols], f"VMD_{indicator}_{norm(province)}.csv")


elif page == "Research Findings":

    st.markdown('<div class="section-title">Research Findings | Thesis-Ready Evidence</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="note">This page converts the supplied CSV statistics into concise, '
        'research-oriented findings. Values are calculated from the source tables and are not manually entered.</div>',
        unsafe_allow_html=True,
    )

    ndvi = data["NDVI"]
    spi = data["SPI3"]
    vssi = data["VSSI"]
    sal = data["Salinity"]
    vul = data["Vulnerability"]

    # NDVI findings
    if ndvi is not None:
        s = period_summary(ndvi)
        pc = province_col(ndvi)
        mc = mean_col(ndvi)
        pcol = period_col(ndvi)

        if s is not None and len(s) >= 2:
            delta = s.iloc[-1]["Regional Mean"] - s.iloc[0]["Regional Mean"]
            direction = "increased" if delta > 0 else "decreased"
            st.markdown(
                f'<div class="finding"><b>NDVI finding:</b> Regional mean NDVI {direction} '
                f'from {fmt(s.iloc[0]["Regional Mean"])} in {s.iloc[0]["_period"]} '
                f'to {fmt(s.iloc[-1]["Regional Mean"])} in {s.iloc[-1]["_period"]}, '
                f'a net change of <b>{delta:+.3f}</b>.</div>',
                unsafe_allow_html=True,
            )

        if pc and mc and pcol:
            latest = ndvi[ndvi[pcol].astype(str) == ndvi[pcol].dropna().astype(str).iloc[-1]].copy()
            latest["_mean"] = safe_numeric(latest, mc)
            if not latest.empty:
                hi = latest.loc[latest["_mean"].idxmax()]
                lo = latest.loc[latest["_mean"].idxmin()]
                st.markdown(
                    f'<div class="finding"><b>Latest NDVI spatial contrast:</b> '
                    f'{hi[pc]} has the highest supplied provincial mean ({fmt(hi["_mean"])}), '
                    f'while {lo[pc]} has the lowest ({fmt(lo["_mean"])}).</div>',
                    unsafe_allow_html=True,
                )

    # SPI findings
    if spi is not None:
        s = period_summary(spi)
        if s is not None and len(s) >= 2:
            delta = s.iloc[-1]["Regional Mean"] - s.iloc[0]["Regional Mean"]
            st.markdown(
                f'<div class="finding"><b>SPI-3 finding:</b> Regional mean SPI-3 changed '
                f'from {fmt(s.iloc[0]["Regional Mean"])} to {fmt(s.iloc[-1]["Regional Mean"])} '
                f'between the first and latest supplied periods, a change of <b>{delta:+.3f}</b>. '
                f'The more negative latest value indicates greater precipitation/moisture stress '
                f'under the SPI interpretation.</div>',
                unsafe_allow_html=True,
            )

    # Salinity findings
    if sal is not None and area_col(sal):
        s = sal.copy()
        pc = province_col(s)
        ac = area_col(s)
        s["_area"] = safe_numeric(s, ac)
        if pc and not s.empty:
            s = s.sort_values("_area", ascending=False)
            total = s["_area"].sum()
            top3 = s.head(3)["_area"].sum() / total * 100 if total else None
            st.markdown(
                f'<div class="finding"><b>Salinity finding:</b> '
                f'{s.iloc[0][pc]} has the largest reported salinity hotspot area '
                f'({fmt(s.iloc[0]["_area"], 0)}). The three leading provinces together '
                f'represent approximately <b>{top3:.1f}%</b> of the supplied total.</div>',
                unsafe_allow_html=True,
            )

    # Vulnerability findings
    if vul is not None and area_col(vul):
        v = vul.copy()
        pc = province_col(v)
        ac = area_col(v)
        v["_area"] = safe_numeric(v, ac)
        if pc and not v.empty:
            v = v.sort_values("_area", ascending=False)
            total = v["_area"].sum()
            share = v.head(2)["_area"].sum() / total * 100 if total else None
            st.markdown(
                f'<div class="finding"><b>Vulnerability finding:</b> '
                f'{v.iloc[0][pc]} records the largest reported agricultural vulnerability '
                f'area ({fmt(v.iloc[0]["_area"], 0)}). The two leading provinces together '
                f'account for approximately <b>{share:.1f}%</b> of the supplied total.</div>',
                unsafe_allow_html=True,
            )

    # VSSI findings
    if vssi is not None:
        s = period_summary(vssi)
        if s is not None and len(s) >= 2:
            delta = s.iloc[-1]["Regional Mean"] - s.iloc[0]["Regional Mean"]
            st.markdown(
                f'<div class="finding"><b>VSSI finding:</b> The supplied VSSI statistics '
                f'show a regional mean change of <b>{delta:+.3f}</b> between the first and '
                f'latest periods. Because the source values are negative, the original CSV scale '
                f'is retained throughout the dashboard.</div>',
                unsafe_allow_html=True,
            )

    st.markdown('<div class="section-subtitle">Suggested Thesis Results Structure</div>', unsafe_allow_html=True)
    thesis_sections = [
        "4.1 Study-area and dataset overview",
        "4.2 NDVI spatial and temporal assessment",
        "4.3 Drought assessment using SPI-3",
        "4.4 Salinity hotspot distribution",
        "4.5 Salinity–drought interaction",
        "4.6 Combined agricultural vulnerability",
        "4.7 VSSI spatial and temporal assessment",
        "4.8 Integrated interpretation of agricultural vulnerability",
    ]
    for s in thesis_sections:
        st.markdown(f"• {s}")



elif page == "Data & Reproducibility":

    st.markdown('<div class="section-title">Data, Files & Reproducibility</div>', unsafe_allow_html=True)

    st.markdown(
        f'<div class="note"><b>Configured source folder:</b> {OUTPUT_FOLDER}<br>'
        f'The dashboard scans this folder recursively for maps and CSV files. '
        f'No data are uploaded to a remote service by this dashboard.</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-subtitle">Deployment Diagnostics</div>', unsafe_allow_html=True)
    diag = pd.DataFrame([
        {
            "Component": "Dashboard file",
            "Status": "Available",
            "Location": str(Path(__file__).resolve()),
        },
        {
            "Component": "Bundled Output folder",
            "Status": "Available" if OUTPUT_FOLDER.exists() else "Missing",
            "Location": str(OUTPUT_FOLDER),
        },
        {
            "Component": "Map files",
            "Status": f"{len(images)} detected",
            "Location": str(MAPS_FOLDER),
        },
        {
            "Component": "CSV files",
            "Status": f"{len(csvs)} detected",
            "Location": str(DATA_FOLDER),
        },
    ])
    st.dataframe(diag, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-subtitle">Core Dataset Status</div>', unsafe_allow_html=True)
    status_rows = []
    for key, wanted in CSV_NAMES.items():
        path = find_csv(csvs, wanted)
        df = data[key]
        status_rows.append({
            "Dataset": key,
            "Expected file": wanted,
            "Status": "Available" if df is not None else "Missing",
            "Rows": len(df) if df is not None else 0,
            "Columns": len(df.columns) if df is not None else 0,
            "Detected file": path.name if path else "—",
        })
    status_df = pd.DataFrame(status_rows)
    st.dataframe(status_df, use_container_width=True, hide_index=True)
    dataframe_download(status_df, "VMD_Core_Dataset_Status.csv")

    st.markdown('<div class="section-subtitle">Detected CSV Files</div>', unsafe_allow_html=True)
    csv_inventory = pd.DataFrame({
        "File": [p.name for p in csvs],
        "Path": [str(p) for p in csvs],
    })
    st.dataframe(csv_inventory, use_container_width=True, hide_index=True)

    st.markdown('<div class="section-subtitle">Methodological Notes</div>', unsafe_allow_html=True)
    st.markdown("""
    **NDVI:** vegetation-condition indicator used for spatial and temporal vegetation assessment.

    **SPI-3:** three-month precipitation/moisture stress indicator. More negative values indicate stronger precipitation deficits.

    **Salinity Hotspots:** spatial areas identified by the supplied GIS analysis as salinity affected.

    **Salinity–Drought Overlap:** spatial coincidence between the supplied salinity and drought conditions.

    **Combined Agricultural Vulnerability:** integrated spatial representation generated by the supplied analysis workflow.

    **VSSI:** salinity-stress indicator evaluated across five study periods. The supplied CSV numerical scale is retained.

    **Trend maps:** Mann–Kendall S and Sen's slope figures are displayed as supplied. The dashboard does not derive numerical trend statistics from map colours.

    **Area units:** reported area fields are displayed without assigning hectares, square kilometres or another unit unless the source dataset explicitly establishes that unit.
    """)

    st.markdown(
        '<div class="warning-note"><b>For final thesis reproducibility:</b> retain the original GIS '
        'processing workflow, source rasters, temporal compositing rules, projection/coordinate reference system, '
        'classification thresholds, masking rules and index formulas alongside the dashboard.</div>',
        unsafe_allow_html=True,
    )


st.markdown("---")
st.markdown(
    '<div class="footer">VMD Thesis GIS Dashboard · Vietnamese Mekong Delta Agricultural Vulnerability Assessment · '
    'Professional Research Interface</div>',
    unsafe_allow_html=True,
)
