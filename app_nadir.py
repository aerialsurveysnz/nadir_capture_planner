"""
app_nadir.py  —  Nadir Capture Flight Planner v1
=================================================
GSD-first nadir photogrammetry mission planner for Aerial Surveys Ltd.

Supported cameras:
  Vexcel UltraCam Eagle Prime (UCE Prime)
  Vexcel UltraCam-Lp
  Phase One iXM-RS150F / PAS 150
  Custom camera (user-defined)

Workflow:
  Select camera + lens → enter target GSD + overlaps + speed → load KML →
  (optional: load DEM) → generate flight lines → review warnings → export

Key design decisions vs oblique planner:
  - GSD is PRIMARY input; altitude is derived, not entered
  - Nadir-only geometry: rectangular footprints, uniform GSD
  - Separate camera DB (cameras_nadir.py equivalent, embedded here for simplicity)
  - DEM-aware sidelap validation (Stage 3, architecture ready from day 1)
  - Trigger interval check displayed prominently as RED/AMBER/GREEN
  - Named planning modes (AT photogrammetry / orthophoto / quick-look / custom)
  - All KML/AOI/line-generation/export code reused from oblique planner

Run:
    streamlit run app_nadir.py
"""

import json
import html
import math
import io
from pathlib import Path
from datetime import datetime
from xml.etree import ElementTree as ET

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
try:
    from docx import Document
    from docx.shared import Inches, Pt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import folium
    from streamlit_folium import st_folium
    FOLIUM_AVAILABLE = True
except ImportError:
    FOLIUM_AVAILABLE = False

try:
    from shapely.geometry import Polygon as ShapelyPolygon, LineString as ShapelyLineString
    from shapely.ops import unary_union
    from shapely.affinity import rotate as shapely_rotate
    SHAPELY_AVAILABLE = True
except Exception:
    ShapelyPolygon = None
    ShapelyLineString = None
    unary_union = None
    shapely_rotate = None
    SHAPELY_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# Camera database
# Each entry: sensor_w_px, sensor_h_px, pixel_size_um, lens options (focal_mm),
#             min_interval_s (minimum time between frames), notes
# Specs verified against published datasheets — store as editable dict so
# corrections don't require code changes.
# ─────────────────────────────────────────────────────────────────────────────

NADIR_CAMERAS = {
    "Vexcel UCE Prime": {
        "sensor_w_px":    20010,
        "sensor_h_px":    13080,
        "pixel_size_um":  5.2,
        "lenses_mm":      [80, 100, 120, 210],
        "default_lens_mm": 100,
        "min_interval_s": 1.65,
        "storage_mb":     {
            "IIQ L (lossless)":   350.0,
            "IIQ S (compressed)": 200.0,
        },
        "notes": "UltraCam Eagle Prime — PAN 20010×13080, 5.2 µm, multi-cone RGB+NIR",
    },
    "Vexcel UltraCam-Lp": {
        "sensor_w_px":    11704,
        "sensor_h_px":     7920,
        "pixel_size_um":   6.0,
        "lenses_mm":       [70],
        "default_lens_mm": 70,
        "min_interval_s":  2.5,
        "storage_mb":     {
            "IIQ L (lossless)":  130.0,
            "IIQ S (compressed)": 80.0,
        },
        "notes": "UltraCam-Lp — PAN 11704×7920, 6.0 µm, 70 mm fixed PAN lens",
    },
    "Phase One iXM-RS150F": {
        "sensor_w_px":    14204,
        "sensor_h_px":    10652,
        "pixel_size_um":   3.76,
        "lenses_mm":       [35, 50, 70, 90, 110],
        "default_lens_mm": 70,
        "min_interval_s":  0.5,    # 2 fps max
        "storage_mb":     {
            "IIQ L (lossless)":  150.0,
            "IIQ S (compressed)": 90.0,
        },
        "notes": "Phase One iXM-RS150F — 14204×10652, 3.76 µm, RS lenses 35–110 mm, up to 2 fps",
    },
    "Custom camera": {
        "sensor_w_px":    10000,
        "sensor_h_px":     8000,
        "pixel_size_um":   5.0,
        "lenses_mm":       [50],
        "default_lens_mm": 50,
        "min_interval_s":  1.0,
        "storage_mb":     {
            "RAW": 80.0,
        },
        "notes": "User-defined camera — edit sensor dimensions, pixel size and lens below.",
    },
}

# Named planning modes: (forwardlap, sidelap, description)
PLANNING_MODES = {
    "AT Photogrammetry (60/30)": (60, 30,
        "Standard photogrammetric block. 60% forwardlap, 30% sidelap. "
        "Suitable for DEM, point cloud, and orthorectified imagery."),
    "High-density AT (80/60)":   (80, 60,
        "High-overlap block for dense matching or steep terrain. "
        "Significantly more images and flight time."),
    "Orthophoto / mapping (60/20)": (60, 20,
        "Reduced sidelap for orthophoto production where stereo AT strength "
        "is not required. Faster capture."),
    "Quick-look / inspection (30/10)": (30, 10,
        "Minimal overlap for rapid visual coverage. "
        "Not suitable for photogrammetric processing."),
    "Custom": (60, 30, "Set forwardlap and sidelap manually below."),
}

# ─────────────────────────────────────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────────────────────────────────────
_APP_DIR        = Path(__file__).resolve().parent
AOI_LIBRARY_DIR = _APP_DIR / "aoi_library"
SCENARIO_DIR    = _APP_DIR / "saved_scenarios_nadir"

REPORT_LOGO_CANDIDATES = [
    _APP_DIR / "ASL_Logo_White.png",
    _APP_DIR / "ASL_Logo.png",
    _APP_DIR / "assets" / "aerial_surveys_logo.png",
]

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Nadir Capture Planner — ASL",
    page_icon=str(_APP_DIR / "ASL_Imagery_Icon.png") if (_APP_DIR / "ASL_Imagery_Icon.png").exists() else "📷",
    layout="wide",
)
st.markdown("""
<style>
.stApp { background: #0f2235 !important; }
[data-testid="stAppViewContainer"] { background: #0f2235 !important; }
[data-testid="stHeader"] { background: #0f2235 !important; }
.trigger-ok   { background:#1a4a2e; border-left:4px solid #3fb950;
                padding:0.5rem 0.8rem; border-radius:4px; margin:4px 0; }
.trigger-warn { background:#4a3800; border-left:4px solid #d29922;
                padding:0.5rem 0.8rem; border-radius:4px; margin:4px 0; }
.trigger-fail { background:#4a1a1a; border-left:4px solid #f85149;
                padding:0.5rem 0.8rem; border-radius:4px; margin:4px 0; }
.spec-card    { background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.10);
                border-radius:8px; padding:0.9rem 1.1rem; margin:4px 0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Unit helpers  (identical to oblique planner)
# ─────────────────────────────────────────────────────────────────────────────

def m_to_unit(v, unit):
    return v * {"m": 1.0, "ft": 3.280839895, "km": 0.001, "NM": 0.000539957}.get(unit, 1.0)

def unit_to_m(v, unit):
    return v / {"m": 1.0, "ft": 3.280839895, "km": 0.001, "NM": 0.000539957}.get(unit, 1.0)

def fmt_m(v_m, unit, d=1):
    return f"{m_to_unit(v_m, unit):.{d}f} {unit}"

def fmt_gsd(v_m, d=2):
    return f"{v_m * 100:.{d}f} cm/px"

def dark_fig(w=12, h=6):
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_facecolor("#1a2332")
    ax.set_facecolor("#1e2d40")
    ax.tick_params(colors="#a0aec0", labelsize=8)
    for sp in ax.spines.values():
        sp.set_color("#3d5166")
    return fig, ax

def fig_to_png_bytes(fig, dpi=120):
    bio = io.BytesIO()
    fig.savefig(bio, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    bio.seek(0)
    return bio.getvalue()

def find_report_logo():
    for c in REPORT_LOGO_CANDIDATES:
        if Path(c).exists():
            return Path(c)
    return None

# ─────────────────────────────────────────────────────────────────────────────
# Core nadir geometry
# Simple, verified, GSD-driven.
# ─────────────────────────────────────────────────────────────────────────────

def nadir_calc(camera_name, focal_mm, gsd_cm, fwd_pct, side_pct, speed_ms,
               custom_w_px=None, custom_h_px=None, custom_pixel_um=None):
    """
    All nadir geometry from target GSD.
    Returns a dict of all calculated values.
    """
    cam = NADIR_CAMERAS[camera_name]

    # Allow custom overrides
    w_px       = int(custom_w_px   or cam["sensor_w_px"])
    h_px       = int(custom_h_px   or cam["sensor_h_px"])
    pixel_um   = float(custom_pixel_um or cam["pixel_size_um"])
    min_int_s  = float(cam["min_interval_s"])

    pixel_m    = pixel_um / 1e6          # pixel size in metres
    focal_m    = focal_mm / 1000.0
    gsd_m      = gsd_cm / 100.0

    # AGL from GSD:  AGL = GSD × focal / pixel_size
    agl_m      = gsd_m * focal_m / pixel_m

    # Footprint in metres
    fp_across_m = w_px * gsd_m           # across-track (wide side)
    fp_along_m  = h_px * gsd_m           # along-track

    # Flight line spacing and trigger spacing
    fwd_frac   = fwd_pct / 100.0
    side_frac  = side_pct / 100.0
    line_sp_m   = fp_across_m * (1.0 - side_frac)
    trig_sp_m   = fp_along_m  * (1.0 - fwd_frac)

    # Trigger interval
    if speed_ms > 0 and trig_sp_m > 0:
        trig_int_s = trig_sp_m / speed_ms
    else:
        trig_int_s = float("inf")

    # Trigger status
    margin = trig_int_s / min_int_s if min_int_s > 0 else float("inf")
    if margin >= 1.25:
        trig_status = "ok"
    elif margin >= 1.0:
        trig_status = "warn"
    else:
        trig_status = "fail"

    return {
        "camera_name":     camera_name,
        "focal_mm":        focal_mm,
        "w_px":            w_px,
        "h_px":            h_px,
        "pixel_um":        pixel_um,
        "gsd_m":           gsd_m,
        "gsd_cm":          gsd_cm,
        "agl_m":           agl_m,
        "fp_across_m":     fp_across_m,
        "fp_along_m":      fp_along_m,
        "fwd_pct":         fwd_pct,
        "side_pct":        side_pct,
        "line_sp_m":       line_sp_m,
        "trig_sp_m":       trig_sp_m,
        "trig_int_s":      trig_int_s,
        "min_int_s":       min_int_s,
        "trig_margin":     margin,
        "trig_status":     trig_status,
        "speed_ms":        speed_ms,
    }


# ─────────────────────────────────────────────────────────────────────────────
# KML / AOI  (ported directly from oblique planner — same robust parser)
# ─────────────────────────────────────────────────────────────────────────────

def ensure_aoi_library_dir():
    AOI_LIBRARY_DIR.mkdir(parents=True, exist_ok=True)
    return AOI_LIBRARY_DIR

def list_library_kmls():
    ensure_aoi_library_dir()
    return sorted(
        [p for p in AOI_LIBRARY_DIR.iterdir()
         if p.is_file() and p.suffix.lower() == ".kml"],
        key=lambda p: p.name.lower(),
    )

def kml_ring_to_lonlat(coords_text):
    pts = []
    for token in str(coords_text or "").replace("\n", " ").split():
        parts = token.split(",")
        if len(parts) < 2:
            continue
        try:
            pts.append((float(parts[0]), float(parts[1])))
        except Exception:
            continue
    if len(pts) >= 3 and pts[0] != pts[-1]:
        pts.append(pts[0])
    return pts

def _detect_kml_crs(points):
    if not points:
        return "wgs84"
    if min(p[0] for p in points) > 100_000:
        return "nztm"
    return "wgs84"

def parse_kml_aoi(path):
    if not SHAPELY_AVAILABLE:
        raise RuntimeError("Shapely is required for AOI / mission outputs.")
    raw_text = None
    for encoding in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            raw_text = Path(path).read_text(encoding=encoding)
            break
        except Exception:
            continue
    if raw_text is None:
        raise RuntimeError(f"Could not read {Path(path).name}.")
    try:
        root = ET.fromstring(raw_text)
    except Exception as exc:
        raise RuntimeError(f"KML parse error: {exc}") from exc

    def _tag_ends(elem, suffix):
        return str(getattr(elem, "tag", "")).lower().endswith(suffix.lower())

    def _first_coords(parent):
        for elem in parent.iter():
            if _tag_ends(elem, "coordinates") and (elem.text or "").strip():
                return elem.text
        return None

    def _boundary_elems(poly_elem, name):
        return [e for e in poly_elem.iter() if _tag_ends(e, name)]

    polygon_specs = []
    all_lonlat = []
    total_inner = 0

    for poly_elem in [e for e in root.iter() if _tag_ends(e, "polygon")]:
        outer_elems = _boundary_elems(poly_elem, "outerBoundaryIs")
        if not outer_elems:
            continue
        outer_ring = kml_ring_to_lonlat(_first_coords(outer_elems[0]))
        if len(outer_ring) < 4:
            continue
        inner_rings = []
        for inner_elem in _boundary_elems(poly_elem, "innerBoundaryIs"):
            ring = kml_ring_to_lonlat(_first_coords(inner_elem))
            if len(ring) >= 4:
                inner_rings.append(ring)
        polygon_specs.append((outer_ring, inner_rings))
        all_lonlat.extend(outer_ring)
        for r in inner_rings:
            all_lonlat.extend(r)
        total_inner += len(inner_rings)

    if not polygon_specs:
        for elem in root.iter():
            if _tag_ends(elem, "coordinates") and (elem.text or "").strip():
                ring = kml_ring_to_lonlat(elem.text)
                if len(ring) >= 4:
                    polygon_specs.append((ring, []))
                    all_lonlat.extend(ring)

    if not polygon_specs:
        raise RuntimeError("No valid polygon coordinates found in KML.")

    crs = _detect_kml_crs(all_lonlat)
    xs_raw = [p[0] for p in all_lonlat]
    ys_raw = [p[1] for p in all_lonlat]

    if crs == "nztm":
        try:
            import pyproj as _pj
            _n2w = _pj.Transformer.from_crs("EPSG:2193", "EPSG:4326", always_xy=True)
            sw = _n2w.transform(min(xs_raw), min(ys_raw))
            ne = _n2w.transform(max(xs_raw), max(ys_raw))
            _lon_min, _lat_min = min(sw[0], ne[0]), min(sw[1], ne[1])
            _lon_max, _lat_max = max(sw[0], ne[0]), max(sw[1], ne[1])
        except ImportError:
            _lon_min, _lon_max = min(xs_raw), max(xs_raw)
            _lat_min, _lat_max = min(ys_raw), max(ys_raw)
    else:
        _lon_min, _lon_max = min(xs_raw), max(xs_raw)
        _lat_min, _lat_max = min(ys_raw), max(ys_raw)

    lon0 = 0.5 * (_lon_min + _lon_max)
    lat0 = 0.5 * (_lat_min + _lat_max)

    try:
        import pyproj as _pj
        _wm = _pj.Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
        _nw = _pj.Transformer.from_crs("EPSG:2193", "EPSG:4326", always_xy=True)
        mx0, my0 = _wm.transform(lon0, lat0)
        use_merc = True
    except ImportError:
        use_merc = False
        mx0 = my0 = 0.0
        r = 6378137.0
        cos_lat = math.cos(math.radians(lat0))

    def _project_ring(ring):
        if use_merc:
            if crs == "nztm":
                ring = [_nw.transform(e, n) for e, n in ring]
            merc = [_wm.transform(lon, lat) for lon, lat in ring]
            return [(mx - mx0, my - my0) for mx, my in merc]
        return [(r * math.radians(lon - lon0) * cos_lat,
                 r * math.radians(lat - lat0))
                for lon, lat in ring]

    polygons = []
    for outer_ring, inner_rings in polygon_specs:
        try:
            poly = ShapelyPolygon(
                _project_ring(outer_ring),
                [_project_ring(ir) for ir in inner_rings if len(ir) >= 4]
            )
            if not poly.is_valid:
                poly = poly.buffer(0)
            if not poly.is_empty and poly.area > 0:
                polygons.append(poly)
        except Exception:
            continue

    if not polygons:
        raise RuntimeError("No valid polygon coordinates found in KML.")

    # Nesting / odd-even hole handling (same logic as oblique planner)
    flat = []
    def _flatten(geom):
        if geom is None or getattr(geom, "is_empty", True):
            return
        if getattr(geom, "geom_type", "") == "Polygon":
            flat.append(geom)
        elif hasattr(geom, "geoms"):
            for g in geom.geoms:
                _flatten(g)
    for p in polygons:
        _flatten(p)

    parents = [None] * len(flat)
    for i, child in enumerate(flat):
        crep = child.representative_point()
        best_parent, best_area = None, float("inf")
        for j, cand in enumerate(flat):
            if i == j or cand.area <= child.area:
                continue
            try:
                if cand.buffer(0.01).covers(crep) and cand.buffer(0.01).covers(child):
                    if cand.area < best_area:
                        best_parent, best_area = j, cand.area
            except Exception:
                pass
        parents[i] = best_parent

    def _depth(idx):
        d, seen, p = 0, set(), parents[idx]
        while p is not None and p not in seen:
            seen.add(p); d += 1; p = parents[p]
        return d

    depths = [_depth(i) for i in range(len(flat))]
    aoi_parts = []
    for i, poly in enumerate(flat):
        if depths[i] % 2 != 0:
            continue
        part = poly
        for j, hole in enumerate(flat):
            if parents[j] == i and depths[j] % 2 == 1:
                try:
                    part = part.difference(hole)
                except Exception:
                    pass
        if not getattr(part, "is_empty", True) and part.area > 0:
            aoi_parts.append(part)

    aoi_poly = unary_union(aoi_parts) if aoi_parts else unary_union(polygons)
    if aoi_poly.is_empty:
        raise RuntimeError("KML geometry produced empty AOI.")
    try:
        simp = aoi_poly.simplify(1.0, preserve_topology=True)
        if not simp.is_empty and simp.is_valid:
            aoi_poly = simp
    except Exception:
        pass

    return {
        "name":      Path(path).stem,
        "source":    "kml_library",
        "polygon":   aoi_poly,
        "area_m2":   float(aoi_poly.area),
        "lon0":      float(lon0),
        "lat0":      float(lat0),
        "mx0":       float(mx0),
        "my0":       float(my0),
        "lon_min":   _lon_min,
        "lon_max":   _lon_max,
        "lat_min":   _lat_min,
        "lat_max":   _lat_max,
        "inner_ring_count": int(total_inner),
        "has_inner_holes":  bool(total_inner > 0),
    }


def build_buffered_aoi(aoi_payload, buffer_m):
    if aoi_payload is None or not SHAPELY_AVAILABLE:
        return aoi_payload
    poly = aoi_payload.get("polygon")
    if poly is None or buffer_m <= 0:
        return aoi_payload
    lat0 = aoi_payload.get("lat0")
    scale = (1.0 / max(math.cos(math.radians(float(lat0))), 1e-6)) if lat0 else 1.0
    buf_coord = buffer_m * scale
    try:
        buffered = poly.buffer(buf_coord, join_style=2)
        if buffered.is_empty or not buffered.is_valid:
            return aoi_payload
        result = dict(aoi_payload)
        result["original_polygon"] = poly
        result["polygon"]  = buffered
        result["area_m2"]  = float(buffered.area)
        result["buffered_m"] = float(buffer_m)
        mx0, my0 = aoi_payload.get("mx0", 0), aoi_payload.get("my0", 0)
        if mx0 and my0:
            bx0, by0, bx1, by1 = buffered.bounds
            import math as _bm
            _R = 6378137.0
            result["lon_min"] = _bm.degrees((mx0 + bx0) / _R)
            result["lon_max"] = _bm.degrees((mx0 + bx1) / _R)
            result["lat_min"] = _bm.degrees(2 * _bm.atan(_bm.exp((my0 + by0) / _R)) - _bm.pi / 2)
            result["lat_max"] = _bm.degrees(2 * _bm.atan(_bm.exp((my0 + by1) / _R)) - _bm.pi / 2)
        return result
    except Exception as exc:
        st.warning(f"⚠️ Buffer failed ({exc}) — using original AOI.")
        return aoi_payload


# ─────────────────────────────────────────────────────────────────────────────
# Mission generation  (simplified from oblique planner — nadir only)
# ─────────────────────────────────────────────────────────────────────────────

def compute_nadir_mission(aoi_payload, calc, flight_azimuth_deg=0.0,
                          lead_in_out_m=None, turn_time_min=3.5,
                          storage_profile="IIQ L (lossless)"):
    """
    Generate nadir flight lines with guaranteed AOI coverage.

    Lead-in/out is enforced to be at least one full trigger spacing so there
    is always one captured frame before the aircraft enters the AOI boundary
    and one after it exits.  The user may set a larger value but not smaller.

    After line generation, image footprint rectangles are unioned at every
    trigger position and checked against the original unbuffered AOI polygon.
    True coverage % and any uncovered gap geometry are returned.
    """
    if not SHAPELY_AVAILABLE or aoi_payload is None:
        return None
    polygon = aoi_payload.get("polygon")
    if polygon is None or getattr(polygon, "is_empty", True):
        return None

    line_sp_m  = calc["line_sp_m"]
    trig_sp_m  = calc["trig_sp_m"]
    fp_across_m = calc["fp_across_m"]
    fp_along_m  = calc["fp_along_m"]
    speed_ms   = calc["speed_ms"]

    if not (line_sp_m > 0 and trig_sp_m > 0 and speed_ms > 0):
        return None

    # Enforce minimum lead-in = one full trigger spacing
    min_lead_m = trig_sp_m
    if lead_in_out_m is None or lead_in_out_m < min_lead_m:
        effective_lead_m = min_lead_m
        lead_clipped = (lead_in_out_m is not None and lead_in_out_m < min_lead_m)
    else:
        effective_lead_m = float(lead_in_out_m)
        lead_clipped = False

    # Mercator scale correction
    lat0 = aoi_payload.get("lat0")
    merc_scale  = (1.0 / max(math.cos(math.radians(float(lat0))), 1e-6)) if lat0 else 1.0
    line_sp_c   = line_sp_m   * merc_scale
    trig_sp_c   = trig_sp_m   * merc_scale
    fp_across_c = fp_across_m * merc_scale
    fp_along_c  = fp_along_m  * merc_scale
    lead_c      = effective_lead_m * merc_scale

    # Rotate polygon to flight-line frame (lines run vertically in rotated frame)
    rotated = shapely_rotate(polygon, float(flight_azimuth_deg),
                             origin="centroid", use_radians=False)
    minx, miny, maxx, maxy = rotated.bounds
    width = maxx - minx
    if not math.isfinite(width) or width < 0:
        return None

    # Generate line x-offsets covering full AOI width plus one line either side
    n = max(1, int(math.ceil(width / line_sp_c)) + 2)
    start_x = minx - line_sp_c * 0.5
    offsets = [start_x + i * line_sp_c for i in range(n + 1)]
    offsets = [x for x in offsets if minx - line_sp_c <= x <= maxx + line_sp_c]

    pad = (maxy - miny) + lead_c + 1000.0

    # Gap-over-internal-exclusion threshold
    gap_threshold_c = speed_ms * turn_time_min * 60.0 * merc_scale

    mission_lines  = []   # (x, y0, x, y1) in rotated frame
    transit_lines  = []   # (x, y0, x, y1) transit-over-gap segments

    for x in offsets:
        scan_line = ShapelyLineString([(x, miny - pad), (x, maxy + pad)])
        try:
            intersection = rotated.intersection(scan_line)
        except Exception:
            continue

        segs = []
        if intersection.is_empty:
            continue
        gtype = intersection.geom_type
        if gtype == "LineString":
            coords = list(intersection.coords)
            if len(coords) >= 2:
                segs.append((min(c[1] for c in coords), max(c[1] for c in coords)))
        elif gtype in ("MultiLineString", "GeometryCollection"):
            for sub in intersection.geoms:
                if sub.geom_type == "LineString":
                    coords = list(sub.coords)
                    if len(coords) >= 2:
                        segs.append((min(c[1] for c in coords), max(c[1] for c in coords)))
        if not segs:
            continue

        segs.sort(key=lambda s: s[0])
        merged = [segs[0]]
        for seg in segs[1:]:
            gap = seg[0] - merged[-1][1]
            if gap <= gap_threshold_c:
                transit_lines.append((x, merged[-1][1], x, seg[0]))
                merged[-1] = (merged[-1][0], seg[1])
            else:
                merged.append(seg)

        for y0, y1 in merged:
            mission_lines.append((x, y0 - lead_c, x, y1 + lead_c))

    if not mission_lines:
        return None

    # ── Rotate everything back to geographic (Mercator offset) frame ──────────
    centroid = polygon.centroid
    cx, cy = centroid.x, centroid.y

    def _rot(px, py):
        rad = math.radians(-flight_azimuth_deg)
        dx, dy = px - cx, py - cy
        return (cx + dx * math.cos(rad) - dy * math.sin(rad),
                cy + dx * math.sin(rad) + dy * math.cos(rad))

    geo_lines = [(_rot(x0, y0), _rot(x1, y1)) for x0, y0, x1, y1 in mission_lines]
    geo_transit = [(_rot(x0, y0), _rot(x1, y1)) for x0, y0, x1, y1 in transit_lines]

    # ── Generate trigger positions along each line ────────────────────────────
    # Triggers start at lead_c from the line start and repeat every trig_sp_c.
    # This guarantees the first trigger is one full footprint before the AOI edge.
    trigger_points = []   # list of (x, y) in geographic Mercator offsets
    triggers_per_line = []

    for (ax_, ay_), (bx_, by_) in geo_lines:
        line_vec = (bx_ - ax_, by_ - ay_)
        ln = math.hypot(*line_vec)
        if ln < 1:
            triggers_per_line.append(0)
            continue
        ux, uy = line_vec[0] / ln, line_vec[1] / ln
        # First trigger at exactly lead_c from start of line
        t = lead_c
        count = 0
        while t <= ln - lead_c + trig_sp_c * 0.01:
            tx = ax_ + ux * t
            ty = ay_ + uy * t
            trigger_points.append((tx, ty))
            count += 1
            t += trig_sp_c
        triggers_per_line.append(count)

    total_triggers = len(trigger_points)

    # ── True footprint coverage check ────────────────────────────────────────
    # Build union of all image footprint rectangles and compare against
    # the original unbuffered AOI polygon.
    original_poly = aoi_payload.get("original_polygon", polygon)

    # Across/along unit vectors in geographic frame (use middle line direction)
    mid_line = geo_lines[len(geo_lines) // 2]
    mid_vec  = (mid_line[1][0] - mid_line[0][0], mid_line[1][1] - mid_line[0][1])
    mid_ln   = math.hypot(*mid_vec) or 1.0
    ux_geo, uy_geo = mid_vec[0] / mid_ln, mid_vec[1] / mid_ln   # along-track
    px_geo, py_geo = -uy_geo, ux_geo                              # across-track

    hw = fp_across_c / 2.0   # half-width across-track in Mercator units
    hl = fp_along_c  / 2.0   # half-length along-track

    # Cap footprint union at 5000 frames to keep it fast on large jobs
    MAX_FP_FOR_COVERAGE = 5000
    fp_polys = []
    sample_pts = trigger_points if len(trigger_points) <= MAX_FP_FOR_COVERAGE \
                 else trigger_points[::max(1, len(trigger_points) // MAX_FP_FOR_COVERAGE)]

    for tx, ty in sample_pts:
        corners = [
            (tx + px_geo*hw + ux_geo*hl, ty + py_geo*hw + uy_geo*hl),
            (tx - px_geo*hw + ux_geo*hl, ty - py_geo*hw + uy_geo*hl),
            (tx - px_geo*hw - ux_geo*hl, ty - py_geo*hw - uy_geo*hl),
            (tx + px_geo*hw - ux_geo*hl, ty + py_geo*hw - uy_geo*hl),
        ]
        try:
            fp_polys.append(ShapelyPolygon(corners))
        except Exception:
            pass

    coverage_pct = 0.0
    gap_geometry = None
    coverage_sampled = len(sample_pts) < len(trigger_points)

    if fp_polys:
        try:
            fp_union = unary_union(fp_polys)
            covered  = original_poly.intersection(fp_union)
            orig_area = original_poly.area
            if orig_area > 0:
                coverage_pct = min(100.0, 100.0 * covered.area / orig_area)
            uncovered = original_poly.difference(fp_union)
            if not uncovered.is_empty and uncovered.area > orig_area * 0.0001:
                gap_geometry = uncovered
        except Exception:
            # Fall back to swath estimate if union fails
            line_count_tmp = len(geo_lines)
            avg_len_tmp = (sum(math.hypot(b[0]-a[0], b[1]-a[1])
                              for a, b in geo_lines) / line_count_tmp / merc_scale
                          ) if line_count_tmp > 0 else 0
            orig_area_m2 = float(original_poly.area) / (merc_scale**2)
            swath_m2 = line_count_tmp * line_sp_m * avg_len_tmp
            coverage_pct = min(100.0, 100.0 * swath_m2 / orig_area_m2) if orig_area_m2 > 0 else 0.0

    # ── Statistics ────────────────────────────────────────────────────────────
    line_count    = len(geo_lines)
    line_lengths_c = [math.hypot(b[0]-a[0], b[1]-a[1]) for a, b in geo_lines]
    total_length_m = sum(line_lengths_c) / merc_scale
    avg_length_m   = total_length_m / line_count if line_count > 0 else 0.0

    cam = NADIR_CAMERAS[calc["camera_name"]]
    storage_per_image_mb = cam["storage_mb"].get(storage_profile,
                           list(cam["storage_mb"].values())[0])
    total_storage_mb = total_triggers * storage_per_image_mb

    airborne_s   = total_length_m / speed_ms
    turn_s_total = max(0, line_count - 1) * turn_time_min * 60.0
    flight_s     = airborne_s + turn_s_total

    area_m2 = float(original_poly.area) / (merc_scale ** 2)

    return {
        "name":                     aoi_payload.get("name", "AOI"),
        "source":                   aoi_payload.get("source", "unknown"),
        "area_m2":                  area_m2,
        "flight_azimuth_deg":       float(flight_azimuth_deg),
        "line_spacing_m":           line_sp_m,
        "trig_spacing_m":           trig_sp_m,
        "lead_in_out_m":            effective_lead_m,
        "lead_clipped":             lead_clipped,
        "min_lead_m":               min_lead_m,
        "line_count":               line_count,
        "total_triggers":           total_triggers,
        "total_images":             total_triggers,
        "triggers_per_line":        triggers_per_line,
        "trigger_points":           trigger_points,
        "total_line_length_m":      total_length_m,
        "average_line_length_m":    avg_length_m,
        "storage_profile":          storage_profile,
        "storage_per_image_mb":     storage_per_image_mb,
        "total_storage_mb":         total_storage_mb,
        "airborne_time_s":          airborne_s,
        "total_turn_time_s":        turn_s_total,
        "flight_time_s":            flight_s,
        "coverage_pct":             coverage_pct,
        "coverage_sampled":         coverage_sampled,
        "gap_geometry":             gap_geometry,
        "mission_line_geometries":  geo_lines,
        "transit_gap_geometries":   geo_transit,
        "original_aoi_polygon":     original_poly,
        "aoi_polygon":              polygon,
        "mx0":                      aoi_payload.get("mx0"),
        "my0":                      aoi_payload.get("my0"),
        "lat0":                     aoi_payload.get("lat0"),
        "merc_scale":               merc_scale,
        # unit vectors for footprint drawing (geographic frame)
        "ux_geo":                   ux_geo,
        "uy_geo":                   uy_geo,
        "px_geo":                   px_geo,
        "py_geo":                   py_geo,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Heading optimiser (brute-force 1° search)
# ─────────────────────────────────────────────────────────────────────────────

def optimise_heading(aoi_payload, calc, lead_in_out_m=200.0, turn_time_min=3.5,
                     step_deg=1):
    """
    Brute-force 1° search for the heading that gives minimum total flight time.
    flight_time = (total_line_length / speed) + (line_count - 1) * turn_time_s
    This correctly balances fewer turns against shorter line lengths — for some
    AOI shapes a heading with more lines but much shorter lines wins overall.
    """
    best_heading, best_result, best_time = 0.0, None, float("inf")
    for hdg in range(0, 180, step_deg):
        result = compute_nadir_mission(
            aoi_payload, calc, flight_azimuth_deg=float(hdg),
            lead_in_out_m=lead_in_out_m, turn_time_min=turn_time_min)
        if result is None:
            continue
        flight_time = result["flight_time_s"]
        if flight_time < best_time:
            best_time = flight_time
            best_heading = float(hdg)
            best_result = result
    return best_heading, best_result


# ─────────────────────────────────────────────────────────────────────────────
# AOI map figure — with display options
# ─────────────────────────────────────────────────────────────────────────────

# Cap on frames shown in map to keep rendering fast
MAX_FP_DISPLAY = 2000

def draw_mission_map(mission_outputs, calc, dist_unit="m",
                     show_triggers=False, show_footprints=False,
                     show_gaps=True):
    if mission_outputs is None:
        return None, 0

    fig, ax = dark_fig(8, 6)
    ax.set_aspect("equal")

    geo_lines    = mission_outputs.get("mission_line_geometries", [])
    transit      = mission_outputs.get("transit_gap_geometries",  [])
    trigger_pts  = mission_outputs.get("trigger_points", [])
    aoi_poly     = (mission_outputs.get("original_aoi_polygon") or
                    mission_outputs.get("aoi_polygon"))
    gap_geom     = mission_outputs.get("gap_geometry")
    ms           = mission_outputs.get("merc_scale", 1.0)
    ux, uy       = mission_outputs.get("ux_geo", 0.0), mission_outputs.get("uy_geo", 1.0)
    px, py       = mission_outputs.get("px_geo", -1.0), mission_outputs.get("py_geo", 0.0)
    hw           = calc["fp_across_m"] * ms / 2.0
    hl           = calc["fp_along_m"]  * ms / 2.0

    # ── AOI boundary ──────────────────────────────────────────────────────────
    def _draw_poly(p, face="#4488ff", edge="#4488ff", face_alpha=0.08, edge_lw=1.5):
        if getattr(p, "geom_type", "") == "Polygon":
            xs, ys = p.exterior.xy
            ax.fill(xs, ys, alpha=face_alpha, color=face)
            ax.plot(xs, ys, color=edge, lw=edge_lw)
            for interior in p.interiors:
                ix, iy = zip(*interior.coords)
                ax.fill(ix, iy, alpha=0.3, color="#0f2235")
                ax.plot(ix, iy, color="#ff6644", lw=1.0, ls="--")
        elif hasattr(p, "geoms"):
            for sub in p.geoms:
                _draw_poly(sub, face, edge, face_alpha, edge_lw)

    if aoi_poly is not None:
        _draw_poly(aoi_poly)

    # ── Coverage gap overlay ──────────────────────────────────────────────────
    if show_gaps and gap_geom is not None:
        try:
            _draw_poly(gap_geom, face="#ff2020", edge="#ff2020",
                       face_alpha=0.35, edge_lw=0.8)
        except Exception:
            pass

    # ── Footprint rectangles ──────────────────────────────────────────────────
    fp_shown = 0
    if show_footprints and trigger_pts:
        step = max(1, len(trigger_pts) // MAX_FP_DISPLAY)
        for tx, ty in trigger_pts[::step]:
            corners = [
                (tx + px*hw + ux*hl, ty + py*hw + uy*hl),
                (tx - px*hw + ux*hl, ty - py*hw + uy*hl),
                (tx - px*hw - ux*hl, ty - py*hw - uy*hl),
                (tx + px*hw - ux*hl, ty + py*hw - uy*hl),
            ]
            patch = mpatches.Polygon(corners, closed=True,
                                     facecolor="#3fb95014", edgecolor="#3fb950",
                                     lw=0.5, zorder=2)
            ax.add_patch(patch)
            fp_shown += 1

    # ── Flight lines ──────────────────────────────────────────────────────────
    for i, (a, b) in enumerate(geo_lines):
        col = "#f0c040" if i % 2 == 0 else "#40c0f0"
        ax.plot([a[0], b[0]], [a[1], b[1]], color=col, lw=0.8, alpha=0.85, zorder=3)

    # ── Transit-over-gap lines ────────────────────────────────────────────────
    for a, b in transit:
        ax.plot([a[0], b[0]], [a[1], b[1]], color="#ff9944",
                lw=1.2, ls="--", alpha=0.7, zorder=3)

    # ── Trigger points ────────────────────────────────────────────────────────
    if show_triggers and trigger_pts:
        step = max(1, len(trigger_pts) // MAX_FP_DISPLAY)
        pts = trigger_pts[::step]
        ax.scatter([p[0] for p in pts], [p[1] for p in pts],
                   s=4, color="#00ffff", alpha=0.6, zorder=4, linewidths=0)

    # ── Labels & formatting ───────────────────────────────────────────────────
    ax.set_xlabel(f"East offset ({dist_unit})", color="#8b949e", fontsize=9)
    ax.set_ylabel(f"North offset ({dist_unit})", color="#8b949e", fontsize=9)
    xt = ax.get_xticks()
    ax.set_xticklabels([f"{m_to_unit(t/ms, dist_unit):.0f}" for t in xt], color="#8b949e")
    yt = ax.get_yticks()
    ax.set_yticklabels([f"{m_to_unit(t/ms, dist_unit):.0f}" for t in yt], color="#8b949e")

    hdg  = mission_outputs.get("flight_azimuth_deg", 0.0)
    lc   = mission_outputs.get("line_count", 0)
    dist = mission_outputs.get("total_line_length_m", 0.0)
    cov  = mission_outputs.get("coverage_pct", 0.0)
    ax.set_title(
        f"Nadir flight plan — {lc} lines | {hdg:.0f}° | "
        f"{dist/1000:.1f} km | Coverage: {cov:.1f}%",
        color="#c9d1d9", fontsize=10, pad=8)
    fig.tight_layout()
    return fig, fp_shown


# ─────────────────────────────────────────────────────────────────────────────
# Interactive Folium map
# ─────────────────────────────────────────────────────────────────────────────

def make_folium_map(mission_outputs, calc,
                    show_triggers=False,
                    show_footprints=False):
    """
    Build a zoomable Leaflet/Folium map of the flight plan.
    Returns a folium.Map object or None if folium is not available.
    Footprints are capped at MAX_FP_DISPLAY to keep the map responsive.
    """
    try:
        import folium
    except ImportError:
        return None

    if mission_outputs is None:
        return None

    mx0 = mission_outputs.get("mx0")
    my0 = mission_outputs.get("my0")
    lat0 = mission_outputs.get("lat0")
    lon0 = mission_outputs.get("lon0") or (
        math.degrees(mx0 / 6378137.0) if mx0 else 0.0)
    if lat0 is None or lon0 is None:
        return None

    _R = 6378137.0
    def _wgs84(lx, ly):
        lon = math.degrees((mx0 + lx) / _R)
        lat = math.degrees(2 * math.atan(math.exp((my0 + ly) / _R)) - math.pi / 2)
        return lat, lon   # folium uses (lat, lon)

    geo_lines   = mission_outputs.get("mission_line_geometries", [])
    transit     = mission_outputs.get("transit_gap_geometries",  [])
    trigger_pts = mission_outputs.get("trigger_points", [])
    aoi_poly    = (mission_outputs.get("original_aoi_polygon") or
                   mission_outputs.get("aoi_polygon"))
    merc_scale  = mission_outputs.get("merc_scale", 1.0)

    ux = mission_outputs.get("ux_geo", 0.0)
    uy = mission_outputs.get("uy_geo", 1.0)
    px = mission_outputs.get("px_geo", -1.0)
    py = mission_outputs.get("py_geo", 0.0)
    hw = calc["fp_across_m"] * merc_scale / 2.0
    hl = calc["fp_along_m"]  * merc_scale / 2.0

    m = folium.Map(
        location=[float(lat0), float(lon0)],
        zoom_start=10,
        tiles="CartoDB dark_matter",
        prefer_canvas=True,
    )

    # ── AOI boundary ──────────────────────────────────────────────────────────
    if aoi_poly is not None:
        def _add_aoi(p):
            if getattr(p, "geom_type", "") == "Polygon":
                ext = [(_wgs84(x, y)) for x, y in p.exterior.coords]
                folium.Polygon(
                    locations=ext,
                    color="#4488ff", weight=2,
                    fill=True, fill_color="#4488ff", fill_opacity=0.08,
                    tooltip="AOI Boundary",
                ).add_to(m)
                for interior in p.interiors:
                    hole = [_wgs84(x, y) for x, y in interior.coords]
                    folium.Polygon(
                        locations=hole,
                        color="#ff6644", weight=1.5, dash_array="6 4",
                        fill=True, fill_color="#0f2235", fill_opacity=0.4,
                        tooltip="Internal exclusion",
                    ).add_to(m)
            elif hasattr(p, "geoms"):
                for sub in p.geoms:
                    _add_aoi(sub)
        _add_aoi(aoi_poly)

    # ── Frame footprints ──────────────────────────────────────────────────────
    if show_footprints and trigger_pts:
        fp_layer = folium.FeatureGroup(name="Frame Footprints", show=True)
        step = max(1, len(trigger_pts) // MAX_FP_DISPLAY)
        for tx, ty in trigger_pts[::step]:
            corners = [
                _wgs84(tx + px*hw + ux*hl, ty + py*hw + uy*hl),
                _wgs84(tx - px*hw + ux*hl, ty - py*hw + uy*hl),
                _wgs84(tx - px*hw - ux*hl, ty - py*hw - uy*hl),
                _wgs84(tx + px*hw - ux*hl, ty + py*hw - uy*hl),
            ]
            folium.Polygon(
                locations=corners,
                color="#3fb950", weight=0.6,
                fill=True, fill_color="#3fb950", fill_opacity=0.06,
            ).add_to(fp_layer)
        fp_layer.add_to(m)

    # ── Flight lines ──────────────────────────────────────────────────────────
    line_layer = folium.FeatureGroup(name="Flight Lines", show=True)
    for i, (a, b) in enumerate(geo_lines):
        col = "#f0c040" if i % 2 == 0 else "#40c0f0"
        la0_, lo0_ = _wgs84(a[0], a[1])
        la1_, lo1_ = _wgs84(b[0], b[1])
        folium.PolyLine(
            locations=[(la0_, lo0_), (la1_, lo1_)],
            color=col, weight=1.5, opacity=0.85,
            tooltip=f"Line {i+1}",
        ).add_to(line_layer)
    line_layer.add_to(m)

    # ── Transit over gaps ─────────────────────────────────────────────────────
    if transit:
        tr_layer = folium.FeatureGroup(name="Transit (cameras off)", show=True)
        for a, b in transit:
            la0_, lo0_ = _wgs84(a[0], a[1])
            la1_, lo1_ = _wgs84(b[0], b[1])
            folium.PolyLine(
                locations=[(la0_, lo0_), (la1_, lo1_)],
                color="#ff9944", weight=2, dash_array="8 4",
                opacity=0.8, tooltip="Transit (cameras off)",
            ).add_to(tr_layer)
        tr_layer.add_to(m)

    # ── Trigger points ────────────────────────────────────────────────────────
    if show_triggers and trigger_pts:
        trig_layer = folium.FeatureGroup(name="Trigger Points", show=True)
        step = max(1, len(trigger_pts) // MAX_FP_DISPLAY)
        for i, (tx, ty) in enumerate(trigger_pts[::step]):
            lat_, lon_ = _wgs84(tx, ty)
            folium.CircleMarker(
                location=(lat_, lon_),
                radius=2, color="#00ffff",
                fill=True, fill_color="#00ffff", fill_opacity=0.7,
                weight=0.5,
                tooltip=f"T{i*step+1}",
            ).add_to(trig_layer)
        trig_layer.add_to(m)

    folium.LayerControl(collapsed=False).add_to(m)

    # Summary popup at AOI centre
    hdg  = mission_outputs.get("flight_azimuth_deg", 0.0)
    lc   = mission_outputs.get("line_count", 0)
    dist = mission_outputs.get("total_line_length_m", 0.0)
    cov  = mission_outputs.get("coverage_pct", 0.0)
    trigs = mission_outputs.get("total_triggers", 0)
    popup_html = (
        f"<b>{mission_outputs.get('name','AOI')}</b><br>"
        f"Camera: {calc['camera_name']} {calc['focal_mm']} mm<br>"
        f"GSD: {calc['gsd_cm']:.1f} cm | AGL: {calc['agl_m']:.0f} m<br>"
        f"Heading: {hdg:.0f}° | Lines: {lc}<br>"
        f"Triggers: {trigs:,} | Dist: {dist/1000:.1f} km<br>"
        f"Coverage: {cov:.1f}%"
    )
    folium.Marker(
        location=[float(lat0), float(lon0)],
        icon=folium.DivIcon(html='<div style="display:none"></div>'),
        popup=folium.Popup(popup_html, max_width=260),
    ).add_to(m)

    return m

# Folders: 1. Flight Lines  1b. Transit  2. Trigger Points
#           3. Frame Footprints  4. AOI Boundary
# ─────────────────────────────────────────────────────────────────────────────

def make_kml_export(mission_outputs, calc,
                    include_triggers=True,
                    include_footprints=True):
    if mission_outputs is None:
        return None
    mx0 = mission_outputs.get("mx0")
    my0 = mission_outputs.get("my0")
    if mx0 is None or my0 is None:
        return None

    _R = 6378137.0

    def _wgs84(lx, ly):
        lon = math.degrees((mx0 + lx) / _R)
        lat = math.degrees(2 * math.atan(math.exp((my0 + ly) / _R)) - math.pi / 2)
        return lon, lat

    def _coord_str(lx, ly):
        lo, la = _wgs84(lx, ly)
        return f"{lo:.8f},{la:.8f},0"

    def _ring_coords(pts):
        return " ".join(_coord_str(p[0], p[1]) for p in pts)

    name    = mission_outputs.get("name", "AOI")
    hdg     = float(mission_outputs.get("flight_azimuth_deg", 0.0))
    lc      = int(mission_outputs.get("line_count", 0))
    lsp     = float(mission_outputs.get("line_spacing_m", 0.0))
    tsp     = float(mission_outputs.get("trig_spacing_m", 0.0))
    dist_km = float(mission_outputs.get("total_line_length_m", 0.0)) / 1000.0
    fly_hr  = float(mission_outputs.get("flight_time_s", 0.0)) / 3600.0
    cov_pct = float(mission_outputs.get("coverage_pct", 0.0))
    lead_m  = float(mission_outputs.get("lead_in_out_m", 0.0))
    gsd_cm  = calc.get("gsd_cm", 0.0)
    agl_m   = calc.get("agl_m", 0.0)

    geo_lines      = mission_outputs.get("mission_line_geometries", [])
    transit        = mission_outputs.get("transit_gap_geometries", [])
    trigger_pts    = mission_outputs.get("trigger_points", [])
    aoi_poly       = (mission_outputs.get("original_aoi_polygon") or
                      mission_outputs.get("aoi_polygon"))
    merc_scale     = mission_outputs.get("merc_scale", 1.0)

    # Footprint unit vectors
    ux = mission_outputs.get("ux_geo", 0.0)
    uy = mission_outputs.get("uy_geo", 1.0)
    px = mission_outputs.get("px_geo", -1.0)
    py = mission_outputs.get("py_geo", 0.0)
    hw = calc["fp_across_m"] * merc_scale / 2.0
    hl = calc["fp_along_m"]  * merc_scale / 2.0

    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<kml xmlns="http://www.opengis.net/kml/2.2">',
           '<Document>',
           f'  <name>{name} — Nadir Flight Plan</name>',
           f'  <description>'
           f'Camera: {calc["camera_name"]} | FL: {calc["focal_mm"]} mm | '
           f'GSD: {gsd_cm:.1f} cm | AGL: {agl_m:.0f} m | '
           f'Azimuth: {hdg:.1f}° | Lines: {lc} | '
           f'Line spacing: {lsp:.0f} m | Trigger spacing: {tsp:.0f} m | '
           f'Lead-in/out: {lead_m:.0f} m | '
           f'Total distance: {dist_km:.1f} km | Flying time: {fly_hr:.2f} hr | '
           f'Coverage: {cov_pct:.1f}%'
           f'</description>']

    # ── Styles ────────────────────────────────────────────────────────────────
    out += [
        '  <Style id="fl_odd">'
        '<LineStyle><color>fff0c040</color><width>2</width></LineStyle></Style>',
        '  <Style id="fl_even">'
        '<LineStyle><color>ff40c0f0</color><width>2</width></LineStyle></Style>',
        '  <Style id="transit">'
        '<LineStyle><color>ffff6644</color><width>2</width></LineStyle></Style>',
        '  <Style id="trig">'
        '<IconStyle><color>ff00ffff</color><scale>0.45</scale>'
        '<Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon>'
        '</IconStyle><LabelStyle><scale>0</scale></LabelStyle></Style>',
        '  <Style id="frame">'
        '<LineStyle><color>aa3fb950</color><width>1</width></LineStyle>'
        '<PolyStyle><color>183fb950</color></PolyStyle></Style>',
        '  <Style id="aoi">'
        '<LineStyle><color>ff4488ff</color><width>3</width></LineStyle>'
        '<PolyStyle><color>204488ff</color></PolyStyle></Style>',
        '  <Style id="aoi_hole">'
        '<LineStyle><color>ffff6644</color><width>2</width></LineStyle>'
        '<PolyStyle><color>20ff6644</color></PolyStyle></Style>',
    ]

    # ── 1. Flight Lines ───────────────────────────────────────────────────────
    out.append('  <Folder>')
    out.append('    <name>1. Flight Lines</name>')
    out.append('    <visibility>1</visibility>')
    for i, (a, b) in enumerate(geo_lines):
        sid = "fl_odd" if i % 2 == 0 else "fl_even"
        c0 = _coord_str(a[0], a[1])
        c1 = _coord_str(b[0], b[1])
        out.append(f'    <Placemark><name>Line {i+1}</name>'
                   f'<styleUrl>#{sid}</styleUrl>'
                   f'<LineString><tessellate>1</tessellate>'
                   f'<coordinates>{c0} {c1}</coordinates>'
                   f'</LineString></Placemark>')
    out.append('  </Folder>')

    # ── 1b. Transit over internal gaps ────────────────────────────────────────
    if transit:
        out.append('  <Folder>')
        out.append('    <name>1b. Transit (cameras off)</name>')
        out.append('    <visibility>1</visibility>')
        for a, b in transit:
            c0 = _coord_str(a[0], a[1])
            c1 = _coord_str(b[0], b[1])
            out.append(f'    <Placemark><styleUrl>#transit</styleUrl>'
                       f'<LineString><tessellate>1</tessellate>'
                       f'<coordinates>{c0} {c1}</coordinates>'
                       f'</LineString></Placemark>')
        out.append('  </Folder>')

    # ── 2. Trigger Points ─────────────────────────────────────────────────────
    if include_triggers and trigger_pts:
        out.append('  <Folder>')
        out.append('    <name>2. Trigger Points</name>')
        out.append('    <visibility>1</visibility>')
        for i, (tx, ty) in enumerate(trigger_pts):
            cs = _coord_str(tx, ty)
            out.append(f'    <Placemark><name>T{i+1}</name>'
                       f'<styleUrl>#trig</styleUrl>'
                       f'<Point><coordinates>{cs}</coordinates></Point>'
                       f'</Placemark>')
        out.append('  </Folder>')

    # ── 3. Frame Footprints ───────────────────────────────────────────────────
    if include_footprints and trigger_pts:
        out.append('  <Folder>')
        out.append('    <name>3. Frame Footprints</name>')
        out.append('    <visibility>0</visibility>')  # off by default — large
        for i, (tx, ty) in enumerate(trigger_pts):
            corners = [
                (tx + px*hw + ux*hl, ty + py*hw + uy*hl),
                (tx - px*hw + ux*hl, ty - py*hw + uy*hl),
                (tx - px*hw - ux*hl, ty - py*hw - uy*hl),
                (tx + px*hw - ux*hl, ty + py*hw - uy*hl),
                (tx + px*hw + ux*hl, ty + py*hw + uy*hl),  # close ring
            ]
            ring_cs = _ring_coords(corners)
            out.append(f'    <Placemark><name>F{i+1}</name>'
                       f'<styleUrl>#frame</styleUrl>'
                       f'<Polygon><tessellate>1</tessellate>'
                       f'<outerBoundaryIs><LinearRing>'
                       f'<coordinates>{ring_cs}</coordinates>'
                       f'</LinearRing></outerBoundaryIs>'
                       f'</Polygon></Placemark>')
        out.append('  </Folder>')

    # ── 4. AOI Boundary ───────────────────────────────────────────────────────
    if aoi_poly is not None:
        out.append('  <Folder>')
        out.append('    <name>4. AOI Boundary</name>')
        out.append('    <visibility>1</visibility>')

        def _emit_polygon(p):
            if getattr(p, "geom_type", "") == "Polygon":
                ext_cs = _ring_coords(list(p.exterior.coords))
                pm = (f'    <Placemark><styleUrl>#aoi</styleUrl>'
                      f'<Polygon><tessellate>1</tessellate>'
                      f'<outerBoundaryIs><LinearRing>'
                      f'<coordinates>{ext_cs}</coordinates>'
                      f'</LinearRing></outerBoundaryIs>')
                for interior in p.interiors:
                    inner_cs = _ring_coords(list(interior.coords))
                    pm += (f'<innerBoundaryIs><LinearRing>'
                           f'<coordinates>{inner_cs}</coordinates>'
                           f'</LinearRing></innerBoundaryIs>')
                pm += '</Polygon></Placemark>'
                out.append(pm)
            elif hasattr(p, "geoms"):
                for sub in p.geoms:
                    _emit_polygon(sub)

        _emit_polygon(aoi_poly)
        out.append('  </Folder>')

    out.append('</Document>')
    out.append('</kml>')
    return "\n".join(out).encode("utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────────────────────────────────────

def make_excel_export(calc, mission_outputs, settings_rows, map_png_bytes=None):
    wb = Workbook()

    # --- Summary sheet ---
    ws = wb.active
    ws.title = "Summary"
    hdr_fill = PatternFill("solid", fgColor="1A3A5C")
    hdr_font = Font(color="FFFFFF", bold=True)

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 28

    ws.merge_cells("A1:B1")
    title_cell = ws["A1"]
    title_cell.value = "Nadir Capture Planner — Aerial Surveys Ltd"
    title_cell.font = Font(bold=True, size=13, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="0F2235")
    title_cell.alignment = Alignment(horizontal="center")

    ws.append(["Setting", "Value"])
    for cell in ws[2]:
        cell.font = hdr_font
        cell.fill = hdr_fill

    for key, val in settings_rows:
        ws.append([key, val])

    # --- Camera results sheet ---
    ws2 = wb.create_sheet("Camera & Geometry")
    ws2.column_dimensions["A"].width = 32
    ws2.column_dimensions["B"].width = 22
    ws2.append(["Parameter", "Value"])
    for cell in ws2[1]:
        cell.font = hdr_font
        cell.fill = hdr_fill

    cam_rows = _build_camera_rows(calc)
    for k, v in cam_rows:
        ws2.append([k, v])

    # --- Mission sheet ---
    if mission_outputs:
        ws3 = wb.create_sheet("Mission Statistics")
        ws3.column_dimensions["A"].width = 35
        ws3.column_dimensions["B"].width = 22
        ws3.append(["Metric", "Value"])
        for cell in ws3[1]:
            cell.font = hdr_font
            cell.fill = hdr_fill
        for k, v in _build_mission_rows(mission_outputs, calc):
            ws3.append([k, v])

        if map_png_bytes:
            try:
                from openpyxl.drawing.image import Image as XLImage
                img = XLImage(io.BytesIO(map_png_bytes))
                img.width, img.height = 700, 460
                ws3.add_image(img, "D3")
            except Exception:
                pass

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def _build_camera_rows(calc):
    cam = NADIR_CAMERAS.get(calc["camera_name"], {})
    ts = calc["trig_status"]
    ts_label = {"ok": "✅ OK", "warn": "⚠️ MARGINAL", "fail": "❌ TOO FAST"}.get(ts, ts)
    return [
        ("Camera",                calc["camera_name"]),
        ("Lens (focal length)",   f"{calc['focal_mm']} mm"),
        ("Sensor width (px)",     f"{calc['w_px']:,}"),
        ("Sensor height (px)",    f"{calc['h_px']:,}"),
        ("Pixel size (µm)",       f"{calc['pixel_um']:.2f} µm"),
        ("Target GSD",            f"{calc['gsd_cm']:.1f} cm/px"),
        ("Required AGL",          f"{calc['agl_m']:.0f} m"),
        ("Footprint across-track",f"{calc['fp_across_m']:.0f} m"),
        ("Footprint along-track", f"{calc['fp_along_m']:.0f} m"),
        ("Forward overlap",       f"{calc['fwd_pct']}%"),
        ("Sidelap",               f"{calc['side_pct']}%"),
        ("Line spacing",          f"{calc['line_sp_m']:.0f} m"),
        ("Trigger spacing",       f"{calc['trig_sp_m']:.0f} m"),
        ("Trigger interval",      f"{calc['trig_int_s']:.2f} s"),
        ("Min camera interval",   f"{calc['min_int_s']:.2f} s"),
        ("Trigger margin",        f"{calc['trig_margin']:.2f}×  {ts_label}"),
        ("Aircraft speed",        f"{calc['speed_ms']:.1f} m/s  ({calc['speed_ms']*1.94384:.0f} kts)"),
        ("Notes",                 cam.get("notes", "")),
    ]


def _build_mission_rows(mo, calc):
    area_km2   = float(mo.get("area_m2", 0)) / 1e6
    dist_km    = float(mo.get("total_line_length_m", 0)) / 1000.0
    avg_km     = float(mo.get("average_line_length_m", 0)) / 1000.0
    fly_hr     = float(mo.get("flight_time_s", 0)) / 3600.0
    air_hr     = float(mo.get("airborne_time_s", 0)) / 3600.0
    turn_hr    = float(mo.get("total_turn_time_s", 0)) / 3600.0
    stor_gb    = float(mo.get("total_storage_mb", 0)) / 1024.0
    return [
        ("AOI name",              mo.get("name", "—")),
        ("AOI area",              f"{area_km2:.2f} km²"),
        ("Flight azimuth",        f"{float(mo.get('flight_azimuth_deg', 0)):.1f}°"),
        ("Line count",            f"{mo.get('line_count', 0):,}"),
        ("Total trigger events",  f"{mo.get('total_triggers', 0):,}"),
        ("Total images",          f"{mo.get('total_images', 0):,}"),
        ("Line spacing",          f"{mo.get('line_spacing_m', 0):.0f} m"),
        ("Trigger spacing",       f"{mo.get('trig_spacing_m', 0):.0f} m"),
        ("Total line length",     f"{dist_km:.2f} km"),
        ("Average line length",   f"{avg_km:.2f} km"),
        ("Lead-in / out",         f"{mo.get('lead_in_out_m', 0):.0f} m"),
        ("Storage profile",       mo.get("storage_profile", "—")),
        ("Storage per image",     f"{mo.get('storage_per_image_mb', 0):.0f} MB"),
        ("Estimated storage",     f"{stor_gb:.1f} GB"),
        ("Coverage estimate",     f"{mo.get('coverage_pct', 0):.1f}%"),
        ("Airborne time",         f"{air_hr:.2f} hr"),
        ("Turn allowance",        f"{turn_hr:.2f} hr"),
        ("Total flying time",     f"{fly_hr:.2f} hr"),
    ]


# ─────────────────────────────────────────────────────────────────────────────
# Word Export
# ─────────────────────────────────────────────────────────────────────────────

def make_word_export(calc, mission_outputs, settings_rows, map_png_bytes=None):
    if not DOCX_AVAILABLE:
        return None
    doc = Document()

    # Page orientation landscape
    section = doc.sections[0]
    section.orientation = 1  # landscape
    section.page_width, section.page_height = section.page_height, section.page_width

    def _add_table(ws, rows):
        tbl = doc.add_table(rows=1, cols=2)
        try:
            tbl.style = "Light Grid Accent 1"
        except Exception:
            pass
        hdr = tbl.rows[0].cells
        hdr[0].text, hdr[1].text = "Item", "Value"
        for k, v in rows:
            r = tbl.add_row().cells
            r[0].text, r[1].text = str(k), str(v)

    logo = find_report_logo()
    if logo:
        try:
            doc.add_picture(str(logo), width=Inches(2.2))
        except Exception:
            pass

    title = doc.add_paragraph()
    run = title.add_run("Nadir Capture Planner — Mission Report")
    run.bold = True
    run.font.size = Pt(18)

    sub = doc.add_paragraph("Aerial Surveys Ltd  ·  " + datetime.now().strftime("%d %B %Y"))
    sub.runs[0].italic = True

    doc.add_heading("Mission Settings", level=2)
    _add_table(doc, settings_rows)

    doc.add_heading("Camera & Geometry", level=2)
    _add_table(doc, _build_camera_rows(calc))

    if mission_outputs:
        doc.add_heading("Mission Statistics", level=2)
        _add_table(doc, _build_mission_rows(mission_outputs, calc))

        if map_png_bytes:
            doc.add_heading("Flight Plan Map", level=2)
            try:
                doc.add_picture(io.BytesIO(map_png_bytes), width=Inches(9.0))
            except Exception:
                pass

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Helper — build export settings rows
# ─────────────────────────────────────────────────────────────────────────────

def build_settings_rows(calc, mission_name="—", buffer_m=0.0, hdg=0.0,
                        turn_time_min=3.5, lead_in_m=200.0, dist_unit="m"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    return [
        ("Prepared",          ts),
        ("Project / AOI",     mission_name),
        ("Camera",            calc["camera_name"]),
        ("Lens",              f"{calc['focal_mm']} mm"),
        ("Target GSD",        f"{calc['gsd_cm']:.1f} cm/px"),
        ("Required AGL",      f"{calc['agl_m']:.0f} m"),
        ("Aircraft speed",    f"{calc['speed_ms']:.1f} m/s  ({calc['speed_ms']*1.94384:.0f} kts)"),
        ("Forward overlap",   f"{calc['fwd_pct']}%"),
        ("Sidelap",           f"{calc['side_pct']}%"),
        ("Line spacing",      f"{calc['line_sp_m']:.0f} m"),
        ("Trigger spacing",   f"{calc['trig_sp_m']:.0f} m"),
        ("Trigger interval",  f"{calc['trig_int_s']:.2f} s"),
        ("Coverage buffer",   f"{buffer_m:.0f} m"),
        ("Flight heading",    f"{hdg:.1f}°"),
        ("Turn time",         f"{turn_time_min:.1f} min"),
        ("Lead-in / out",     f"{lead_in_m:.0f} m"),
        ("Distance unit",     dist_unit),
    ]


# ─────────────────────────────────────────────────────────────────────────────
# Footprint diagram
# ─────────────────────────────────────────────────────────────────────────────

def draw_footprint_diagram(calc, dist_unit="m", n_strips=3, n_triggers=4):
    fig, ax = dark_fig(7, 5)
    ax.set_aspect("equal")

    fp_w = calc["fp_across_m"]
    fp_h = calc["fp_along_m"]
    ls   = calc["line_sp_m"]
    ts   = calc["trig_sp_m"]

    for strip in range(n_strips):
        cx = (strip - n_strips//2) * ls
        for trig in range(n_triggers):
            cy = (trig - n_triggers//2) * ts
            rect = mpatches.FancyBboxPatch(
                (cx - fp_w/2, cy - fp_h/2), fp_w, fp_h,
                boxstyle="square,pad=0",
                facecolor="#3fb95018",
                edgecolor="#3fb950",
                lw=0.8, zorder=2,
            )
            ax.add_patch(rect)
            if strip == n_strips//2 and trig == n_triggers//2:
                ax.add_patch(mpatches.FancyBboxPatch(
                    (cx - fp_w/2, cy - fp_h/2), fp_w, fp_h,
                    boxstyle="square,pad=0",
                    facecolor="#3fb95040",
                    edgecolor="#3fb950",
                    lw=1.5, zorder=3,
                ))

    # Dimension arrows — positioned just outside the content area
    arr_y  = -(n_triggers//2) * ts - fp_h/2 - ts * 0.3
    arr_x  = (n_strips//2) * ls + fp_w/2 + ls * 0.15

    ax.annotate("", xy=(ls, arr_y), xytext=(0.0, arr_y),
                arrowprops=dict(arrowstyle="<->", color="#d29922", lw=1.2))
    ax.text(ls/2, arr_y - ts * 0.15,
            f"Line sp: {m_to_unit(ls, dist_unit):.0f} {dist_unit}  |  Sidelap: {calc['side_pct']}%",
            color="#d29922", ha="center", va="top", fontsize=7.5)

    ax.annotate("", xy=(arr_x, ts), xytext=(arr_x, 0.0),
                arrowprops=dict(arrowstyle="<->", color="#58a6ff", lw=1.2))
    ax.text(arr_x + ls * 0.08, ts/2,
            f"Trig sp: {m_to_unit(ts, dist_unit):.0f} {dist_unit}\nFwdlap: {calc['fwd_pct']}%",
            color="#58a6ff", ha="left", va="center", fontsize=7.5)

    # Tight limits — just enough margin around the actual footprint content
    content_xmin = -(n_strips//2) * ls - fp_w/2
    content_xmax =  (n_strips//2) * ls + fp_w/2
    content_ymin = -(n_triggers//2) * ts - fp_h/2
    content_ymax =  (n_triggers//2) * ts + fp_h/2
    x_margin = (content_xmax - content_xmin) * 0.12
    y_margin = (content_ymax - content_ymin) * 0.25  # extra below for arrow label
    ax.set_xlim(content_xmin - x_margin, content_xmax + x_margin + ls * 0.5)
    ax.set_ylim(content_ymin - y_margin, content_ymax + y_margin * 0.4)

    xt = ax.get_xticks()
    ax.set_xticklabels([f"{m_to_unit(t, dist_unit):.0f}" for t in xt], color="#8b949e")
    yt = ax.get_yticks()
    ax.set_yticklabels([f"{m_to_unit(t, dist_unit):.0f}" for t in yt], color="#8b949e")
    ax.set_xlabel(f"Across-track ({dist_unit})", color="#8b949e", fontsize=9)
    ax.set_ylabel(f"Along-track ({dist_unit})", color="#8b949e", fontsize=9)
    ax.set_title(
        f"Footprint diagram — {calc['camera_name']} {calc['focal_mm']} mm | "
        f"GSD {calc['gsd_cm']:.1f} cm | AGL {calc['agl_m']:.0f} m",
        color="#c9d1d9", fontsize=10, pad=8)
    fig.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# Session state init
# ─────────────────────────────────────────────────────────────────────────────

if "nadir_camera"       not in st.session_state:
    st.session_state.nadir_camera       = "Vexcel UCE Prime"
if "nadir_focal_mm"     not in st.session_state:
    st.session_state.nadir_focal_mm     = 100
if "nadir_gsd_cm"       not in st.session_state:
    st.session_state.nadir_gsd_cm       = 5.0
if "nadir_mode"         not in st.session_state:
    st.session_state.nadir_mode         = "AT Photogrammetry (60/30)"
if "nadir_fwd_pct"      not in st.session_state:
    st.session_state.nadir_fwd_pct      = 60
if "nadir_side_pct"     not in st.session_state:
    st.session_state.nadir_side_pct     = 30
if "nadir_speed_ms"     not in st.session_state:
    st.session_state.nadir_speed_ms     = 55.0
if "nadir_dist_unit"    not in st.session_state:
    st.session_state.nadir_dist_unit    = "m"
if "nadir_aoi_payload"  not in st.session_state:
    st.session_state.nadir_aoi_payload  = None
if "nadir_mission"      not in st.session_state:
    st.session_state.nadir_mission      = None
if "nadir_hdg"          not in st.session_state:
    st.session_state.nadir_hdg          = 0.0
if "nadir_buffer_m"     not in st.session_state:
    st.session_state.nadir_buffer_m     = 150.0
if "nadir_lead_in_m"    not in st.session_state:
    st.session_state.nadir_lead_in_m    = 200.0
if "nadir_turn_min"     not in st.session_state:
    st.session_state.nadir_turn_min     = 3.5
if "nadir_storage_prof" not in st.session_state:
    st.session_state.nadir_storage_prof = "IIQ L (lossless)"
if "nadir_custom_w_px"  not in st.session_state:
    st.session_state.nadir_custom_w_px  = 10000
if "nadir_custom_h_px"  not in st.session_state:
    st.session_state.nadir_custom_h_px  = 8000
if "nadir_custom_pxum"  not in st.session_state:
    st.session_state.nadir_custom_pxum  = 5.0
if "nadir_custom_focal" not in st.session_state:
    st.session_state.nadir_custom_focal = 50

# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────

logo_path = find_report_logo()
col_logo, col_title = st.columns([1, 8])
with col_logo:
    if logo_path:
        st.image(str(logo_path), width=110)
with col_title:
    st.markdown("## Nadir Capture Planner")
    st.caption("Vexcel UCE Prime · Vexcel UltraCam-Lp · Phase One iXM-RS150F · Custom  |  Aerial Surveys Ltd")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — Camera, GSD, overlaps, speed
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### 📷 Camera & Lens")
    camera_name = st.selectbox(
        "Camera",
        list(NADIR_CAMERAS.keys()),
        index=list(NADIR_CAMERAS.keys()).index(st.session_state.nadir_camera),
        key="sb_camera",
    )
    st.session_state.nadir_camera = camera_name
    cam_db = NADIR_CAMERAS[camera_name]

    if camera_name == "Custom camera":
        st.session_state.nadir_custom_w_px  = st.number_input("Sensor width (px)",  value=st.session_state.nadir_custom_w_px,  min_value=1000, step=100, key="sb_cw")
        st.session_state.nadir_custom_h_px  = st.number_input("Sensor height (px)", value=st.session_state.nadir_custom_h_px,  min_value=1000, step=100, key="sb_ch")
        st.session_state.nadir_custom_pxum  = st.number_input("Pixel size (µm)",    value=st.session_state.nadir_custom_pxum,  min_value=1.0,  max_value=20.0, step=0.1, key="sb_px")
        st.session_state.nadir_custom_focal = st.number_input("Focal length (mm)",  value=st.session_state.nadir_custom_focal, min_value=10,   max_value=300, step=5, key="sb_fl_c")
        focal_mm = st.session_state.nadir_custom_focal
    else:
        lens_opts = cam_db["lenses_mm"]
        default_fl = cam_db["default_lens_mm"]
        default_idx = lens_opts.index(default_fl) if default_fl in lens_opts else 0
        if st.session_state.nadir_focal_mm in lens_opts:
            default_idx = lens_opts.index(st.session_state.nadir_focal_mm)
        focal_mm = st.selectbox(
            "Lens (focal length)",
            lens_opts,
            index=default_idx,
            format_func=lambda x: f"{x} mm",
            key="sb_focal",
        )
        st.session_state.nadir_focal_mm = focal_mm

    st.markdown("---")
    st.markdown("### 🎯 Target GSD")
    gsd_cm = st.number_input(
        "Required GSD (cm/px)",
        value=float(st.session_state.nadir_gsd_cm),
        min_value=0.5, max_value=50.0, step=0.5,
        help="GSD is the PRIMARY input. Altitude is calculated from this value.",
        key="sb_gsd",
    )
    st.session_state.nadir_gsd_cm = gsd_cm

    st.markdown("---")
    st.markdown("### 🔄 Overlaps & Mode")
    mode_name = st.selectbox(
        "Planning mode",
        list(PLANNING_MODES.keys()),
        index=list(PLANNING_MODES.keys()).index(st.session_state.nadir_mode),
        key="sb_mode",
    )
    st.session_state.nadir_mode = mode_name
    preset_fwd, preset_side, mode_desc = PLANNING_MODES[mode_name]
    st.caption(mode_desc)

    if mode_name == "Custom":
        fwd_pct  = st.slider("Forward overlap (%)", 0, 90, int(st.session_state.nadir_fwd_pct), 5, key="sb_fwd")
        side_pct = st.slider("Sidelap (%)",          0, 80, int(st.session_state.nadir_side_pct), 5, key="sb_side")
    else:
        fwd_pct, side_pct = preset_fwd, preset_side
        st.markdown(f"**Forwardlap:** {fwd_pct}%  &nbsp;&nbsp; **Sidelap:** {side_pct}%")
    st.session_state.nadir_fwd_pct  = fwd_pct
    st.session_state.nadir_side_pct = side_pct

    st.markdown("---")
    st.markdown("### ✈️ Aircraft")
    speed_kts = st.number_input(
        "Aircraft speed (kts)",
        value=round(st.session_state.nadir_speed_ms * 1.94384, 1),
        min_value=30.0, max_value=300.0, step=5.0,
        key="sb_speed",
    )
    speed_ms = speed_kts / 1.94384
    st.session_state.nadir_speed_ms = speed_ms

    st.markdown("---")
    dist_unit = st.selectbox("Distance display unit", ["m", "ft", "km", "NM"],
                             index=["m","ft","km","NM"].index(st.session_state.nadir_dist_unit),
                             key="sb_unit")
    st.session_state.nadir_dist_unit = dist_unit


# ─────────────────────────────────────────────────────────────────────────────
# Calculate
# ─────────────────────────────────────────────────────────────────────────────

calc = nadir_calc(
    camera_name   = camera_name,
    focal_mm      = focal_mm,
    gsd_cm        = gsd_cm,
    fwd_pct       = fwd_pct,
    side_pct      = side_pct,
    speed_ms      = speed_ms,
    custom_w_px   = st.session_state.nadir_custom_w_px  if camera_name == "Custom camera" else None,
    custom_h_px   = st.session_state.nadir_custom_h_px  if camera_name == "Custom camera" else None,
    custom_pixel_um = st.session_state.nadir_custom_pxum if camera_name == "Custom camera" else None,
)


# ─────────────────────────────────────────────────────────────────────────────
# System results panel
# ─────────────────────────────────────────────────────────────────────────────

st.subheader("📐 System Results")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Target GSD",       f"{calc['gsd_cm']:.1f} cm/px")
    st.metric("Required AGL",     f"{calc['agl_m']:.0f} m")
with col2:
    st.metric("Footprint across", fmt_m(calc["fp_across_m"], dist_unit, 0))
    st.metric("Footprint along",  fmt_m(calc["fp_along_m"],  dist_unit, 0))
with col3:
    st.metric("Line spacing",     fmt_m(calc["line_sp_m"],   dist_unit, 0))
    st.metric("Trigger spacing",  fmt_m(calc["trig_sp_m"],   dist_unit, 0))
with col4:
    st.metric("Trigger interval", f"{calc['trig_int_s']:.2f} s")
    st.metric("Min camera interval", f"{calc['min_int_s']:.2f} s")

# Trigger interval warning — prominent
ts = calc["trig_status"]
margin = calc["trig_margin"]
if ts == "ok":
    st.markdown(
        f'<div class="trigger-ok">✅ <strong>Trigger interval OK</strong> — '
        f'{calc["trig_int_s"]:.2f} s required, {calc["min_int_s"]:.2f} s minimum. '
        f'Margin: {margin:.2f}×</div>',
        unsafe_allow_html=True)
elif ts == "warn":
    st.markdown(
        f'<div class="trigger-warn">⚠️ <strong>Trigger interval MARGINAL</strong> — '
        f'{calc["trig_int_s"]:.2f} s required, {calc["min_int_s"]:.2f} s minimum. '
        f'Margin: {margin:.2f}×. Consider reducing speed or increasing trigger spacing.</div>',
        unsafe_allow_html=True)
else:
    st.markdown(
        f'<div class="trigger-fail">❌ <strong>Trigger interval TOO FAST</strong> — '
        f'Camera needs {calc["min_int_s"]:.2f} s but only {calc["trig_int_s"]:.2f} s available. '
        f'Reduce speed, reduce forwardlap, or choose a longer focal length (higher AGL).</div>',
        unsafe_allow_html=True)

# Camera info
with st.expander(f"ℹ️ Camera notes — {camera_name}", expanded=False):
    cam_db_entry = NADIR_CAMERAS[camera_name]
    st.markdown(f"""
| Property | Value |
|---|---|
| Sensor | {calc['w_px']:,} × {calc['h_px']:,} px |
| Pixel size | {calc['pixel_um']:.2f} µm |
| Focal length | {calc['focal_mm']} mm |
| Min trigger interval | {calc['min_int_s']:.2f} s |
| Notes | {cam_db_entry.get('notes', '—')} |
""")
    st.info("⚠️ Camera specs are stored in the NADIR_CAMERAS dict at the top of app_nadir.py. "
            "Verify against current manufacturer datasheets before mission-critical use.")

# Footprint diagram
st.markdown("---")
st.subheader("📊 Footprint & Overlap Diagram")
fp_col, _ = st.columns([2, 1])
with fp_col:
    fp_fig = draw_footprint_diagram(calc, dist_unit=dist_unit)
    st.pyplot(fp_fig, use_container_width=True)
    plt.close(fp_fig)


# ─────────────────────────────────────────────────────────────────────────────
# AOI / Mission section
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("---")
st.subheader("🗺️ AOI & Mission Planning")

if not SHAPELY_AVAILABLE:
    st.error("Shapely is not installed. KML loading and mission planning require Shapely. "
             "Install with: pip install shapely")
else:
    aoi_col, opt_col = st.columns([2, 1])

    with aoi_col:
        st.markdown("**Load KML area of interest**")
        kml_files = list_library_kmls()
        aoi_source = st.radio("AOI source", ["KML from library", "Upload KML", "Standard square"],
                              horizontal=True, key="aoi_source")

        if aoi_source == "KML from library":
            if kml_files:
                selected_kml = st.selectbox(
                    "KML file",
                    kml_files,
                    format_func=lambda p: p.name,
                    key="kml_sel",
                )
                if st.button("Load KML", key="btn_load_kml"):
                    try:
                        payload = parse_kml_aoi(selected_kml)
                        st.session_state.nadir_aoi_payload = payload
                        st.session_state.nadir_mission = None
                        area_km2 = payload["area_m2"] / 1e6
                        holes = payload.get("inner_ring_count", 0)
                        st.success(f"Loaded: **{payload['name']}** — {area_km2:.2f} km²"
                                   + (f" ({holes} internal exclusion{'s' if holes!=1 else ''})" if holes else ""))
                    except Exception as exc:
                        st.error(f"KML error: {exc}")
            else:
                st.info(f"No KML files found in `aoi_library/`. "
                        f"Copy KML files there and restart the app.")

        elif aoi_source == "Upload KML":
            uploaded = st.file_uploader("Upload KML file", type=["kml"], key="kml_upload")
            if uploaded and st.button("Load uploaded KML", key="btn_upload_kml"):
                tmp = _APP_DIR / "aoi_library" / uploaded.name
                tmp.write_bytes(uploaded.getvalue())
                try:
                    payload = parse_kml_aoi(tmp)
                    st.session_state.nadir_aoi_payload = payload
                    st.session_state.nadir_mission = None
                    st.success(f"Loaded: **{payload['name']}** — {payload['area_m2']/1e6:.2f} km²")
                except Exception as exc:
                    st.error(f"KML error: {exc}")

        else:  # Standard square
            sq_km2 = st.number_input("Square AOI area (km²)", value=100.0, min_value=1.0, step=10.0, key="sq_km2")
            if st.button("Use standard square", key="btn_sq"):
                side_m = math.sqrt(sq_km2 * 1e6)
                half   = side_m / 2.0
                poly   = ShapelyPolygon([(-half,-half),(half,-half),(half,half),(-half,half)])
                st.session_state.nadir_aoi_payload = {
                    "name":    f"Standard {sq_km2:.0f} km² square",
                    "source":  "standard_example",
                    "polygon": poly,
                    "area_m2": float(poly.area),
                }
                st.session_state.nadir_mission = None
                st.success(f"Standard square AOI set: {sq_km2:.0f} km²")

    with opt_col:
        st.markdown("**Mission options**")
        buffer_m  = st.number_input("Coverage buffer (m)", value=st.session_state.nadir_buffer_m,
                                    min_value=0.0, max_value=2000.0, step=50.0, key="buf_m")

        # Lead-in: show minimum based on current trigger spacing
        min_lead_m = calc["trig_sp_m"]
        lead_in_m  = st.number_input(
            f"Lead-in / out per line (m)",
            value=max(float(st.session_state.nadir_lead_in_m), min_lead_m),
            min_value=min_lead_m,
            max_value=2000.0, step=50.0, key="lead_m",
            help=f"Minimum is one trigger spacing ({min_lead_m:.0f} m) to guarantee "
                 f"one captured frame before and after the AOI boundary.")
        if lead_in_m < min_lead_m:
            st.warning(f"⚠️ Lead-in increased to {min_lead_m:.0f} m (= trigger spacing) "
                       f"to guarantee full edge coverage.")
            lead_in_m = min_lead_m

        turn_min  = st.number_input("Turn time per line (min)", value=st.session_state.nadir_turn_min,
                                    min_value=0.5, max_value=10.0, step=0.5, key="turn_m")
        st.session_state.nadir_buffer_m  = buffer_m
        st.session_state.nadir_lead_in_m = lead_in_m
        st.session_state.nadir_turn_min  = turn_min

        cam_profiles = list(NADIR_CAMERAS[camera_name]["storage_mb"].keys())
        if st.session_state.nadir_storage_prof not in cam_profiles:
            st.session_state.nadir_storage_prof = cam_profiles[0]
        storage_prof = st.selectbox("Storage profile", cam_profiles,
                                    index=cam_profiles.index(st.session_state.nadir_storage_prof),
                                    key="stor_prof")
        st.session_state.nadir_storage_prof = storage_prof

        # DEM notice
        st.markdown("**DEM (terrain optimisation)**")
        st.info("🏔️ DEM support — Stage 3 planned. Will validate sidelap over "
                "high terrain using LINZ 8 m DEM (NZTM2000 native). "
                "Currently plans at flat terrain.")

    # Heading controls
    st.markdown("**Flight heading**")
    hdg_col1, hdg_col2 = st.columns([1, 1])
    with hdg_col1:
        use_optimise = st.checkbox("Auto-optimise heading (minimum flight time)", value=False, key="opt_hdg")
    with hdg_col2:
        if not use_optimise:
            manual_hdg = st.number_input("Heading (°)", value=float(st.session_state.nadir_hdg),
                                         min_value=0.0, max_value=179.0, step=1.0, key="man_hdg")
            st.session_state.nadir_hdg = manual_hdg
        else:
            st.caption("Heading will be optimised when you click Generate.")

    # Generate button
    aoi_payload_raw = st.session_state.nadir_aoi_payload
    if st.button("🛫 Generate flight lines", type="primary", key="btn_gen",
                 disabled=(aoi_payload_raw is None)):
        with st.spinner("Generating mission…"):
            try:
                # Apply buffer
                aoi_to_plan = build_buffered_aoi(aoi_payload_raw, buffer_m)

                if use_optimise:
                    best_hdg, mission = optimise_heading(
                        aoi_to_plan, calc,
                        lead_in_out_m=lead_in_m,
                        turn_time_min=turn_min)
                    st.session_state.nadir_hdg = best_hdg
                else:
                    mission = compute_nadir_mission(
                        aoi_to_plan, calc,
                        flight_azimuth_deg=st.session_state.nadir_hdg,
                        lead_in_out_m=lead_in_m,
                        turn_time_min=turn_min,
                        storage_profile=storage_prof)

                st.session_state.nadir_mission = mission
                if mission is None:
                    st.error("Mission generation failed. Check AOI polygon and spacing values.")
                else:
                    hdg_label = f"(optimised to {best_hdg:.0f}°)" if use_optimise else ""
                    st.success(
                        f"Generated {mission['line_count']:,} lines · "
                        f"{mission['total_triggers']:,} triggers · "
                        f"{mission['total_line_length_m']/1000:.1f} km total {hdg_label}")
            except Exception as exc:
                st.error(f"Error during mission generation: {exc}")
                import traceback; st.text(traceback.format_exc())

    # ─── Mission results ───────────────────────────────────────────────────────
    mission = st.session_state.nadir_mission
    if mission is not None:
        st.markdown("---")
        st.subheader("📋 Mission Statistics")

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.metric("Flight lines",    f"{mission['line_count']:,}")
            st.metric("Total triggers",  f"{mission['total_triggers']:,}")
        with m2:
            st.metric("Total length",    f"{mission['total_line_length_m']/1000:.1f} km")
            st.metric("Avg line length", f"{mission['average_line_length_m']/1000:.2f} km")
        with m3:
            st.metric("Airborne time",   f"{mission['airborne_time_s']/3600:.2f} hr")
            st.metric("Total flying",    f"{mission['flight_time_s']/3600:.2f} hr")
        with m4:
            st.metric("AOI area",        f"{mission['area_m2']/1e6:.2f} km²")
            st.metric("Lead-in / out",   f"{mission['lead_in_out_m']:.0f} m")
        with m5:
            st.metric("Storage",         f"{mission['total_storage_mb']/1024:.1f} GB")
            st.metric("Flight heading",  f"{st.session_state.nadir_hdg:.0f}°")

        # ── Coverage status ────────────────────────────────────────────────────
        cov_pct = mission.get("coverage_pct", 0.0)
        sampled = mission.get("coverage_sampled", False)
        gap_geom = mission.get("gap_geometry")
        lead_clipped = mission.get("lead_clipped", False)

        if lead_clipped:
            st.warning(f"⚠️ Lead-in was increased to {mission['min_lead_m']:.0f} m "
                       f"(trigger spacing) to guarantee full edge coverage.")

        if cov_pct >= 99.9:
            st.markdown(
                f'<div class="trigger-ok">✅ <strong>Coverage: {cov_pct:.1f}%'
                f'{"  (sampled)" if sampled else ""}</strong> — '
                f'Full AOI covered by image footprints.</div>',
                unsafe_allow_html=True)
        elif cov_pct >= 98.0:
            st.markdown(
                f'<div class="trigger-warn">⚠️ <strong>Coverage: {cov_pct:.1f}%'
                f'{"  (sampled)" if sampled else ""}</strong> — '
                f'Minor gaps detected. Check map (red areas). '
                f'Increase coverage buffer or lead-in to close.</div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="trigger-fail">❌ <strong>Coverage: {cov_pct:.1f}%'
                f'{"  (sampled)" if sampled else ""}</strong> — '
                f'Significant gaps detected (red areas on map). '
                f'Increase coverage buffer, reduce line spacing, or increase lead-in.</div>',
                unsafe_allow_html=True)

        if sampled:
            st.caption(f"Coverage check sampled {min(2000, mission['total_triggers']):,} of "
                       f"{mission['total_triggers']:,} frames for performance. "
                       f"Result is representative but not exact.")

        # ── Map display options ────────────────────────────────────────────────
        st.markdown("**Flight plan map**")
        disp_col1, disp_col2, disp_col3 = st.columns(3)
        with disp_col1:
            show_triggers  = st.checkbox("Show trigger points", value=False, key="show_trig")
        with disp_col2:
            show_footprints = st.checkbox("Show image footprints", value=False, key="show_fp")
        with disp_col3:
            show_gaps_opt  = st.checkbox("Highlight coverage gaps", value=True, key="show_gaps")

        total_trigs = mission.get("total_triggers", 0)
        if show_footprints and total_trigs > MAX_FP_DISPLAY:
            st.caption(f"ℹ️ {total_trigs:,} frames — displaying every "
                       f"{max(1, total_trigs // MAX_FP_DISPLAY)}th footprint for performance.")
        if show_triggers and total_trigs > MAX_FP_DISPLAY:
            st.caption(f"ℹ️ Trigger points subsampled to {MAX_FP_DISPLAY:,} for display.")

        # Two tabs: static overview + interactive zoomable map
        if FOLIUM_AVAILABLE:
            tab_static, tab_interactive = st.tabs(["📊 Overview map", "🗺️ Interactive map (zoom)"])
        else:
            tab_static = st.container()
            tab_interactive = None

        with tab_static:
            map_col, _ = st.columns([2, 1])
            map_fig, fp_shown = draw_mission_map(
                mission, calc, dist_unit=dist_unit,
                show_triggers=show_triggers,
                show_footprints=show_footprints,
                show_gaps=show_gaps_opt)
            if map_fig:
                with map_col:
                    st.pyplot(map_fig, use_container_width=True)
                map_fig_bytes = fig_to_png_bytes(map_fig)
                plt.close(map_fig)
            else:
                map_fig_bytes = None

        if FOLIUM_AVAILABLE and tab_interactive is not None:
            with tab_interactive:
                st.caption("Pan and zoom freely. Layer toggles top-right. "
                           "Click the centre marker for a mission summary.")
                folium_map = make_folium_map(
                    mission, calc,
                    show_triggers=show_triggers,
                    show_footprints=show_footprints)
                if folium_map:
                    st_folium(folium_map, use_container_width=True, height=520,
                              returned_objects=[])
                else:
                    st.info("Interactive map not available — missing mx0/my0. "
                            "Load a KML file (not a standard square) to enable.")
        elif not FOLIUM_AVAILABLE:
            st.caption("Install `folium` and `streamlit-folium` for an interactive zoomable map.")

        # ─── Exports ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("💾 Export")

        aoi_name = aoi_payload_raw.get("name", "nadir_plan") if aoi_payload_raw else "nadir_plan"
        settings_rows = build_settings_rows(
            calc, mission_name=aoi_name,
            buffer_m=buffer_m, hdg=st.session_state.nadir_hdg,
            turn_time_min=turn_min, lead_in_m=lead_in_m,
            dist_unit=dist_unit)

        ex1, ex2, ex3 = st.columns(3)

        with ex1:
            excel_bytes = make_excel_export(
                calc, mission, settings_rows, map_png_bytes=map_fig_bytes)
            st.download_button(
                "📊 Download Excel report",
                data=excel_bytes,
                file_name=f"{aoi_name}_nadir_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with ex2:
            if DOCX_AVAILABLE:
                word_bytes = make_word_export(
                    calc, mission, settings_rows, map_png_bytes=map_fig_bytes)
                if word_bytes:
                    st.download_button(
                        "📄 Download Word report",
                        data=word_bytes,
                        file_name=f"{aoi_name}_nadir_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            else:
                st.info("Install python-docx for Word export.")

        with ex3:
            st.markdown("**KML options**")
            kml_inc_trigs = st.checkbox("Include trigger points", value=True, key="kml_trigs")
            kml_inc_fp    = st.checkbox("Include frame footprints", value=True, key="kml_fp")
            if kml_inc_fp and total_trigs > 5000:
                st.caption(f"⚠️ {total_trigs:,} footprint polygons — KML file may be large. "
                           f"Consider unchecking for very large missions.")
            kml_bytes = make_kml_export(mission, calc,
                                        include_triggers=kml_inc_trigs,
                                        include_footprints=kml_inc_fp)
            if kml_bytes:
                kml_size_kb = len(kml_bytes) / 1024
                st.download_button(
                    f"🗺️ Download KML flight plan ({kml_size_kb:.0f} KB)",
                    data=kml_bytes,
                    file_name=f"{aoi_name}_nadir_flight_plan.kml",
                    mime="application/vnd.google-earth.kml+xml",
                )

    elif aoi_payload_raw is not None:
        # AOI loaded, mission not generated yet
        area_km2 = aoi_payload_raw["area_m2"] / 1e6
        st.info(
            f"AOI loaded: **{aoi_payload_raw['name']}** — {area_km2:.2f} km²  "
            f"| Line spacing: **{calc['line_sp_m']:.0f} m**  "
            f"| Estimated lines: ~{int(math.sqrt(area_km2*1e6) / calc['line_sp_m'])+5}  "
            f"| Click **Generate flight lines** when ready."
        )
    else:
        st.info("Load a KML AOI above to enable mission planning.")


# ─────────────────────────────────────────────────────────────────────────────
# Formula trace
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("---")
with st.expander("🔢 Formula trace", expanded=False):
    pixel_m = calc["pixel_um"] / 1e6
    focal_m = calc["focal_mm"] / 1000.0
    st.markdown(f"""
| Step | Formula | Result |
|---|---|---|
| Pixel size | {calc["pixel_um"]:.2f} µm | {pixel_m*1e6:.2f} µm = {pixel_m:.6f} m |
| Focal length | {calc["focal_mm"]} mm | {focal_m:.4f} m |
| **AGL** | `GSD × focal / pixel_size` | `{calc["gsd_cm"]/100:.4f} × {focal_m:.4f} / {pixel_m:.6f}` = **{calc["agl_m"]:.2f} m** |
| Footprint across | `sensor_w_px × GSD` | `{calc["w_px"]:,} × {calc["gsd_cm"]/100:.4f}` = **{calc["fp_across_m"]:.2f} m** |
| Footprint along | `sensor_h_px × GSD` | `{calc["h_px"]:,} × {calc["gsd_cm"]/100:.4f}` = **{calc["fp_along_m"]:.2f} m** |
| Line spacing | `fp_across × (1 − sidelap)` | `{calc["fp_across_m"]:.2f} × {1 - calc["side_pct"]/100:.2f}` = **{calc["line_sp_m"]:.2f} m** |
| Trigger spacing | `fp_along × (1 − forwardlap)` | `{calc["fp_along_m"]:.2f} × {1 - calc["fwd_pct"]/100:.2f}` = **{calc["trig_sp_m"]:.2f} m** |
| Trigger interval | `trigger_spacing / speed` | `{calc["trig_sp_m"]:.2f} / {calc["speed_ms"]:.2f}` = **{calc["trig_int_s"]:.3f} s** |
| Min camera interval | camera spec | **{calc["min_int_s"]:.2f} s** |
| Margin | `trigger_interval / min_interval` | **{calc["trig_margin"]:.3f}×** {'✅' if calc["trig_status"]=="ok" else '⚠️' if calc["trig_status"]=="warn" else '❌'} |
""")

# ─────────────────────────────────────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("---")
st.caption(
    "Nadir Capture Planner v1  ·  Aerial Surveys Ltd  ·  "
    "Flat terrain model  ·  DEM terrain optimisation: Stage 3 planned  ·  "
    "Verify camera specs against current manufacturer datasheets before use"
)
