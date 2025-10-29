# -*- coding: utf-8 -*-
"""
Created on Wed Oct 29 20:21:42 2025

@author: Hugo
"""

# -*- coding: utf-8 -*-
"""
FastAPI · Firestore (censaPalma) → Excel (descarga)
---------------------------------------------------
GET /export  →  {"download_url": ".../downloads/<archivo>.xlsx"}

• Lee FIREBASE_KEY_B64 desde variable de entorno (no hardcodeado).
• Campos esperados en la colección 'censaPalma': email, coords, comentario, dateTime, Nombre, Apellido.
• Coordenadas: detecta varias formas, crea columnas _lat/_lon y UTM (zona automática).
• Tiempo: 'dateTime' → columnas dt_* en America/Santiago.
• Quita columnas 'doc_id', 'coords' y 'coords__*'.
• Guarda Excel en /tmp/downloads (filesystem efímero en Render), servido en /downloads.
"""

import os, re, json, base64, math
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from pyproj import Transformer

import firebase_admin
from firebase_admin import credentials, firestore

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

# ─────────────────────────────────────────────── CONFIG
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "/tmp/downloads"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

def _get_sa_info() -> dict:
    b64 = os.getenv("FIREBASE_KEY_B64", "").strip()
    if not b64:
        raise RuntimeError(
            "FIREBASE_KEY_B64 no está definida. Sube tu service-account.json como base64 a esa variable."
        )
    try:
        return json.loads(base64.b64decode(b64).decode("utf-8"))
    except Exception as e:
        raise RuntimeError(f"FIREBASE_KEY_B64 inválida: {e}")

# ─────────────────────────────────────────────── Firebase
def init_fs_client() -> firestore.Client:
    sa = _get_sa_info()
    cred = credentials.Certificate(sa)
    if not firebase_admin._apps:
        firebase_admin.initialize_app(cred)
    return firestore.client()

# ─────────────────────────────────────────────── Coordenadas: parsers
_num_re = re.compile(r"-?\d+(?:[.,]\d+)?")

def _to_float_safe(x: Any) -> Optional[float]:
    try:
        if isinstance(x, (int, float)):
            return float(x)
        if isinstance(x, str):
            return float(x.strip().replace(",", "."))
    except Exception:
        return None
    return None

def _from_wkt_point(s: str) -> Optional[Tuple[float, float]]:
    s2 = s.strip().upper()
    if not s2.startswith("POINT"):
        return None
    m = re.search(r"POINT\s+Z?\s*\(([^)]+)\)", s2)
    if not m:
        return None
    nums = _num_re.findall(m.group(1))
    if len(nums) < 2:
        return None
    lon, lat = [float(n.replace(",", ".")) for n in nums[:2]]
    if -90 <= lat <= 90 and -180 <= lon <= 180:
        return (lat, lon)
    return None

def _from_string_pair(s: str) -> Optional[Tuple[float, float]]:
    nums = _num_re.findall(s)
    if len(nums) < 2:
        return None
    a = float(nums[0].replace(",", "."))
    b = float(nums[1].replace(",", "."))
    if -90 <= a <= 90 and -180 <= b <= 180:
        return (a, b)
    if -90 <= b <= 90 and -180 <= a <= 180:
        return (b, a)
    return None

def _from_geopoint(obj: Any) -> Optional[Tuple[float, float]]:
    try:
        if hasattr(obj, "latitude") and hasattr(obj, "longitude"):
            lat = float(obj.latitude); lon = float(obj.longitude)
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                return (lat, lon)
    except Exception:
        pass
    return None

def _from_sequence(seq: Any) -> Optional[Tuple[float, float]]:
    try:
        if isinstance(seq, (list, tuple)) and len(seq) >= 2:
            a = _to_float_safe(seq[0]); b = _to_float_safe(seq[1])
            if a is None or b is None: return None
            if -90 <= a <= 90 and -180 <= b <= 180: return (a, b)
            if -90 <= b <= 90 and -180 <= a <= 180: return (b, a)
    except Exception:
        pass
    return None

def _from_mapping(d: Dict[str, Any]) -> Optional[Tuple[float, float]]:
    if not isinstance(d, dict): return None
    candidates = [
        ("lat", "lon"), ("lat", "lng"),
        ("latitude", "longitude"), ("_lat", "_long"),
        ("y", "x"),
    ]
    for la, lo in candidates:
        if la in d and lo in d:
            lat = _to_float_safe(d.get(la)); lon = _to_float_safe(d.get(lo))
            if lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180:
                return (lat, lon)
    if "lat" in d and any(k in d for k in ("lon","lng","long","longitude")):
        lat = _to_float_safe(d.get("lat"))
        lon = _to_float_safe(d.get("lon", d.get("lng", d.get("long", d.get("longitude")))))
        if lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180:
            return (lat, lon)
    return None

def extract_lat_lon(value: Any) -> Optional[Tuple[float, float]]:
    if value is None:
        return None
    got = _from_geopoint(value)
    if got: return got
    if isinstance(value, dict):
        got = _from_mapping(value)
        if got: return got
    if isinstance(value, (list, tuple)):
        got = _from_sequence(value)
        if got: return got
    if isinstance(value, str):
        got = _from_wkt_point(value)
        if got: return got
        got = _from_string_pair(value)
        if got: return got
    if isinstance(value, dict):
        lat = _to_float_safe(value.get("lat") or value.get("latitude") or value.get("_lat"))
        lon = _to_float_safe(value.get("lon") or value.get("lng") or value.get("longitude") or value.get("_long"))
        if lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180:
            return (lat, lon)
    return None

def find_coord_in_record(record: Dict[str, Any]) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    pref_keys = [
        "coords","coordenadas","coordenada","ubicacion","ubicación","location",
        "latlng","lat_lng","geom","geometry","point","punto","gps",
        "posicion","posición"
    ]
    for k in pref_keys:
        if k in record:
            got = extract_lat_lon(record.get(k))
            if got: return got[0], got[1], k
    for k, v in record.items():
        got = extract_lat_lon(v)
        if got: return got[0], got[1], k
    lat = _to_float_safe(record.get("lat") or record.get("latitude") or record.get("_lat"))
    lon = _to_float_safe(record.get("lon") or record.get("lng") or record.get("longitude") or record.get("_long"))
    if lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180:
        return lat, lon, None
    return None, None, None

# ─────────────────────────────────────────────── WGS84 → UTM
_transformers: Dict[int, Transformer] = {}

def latlon_to_utm(lat: float, lon: float) -> Tuple[Optional[float], Optional[float], Optional[int], str, Optional[int]]:
    if lat is None or lon is None:
        return None, None, None, "", None
    zone = int(math.floor((lon + 180.0) / 6.0) + 1)  # 1..60
    south = lat < 0
    epsg = 32700 + zone if south else 32600 + zone
    if epsg not in _transformers:
        _transformers[epsg] = Transformer.from_crs("EPSG:4326", f"EPSG:{epsg}", always_xy=True)
    e, n = _transformers[epsg].transform(lon, lat)  # (lon, lat)
    hemi = "S" if south else "N"
    return float(e), float(n), zone, hemi, epsg

# ─────────────────────────────────────────────── Tiempo (America/Santiago)
SCL = ZoneInfo("America/Santiago")
_TZ_SUFFIX = re.compile(r"(Z|[+-]\d{2}:?\d{2})$")

def _to_scl_timestamp(val):
    """Devuelve pandas.Timestamp tz-aware en America/Santiago."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return pd.NaT
    try:
        if isinstance(val, str):
            s = val.strip()
            if _TZ_SUFFIX.search(s):
                ts = pd.to_datetime(s, errors="coerce", utc=True).tz_convert(SCL)
            else:
                ts = pd.to_datetime(s, errors="coerce").tz_localize(SCL)
            return ts
        ts = pd.to_datetime(val, errors="coerce")
        if ts is pd.NaT:
            return pd.NaT
        if ts.tzinfo is None:
            return ts.tz_localize(SCL)
        return ts.tz_convert(SCL)
    except Exception:
        return pd.NaT

# ─────────────────────────────────────────────── Data builder (censaPalma)
def build_dataframe() -> pd.DataFrame:
    fs = init_fs_client()
    rows: List[Dict[str, Any]] = []

    # ⚠️ Colección objetivo: censaPalma
    for doc in fs.collection("censaPalma").stream():
        data = doc.to_dict() or {}
        # Solo tomamos/normalizamos campos esperados, permitiendo extras sin romper:
        base = {
            "email": data.get("email"),
            "coords": data.get("coords"),
            "comentario": data.get("comentario"),
            "dateTime": data.get("dateTime"),
            "Nombre": data.get("Nombre"),
            "Apellido": data.get("Apellido"),
        }
        lat, lon, _ = find_coord_in_record(base)
        row = {"doc_id": doc.id, "_lat": lat, "_lon": lon}
        row.update(base)
        rows.append(row)

    df = pd.json_normalize(rows, sep="__")

    # UTM columns
    def _convert_row(r):
        lat = r.get("_lat")
        lon = r.get("_lon")
    
        if pd.isna(lat) or pd.isna(lon):
            return [None, None, None, None, None]
    
        try:
            e, n, zone, hemi, epsg = latlon_to_utm(float(lat), float(lon))
            return [e, n, zone, hemi, epsg]
        except Exception:
            return [None, None, None, None, None]

    df[["utm_e","utm_n","utm_zone","utm_hemisphere","utm_epsg"]] = df.apply(_convert_row, axis=1)

    # dateTime → componentes locales
    if "dateTime" in df.columns:
        dt_scl = df["dateTime"].apply(_to_scl_timestamp)
        df["dt_year"]   = dt_scl.dt.year
        df["dt_month"]  = dt_scl.dt.month
        df["dt_day"]    = dt_scl.dt.day
        df["dt_hour"]   = dt_scl.dt.hour
        df["dt_minute"] = dt_scl.dt.minute
        df["dt_second"] = dt_scl.dt.second

    # Drop columnas pedidas
    drop_cols = [c for c in df.columns if c == "doc_id" or c == "coords" or c.startswith("coords__")]
    df = df.drop(columns=drop_cols, errors="ignore")
    return df

# ─────────────────────────────────────────────── FastAPI app
app = FastAPI(title="CensaPalma → Excel", version="1.0.0")

# CORS básico (ajusta orígenes si quieres restringir)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Sirve descargas estáticas desde /tmp/downloads
app.mount("/downloads", StaticFiles(directory=str(DOWNLOAD_DIR)), name="downloads")

@app.get("/", tags=["health"])
def root():
    return {"ok": True, "project": "CensaPalma", "endpoints": ["/export"]}

@app.get("/export", tags=["export"])
def export_excel(request: Request):
    """
    Genera un Excel con el DF y devuelve la URL de descarga.
    """
    try:
        df = build_dataframe()

        # Nombre de archivo
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"CensaPalma_{ts}.xlsx"
        fpath = DOWNLOAD_DIR / fname

        # ── Excel-safe: convertir datetimes con tz a texto ISO-8601 (+00:00)
        from pandas.api.types import is_datetime64tz_dtype
        for col in df.columns:
            if is_datetime64tz_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S%z")
                df[col] = df[col].str.replace(r"([+-]\d{2})(\d{2})$", r"\1:\2", regex=True)

        # Escribir Excel
        df.to_excel(fpath, index=False)  # requiere openpyxl

        # URL absoluta
        base = str(request.base_url).rstrip("/")
        download_url = f"{base}/downloads/{fname}"
        return JSONResponse({"download_url": download_url})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export error: {e}")


