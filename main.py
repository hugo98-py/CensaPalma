# -*- coding: utf-8 -*-
"""
FastAPI · Firestore (censaPalma) → Excel (descarga)
---------------------------------------------------
GET /export  →  {"download_url": ".../downloads/<archivo>.xlsx"}

• Lee FIREBASE_KEY_B64 desde variable de entorno (no hardcodeado).
• Coordenadas: detecta varias formas, crea _lat/_lon y UTM (zona automática).
• Tiempo: 'dateTime' → columnas dt_* en America/Santiago.
• Quita columnas 'doc_id', 'coords' y 'coords__*'.
• Guarda Excel en /tmp (filesystem efímero en Render), servido en /downloads.
"""

import os, re, json, base64
from typing import Any, Dict, Optional, Tuple
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from pyproj import Transformer
from pandas.api.types import is_datetime64_any_dtype

import firebase_admin
from firebase_admin import credentials, firestore

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles


# ─────────────────────────────────────────────── CONFIG
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "/tmp/downloads"))
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)


def _get_sa_info() -> dict:
    b64 = os.getenv("FIREBASE_KEY_B64", "").strip()
    if not b64:
        raise RuntimeError("FIREBASE_KEY_B64 no está definida.")
    try:
        return json.loads(base64.b64decode(b64).decode("utf-8"))
    except Exception:
        raise RuntimeError("FIREBASE_KEY_B64 inválida")


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
    return (lat, lon)


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
            return (float(obj.latitude), float(obj.longitude))
    except Exception:
        pass
    return None


def _from_sequence(seq: Any) -> Optional[Tuple[float, float]]:
    if isinstance(seq, (list, tuple)) and len(seq) >= 2:
        a = _to_float_safe(seq[0])
        b = _to_float_safe(seq[1])
        if a is None or b is None:
            return None
        if -90 <= a <= 90 and -180 <= b <= 180:
            return (a, b)
        if -90 <= b <= 90 and -180 <= a <= 180:
            return (b, a)
    return None


def _from_mapping(d: Dict[str, Any]) -> Optional[Tuple[float, float]]:
    if not isinstance(d, dict):
        return None
    for la, lo in [
        ("lat", "lon"),
        ("lat", "lng"),
        ("latitude", "longitude"),
        ("_lat", "_long"),
        ("y", "x"),
    ]:
        if la in d and lo in d:
            a = _to_float_safe(d.get(la))
            b = _to_float_safe(d.get(lo))
            if a is not None and b is not None:
                return (a, b)
    return None


def extract_lat_lon(value: Any) -> Optional[Tuple[float, float]]:
    if value is None:
        return None
    for fn in (_from_geopoint, _from_mapping, _from_sequence):
        got = fn(value)
        if got:
            return got
    if isinstance(value, str):
        got = _from_wkt_point(value) or _from_string_pair(value)
        if got:
            return got
    return None


def find_coord_in_record(record: Dict[str, Any]) -> Tuple[Optional[float], Optional[float]]:
    # Clave principal
    if "coords" in record:
        got = extract_lat_lon(record["coords"])
        if got:
            return got

    # fallback extra
    lat = _to_float_safe(record.get("lat") or record.get("latitude"))
    lon = _to_float_safe(record.get("lon") or record.get("lng") or record.get("longitude"))
    if lat is not None and lon is not None:
        return lat, lon

    return None, None


# ─────────────────────────────────────────────── WGS84 → UTM
_transformers: Dict[int, Transformer] = {}

def latlon_to_utm(lat: float, lon: float):
    if lat is None or lon is None:
        return None, None, None, None, None
    zone = int((lon + 180) / 6) + 1
    epsg = 32700 + zone
    if epsg not in _transformers:
        _transformers[epsg] = Transformer.from_crs(
            "EPSG:4326", f"EPSG:{epsg}", always_xy=True
        )
    e, n = _transformers[epsg].transform(lon, lat)
    return float(e), float(n), zone, "S", epsg


# ─────────────────────────────────────────────── Tiempo
SCL = ZoneInfo("America/Santiago")

def _to_scl_timestamp(val):
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if ts.tzinfo is None:
            return ts.tz_localize(SCL)
        return ts.tz_convert(SCL)
    except Exception:
        return pd.NaT


# ─────────────────────────────────────────────── Data builder
def build_dataframe() -> pd.DataFrame:
    fs = init_fs_client()
    rows = []

    for doc in fs.collection("DataCenso").stream():
        d = doc.to_dict() or {}

        lat, lon = find_coord_in_record(d)

        # ✅ Campos originales + nuevos
        d_clean = {
            "doc_id": doc.id,
            "email": d.get("email"),
            "Nombre": d.get("Nombre"),
            "Apellido": d.get("Apellido"),
            "comentario": d.get("comentario"),
            "tamano": d.get("tamano"),
            "condicionSanitaria": d.get("condicionSanitaria"),
            "dateTime": d.get("dateTime"),
            "_lat": lat,
            "_lon": lon,
        }
        rows.append(d_clean)

    df = pd.DataFrame(rows)

    # UTM expandido
    def _conv(r):
        return pd.Series(
            latlon_to_utm(r["_lat"], r["_lon"]),
            index=["utm_e", "utm_n", "utm_zone", "utm_hemisphere", "utm_epsg"],
        )

    if "_lat" in df.columns and "_lon" in df.columns:
        utm_df = df.apply(_conv, axis=1)
        df = pd.concat([df, utm_df], axis=1)

    # fecha → columnas dt_*
    if "dateTime" in df.columns:
        dt = df["dateTime"].apply(_to_scl_timestamp)
        df["dt_year"] = dt.dt.year
        df["dt_month"] = dt.dt.month
        df["dt_day"] = dt.dt.day
        df["dt_hour"] = dt.dt.hour
        df["dt_minute"] = dt.dt.minute
        df["dt_second"] = dt.dt.second

    df = df.drop(columns=["doc_id"], errors="ignore")
    return df


# ─────────────────────────────────────────────── FastAPI
app = FastAPI(title="CensaPalma → Excel", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/downloads", StaticFiles(directory=str(DOWNLOAD_DIR)), name="downloads")


@app.get("/", tags=["health"])
def root():
    return {"ok": True, "endpoints": ["/export"]}


@app.get("/export", tags=["export"])
def export_excel(request: Request):
    df = build_dataframe()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"CensaPalma_{ts}.xlsx"
    fpath = DOWNLOAD_DIR / fname

    for col in df.columns:
        if is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")

    df.to_excel(fpath, index=False)

    base = str(request.base_url).rstrip("/")
    return JSONResponse({"download_url": f"{base}/downloads/{fname}"})





