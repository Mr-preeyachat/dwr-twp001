import os
from typing import Optional

import geopandas as gpd
from shapely.geometry import Point
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field


# =============================================================================
# FastAPI
# =============================================================================
app = FastAPI(
    title="Spatial Map API",
    description="รับ Lat/Long แล้วค้นหาข้อมูลพื้นที่จาก GeoPackage ด้วย GeoPandas",
    version="3.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_FILE = os.path.join(BASE_DIR, "index_map.html")
ASSET_DIR = os.path.join(BASE_DIR, "assets")

# ไฟล์เชิงพื้นที่
FILES = {
    "tambon": os.path.join(ASSET_DIR, "AreaDWR_5.gpkg"),
    "basin": os.path.join(ASSET_DIR, "SB_ONWR_2m.gpkg"),
    "area_based": os.path.join(ASSET_DIR, "Area Based.gpkg"),
}

# จะใช้ตำแหน่งคอลัมน์ตามชุดข้อมูลเดิมของผู้ใช้:
# AreaDWR_5.gpkg : [.., TAMBON, AMPHOE, PROVINCE, ...]
# SB_ONWR_2m.gpkg: [SUB_BASIN, MAIN_BASIN, ...]
# Area Based.gpkg: [NAME_AB, CODE_AB, TYPE_AB, ...]
# แต่ทำให้ปลอดภัยขึ้นด้วยการค้นหาชื่อคอลัมน์ก่อน และ fallback เป็นตำแหน่งเดิม
COLUMN_HINTS = {
    "tambon": {
        "province": ["PROV_NAM_T"],
        "amphoe": ["AMPHOE_T"],
        "tambon": ["TAM_NAM_T"],
    },
    "basin": {
        "sub_basin": ["SUB_BASIN", "SUBBASIN", "ลุ่มน้ำสาขา", "SUB_NAME", "SB_NAME"],
        "main_basin": ["MAIN_BASIN", "MAINBASIN", "ลุ่มน้ำหลัก", "BASIN", "MB_NAME"],
    },
    "area_based": {
        "name_AB": ["NAME_AB", "AB_NAME", "NAME", "ชื่อ", "AREA_NAME"],
        "code_AB": ["CODE_AB", "AB_CODE", "CODE", "รหัส", "AREA_CODE"],
        "type_AB": ["TYPE_AB", "AB_TYPE", "TYPE", "ประเภท", "AREA_TYPE"],
    },
}


# =============================================================================
# โหลดข้อมูลครั้งเดียวตอนเริ่มเซิร์ฟเวอร์
# =============================================================================
def load_layer(path: str, layer_name: str) -> tuple[Optional[gpd.GeoDataFrame], object]:
    if not os.path.exists(path):
        print(f"⚠️ ไม่พบไฟล์ {layer_name}: {path}")
        return None, None

    try:
        gdf = gpd.read_file(path)
        if gdf.empty:
            print(f"⚠️ {layer_name} ไม่มีข้อมูล")
            return None, None

        if gdf.crs is None:
            raise ValueError("ไฟล์ไม่มี CRS จึงไม่สามารถแปลงพิกัดเป็น WGS84 ได้")

        # จุดจาก HTML เป็น WGS84 (EPSG:4326)
        if gdf.crs.to_epsg() != 4326:
            gdf = gdf.to_crs(epsg=4326)

        # ตัด geometry ที่ว่าง/เสียออก เพื่อป้องกันปัญหาตอนค้นหา
        gdf = gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty].copy()
        gdf.reset_index(drop=True, inplace=True)

        sindex = gdf.sindex
        print(f"✅ โหลด {layer_name}: {len(gdf):,} features | CRS={gdf.crs}")
        print(f"   columns: {list(gdf.columns)}")
        return gdf, sindex

    except Exception as e:
        print(f"❌ โหลด {layer_name} ไม่สำเร็จ: {e}")
        return None, None


gdf_tambon, sindex_tambon = load_layer(FILES["tambon"], "AreaDWR_5")
gdf_basin, sindex_basin = load_layer(FILES["basin"], "SB_ONWR_2m")
gdf_area, sindex_area = load_layer(FILES["area_based"], "Area Based")


# =============================================================================
# Utility
# =============================================================================
def clean_value(value) -> str:
    if value is None:
        return "ไม่พบ"
    try:
        if value != value:  # NaN
            return "ไม่พบ"
    except Exception:
        pass
    text = str(value).strip()
    return text if text and text.lower() != "nan" else "ไม่พบ"


def find_column(gdf: Optional[gpd.GeoDataFrame], hints: list[str], fallback_idx: int) -> Optional[str]:
    if gdf is None:
        return None

    # exact match ก่อน
    upper_map = {str(col).strip().upper(): col for col in gdf.columns}
    for hint in hints:
        if hint.upper() in upper_map:
            return upper_map[hint.upper()]

    # match แบบ contains กรณีชื่อคอลัมน์มี prefix/suffix
    for col in gdf.columns:
        col_upper = str(col).strip().upper()
        if any(h.upper() in col_upper for h in hints):
            return col

    # fallback ตามโครงสร้างไฟล์เดิม
    columns_without_geometry = [c for c in gdf.columns if c != gdf.geometry.name]
    if 0 <= fallback_idx < len(columns_without_geometry):
        return columns_without_geometry[fallback_idx]
    return None


def value_from_row(row, column: Optional[str]) -> str:
    return clean_value(row[column]) if column is not None else "ไม่พบ"


def point_in_layer(point: Point, gdf: Optional[gpd.GeoDataFrame], sindex) -> Optional[object]:
    """ค้นหา feature ที่ครอบ/แตะจุด โดยใช้ spatial index ก่อนเพื่อให้เร็ว"""
    if gdf is None or sindex is None:
        return None

    candidates = list(sindex.intersection(point.bounds))
    if not candidates:
        return None

    for idx in candidates:
        geom = gdf.geometry.iloc[idx]
        if geom is None or geom.is_empty:
            continue

        # intersects ดีกว่า contains สำหรับจุดที่ตกบนเส้นขอบ polygon
        try:
            if geom.intersects(point):
                return gdf.iloc[idx]
        except Exception:
            continue

    return None


# =============================================================================
# เตรียม column mapping เพียงครั้งเดียว
# =============================================================================
COL = {
    "province": find_column(gdf_tambon, COLUMN_HINTS["tambon"]["province"], 4),
    "amphoe": find_column(gdf_tambon, COLUMN_HINTS["tambon"]["amphoe"], 3),
    "tambon": find_column(gdf_tambon, COLUMN_HINTS["tambon"]["tambon"], 2),

    "sub_basin": find_column(gdf_basin, COLUMN_HINTS["basin"]["sub_basin"], 1),
    "main_basin": find_column(gdf_basin, COLUMN_HINTS["basin"]["main_basin"], 2),

    "name_AB": find_column(gdf_area, COLUMN_HINTS["area_based"]["name_AB"], 0),
    "code_AB": find_column(gdf_area, COLUMN_HINTS["area_based"]["code_AB"], 1),
    "type_AB": find_column(gdf_area, COLUMN_HINTS["area_based"]["type_AB"], 2),
}

print("\n📌 Column mapping")
for key, col in COL.items():
    print(f"   {key:<12} -> {col}")


# =============================================================================
# Spatial lookup
# =============================================================================
def lookup_spatial_point(lat: float, lon: float) -> dict:
    if not (-90 <= lat <= 90):
        raise ValueError("Latitude ต้องอยู่ระหว่าง -90 ถึง 90")
    if not (-180 <= lon <= 180):
        raise ValueError("Longitude ต้องอยู่ระหว่าง -180 ถึง 180")

    point = Point(lon, lat)

    tambon_row = point_in_layer(point, gdf_tambon, sindex_tambon)
    basin_row = point_in_layer(point, gdf_basin, sindex_basin)
    area_row = point_in_layer(point, gdf_area, sindex_area)

    data = {
        # บังคับคาสต์ float ให้ชัวร์ก่อนส่งกลับ
        "lat": float(lat),
        "lon": float(lon),
        "province": value_from_row(tambon_row, COL["province"]) if tambon_row is not None else "ไม่พบ",
        "amphoe": value_from_row(tambon_row, COL["amphoe"]) if tambon_row is not None else "ไม่พบ",
        "tambon": value_from_row(tambon_row, COL["tambon"]) if tambon_row is not None else "ไม่พบ",
        "main_basin": value_from_row(basin_row, COL["main_basin"]) if basin_row is not None else "ไม่พบ",
        "sub_basin": value_from_row(basin_row, COL["sub_basin"]) if basin_row is not None else "ไม่พบ",
        "code_AB": value_from_row(area_row, COL["code_AB"]) if area_row is not None else "ไม่พบ",
        "name_AB": value_from_row(area_row, COL["name_AB"]) if area_row is not None else "ไม่พบ",
        "type_AB": value_from_row(area_row, COL["type_AB"]) if area_row is not None else "ไม่พบ",
    }

    data["found"] = {
        "tambon": tambon_row is not None,
        "basin": basin_row is not None,
        "area_based": area_row is not None,
    }
    return data


# =============================================================================
# Request model
# =============================================================================
class CoordRequest(BaseModel):
    lat: float = Field(..., ge=-90, le=90)
    lon: float = Field(..., ge=-180, le=180)


# =============================================================================
# Routes
# =============================================================================
@app.get("/", response_class=HTMLResponse)
async def serve_index():
    if not os.path.exists(INDEX_FILE):
        raise HTTPException(status_code=404, detail="ไม่พบ index_map.html")
    with open(INDEX_FILE, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/health")
async def health_check():
    return {
        "success": True,
        "layers": {
            "tambon": gdf_tambon is not None,
            "basin": gdf_basin is not None,
            "area_based": gdf_area is not None,
        },
    }


@app.post("/api/check_coords")
async def check_coordinates(req: CoordRequest):
    try:
        result = lookup_spatial_point(req.lat, req.lon)
        return {"success": True, "data": result}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        print(f"❌ Spatial lookup error: {e}")
        raise HTTPException(status_code=500, detail="เกิดข้อผิดพลาดระหว่างค้นหาข้อมูลพื้นที่")


# =============================================================================
# Run
# =============================================================================
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main_map:app",
        host="0.0.0.0",
        port=8000,
        reload=False,
    )
