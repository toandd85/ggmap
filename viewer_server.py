import os
import re
import glob
import json
import asyncio
import datetime
import subprocess
import threading
import pandas as pd
import sys
from pathlib import Path
from typing import Optional
from fastapi import FastAPI, BackgroundTasks, Query
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import httpx

app = FastAPI(title="Google Maps Scraper - Viet UI Pro", root_path="/ggmdata" if os.environ.get("USE_ROOT_PATH") == "1" else "")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

WORKSPACE_DIR = Path(os.environ.get("WORKSPACE_DIR", Path(__file__).parent.resolve())).resolve()
SCRAPER_BIN = WORKSPACE_DIR / ("google-maps-scraper.exe" if sys.platform == "win32" else "google-maps-scraper")
RESULTS_DIR = WORKSPACE_DIR / "outputs"
CONFIG_FILE = WORKSPACE_DIR / "v9tech_config.json"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# Trạng thái job cào hiện tại
current_scrape_job = {
    "is_running": False,
    "job_id": None,
    "logs": [],
    "status": "idle", # idle, running, completed, error
    "output_file": None,
    "error_message": None
}

def clean_vietnamese_phone(raw_phone: str) -> dict:
    """Chuẩn hóa và nhận diện số di động Việt Nam để tạo link Zalo & Gọi PBX"""
    if not raw_phone or pd.isna(raw_phone):
        return {"raw": "", "clean": "", "is_mobile": False, "zalo_url": "", "type": "none"}
    
    # Chỉ giữ các chữ số
    digits = "".join(c for c in str(raw_phone) if c.isdigit())
    if digits.startswith("84") and len(digits) >= 11:
        digits = "0" + digits[2:]
    elif not digits.startswith("0") and len(digits) == 9:
        digits = "0" + digits

    # Nhận diện đầu số di động Việt Nam 10 số (03x, 05x, 07x, 08x, 09x)
    is_mobile = bool(re.match(r"^0(3[2-9]|5[2689]|7[06-9]|8[1-9]|9[0-9])[0-9]{7}$", digits))
    zalo_url = f"https://zalo.me/{digits}" if is_mobile else ""
    phone_type = "mobile" if is_mobile else ("landline" if len(digits) >= 10 else "other")
    
    return {
        "raw": str(raw_phone).strip(),
        "clean": digits,
        "is_mobile": is_mobile,
        "zalo_url": zalo_url,
        "type": phone_type
    }

def fix_and_save_files(csv_path: Path):
    """Chuẩn hóa file CSV sang UTF-8 BOM và tạo thêm file Excel .xlsx với link Zalo"""
    if not csv_path.exists():
        return
    try:
        df = pd.read_csv(csv_path, encoding='utf-8')
        
        # Thêm cột Zalo và Loại số điện thoại
        if 'phone' in df.columns:
            phone_infos = [clean_vietnamese_phone(p) for p in df['phone']]
            df['phone_clean'] = [p['clean'] for p in phone_infos]
            df['is_mobile'] = ['Di động' if p['is_mobile'] else ('Cố định' if p['clean'] else '') for p in phone_infos]
            df['zalo_link'] = [p['zalo_url'] for p in phone_infos]

        # Lưu bản UTF-8 with BOM
        bom_csv_path = csv_path.with_name(f"{csv_path.stem}_excel.csv")
        df.to_csv(bom_csv_path, index=False, encoding='utf-8-sig')

        # Lưu bản Excel .xlsx
        excel_path = csv_path.with_suffix('.xlsx')
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='GoogleMapsData')
    except Exception as e:
        print(f"Lỗi chuẩn hóa file: {e}")

def run_scraper_thread(cmd: list, raw_csv_path: Path):
    global current_scrape_job
    try:
        current_scrape_job["is_running"] = True
        current_scrape_job["status"] = "running"
        current_scrape_job["logs"] = ["Bắt đầu tiến trình cào dữ liệu Google Maps..."]
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=str(WORKSPACE_DIR)
        )

        for line in process.stdout:
            clean_line = line.strip()
            if clean_line:
                current_scrape_job["logs"].append(clean_line)
                if len(current_scrape_job["logs"]) > 250:
                    current_scrape_job["logs"].pop(0)

        process.wait()

        if process.returncode == 0:
            current_scrape_job["status"] = "completed"
            current_scrape_job["logs"].append(">>> Cào dữ liệu hoàn tất thành công!")
            fix_and_save_files(raw_csv_path)
        else:
            current_scrape_job["status"] = "error"
            current_scrape_job["logs"].append(f">>> Tiến trình kết thúc với mã lỗi: {process.returncode}")
    except Exception as e:
        current_scrape_job["status"] = "error"
        current_scrape_job["error_message"] = str(e)
        current_scrape_job["logs"].append(f">>> Lỗi ngoại lệ: {str(e)}")
    finally:
        current_scrape_job["is_running"] = False

@app.get("/api/status")
def get_scrape_status():
    return {
        "is_running": current_scrape_job["is_running"],
        "status": current_scrape_job["status"],
        "logs": current_scrape_job["logs"][-30:],
        "output_file": current_scrape_job["output_file"]
    }

# API Cấu hình V9 Tech
@app.get("/api/v9tech/config")
def get_v9tech_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "api_url": "https://api.v9tech.vn/api/v1/call/click-to-call",
        "extension": "101",
        "api_key": ""
    }

@app.post("/api/v9tech/config")
def save_v9tech_config(data: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return {"success": True, "message": "Đã lưu cấu hình V9 Tech thành công"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# API Click-to-call V9 Tech
@app.post("/api/v9tech/call")
async def v9tech_call(data: dict):
    phone = data.get("phone", "").strip()
    phone_info = clean_vietnamese_phone(phone)
    clean_phone = phone_info["clean"] or phone

    # Đọc config
    config = get_v9tech_config()
    api_url = data.get("api_url") or config.get("api_url", "")
    extension = data.get("extension") or config.get("extension", "101")
    api_key = data.get("api_key") or config.get("api_key", "")

    if not extension:
        return JSONResponse(status_code=400, content={"error": "Vui lòng nhập Số máy nhánh (Extension) của bạn."})

    # Gọi tới tổng đài V9 Tech
    try:
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"
            headers["X-API-KEY"] = api_key

        payload = {
            "extension": extension,
            "caller": extension,
            "callee": clean_phone,
            "phone": clean_phone,
            "phone_number": clean_phone
        }

        async with httpx.AsyncClient(timeout=8.0, verify=False) as client:
            resp = await client.post(api_url, json=payload, headers=headers)
            status_code = resp.status_code
            resp_body = resp.text

            if status_code in (200, 201):
                return {
                    "success": True,
                    "message": f"Tổng đài V9 Tech đang đổ chuông máy nhánh [{extension}]. Vui lòng nhấc máy để kết nối tới [{clean_phone}]!",
                    "details": resp_body
                }
            else:
                return {
                    "success": False,
                    "status_code": status_code,
                    "message": f"Tổng đài V9 Tech phản hồi HTTP {status_code}: {resp_body[:150]}",
                    "clean_phone": clean_phone
                }
    except Exception as e:
        return {
            "success": False,
            "fallback": True,
            "message": f"Không thể kết nối trực tiếp API V9 Tech ({str(e)}). Bạn có thể bấm mở qua ứng dụng Softphone/Zoiper.",
            "clean_phone": clean_phone
        }

@app.post("/api/start-scrape")
def start_scrape(data: dict):
    global current_scrape_job
    if current_scrape_job["is_running"]:
        return JSONResponse(status_code=400, content={"error": "Một tác vụ cào dữ liệu đang chạy, vui lòng chờ."})

    keywords = data.get("keywords", "").strip()
    if not keywords:
        return JSONResponse(status_code=400, content={"error": "Vui lòng nhập từ khóa tìm kiếm."})

    mode = data.get("mode", "radius") # "radius", "admin" (phường/quận/tỉnh)
    lat = data.get("lat")
    lon = data.get("lon")
    radius_km = float(data.get("radius_km", 2.0))
    zoom = int(data.get("zoom", 15))
    depth = int(data.get("depth", 30))
    fetch_email = bool(data.get("email", False))
    use_grid = bool(data.get("use_grid", False))
    bbox = data.get("bbox") # "minLat,minLon,maxLat,maxLon"

    # Danh sách truy vấn
    query_lines = []

    if mode == "admin" and data.get("admin_areas"):
        # Chế độ quét theo từng đơn vị hành chính
        # admin_areas: danh sách các phường/xã hoặc quận/huyện
        areas = data.get("admin_areas", [])
        base_kws = [k.strip() for k in keywords.split("\n") if k.strip()]
        for kw in base_kws:
            for area in areas:
                query_lines.append(f"{kw} tại {area}")
    else:
        for kw in keywords.split("\n"):
            kw_clean = kw.strip()
            if kw_clean:
                query_lines.append(kw_clean)

    if not query_lines:
        return JSONResponse(status_code=400, content={"error": "Không có từ khóa hợp lệ."})

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    first_kw = query_lines[0].split()[0] if query_lines else "gmaps"
    safe_name = "".join(c for c in first_kw if c.isalnum() or c in ('_', '-')) or "gmaps"
    output_filename = f"{safe_name}_{timestamp}.csv"
    raw_csv_path = RESULTS_DIR / output_filename

    query_file_path = RESULTS_DIR / f"queries_{timestamp}.txt"
    with open(query_file_path, "w", encoding="utf-8") as f:
        for q in query_lines:
            f.write(q + "\n")

    cmd = [
        str(SCRAPER_BIN),
        "-input", str(query_file_path),
        "-results", str(raw_csv_path),
        "-lang", "vi",
        "-depth", str(depth),
        "-c", "2",
        "-exit-on-inactivity", "2m"
    ]

    if fetch_email:
        cmd.append("-email")

    if use_grid and bbox:
        cmd.extend(["-grid-bbox", str(bbox), "-grid-cell", "1.0", "-zoom", str(zoom)])
    elif lat and lon:
        cmd.extend(["-geo", f"{lat},{lon}", "-zoom", str(zoom)])
        if radius_km:
            cmd.extend(["-radius", str(radius_km * 1000)])

    current_scrape_job["is_running"] = True
    current_scrape_job["status"] = "running"
    current_scrape_job["output_file"] = output_filename
    current_scrape_job["logs"] = []

    threading.Thread(target=run_scraper_thread, args=(cmd, raw_csv_path), daemon=True).start()

    return {"message": "Đã bắt đầu cào dữ liệu", "output_file": output_filename, "queries_count": len(query_lines)}

@app.get("/api/files")
def list_result_files():
    files = []
    search_paths = [RESULTS_DIR, WORKSPACE_DIR]
    seen = set()
    for sp in search_paths:
        for p in sp.glob("*.csv"):
            if p.name.endswith("_excel.csv"):
                continue
            if p.name in seen:
                continue
            seen.add(p.name)
            stat = p.stat()
            files.append({
                "name": p.name,
                "path": str(p),
                "size_kb": round(stat.st_size / 1024, 1),
                "modified": datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
            })
    files.sort(key=lambda x: x["modified"], reverse=True)
    return files

@app.get("/api/data")
def get_data(filename: str = Query(...)):
    p = RESULTS_DIR / filename
    if not p.exists():
        p = WORKSPACE_DIR / filename
    if not p.exists():
        return JSONResponse(status_code=404, content={"error": "Không tìm thấy file"})

    try:
        df = pd.read_csv(p, encoding="utf-8")
        df = df.fillna("")
        records = df.to_dict(orient="records")
        
        # Thêm thông tin Zalo và số điện thoại chi tiết cho từng bản ghi
        for r in records:
            p_info = clean_vietnamese_phone(r.get("phone", ""))
            r["phone_info"] = p_info

        return {
            "total": len(records),
            "filename": filename,
            "data": records
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Lỗi đọc file: {str(e)}"})

@app.get("/api/download")
def download_file(filename: str = Query(...), format: str = Query("excel")):
    p = RESULTS_DIR / filename
    if not p.exists():
        p = WORKSPACE_DIR / filename
    if not p.exists():
        return JSONResponse(status_code=404, content={"error": "Không tìm thấy file"})

    if format == "excel":
        xlsx_path = p.with_suffix(".xlsx")
        if not xlsx_path.exists():
            fix_and_save_files(p)
        if xlsx_path.exists():
            return FileResponse(xlsx_path, filename=xlsx_path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    bom_csv = p.with_name(f"{p.stem}_excel.csv")
    if not bom_csv.exists():
        fix_and_save_files(p)
    target = bom_csv if bom_csv.exists() else p
    return FileResponse(target, filename=f"{p.stem}_tieng_viet.csv", media_type="text/csv")

# Giao diện HTML Tiếng Việt Chuyên Nghiệp
@app.get("/", response_class=HTMLResponse)
def index_html():
    html = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Google Maps Scraper - Tiếng Việt Pro (Bán Kính, Phường/Xã, PBX V9 Tech, Zalo)</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Leaflet Map CSS & JS -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; }
        #picker-map, #results-map { height: 460px; border-radius: 0.75rem; }
        .custom-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        .zalo-badge { background-color: #0068FF; color: white; }
        .zalo-badge:hover { background-color: #0053cc; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen flex flex-col">
    <!-- Header -->
    <header class="bg-indigo-700 text-white shadow-md">
        <div class="max-w-7xl mx-auto px-4 py-3 flex flex-wrap items-center justify-between gap-4">
            <div class="flex items-center space-x-3">
                <div class="p-2 bg-indigo-800 rounded-lg">
                    <i data-lucide="map-pin" class="w-6 h-6 text-amber-300"></i>
                </div>
                <div>
                    <h1 class="text-xl font-bold tracking-tight flex items-center gap-2">
                        Google Maps Scraper Pro
                        <span class="text-[11px] bg-amber-400 text-slate-900 px-2 py-0.5 rounded font-extrabold uppercase">V9 PBX & Zalo</span>
                    </h1>
                    <p class="text-xs text-indigo-200">Quét theo Bán Kính / Phường / Xã / Quận / Huyện • Click-to-Call V9 Tech • Nhắn Zalo 1-Chạm</p>
                </div>
            </div>
            
            <div class="flex items-center space-x-2">
                <button onclick="switchTab('tab-scrape')" id="btn-tab-scrape" class="px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5">
                    <i data-lucide="crosshair" class="w-4 h-4"></i> Chọn vùng & Cào
                </button>
                <button onclick="switchTab('tab-view')" id="btn-tab-view" class="px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5">
                    <i data-lucide="table" class="w-4 h-4"></i> Xem dữ liệu & Gọi
                </button>
                <button onclick="openV9Modal()" class="px-3 py-2 text-xs font-bold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white flex items-center gap-1.5 shadow" title="Cài đặt tổng đài V9 Tech">
                    <i data-lucide="phone-call" class="w-4 h-4"></i> Cấu hình V9 Tech
                </button>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 py-6 flex-1 w-full">
        
        <!-- TAB 1: CHỌN VỊ TRÍ, BÁN KÍNH, PHƯỜNG/XÃ & CÀO -->
        <div id="tab-scrape" class="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <!-- Cột trái: Form cấu hình tìm kiếm -->
            <div class="lg:col-span-5 bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col gap-4">
                
                <!-- Chế độ quét: Bán kính hoặc Đơn vị hành chính -->
                <div class="border-b border-slate-200 pb-3">
                    <label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Chế độ quét vị trí:</label>
                    <div class="grid grid-cols-2 gap-2 bg-slate-100 p-1 rounded-lg text-xs font-semibold">
                        <button type="button" id="btn-mode-radius" onclick="setScrapeMode('radius')" class="py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1">
                            <i data-lucide="target" class="w-3.5 h-3.5"></i> Theo Bán kính (Km)
                        </button>
                        <button type="button" id="btn-mode-admin" onclick="setScrapeMode('admin')" class="py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1">
                            <i data-lucide="building" class="w-3.5 h-3.5"></i> Theo Phường/Quận/Tỉnh
                        </button>
                    </div>
                </div>

                <!-- Từ khóa tìm kiếm -->
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Từ khóa tìm kiếm (mỗi từ 1 dòng):</label>
                    <textarea id="scrape-keywords" rows="2" placeholder="quán cà phê&#10;cửa hàng quà tặng" 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">quán cà phê</textarea>
                </div>

                <!-- KHU VỰC CHẾ ĐỘ 1: BÁN KÍNH -->
                <div id="panel-mode-radius" class="flex flex-col gap-3">
                    <div class="flex items-center justify-between">
                        <label class="text-xs font-bold text-slate-700">Bán kính quét (Radius):</label>
                        <span id="radius-val-text" class="text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded">2.0 Km (~12 km²)</span>
                    </div>
                    <input type="range" id="scrape-radius-slider" min="0.5" max="25" step="0.5" value="2.0" oninput="updateRadiusFromSlider()" class="w-full accent-indigo-600 cursor-pointer">
                    
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-[11px] font-medium text-slate-500 mb-1">Vĩ độ tâm (Lat):</label>
                            <input type="text" id="scrape-lat" readonly class="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded text-xs font-mono font-semibold text-slate-800">
                        </div>
                        <div>
                            <label class="block text-[11px] font-medium text-slate-500 mb-1">Kinh độ tâm (Lon):</label>
                            <input type="text" id="scrape-lon" readonly class="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded text-xs font-mono font-semibold text-slate-800">
                        </div>
                    </div>

                    <div class="flex items-center gap-2 pt-1">
                        <input type="checkbox" id="scrape-grid-auto" checked class="w-4 h-4 text-indigo-600 rounded">
                        <label for="scrape-grid-auto" class="text-xs text-slate-600">
                            <strong>Quét lưới BBox:</strong> Quét phủ toàn bộ diện tích bán kính để lấy số lượng tối đa
                        </label>
                    </div>
                </div>

                <!-- KHU VỰC CHẾ ĐỘ 2: ĐƠN VỊ HÀNH CHÍNH (PHƯỜNG/XÃ/QUẬN/HUYỆN/TỈNH) -->
                <div id="panel-mode-admin" class="hidden flex flex-col gap-3">
                    <div class="grid grid-cols-2 gap-2">
                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">Tỉnh / Thành phố:</label>
                            <select id="select-province" onchange="onProvinceChange()" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-medium">
                                <option value="Hà Nội" selected>Hà Nội</option>
                                <option value="TP. Hồ Chí Minh">TP. Hồ Chí Minh</option>
                                <option value="Đà Nẵng">Đà Nẵng</option>
                                <option value="Hải Phòng">Hải Phòng</option>
                                <option value="Cần Thơ">Cần Thơ</option>
                                <option value="Bình Dương">Bình Dương</option>
                                <option value="Đồng Nai">Đồng Nai</option>
                                <option value="Khánh Hòa">Khánh Hòa</option>
                                <option value="Bà Rịa - Vũng Tàu">Bà Rịa - Vũng Tàu</option>
                                <option value="Quảng Ninh">Quảng Ninh</option>
                                <option value="Nghệ An">Nghệ An</option>
                                <option value="Thừa Thiên Huế">Thừa Thiên Huế</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">Quận / Huyện:</label>
                            <select id="select-district" onchange="onDistrictChange()" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-medium">
                            </select>
                        </div>
                    </div>

                    <div>
                        <div class="flex items-center justify-between mb-1">
                            <label class="text-xs font-bold text-slate-700">Tách cào theo từng Phường/Xã:</label>
                            <button type="button" onclick="selectAllWards(true)" class="text-[11px] text-indigo-600 hover:underline">Chọn tất cả</button>
                        </div>
                        <div id="wards-checkbox-list" class="max-h-28 overflow-y-auto border border-slate-200 rounded-lg p-2 bg-slate-50 flex flex-wrap gap-1.5 custom-scroll text-xs">
                        </div>
                        <p class="text-[11px] text-slate-500 mt-1">Hệ thống sẽ tự động ghép từ khóa cào từng phường để vét cạn số lượng.</p>
                    </div>
                </div>

                <!-- Tùy chọn nâng cao chung -->
                <div class="grid grid-cols-2 gap-3 pt-1 border-t border-slate-100">
                    <div>
                        <label class="block text-[11px] font-medium text-slate-600 mb-1">Độ sâu cuộn (Depth):</label>
                        <select id="scrape-depth" class="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs">
                            <option value="15">Nhanh (~30 địa điểm/vùng)</option>
                            <option value="35" selected>Trung bình (~60-90 chỗ)</option>
                            <option value="70">Tối đa (~120 chỗ/vùng)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-[11px] font-medium text-slate-600 mb-1">Quét Email:</label>
                        <div class="flex items-center h-8">
                            <input type="checkbox" id="scrape-email" class="w-4 h-4 text-indigo-600 rounded">
                            <label for="scrape-email" class="ml-1.5 text-xs text-slate-700">Tìm email website</label>
                        </div>
                    </div>
                </div>

                <!-- Nút cào -->
                <button id="btn-start-scrape" onclick="startScrape()" 
                    class="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg shadow-md transition flex items-center justify-center gap-2 text-sm mt-1">
                    <i data-lucide="play" class="w-4 h-4"></i> Bắt đầu cào dữ liệu
                </button>

                <!-- Hộp Log thời gian thực -->
                <div id="log-box-container" class="hidden">
                    <div class="flex items-center justify-between text-xs font-semibold text-slate-500 mb-1">
                        <span class="flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Trạng thái:</span>
                        <span id="log-status" class="text-indigo-600 font-bold">Đang chạy...</span>
                    </div>
                    <div id="scrape-logs" class="bg-slate-900 text-slate-200 text-[11px] font-mono p-2.5 rounded-lg h-32 overflow-y-auto custom-scroll">
                    </div>
                </div>

            </div>

            <!-- Cột phải: Bản đồ chọn vị trí và xem bán kính / bounding box -->
            <div class="lg:col-span-7 bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col">
                <div class="flex flex-wrap items-center justify-between mb-2 gap-2">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i data-lucide="map" class="w-4 h-4 text-indigo-600"></i> Bản đồ chọn vị trí & Vòng tròn Bán kính
                    </h2>
                    <div class="flex items-center gap-2">
                        <!-- Tìm kiếm Geocoding -->
                        <div class="relative">
                            <input type="text" id="geo-search-input" placeholder="Gõ địa chỉ (VD: Cầu Giấy, Hà Nội)..." 
                                class="pl-2 pr-7 py-1 text-xs border border-slate-300 rounded-lg w-52 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                                onkeypress="if(event.key==='Enter') searchAddressGeo()">
                            <button onclick="searchAddressGeo()" class="absolute right-1.5 top-1 text-slate-400 hover:text-indigo-600" title="Tìm vị trí">
                                <i data-lucide="search" class="w-3.5 h-3.5"></i>
                            </button>
                        </div>
                        <button onclick="locateUser()" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-lg flex items-center gap-1" title="Định vị GPS">
                            <i data-lucide="navigation" class="w-3.5 h-3.5 text-indigo-600"></i> GPS
                        </button>
                    </div>
                </div>

                <div id="picker-map" class="shadow-inner border border-slate-300 flex-1"></div>

                <div class="mt-2.5 flex items-center justify-between text-xs text-slate-500">
                    <span>Tâm: <strong id="picked-location-text" class="text-slate-800">21.0285, 105.8542</strong></span>
                    <span id="radius-desc-text" class="text-indigo-700 font-medium">Bán kính quét: 2.0 km</span>
                </div>
            </div>
        </div>

        <!-- TAB 2: XEM KẾT QUẢ, GỌI ĐIỆN V9 TECH & NHẮN ZALO -->
        <div id="tab-view" class="hidden flex flex-col gap-6">
            <!-- Thanh điều khiển kết quả -->
            <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap items-center justify-between gap-4">
                <div class="flex items-center gap-2">
                    <label class="text-xs font-bold text-slate-700">Tệp dữ liệu:</label>
                    <select id="file-select" onchange="loadFileData()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold min-w-[220px]">
                        <option value="">Đang tải...</option>
                    </select>
                    <button onclick="refreshFileList()" class="p-1.5 text-slate-600 hover:bg-slate-100 rounded-lg" title="Làm mới">
                        <i data-lucide="rotate-cw" class="w-4 h-4"></i>
                    </button>
                </div>

                <!-- Thống kê nhanh -->
                <div class="flex items-center gap-3 text-xs font-semibold">
                    <div class="bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-lg border border-indigo-100">
                        Tổng số: <span id="stat-total" class="font-bold text-sm">0</span>
                    </div>
                    <div class="bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-lg border border-emerald-100">
                        Có số ĐT: <span id="stat-phone" class="font-bold text-sm">0</span>
                    </div>
                    <div class="bg-blue-50 text-blue-700 px-3 py-1.5 rounded-lg border border-blue-100">
                        📱 Số di động (Zalo): <span id="stat-zalo" class="font-bold text-sm">0</span>
                    </div>
                </div>

                <!-- Nút tải file -->
                <div class="flex items-center gap-2">
                    <button onclick="downloadActiveFile('excel')" class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5">
                        <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Tải Excel (.xlsx)
                    </button>
                    <button onclick="downloadActiveFile('csv')" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-800 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5">
                        <i data-lucide="download" class="w-4 h-4"></i> Tải CSV (Tiếng Việt)
                    </button>
                </div>
            </div>

            <!-- Bản đồ hiển thị địa điểm -->
            <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                <div class="flex items-center justify-between mb-2">
                    <h3 class="text-xs font-bold text-slate-700 flex items-center gap-1.5 uppercase tracking-wider">
                        <i data-lucide="map-pinned" class="w-4 h-4 text-indigo-600"></i> Bản đồ phân bố kết quả (Ghim điểm)
                    </h3>
                    <span class="text-xs text-slate-400">Click vào từng ghim để xem thông tin, Gọi PBX hoặc nhắn Zalo</span>
                </div>
                <div id="results-map" class="border border-slate-300"></div>
            </div>

            <!-- Bảng danh sách chi tiết có nút Call V9 PBX & Link Zalo -->
            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
                <div class="p-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 bg-slate-50">
                    <div class="flex items-center gap-3">
                        <div class="relative w-72">
                            <i data-lucide="search" class="w-4 h-4 absolute left-3 top-2.5 text-slate-400"></i>
                            <input type="text" id="table-search" oninput="filterTable()" placeholder="Tìm theo tên, địa chỉ, số ĐT..." 
                                class="w-full pl-9 pr-3 py-1.5 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white">
                        </div>
                        <!-- Lọc chỉ số di động có Zalo -->
                        <label class="flex items-center gap-1.5 text-xs text-slate-600 font-medium cursor-pointer">
                            <input type="checkbox" id="filter-zalo-only" onchange="filterTable()" class="rounded text-indigo-600">
                            Chỉ xem có Zalo (Số di động)
                        </label>
                    </div>
                    <span class="text-xs text-slate-500 font-medium" id="table-showing-info">Đang hiển thị 0 địa điểm</span>
                </div>
                
                <div class="overflow-x-auto max-h-[550px] custom-scroll">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-slate-100 text-slate-700 uppercase font-bold sticky top-0 shadow-sm border-b border-slate-200">
                            <tr>
                                <th class="p-3 w-10">#</th>
                                <th class="p-3 min-w-[180px]">Tên địa điểm</th>
                                <th class="p-3 min-w-[110px]">Danh mục</th>
                                <th class="p-3 min-w-[230px]">Liên hệ / Gọi / Zalo</th>
                                <th class="p-3 min-w-[200px]">Địa chỉ</th>
                                <th class="p-3 text-center">Đánh giá</th>
                                <th class="p-3 text-center">Reviews</th>
                                <th class="p-3">Website</th>
                                <th class="p-3 text-center">Maps</th>
                            </tr>
                        </thead>
                        <tbody id="table-body" class="divide-y divide-slate-100 text-slate-700">
                            <tr>
                                <td colspan="9" class="p-6 text-center text-slate-400">Chưa có dữ liệu</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

    </main>

    <!-- MODAL CẤU HÌNH V9 TECH PBX -->
    <div id="v9-modal" class="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-xl shadow-2xl max-w-md w-full p-6 border border-slate-200">
            <div class="flex items-center justify-between pb-3 border-b border-slate-200">
                <div class="flex items-center gap-2">
                    <i data-lucide="phone-call" class="w-5 h-5 text-emerald-600"></i>
                    <h3 class="font-bold text-slate-800 text-base">Cấu hình Tổng đài V9 Tech</h3>
                </div>
                <button onclick="closeV9Modal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>

            <div class="flex flex-col gap-3 py-4 text-xs">
                <div>
                    <label class="block font-bold text-slate-700 mb-1">Số máy nhánh của bạn (Extension):</label>
                    <input type="text" id="v9-cfg-ext" placeholder="VD: 101, 102, 1001..." 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-semibold text-emerald-700">
                    <p class="text-[11px] text-slate-500 mt-1">Khi bấm gọi, tổng đài sẽ đổ chuông vào máy nhánh này của bạn trước.</p>
                </div>

                <div>
                    <label class="block font-bold text-slate-700 mb-1">URL API Click-to-call V9 Tech:</label>
                    <input type="text" id="v9-cfg-url" placeholder="https://api.v9tech.vn/api/v1/call/click-to-call" 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs font-mono">
                    <p class="text-[11px] text-slate-500 mt-1">Hoặc đường dẫn API tổng đài do V9 Tech cung cấp cho bạn.</p>
                </div>

                <div>
                    <label class="block font-bold text-slate-700 mb-1">API Key / Token (Nếu có):</label>
                    <input type="password" id="v9-cfg-key" placeholder="Nhập mã token xác thực (tùy chọn)..." 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs">
                </div>
            </div>

            <div class="flex items-center justify-end gap-2 pt-3 border-t border-slate-200">
                <button onclick="closeV9Modal()" class="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg">Đóng</button>
                <button onclick="saveV9Config()" class="px-4 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow">Lưu cấu hình</button>
            </div>
        </div>
    </div>

    <!-- TOAST THÔNG BÁO CUỘC GỌI V9 TECH -->
    <div id="call-toast" class="fixed bottom-6 right-6 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl z-50 flex items-center gap-3 border border-slate-700 hidden max-w-sm">
        <div class="p-2 bg-emerald-600 rounded-full animate-bounce">
            <i data-lucide="phone-outgoing" class="w-4 h-4 text-white"></i>
        </div>
        <div class="flex-1 text-xs">
            <div id="call-toast-title" class="font-bold text-emerald-400">Đang khởi tạo cuộc gọi...</div>
            <div id="call-toast-msg" class="text-slate-300 text-[11px] mt-0.5">Vui lòng nhấc máy nhánh...</div>
        </div>
        <button onclick="document.getElementById('call-toast').classList.add('hidden')" class="text-slate-400 hover:text-white">
            <i data-lucide="x" class="w-4 h-4"></i>
        </button>
    </div>

    <script>
        const API_BASE = window.location.pathname.replace(/\/+$/, '');
        function apiUrl(path) {
            if (!path.startsWith('/')) path = '/' + path;
            return API_BASE ? (API_BASE + path) : path;
        }
        let pickerMap, resultsMap;
        let pickerMarker = null;
        let pickerCircle = null;
        let pickerBBox = null;
        let resultsMarkers = [];
        let allPlacesData = [];
        let filteredPlacesData = [];
        let pollInterval = null;
        let currentScrapeMode = 'radius'; // 'radius' hoặc 'admin'

        // Dữ liệu Hành chính mẫu Tỉnh -> Quận -> Phường
        const VIETNAM_ADMIN_DATA = {
            "Hà Nội": {
                lat: 21.0285, lon: 105.8542,
                districts: {
                    "Cầu Giấy": { lat: 21.0362, lon: 105.7905, wards: ["Dịch Vọng", "Dịch Vọng Hậu", "Yên Hòa", "Trung Hòa", "Nghĩa Tân", "Nghĩa Đô", "Quan Hoa", "Mai Dịch"] },
                    "Hoàn Kiếm": { lat: 21.0307, lon: 105.8524, wards: ["Hàng Bạc", "Hàng Bồ", "Hàng Gai", "Hàng Buồm", "Cửa Đông", "Tràng Tiền", "Lý Thái Tổ", "Phan Chu Trinh"] },
                    "Đống Đa": { lat: 21.0181, lon: 105.8286, wards: ["Láng Hạ", "Láng Thượng", "Ô Chợ Dừa", "Văn Miếu", "Kim Liên", "Cát Linh", "Trung Liệt", "Khâm Thiên"] },
                    "Ba Đình": { lat: 21.0353, lon: 105.8262, wards: ["Kim Mã", "Liễu Giai", "Điện Biên", "Quán Thánh", "Đội Cấn", "Giảng Võ", "Ngọc Khánh"] },
                    "Hai Bà Trưng": { lat: 21.0069, lon: 105.8519, wards: ["Bách Khoa", "Lê Đại Hành", "Phố Huế", "Minh Khai", "Trương Định", "Vĩnh Tuy"] },
                    "Thanh Xuân": { lat: 20.9980, lon: 105.8080, wards: ["Nhân Chính", "Thanh Xuân Bắc", "Thanh Xuân Trung", "Khương Trung", "Khương Mai"] },
                    "Nam Từ Liêm": { lat: 21.0134, lon: 105.7656, wards: ["Mỹ Đình 1", "Mỹ Đình 2", "Mễ Trì", "Cầu Diễn", "Phú Đô", "Tây Mỗ"] },
                    "Bắc Từ Liêm": { lat: 21.0592, lon: 105.7602, wards: ["Cổ Nhuế 1", "Cổ Nhuế 2", "Xuân Tảo", "Phú Diễn", "Minh Khai"] },
                    "Hà Đông": { lat: 20.9715, lon: 105.7766, wards: ["Mộ Lao", "Quang Trung", "Văn Quán", "Hà Cầu", "La Khê", "Yết Kiêu"] },
                    "Hoàng Mai": { lat: 20.9754, lon: 105.8488, wards: ["Hoàng Liệt", "Định Công", "Giáp Bát", "Tân Mai", "Mai Động"] },
                    "Tây Hồ": { lat: 21.0660, lon: 105.8220, wards: ["Quảng An", "Xuân La", "Nhật Tân", "Bưởi", "Thụy Khuê", "Yên Phụ"] }
                }
            },
            "TP. Hồ Chí Minh": {
                lat: 10.7769, lon: 106.7009,
                districts: {
                    "Quận 1": { lat: 10.7756, lon: 106.7004, wards: ["Bến Nghé", "Bến Thành", "Đa Kao", "Tân Định", "Phạm Ngũ Lão", "Nguyễn Thái Bình", "Cầu Kho"] },
                    "Quận 3": { lat: 10.7844, lon: 106.6844, wards: ["Phường Võ Thị Sáu", "Phường 1", "Phường 2", "Phường 3", "Phường 4", "Phường 5", "Phường 9"] },
                    "TP. Thủ Đức": { lat: 10.8494, lon: 106.7717, wards: ["Thảo Điền", "An Phú", "Hiệp Bình Chánh", "Linh Trung", "Linh Chiểu", "Tăng Nhơn Phú A"] },
                    "Bình Thạnh": { lat: 10.8106, lon: 106.7091, wards: ["Phường 1", "Phường 2", "Phường 19", "Phường 25", "Phường 26", "Phường 27"] },
                    "Tân Bình": { lat: 10.8015, lon: 106.6558, wards: ["Phường 2", "Phường 4", "Phường 12", "Phường 13", "Phường 15"] },
                    "Quận 7": { lat: 10.7340, lon: 106.7218, wards: ["Tân Phong", "Tân Phú", "Phú Mỹ", "Tân Hưng", "Tân Kiểng"] },
                    "Gò Vấp": { lat: 10.8388, lon: 106.6653, wards: ["Phường 1", "Phường 3", "Phường 5", "Phường 8", "Phường 10", "Phường 11"] },
                    "Phú Nhuận": { lat: 10.7992, lon: 106.6803, wards: ["Phường 1", "Phường 2", "Phường 7", "Phường 9", "Phường 15"] }
                }
            },
            "Đà Nẵng": {
                lat: 16.0544, lon: 108.2022,
                districts: {
                    "Hải Châu": { lat: 16.0620, lon: 108.2208, wards: ["Thạch Thang", "Hải Châu 1", "Hải Châu 2", "Thuận Phước", "Bình Hiên"] },
                    "Thanh Khê": { lat: 16.0617, lon: 108.1883, wards: ["Vĩnh Trung", "Tân Chính", "Thạc Gián", "Chính Gián", "Xuân Hà"] },
                    "Sơn Trà": { lat: 16.0820, lon: 108.2439, wards: ["An Hải Bắc", "An Hải Tây", "Phước Mỹ", "Mân Thái", "Thọ Quang"] },
                    "Ngũ Hành Sơn": { lat: 16.0125, lon: 108.2525, wards: ["Mỹ An", "Khuê Mỹ", "Hòa Quý", "Hòa Hải"] }
                }
            }
        };

        // Khởi tạo bản đồ chọn vị trí
        function initPickerMap() {
            const defaultLat = 21.0285;
            const defaultLon = 105.8542;

            pickerMap = L.map('picker-map').setView([defaultLat, defaultLon], 14);
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                maxZoom: 19,
                attribution: '© OpenStreetMap'
            }).addTo(pickerMap);

            setPickerLocation(defaultLat, defaultLon);

            pickerMap.on('click', function(e) {
                setPickerLocation(e.latlng.lat, e.latlng.lng);
            });
        }

        function setPickerLocation(lat, lon) {
            lat = Number(lat.toFixed(6));
            lon = Number(lon.toFixed(6));

            document.getElementById('scrape-lat').value = lat;
            document.getElementById('scrape-lon').value = lon;
            document.getElementById('picked-location-text').textContent = `${lat}, ${lon}`;

            if (pickerMarker) {
                pickerMarker.setLatLng([lat, lon]);
            } else {
                pickerMarker = L.marker([lat, lon], { draggable: true }).addTo(pickerMap);
                pickerMarker.on('dragend', function(event) {
                    const pos = event.target.getLatLng();
                    setPickerLocation(pos.lat, pos.lng);
                });
            }

            updateRadiusOverlay();
        }

        // Cập nhật Vòng tròn Bán kính & Bounding Box
        function updateRadiusOverlay() {
            if (!pickerMap) return;
            const lat = parseFloat(document.getElementById('scrape-lat').value) || 21.0285;
            const lon = parseFloat(document.getElementById('scrape-lon').value) || 105.8542;
            const radiusKm = parseFloat(document.getElementById('scrape-radius-slider').value) || 2.0;
            const radiusM = radiusKm * 1000;

            // Xóa overlay cũ
            if (pickerCircle) pickerMap.removeLayer(pickerCircle);
            if (pickerBBox) pickerMap.removeLayer(pickerBBox);

            if (currentScrapeMode === 'radius') {
                // Vẽ vòng tròn bán kính
                pickerCircle = L.circle([lat, lon], {
                    color: '#4f46e5',
                    fillColor: '#6366f1',
                    fillOpacity: 0.15,
                    radius: radiusM
                }).addTo(pickerMap);

                // Tính Bounding Box
                const deltaLat = radiusKm / 111.0;
                const deltaLon = radiusKm / (111.0 * Math.cos(lat * Math.PI / 180));
                const minLat = (lat - deltaLat).toFixed(6);
                const maxLat = (lat + deltaLat).toFixed(6);
                const minLon = (lon - deltaLon).toFixed(6);
                const maxLon = (lon + deltaLon).toFixed(6);

                const bounds = [[minLat, minLon], [maxLat, maxLon]];
                pickerBBox = L.rectangle(bounds, {
                    color: '#0284c7',
                    weight: 1,
                    dashArray: '4, 4',
                    fillOpacity: 0.05
                }).addTo(pickerMap);

                window.__calculatedBBox = `${minLat},${minLon},${maxLat},${maxLon}`;
            }
        }

        function updateRadiusFromSlider() {
            const r = parseFloat(document.getElementById('scrape-radius-slider').value);
            const area = (Math.PI * r * r).toFixed(1);
            document.getElementById('radius-val-text').textContent = `${r} Km (~${area} km²)`;
            document.getElementById('radius-desc-text').textContent = `Bán kính: ${r} km (~${area} km²)`;
            updateRadiusOverlay();
        }

        function setScrapeMode(mode) {
            currentScrapeMode = mode;
            if (mode === 'radius') {
                document.getElementById('panel-mode-radius').classList.remove('hidden');
                document.getElementById('panel-mode-admin').classList.add('hidden');
                document.getElementById('btn-mode-radius').className = "py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1 font-bold";
                document.getElementById('btn-mode-admin').className = "py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1";
                updateRadiusOverlay();
            } else {
                document.getElementById('panel-mode-radius').classList.add('hidden');
                document.getElementById('panel-mode-admin').classList.remove('hidden');
                document.getElementById('btn-mode-admin').className = "py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1 font-bold";
                document.getElementById('btn-mode-radius').className = "py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1";
                if (pickerCircle) pickerMap.removeLayer(pickerCircle);
                if (pickerBBox) pickerMap.removeLayer(pickerBBox);
                initAdminDropdowns();
            }
        }

        function initAdminDropdowns() {
            onProvinceChange();
        }

        function onProvinceChange() {
            const provName = document.getElementById('select-province').value;
            const provData = VIETNAM_ADMIN_DATA[provName];
            const distSelect = document.getElementById('select-district');
            distSelect.innerHTML = '';

            if (provData && provData.districts) {
                Object.keys(provData.districts).forEach(distName => {
                    const opt = document.createElement('option');
                    opt.value = distName;
                    opt.textContent = distName;
                    distSelect.appendChild(opt);
                });
                onDistrictChange();
            }
        }

        function onDistrictChange() {
            const provName = document.getElementById('select-province').value;
            const distName = document.getElementById('select-district').value;
            const provData = VIETNAM_ADMIN_DATA[provName];
            if (!provData || !provData.districts[distName]) return;

            const distData = provData.districts[distName];
            // Di chuyển bản đồ đến quận/huyện
            pickerMap.flyTo([distData.lat, distData.lon], 15);
            setPickerLocation(distData.lat, distData.lon);

            // Nạp danh sách phường/xã
            const wardBox = document.getElementById('wards-checkbox-list');
            wardBox.innerHTML = '';
            (distData.wards || []).forEach(w => {
                const label = document.createElement('label');
                label.className = "flex items-center gap-1 bg-white px-2 py-1 rounded border border-slate-200 cursor-pointer hover:bg-indigo-50";
                label.innerHTML = `<input type="checkbox" name="ward-item" value="${w}, ${distName}, ${provName}" checked class="rounded text-indigo-600"> ${w}`;
                wardBox.appendChild(label);
            });
        }

        function selectAllWards(check) {
            document.querySelectorAll('input[name="ward-item"]').forEach(cb => cb.checked = check);
        }

        // Tìm kiếm địa chỉ bằng Nominatim OSM Geocoding
        async function searchAddressGeo() {
            const query = document.getElementById('geo-search-input').value.trim();
            if (!query) return;
            try {
                const res = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(query + ', Vietnam')}&limit=1`);
                const results = await res.json();
                if (results && results.length > 0) {
                    const lat = parseFloat(results[0].lat);
                    const lon = parseFloat(results[0].lon);
                    pickerMap.flyTo([lat, lon], 15);
                    setPickerLocation(lat, lon);
                } else {
                    alert("Không tìm thấy tọa độ cho địa chỉ này!");
                }
            } catch (e) {
                alert("Lỗi tìm kiếm: " + e.message);
            }
        }

        function locateUser() {
            if (navigator.geolocation) {
                navigator.geolocation.getCurrentPosition(function(pos) {
                    const lat = pos.coords.latitude;
                    const lon = pos.coords.longitude;
                    pickerMap.flyTo([lat, lon], 15);
                    setPickerLocation(lat, lon);
                }, function(err) {
                    alert("Không thể lấy vị trí hiện tại: " + err.message);
                });
            } else {
                alert("Trình duyệt không hỗ trợ Geolocation.");
            }
        }

        function switchTab(tabId) {
            document.getElementById('tab-scrape').classList.add('hidden');
            document.getElementById('tab-view').classList.add('hidden');

            document.getElementById('btn-tab-scrape').className = "px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5";
            document.getElementById('btn-tab-view').className = "px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5";

            if (tabId === 'tab-scrape') {
                document.getElementById('tab-scrape').classList.remove('hidden');
                document.getElementById('btn-tab-scrape').className = "px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5";
                setTimeout(() => pickerMap.invalidateSize(), 200);
            } else {
                document.getElementById('tab-view').classList.remove('hidden');
                document.getElementById('btn-tab-view').className = "px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5";
                initResultsMap();
                refreshFileList();
            }
        }

        async function startScrape() {
            const keywords = document.getElementById('scrape-keywords').value.trim();
            if (!keywords) {
                alert("Vui lòng nhập từ khóa tìm kiếm!");
                return;
            }

            const lat = document.getElementById('scrape-lat').value;
            const lon = document.getElementById('scrape-lon').value;
            const radius_km = document.getElementById('scrape-radius-slider').value;
            const depth = document.getElementById('scrape-depth').value;
            const email = document.getElementById('scrape-email').checked;
            const use_grid = document.getElementById('scrape-grid-auto') ? document.getElementById('scrape-grid-auto').checked : false;

            let admin_areas = [];
            if (currentScrapeMode === 'admin') {
                document.querySelectorAll('input[name="ward-item"]:checked').forEach(cb => {
                    admin_areas.push(cb.value);
                });
                if (admin_areas.length === 0) {
                    const prov = document.getElementById('select-province').value;
                    const dist = document.getElementById('select-district').value;
                    admin_areas.push(`${dist}, ${prov}`);
                }
            }

            const btn = document.getElementById('btn-start-scrape');
            btn.disabled = true;
            btn.classList.add('opacity-50');

            document.getElementById('log-box-container').classList.remove('hidden');
            document.getElementById('scrape-logs').innerHTML = "<div>Đang khởi tạo tiến trình cào dữ liệu Google Maps...</div>";
            document.getElementById('log-status').textContent = "Đang chạy...";

            try {
                const res = await fetch(apiUrl('/api/start-scrape'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ 
                        keywords, 
                        mode: currentScrapeMode,
                        lat, 
                        lon, 
                        radius_km, 
                        zoom: 15, 
                        depth, 
                        email, 
                        use_grid,
                        bbox: window.__calculatedBBox,
                        admin_areas
                    })
                });
                const data = await res.json();
                if (!res.ok) {
                    alert(data.error || "Có lỗi xảy ra");
                    btn.disabled = false;
                    btn.classList.remove('opacity-50');
                    return;
                }

                if (pollInterval) clearInterval(pollInterval);
                pollInterval = setInterval(checkScrapeStatus, 2000);
            } catch (e) {
                alert("Lỗi kết nối: " + e.message);
                btn.disabled = false;
                btn.classList.remove('opacity-50');
            }
        }

        async function checkScrapeStatus() {
            try {
                const res = await fetch(apiUrl('/api/status'));
                const data = await res.json();
                
                const logBox = document.getElementById('scrape-logs');
                logBox.innerHTML = data.logs.map(l => `<div>${escapeHtml(l)}</div>`).join('');
                logBox.scrollTop = logBox.scrollHeight;

                if (!data.is_running) {
                    clearInterval(pollInterval);
                    const btn = document.getElementById('btn-start-scrape');
                    btn.disabled = false;
                    btn.classList.remove('opacity-50');

                    if (data.status === 'completed') {
                        document.getElementById('log-status').textContent = "Hoàn tất!";
                        setTimeout(() => {
                            if (confirm("Cào dữ liệu hoàn tất! Bạn có muốn mở xem kết quả ngay không?")) {
                                switchTab('tab-view');
                            }
                        }, 500);
                    } else if (data.status === 'error') {
                        document.getElementById('log-status').textContent = "Có lỗi xảy ra";
                    }
                }
            } catch (e) {
                console.error(e);
            }
        }

        // ================= TAB XEM KẾT QUẢ, V9 PBX & ZALO =================
        function initResultsMap() {
            if (!resultsMap) {
                resultsMap = L.map('results-map').setView([21.0285, 105.8542], 13);
                L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                    maxZoom: 19,
                    attribution: '© OpenStreetMap'
                }).addTo(resultsMap);
            } else {
                setTimeout(() => resultsMap.invalidateSize(), 200);
            }
        }

        async function refreshFileList() {
            const select = document.getElementById('file-select');
            try {
                const res = await fetch(apiUrl('/api/files'));
                const files = await res.json();
                select.innerHTML = '';
                if (files.length === 0) {
                    select.innerHTML = '<option value="">Không có file kết quả nào</option>';
                    return;
                }
                files.forEach((f, idx) => {
                    const opt = document.createElement('option');
                    opt.value = f.name;
                    opt.textContent = `${f.name} (${f.size_kb} KB - ${f.modified})`;
                    if (idx === 0) opt.selected = true;
                    select.appendChild(opt);
                });
                loadFileData();
            } catch (e) {
                console.error(e);
            }
        }

        async function loadFileData() {
            const filename = document.getElementById('file-select').value;
            if (!filename) return;

            try {
                const res = await fetch(apiUrl(`/api/data?filename=${encodeURIComponent(filename)}`));
                const json = await res.json();
                allPlacesData = json.data || [];
                filteredPlacesData = [...allPlacesData];

                // Cập nhật thống kê
                document.getElementById('stat-total').textContent = allPlacesData.length;
                const phoneCount = allPlacesData.filter(p => p.phone_info && p.phone_info.clean !== "").length;
                const zaloCount = allPlacesData.filter(p => p.phone_info && p.phone_info.is_mobile).length;
                document.getElementById('stat-phone').textContent = phoneCount;
                document.getElementById('stat-zalo').textContent = zaloCount;

                filterTable();
            } catch (e) {
                console.error("Lỗi tải file:", e);
            }
        }

        function renderResultsMap() {
            if (!resultsMap) return;

            resultsMarkers.forEach(m => resultsMap.removeLayer(m));
            resultsMarkers = [];

            const bounds = [];
            filteredPlacesData.forEach(p => {
                const lat = parseFloat(p.latitude);
                const lon = parseFloat(p.longitude);
                if (!isNaN(lat) && !isNaN(lon)) {
                    bounds.push([lat, lon]);
                    const marker = L.marker([lat, lon]).addTo(resultsMap);
                    
                    const pInfo = p.phone_info || {};
                    let phoneActions = '';
                    if (pInfo.clean) {
                        phoneActions += `
                            <div class="mt-2 flex items-center gap-1.5">
                                <button onclick="triggerV9Call('${pInfo.clean}', '${escapeHtml(p.title || '')}')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-2 py-1 rounded text-[11px] font-bold flex items-center gap-1 shadow-sm">
                                    📞 Gọi V9 PBX
                                </button>
                                ${pInfo.is_mobile ? `
                                    <a href="${pInfo.zalo_url}" target="_blank" class="zalo-badge px-2 py-1 rounded text-[11px] font-bold flex items-center gap-1 shadow-sm">
                                        💬 Nhắn Zalo
                                    </a>
                                ` : ''}
                            </div>
                        `;
                    }

                    marker.bindPopup(`
                        <div class="text-xs p-1">
                            <strong class="text-sm font-bold text-indigo-700">${escapeHtml(p.title || '')}</strong><br>
                            <span class="text-slate-500">${escapeHtml(p.category || '')}</span><br>
                            ${p.address ? `<b>Đ/C:</b> ${escapeHtml(p.address)}<br>` : ''}
                            ${p.phone ? `<b>ĐT:</b> <a href="tel:${escapeHtml(p.phone)}" class="text-indigo-600 font-bold">${escapeHtml(p.phone)}</a><br>` : ''}
                            ${p.review_rating ? `⭐ <b>${p.review_rating}</b> (${p.review_count || 0} đánh giá)<br>` : ''}
                            ${phoneActions}
                            <div class="mt-2 border-t pt-1">
                                ${p.link ? `<a href="${p.link}" target="_blank" class="text-blue-600 underline">Mở Google Maps</a>` : ''}
                            </div>
                        </div>
                    `);
                    resultsMarkers.push(marker);
                }
            });

            if (bounds.length > 0) {
                resultsMap.fitBounds(bounds, { padding: [30, 30] });
            }
            setTimeout(() => resultsMap.invalidateSize(), 200);
        }

        function filterTable() {
            const query = document.getElementById('table-search').value.toLowerCase().trim();
            const zaloOnly = document.getElementById('filter-zalo-only').checked;

            filteredPlacesData = allPlacesData.filter(p => {
                if (zaloOnly && (!p.phone_info || !p.phone_info.is_mobile)) {
                    return false;
                }
                if (!query) return true;
                const title = (p.title || '').toLowerCase();
                const addr = (p.address || '').toLowerCase();
                const phone = (p.phone || '').toLowerCase();
                const cat = (p.category || '').toLowerCase();
                return title.includes(query) || addr.includes(query) || phone.includes(query) || cat.includes(query);
            });

            renderTable();
            renderResultsMap();
        }

        function renderTable() {
            const tbody = document.getElementById('table-body');
            document.getElementById('table-showing-info').textContent = `Đang hiển thị ${filteredPlacesData.length} / ${allPlacesData.length} địa điểm`;

            if (filteredPlacesData.length === 0) {
                tbody.innerHTML = `<tr><td colspan="9" class="p-6 text-center text-slate-400">Không tìm thấy địa điểm phù hợp</td></tr>`;
                return;
            }

            tbody.innerHTML = filteredPlacesData.map((p, idx) => {
                const pInfo = p.phone_info || clean_vietnamese_phone(p.phone);
                let phoneHtml = `<span class="text-slate-400 italic">Không có SĐT</span>`;
                
                if (pInfo.clean) {
                    phoneHtml = `
                        <div class="flex flex-col gap-1">
                            <div class="flex items-center gap-1.5 font-mono font-bold text-slate-900">
                                <span>${escapeHtml(pInfo.raw || pInfo.clean)}</span>
                                ${pInfo.is_mobile ? `<span class="text-[10px] bg-blue-100 text-blue-800 px-1 py-0.2 rounded font-semibold">Di động</span>` : `<span class="text-[10px] bg-slate-100 text-slate-600 px-1 py-0.2 rounded font-semibold">Cố định</span>`}
                            </div>
                            <div class="flex items-center gap-1.5 mt-0.5">
                                <button onclick="triggerV9Call('${pInfo.clean}', '${escapeHtml(p.title || '')}')" 
                                    class="bg-emerald-600 hover:bg-emerald-700 text-white px-2 py-0.5 rounded text-[11px] font-bold flex items-center gap-1 shadow-sm"
                                    title="Click to Call V9 Tech PBX">
                                    <i data-lucide="phone-call" class="w-3 h-3"></i> Gọi V9
                                </button>
                                ${pInfo.is_mobile ? `
                                    <a href="${pInfo.zalo_url}" target="_blank" 
                                        class="zalo-badge px-2 py-0.5 rounded text-[11px] font-bold flex items-center gap-1 shadow-sm"
                                        title="Nhắn tin Zalo với số ${pInfo.clean}">
                                        <i data-lucide="message-circle" class="w-3 h-3"></i> Zalo
                                    </a>
                                ` : ''}
                                <a href="tel:${pInfo.clean}" class="text-slate-500 hover:text-slate-800 p-1" title="Gọi bằng ứng dụng máy tính (Softphone)">
                                    <i data-lucide="phone" class="w-3 h-3"></i>
                                </a>
                            </div>
                        </div>
                    `;
                }

                return `
                    <tr class="hover:bg-slate-50 transition border-b border-slate-100">
                        <td class="p-3 font-semibold text-slate-400">${idx + 1}</td>
                        <td class="p-3 font-bold text-indigo-950">${escapeHtml(p.title || 'Không tên')}</td>
                        <td class="p-3"><span class="bg-slate-100 text-slate-600 px-2 py-0.5 rounded text-[11px]">${escapeHtml(p.category || '')}</span></td>
                        <td class="p-3">${phoneHtml}</td>
                        <td class="p-3 text-slate-600 max-w-[220px] truncate" title="${escapeHtml(p.address || '')}">${escapeHtml(p.address || '-')}</td>
                        <td class="p-3 text-center">${p.review_rating ? `<span class="bg-amber-50 text-amber-700 px-1.5 py-0.5 rounded font-bold">★ ${p.review_rating}</span>` : '-'}</td>
                        <td class="p-3 text-center text-slate-500">${p.review_count || 0}</td>
                        <td class="p-3">${p.website ? `<a href="${p.website}" target="_blank" class="text-indigo-600 hover:underline max-w-[120px] truncate block" title="${p.website}">Website</a>` : '-'}</td>
                        <td class="p-3 text-center">${p.link ? `<a href="${p.link}" target="_blank" class="text-blue-600 hover:underline font-semibold">Map</a>` : '-'}</td>
                    </tr>
                `;
            }).join('');
            lucide.createIcons();
        }

        // ================= CLICK TO CALL V9 TECH =================
        async function triggerV9Call(phone, name) {
            const toast = document.getElementById('call-toast');
            const titleEl = document.getElementById('call-toast-title');
            const msgEl = document.getElementById('call-toast-msg');

            toast.classList.remove('hidden');
            titleEl.textContent = `Đang gọi ${phone}...`;
            msgEl.textContent = `Đang kích hoạt tổng đài V9 Tech tới ${name || 'khách hàng'}...`;

            try {
                const res = await fetch(apiUrl('/api/v9tech/call'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ phone })
                });
                const data = await res.json();

                if (data.success) {
                    titleEl.textContent = `Đang đổ chuông máy nhánh!`;
                    msgEl.textContent = data.message;
                } else if (data.fallback) {
                    titleEl.textContent = `Khởi tạo cuộc gọi Softphone`;
                    msgEl.textContent = `Bấm xác nhận để gọi ${phone} qua Softphone trên máy tính.`;
                    window.location.href = `tel:${phone}`;
                } else {
                    titleEl.textContent = `Thông báo V9 Tech`;
                    msgEl.textContent = data.message;
                }
            } catch (e) {
                titleEl.textContent = `Gọi trực tiếp Softphone`;
                msgEl.textContent = `Đang mở trình quay số tới ${phone}...`;
                window.location.href = `tel:${phone}`;
            }

            setTimeout(() => {
                toast.classList.add('hidden');
            }, 7000);
        }

        // Cấu hình V9 Tech
        async function openV9Modal() {
            try {
                const res = await fetch(apiUrl('/api/v9tech/config'));
                const cfg = await res.json();
                document.getElementById('v9-cfg-ext').value = cfg.extension || '101';
                document.getElementById('v9-cfg-url').value = cfg.api_url || 'https://api.v9tech.vn/api/v1/call/click-to-call';
                document.getElementById('v9-cfg-key').value = cfg.api_key || '';
            } catch (e) {}
            document.getElementById('v9-modal').classList.remove('hidden');
        }

        function closeV9Modal() {
            document.getElementById('v9-modal').classList.add('hidden');
        }

        async function saveV9Config() {
            const ext = document.getElementById('v9-cfg-ext').value.trim();
            const url = document.getElementById('v9-cfg-url').value.trim();
            const key = document.getElementById('v9-cfg-key').value.trim();

            if (!ext) {
                alert("Vui lòng nhập Số máy nhánh (Extension) của bạn!");
                return;
            }

            try {
                await fetch(apiUrl('/api/v9tech/config'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ extension: ext, api_url: url, api_key: key })
                });
                alert("Đã lưu cấu hình tổng đài V9 Tech thành công!");
                closeV9Modal();
            } catch (e) {
                alert("Lỗi lưu cấu hình: " + e.message);
            }
        }

        function downloadActiveFile(format) {
            const filename = document.getElementById('file-select').value;
            if (!filename) {
                alert("Vui lòng chọn một tệp!");
                return;
            }
            window.location.href = apiUrl(`/api/download?filename=${encodeURIComponent(filename)}&format=${format}`);
        }

        function escapeHtml(str) {
            if (!str) return '';
            const div = document.createElement('div');
            div.textContent = str;
            return div.innerHTML;
        }

        window.addEventListener('DOMContentLoaded', () => {
            initPickerMap();
            lucide.createIcons();
        });
    </script>
</body>
</html>
"""
    return html

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=5000)
