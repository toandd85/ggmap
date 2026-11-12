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
from fastapi.staticfiles import StaticFiles
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
STATIC_DIR = WORKSPACE_DIR / "static"
CONFIG_FILE = WORKSPACE_DIR / "v9tech_config.json"

RESULTS_DIR.mkdir(parents=True, exist_ok=True)
STATIC_DIR.mkdir(parents=True, exist_ok=True)

# Mount static folder for V9 Softphone SDK
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

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
    """Chuẩn hóa và nhận diện số di động Việt Nam để tạo link Zalo & Gọi PBX/WebRTC"""
    if not raw_phone or pd.isna(raw_phone):
        return {"raw": "", "clean": "", "is_mobile": False, "zalo_url": "", "type": "none"}
    
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
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_path, encoding='utf-8-sig', errors='replace')
        
        # Thêm cột Zalo và Loại số điện thoại nếu có cột phone
        if 'phone' in df.columns:
            phone_infos = [clean_vietnamese_phone(p) for p in df['phone']]
            df['phone_clean'] = [p['clean'] for p in phone_infos]
            df['is_mobile'] = ['Di động' if p['is_mobile'] else ('Cố định' if p['clean'] else '') for p in phone_infos]
            df['zalo_link'] = [p['zalo_url'] for p in phone_infos]

        # 1. Ghi đè chính file gốc với UTF-8-SIG (có UTF-8 BOM để Excel mở không bị bể font tiếng Việt)
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')

        # 2. Lưu bản _excel.csv dự phòng
        bom_csv_path = csv_path.with_name(f"{csv_path.stem}_excel.csv")
        df.to_csv(bom_csv_path, index=False, encoding='utf-8-sig')

        # 3. Lưu bản Excel .xlsx chuẩn 100%
        excel_path = csv_path.with_suffix('.xlsx')
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='GoogleMapsData')
    except Exception as e:
        print(f"Lỗi chuẩn hóa file {csv_path}: {e}")

@app.post("/api/fix-all-files")
def fix_all_files():
    """Tự động quét và chuẩn hóa toàn bộ file CSV sang UTF-8 BOM và file Excel .xlsx"""
    fixed = []
    errors = []
    for sp in [RESULTS_DIR, WORKSPACE_DIR]:
        for p in sp.glob("*.csv"):
            if p.name.endswith("_excel.csv"):
                continue
            try:
                fix_and_save_files(p)
                fixed.append(p.name)
            except Exception as e:
                errors.append(f"{p.name}: {str(e)}")
    return {"success": True, "fixed_count": len(fixed), "fixed_files": fixed, "errors": errors}

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

# API Cấu hình V9 Tech (Hỗ trợ cả Softphone SDK WebRTC & PBX API)
@app.get("/api/v9tech/config")
def get_v9tech_config():
    default_config = {
        "mode": "sdk",  # "sdk" (WebRTC trực tiếp trên trình duyệt) hoặc "api" (Click-to-Call PBX)
        "token": "",
        "domain": "demo.v9.com.vn",
        "client_id": "v9-web-client",
        "sip_server": "demo.v9.com.vn",
        "sip_port": 8089,
        "sip_username": "101",
        "sip_password": "",
        "sip_display_name": "Nhân viên kinh doanh",
        "api_url": "https://api.v9tech.vn/api/v1/call/click-to-call",
        "extension": "101",
        "api_key": ""
    }
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default_config.update(saved)
        except Exception:
            pass
    return default_config

@app.post("/api/v9tech/config")
def save_v9tech_config(data: dict):
    try:
        current = get_v9tech_config()
        current.update(data)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
        return {"success": True, "message": "Đã lưu cấu hình V9 Tech thành công", "config": current}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# API Click-to-call V9 Tech (dùng khi ở chế độ API fallback)
@app.post("/api/v9tech/call")
async def v9tech_call(data: dict):
    phone = data.get("phone", "").strip()
    phone_info = clean_vietnamese_phone(phone)
    clean_phone = phone_info["clean"] or phone

    config = get_v9tech_config()
    api_url = data.get("api_url") or config.get("api_url", "")
    extension = data.get("extension") or config.get("extension", "101")
    api_key = data.get("api_key") or config.get("api_key", "")

    if not extension:
        return JSONResponse(status_code=400, content={"error": "Vui lòng nhập Số máy nhánh (Extension) của bạn."})

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

    mode = data.get("mode", "radius")
    lat = data.get("lat")
    lon = data.get("lon")
    radius_km = float(data.get("radius_km", 2.0))
    zoom = int(data.get("zoom", 15))
    depth = int(data.get("depth", 30))
    fetch_email = bool(data.get("email", False))
    use_grid = bool(data.get("use_grid", False))
    bbox = data.get("bbox")

    query_lines = []

    if mode == "admin" and data.get("admin_areas"):
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
        str(WORKSPACE_DIR / "google-maps-scraper.exe"),
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
            has_xlsx = p.with_suffix(".xlsx").exists()
            files.append({
                "name": p.name,
                "path": str(p),
                "size_kb": round(stat.st_size / 1024, 1),
                "has_xlsx": has_xlsx,
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
        try:
            df = pd.read_csv(p, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(p, encoding="utf-8-sig", errors="replace")
        
        df = df.fillna("")
        records = df.to_dict(orient="records")
        
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
    
    # Format CSV: đảm bảo xuất file UTF-8 BOM
    fix_and_save_files(p)
    bom_csv = p.with_name(f"{p.stem}_excel.csv")
    target = bom_csv if bom_csv.exists() else p
    return FileResponse(target, filename=f"{p.stem}_utf8_bom.csv", media_type="text/csv")

# Giao diện HTML Tiếng Việt Chuyên Nghiệp
@app.get("/", response_class=HTMLResponse)
def index_html():
    html = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Google Maps Scraper - Tiếng Việt Pro (WebRTC V9 Tech & Zalo)</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Leaflet Map CSS & JS -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css"/>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- V9 Tech Softphone SDK (WebRTC SIP Dialer) -->
    <script src="/static/v9_softphone_sdk.js"></script>
    <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; }
        #picker-map, #results-map { height: 460px; border-radius: 0.75rem; }
        .custom-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        .zalo-badge { background-color: #0068FF; color: white; }
        .zalo-badge:hover { background-color: #0053cc; }
        .call-pulse { animation: pulse 1.5s cubic-bezier(0.4, 0, 0.6, 1) infinite; }
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
                        <span class="text-[11px] bg-amber-400 text-slate-900 px-2 py-0.5 rounded font-extrabold uppercase">V9 WebRTC & Zalo</span>
                    </h1>
                    <p class="text-xs text-indigo-200">Quét Bán Kính / Phường / Xã • Gọi WebRTC V9 Tech qua Loa/Mic • Nhắn Zalo 1-Chạm • Xuất Excel Chuẩn Tiếng Việt</p>
                </div>
            </div>
            
            <div class="flex items-center space-x-2">
                <!-- SIP Status Badge -->
                <div id="v9-sip-badge" onclick="openV9Modal()" class="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-indigo-900/80 hover:bg-indigo-900 text-xs font-medium cursor-pointer border border-indigo-600/50 shadow-inner" title="Bấm để cấu hình tổng đài V9 Tech">
                    <span id="v9-sip-dot" class="w-2.5 h-2.5 rounded-full bg-slate-400"></span>
                    <span id="v9-sip-text">Chưa kết nối SIP</span>
                </div>

                <button onclick="switchTab('tab-scrape')" id="btn-tab-scrape" class="px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5">
                    <i data-lucide="crosshair" class="w-4 h-4"></i> Chọn vùng & Cào
                </button>
                <button onclick="switchTab('tab-view')" id="btn-tab-view" class="px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5">
                    <i data-lucide="table" class="w-4 h-4"></i> Xem dữ liệu & Gọi
                </button>
                <button onclick="openV9Modal()" class="px-3 py-2 text-xs font-bold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white flex items-center gap-1.5 shadow" title="Cài đặt tổng đài V9 Tech">
                    <i data-lucide="settings" class="w-4 h-4"></i> Cấu hình V9
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
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">quà tặng</textarea>
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
                    </div>
                </div>

                <!-- Tùy chọn nâng cao -->
                <div class="grid grid-cols-2 gap-3 pt-2 border-t border-slate-100">
                    <div>
                        <label class="block text-[11px] font-bold text-slate-600 mb-1">Số lượng tối đa (Depth):</label>
                        <input type="number" id="scrape-depth" value="50" min="5" max="500" step="10" 
                            class="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs font-semibold">
                    </div>
                    <div>
                        <label class="block text-[11px] font-bold text-slate-600 mb-1">Mức Zoom:</label>
                        <select id="scrape-zoom" class="w-full px-2.5 py-1.5 border border-slate-300 rounded text-xs">
                            <option value="14">14 - Rộng</option>
                            <option value="15" selected>15 - Chuẩn</option>
                            <option value="16">16 - Chi tiết khu phố</option>
                        </select>
                    </div>
                </div>

                <div class="flex items-center gap-2">
                    <input type="checkbox" id="scrape-email" class="w-4 h-4 text-indigo-600 rounded">
                    <label for="scrape-email" class="text-xs text-slate-600">
                        Cào thêm Email từ trang web của địa điểm
                    </label>
                </div>

                <!-- Nút bấm bắt đầu cào -->
                <button type="button" id="btn-start-scrape" onclick="triggerScrape()" 
                    class="w-full py-2.5 px-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg shadow-md flex items-center justify-center gap-2 text-sm transition">
                    <i data-lucide="play" class="w-4 h-4"></i> Bắt đầu Quét Google Maps
                </button>

                <!-- Khung hiển thị log tiến trình -->
                <div id="progress-box" class="hidden flex flex-col gap-2 mt-1">
                    <div class="flex items-center justify-between text-xs">
                        <span class="font-bold text-indigo-700 flex items-center gap-1.5">
                            <span class="w-2 h-2 rounded-full bg-indigo-600 animate-ping"></span>
                            Tiến trình cào dữ liệu:
                        </span>
                        <span id="scrape-status-tag" class="px-2 py-0.5 rounded text-[11px] font-bold bg-amber-100 text-amber-800">Đang chạy</span>
                    </div>
                    <div id="scrape-logs" class="bg-slate-900 text-emerald-400 p-2.5 rounded-lg text-[11px] font-mono h-32 overflow-y-auto custom-scroll leading-relaxed whitespace-pre-wrap"></div>
                </div>

            </div>

            <!-- Cột phải: Bản đồ chọn vị trí và xem bán kính / bounding box -->
            <div class="lg:col-span-7 bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col">
                <div class="flex flex-wrap items-center justify-between mb-2 gap-2">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i data-lucide="map" class="w-4 h-4 text-indigo-600"></i> Bản đồ chọn vị trí & Vòng tròn Bán kính
                    </h2>
                    <div class="flex items-center gap-2">
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
                    <button onclick="refreshFileList()" class="p-1.5 text-slate-600 hover:bg-slate-100 rounded-lg" title="Làm mới danh sách file">
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

                <!-- Nút thao tác tải file & sửa lỗi font -->
                <div class="flex items-center gap-2">
                    <button onclick="fixAllFilesVietnamese()" class="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5" title="Sửa triệt để lỗi font tiếng Việt khi mở Excel">
                        <i data-lucide="wrench" class="w-4 h-4"></i> Sửa lỗi font Excel
                    </button>
                    <button onclick="downloadActiveFile('excel')" class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5" title="Tải file Excel .xlsx chuẩn không bao giờ lỗi font">
                        <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Tải Excel (.xlsx)
                    </button>
                    <button onclick="downloadActiveFile('csv')" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-800 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5" title="Tải file CSV đã chèn sẵn UTF-8 BOM">
                        <i data-lucide="download" class="w-4 h-4"></i> Tải CSV (UTF-8 BOM)
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

    <!-- FLOATING WEBRTC SOFTPHONE CALL WIDGET (Góc dưới bên phải) -->
    <div id="v9-call-widget" class="fixed bottom-5 right-5 z-50 bg-white rounded-2xl shadow-2xl border border-slate-200 w-80 p-5 hidden flex-col transition-all">
        <div class="flex items-center justify-between pb-3 border-b border-slate-100">
            <div class="flex items-center gap-2">
                <span class="relative flex h-3 w-3">
                    <span id="widget-pulse" class="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                    <span id="widget-dot" class="relative inline-flex rounded-full h-3 w-3 bg-emerald-500"></span>
                </span>
                <span class="font-bold text-xs text-slate-700 uppercase tracking-wider">V9 WebRTC Softphone</span>
            </div>
            <button onclick="closeCallWidget()" class="text-slate-400 hover:text-slate-600" title="Đóng widget">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>
        </div>
        <div class="py-4 text-center flex flex-col items-center">
            <div id="call-icon-box" class="w-16 h-16 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center mb-2 shadow-inner">
                <i data-lucide="phone-outgoing" class="w-8 h-8 call-pulse" id="call-icon"></i>
            </div>
            <h4 class="font-bold text-base text-slate-800 font-mono tracking-wide" id="call-phone-number">0988 896 885</h4>
            <p class="text-xs text-slate-500 font-medium mt-0.5" id="call-status-label">Đang kết nối...</p>
            <p class="text-sm font-mono font-bold text-emerald-600 mt-1" id="call-timer">00:00</p>
        </div>
        <div class="flex items-center justify-center gap-4 pt-2 border-t border-slate-100">
            <button id="btn-call-mute" onclick="toggleCallMute()" class="p-3 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 transition" title="Bật/Tắt Mic">
                <i data-lucide="mic" class="w-5 h-5" id="call-mute-icon"></i>
            </button>
            <button id="btn-call-hold" onclick="toggleCallHold()" class="p-3 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 transition" title="Tạm giữ cuộc gọi">
                <i data-lucide="pause" class="w-5 h-5" id="call-hold-icon"></i>
            </button>
            <button onclick="hangupCurrentCall()" class="p-3.5 rounded-full bg-rose-600 hover:bg-rose-700 text-white shadow-lg transition" title="Cúp máy">
                <i data-lucide="phone-off" class="w-5 h-5"></i>
            </button>
        </div>
    </div>

    <!-- MODAL CẤU HÌNH TỔNG ĐÀI V9 TECH -->
    <div id="v9-modal" class="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-xl shadow-2xl max-w-lg w-full p-6 border border-slate-200">
            <div class="flex items-center justify-between pb-3 border-b border-slate-200">
                <div class="flex items-center gap-2">
                    <i data-lucide="phone-call" class="w-5 h-5 text-emerald-600"></i>
                    <h3 class="font-bold text-slate-800 text-base">Cấu hình Tổng đài V9 Tech</h3>
                </div>
                <button onclick="closeV9Modal()" class="text-slate-400 hover:text-slate-600">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>

            <!-- Tab chọn chế độ: WebRTC Softphone hoặc API Click-to-Call -->
            <div class="grid grid-cols-2 gap-2 mt-4 bg-slate-100 p-1 rounded-lg text-xs font-semibold">
                <button type="button" id="tab-v9-sdk" onclick="switchV9ModeTab('sdk')" class="py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1.5">
                    <i data-lucide="mic" class="w-3.5 h-3.5"></i> WebRTC Gọi trên Web
                </button>
                <button type="button" id="tab-v9-api" onclick="switchV9ModeTab('api')" class="py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1.5">
                    <i data-lucide="phone-forwarded" class="w-3.5 h-3.5"></i> Click-to-Call PBX API
                </button>
            </div>

            <!-- PHẦN 1: CẤU HÌNH WEBRTC SDK -->
            <div id="panel-v9-sdk" class="flex flex-col gap-3 py-4 text-xs">
                <div class="bg-indigo-50 border border-indigo-100 text-indigo-700 p-2.5 rounded-lg text-[11px] leading-relaxed">
                    <strong>WebRTC Softphone V9 Tech:</strong> Gọi điện trực tiếp từ trình duyệt qua Tai nghe & Micro máy tính, không cần cài đặt Zoiper hay phần mềm bên ngoài.
                </div>
                <div>
                    <label class="block font-bold text-slate-700 mb-1">JWT Token (Xác thực V9 Tech):</label>
                    <textarea id="v9-sdk-token" rows="2" placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..." class="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-mono"></textarea>
                </div>
                <div class="grid grid-cols-2 gap-2">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Domain PBX:</label>
                        <input type="text" id="v9-sdk-domain" placeholder="demo.v9.com.vn" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Client ID:</label>
                        <input type="text" id="v9-sdk-client-id" placeholder="v9-web-client" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-2">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">SIP Server (WSS/Host):</label>
                        <input type="text" id="v9-sdk-server" placeholder="demo.v9.com.vn" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">SIP Port:</label>
                        <input type="number" id="v9-sdk-port" value="8089" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-2">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Máy nhánh (Extension):</label>
                        <input type="text" id="v9-sdk-username" placeholder="101" class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-bold text-indigo-700">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Mật khẩu SIP:</label>
                        <input type="password" id="v9-sdk-password" placeholder="Mật khẩu máy nhánh..." class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs">
                    </div>
                </div>
            </div>

            <!-- PHẦN 2: CẤU HÌNH API CLICK-TO-CALL (PBX TỔNG ĐÀI) -->
            <div id="panel-v9-api" class="hidden flex flex-col gap-3 py-4 text-xs">
                <div>
                    <label class="block font-bold text-slate-700 mb-1">Số máy nhánh của bạn (Extension):</label>
                    <input type="text" id="v9-cfg-ext" placeholder="VD: 101, 102..." class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs font-semibold text-emerald-700">
                </div>
                <div>
                    <label class="block font-bold text-slate-700 mb-1">URL API Click-to-call V9 Tech:</label>
                    <input type="text" id="v9-cfg-url" placeholder="https://api.v9tech.vn/api/v1/call/click-to-call" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs font-mono">
                </div>
                <div>
                    <label class="block font-bold text-slate-700 mb-1">API Key / Token:</label>
                    <input type="password" id="v9-cfg-key" placeholder="Nhập API Key nếu có..." class="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs">
                </div>
            </div>

            <div class="flex items-center justify-end gap-2 pt-3 border-t border-slate-200">
                <button onclick="closeV9Modal()" class="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg">Đóng</button>
                <button onclick="saveV9ConfigAndConnect()" class="px-4 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow flex items-center gap-1.5">
                    <i data-lucide="check" class="w-4 h-4"></i> Lưu & Kết nối
                </button>
            </div>
        </div>
    </div>

    <!-- TOAST THÔNG BÁO CUỘC GỌI V9 TECH -->
    <div id="v9-toast" class="fixed bottom-5 left-5 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-xl flex items-center gap-3 hidden border border-slate-700 max-w-sm">
        <div class="p-2 bg-emerald-600 rounded-lg">
            <i data-lucide="phone-outgoing" class="w-5 h-5"></i>
        </div>
        <div class="flex-1 text-xs">
            <h4 class="font-bold text-emerald-400" id="toast-title">Đang gọi V9 Tech</h4>
            <p class="text-slate-300 mt-0.5" id="toast-message">Đang khởi tạo...</p>
        </div>
        <button onclick="document.getElementById('v9-toast').classList.add('hidden')" class="text-slate-400 hover:text-white">
            <i data-lucide="x" class="w-4 h-4"></i>
        </button>
    </div>

    <!-- Container ẩn cho Audio WebRTC -->
    <div id="sip-dialer-container" style="display:none;"></div>

    <script>
        const BASE_PATH = window.location.pathname.startsWith('/ggmdata') ? '/ggmdata' : '';
        function apiUrl(endpoint) { return BASE_PATH + endpoint; }

        let pickerMap, pickerMarker, pickerCircle;
        let resultsMap, resultsLayerGroup;
        let currentData = [];
        let currentMode = 'radius';
        let currentV9Mode = 'sdk';

        // V9 Softphone WebRTC State
        let v9Dialer = null;
        let v9Config = {};
        let callTimerInterval = null;
        let callSeconds = 0;
        let currentActivePhone = '';
        let isMuted = false;
        let isHeld = false;

        // Dữ liệu Hành chính mẫu của Việt Nam
        const VIETNAM_ADMIN_DATA = {
            "Hà Nội": {
                "Ba Đình": ["Phúc Xá", "Trúc Bạch", "Vĩnh Phúc", "Cống Vị", "Liễu Giai", "Nguyễn Trung Trực", "Quán Thánh", "Điện Biên", "Đội Cấn", "Ngọc Hà", "Kim Mã", "Giảng Võ", "Thành Công", "Ngọc Khánh"],
                "Hoàn Kiếm": ["Hàng Bạc", "Hàng Bồ", "Hàng Buồm", "Hàng Đào", "Hàng Gai", "Hàng Mã", "Hàng Trống", "Cửa Đông", "Cửa Nam", "Đồng Xuân", "Lý Thái Tổ", "Phan Chu Trinh", "Phúc Tân", "Trần Hưng Đạo", "Tràng Tiền", "Chương Dương"],
                "Cầu Giấy": ["Dịch Vọng", "Dịch Vọng Hậu", "Mai Dịch", "Nghĩa Đô", "Nghĩa Tân", "Quan Hoa", "Trung Hòa", "Yên Hòa"],
                "Đống Đa": ["Cát Linh", "Hàng Bột", "Khâm Thiên", "Khương Thượng", "Kim Liên", "Láng Hạ", "Láng Thượng", "Nam Đồng", "Ngã Tư Sở", "Ô Chợ Dừa", "Phương Liệt", "Phương Mai", "Quang Trung", "Quốc Tử Giám", "Thịnh Quang", "Thổ Quan", "Trung Liệt", "Trung Phụng", "Trung Tự", "Văn Chương", "Văn Miếu"],
                "Hai Bà Trưng": ["Bạch Đằng", "Bách Khoa", "Bạch Mai", "Cầu Dền", "Đống Mác", "Đồng Nhân", "Đồng Tâm", "Lê Đại Hành", "Minh Khai", "Nguyễn Du", "Phạm Đình Hổ", "Phố Huế", "Quỳnh Lôi", "Quỳnh Mai", "Thanh Lương", "Thanh Nhàn", "Trương Định", "Vĩnh Tuy"],
                "Thanh Xuân": ["Hạ Đình", "Khương Đình", "Khương Mai", "Khương Trung", "Kim Giang", "Nhân Chính", "Phương Liệt", "Thanh Xuân Bắc", "Thanh Xuân Nam", "Thanh Xuân Trung", "Thượng Đình"],
                "Nam Từ Liêm": ["Cầu Diễn", "Đại Mỗ", "Mễ Trì", "Mỹ Đình 1", "Mỹ Đình 2", "Phú Đô", "Phương Canh", "Tây Mỗ", "Trung Văn", "Xuân Phương"],
                "Bắc Từ Liêm": ["Cổ Nhuế 1", "Cổ Nhuế 2", "Đông Ngạc", "Đức Thắng", "Liên Mạc", "Minh Khai", "Phú Diễn", "Phúc Diễn", "Tây Tựu", "Thượng Cát", "Thụy Phương", "Xuân Đỉnh", "Xuân Tảo"],
                "Tây Hồ": ["Bưởi", "Nhật Tân", "Phú Thượng", "Quảng An", "Thụy Khuê", "Tứ Liên", "Xuân La", "Yên Phụ"],
                "Long Biên": ["Bồ Đề", "Cự Khối", "Đức Giang", "Gia Thụy", "Giang Biên", "Long Biên", "Ngọc Lâm", "Ngọc Thụy", "Phúc Đồng", "Phúc Lợi", "Sài Đồng", "Thạch Bàn", "Thượng Thanh", "Việt Hưng"],
                "Hà Đông": ["Biên Giang", "Dương Nội", "Đồng Mai", "Hà Cầu", "Kiến Hưng", "La Khê", "Mộ Lao", "Nguyễn Trãi", "Phú La", "Phú Lãm", "Phú Lương", "Phúc La", "Quang Trung", "Vạn Phúc", "Văn Quán", "Yên Nghĩa", "Yết Kiêu"],
                "Hoàng Mai": ["Đại Kim", "Định Công", "Giáp Bát", "Hoàng Liệt", "Hoàng Văn Thụ", "Lĩnh Nam", "Mai Động", "Tân Mai", "Thanh Trì", "Thịnh Liệt", "Trần Phú", "Tương Mai", "Vĩnh Hưng", "Yên Sở"]
            },
            "TP. Hồ Chí Minh": {
                "Quận 1": ["Bến Nghé", "Bến Thành", "Cô Giang", "Cầu Kho", "Cầu Ông Lãnh", "Đa Kao", "Nguyễn Cư Trinh", "Nguyễn Thái Bình", "Phạm Ngũ Lão", "Tân Định"],
                "Quận 3": ["Phường 1", "Phường 2", "Phường 3", "Phường 4", "Phường 5", "Phường 9", "Phường 10", "Phường 11", "Phường 12", "Phường 13", "Phường 14", "Võ Thị Sáu"],
                "Quận 7": ["Bình Thuận", "Phú Mỹ", "Phú Thuận", "Tân Hưng", "Tân Kiểng", "Tân Phong", "Tân Phú", "Tân Quy", "Tân Thuận Đông", "Tân Thuận Tây"],
                "Quận 10": ["Phường 1", "Phường 2", "Phường 4", "Phường 5", "Phường 6", "Phường 7", "Phường 8", "Phường 9", "Phường 10", "Phường 11", "Phường 12", "Phường 13", "Phường 14", "Phường 15"],
                "Bình Thạnh": ["Phường 1", "Phường 2", "Phường 3", "Phường 5", "Phường 6", "Phường 7", "Phường 11", "Phường 12", "Phường 13", "Phường 14", "Phường 15", "Phường 17", "Phường 19", "Phường 21", "Phường 22", "Phường 24", "Phường 25", "Phường 26", "Phường 27", "Phường 28"],
                "Tân Bình": ["Phường 1", "Phường 2", "Phường 3", "Phường 4", "Phường 5", "Phường 6", "Phường 7", "Phường 8", "Phường 9", "Phường 10", "Phường 11", "Phường 12", "Phường 13", "Phường 14", "Phường 15"],
                "TP. Thủ Đức": ["An Khánh", "An Lợi Đông", "An Phú", "Bình Chiểu", "Bình Thọ", "Hiệp Bình Chánh", "Hiệp Bình Phước", "Hiệp Phú", "Linh Chiểu", "Linh Đông", "Linh Trung", "Linh Xuân", "Long Bình", "Long Phước", "Long Thạnh Mỹ", "Long Trường", "Phú Hữu", "Phước Bình", "Phước Long A", "Phước Long B", "Tam Bình", "Tam Phú", "Tăng Nhơn Phú A", "Tăng Nhơn Phú B", "Thảo Điền", "Thạnh Mỹ Lợi", "Thủ Thiêm", "Trường Thạnh", "Trường Thọ"]
            },
            "Đà Nẵng": {
                "Hải Châu": ["Hải Châu 1", "Hải Châu 2", "Thạch Thang", "Thanh Bình", "Thuận Phước", "Hòa Thuận Tây", "Hòa Thuận Đông", "Nam Dương", "Bình Hiên", "Bình Thuận", "Hòa Cường Bắc", "Hòa Cường Nam"],
                "Thanh Khê": ["Tam Thuận", "Thanh Khê Tây", "Thanh Khê Đông", "Xuân Hà", "Tân Chính", "Chính Gián", "Vĩnh Trung", "Thạc Gián", "An Khê", "Hòa Khê"],
                "Sơn Trà": ["An Hải Bắc", "An Hải Đông", "An Hải Tây", "Mân Thái", "Nại Hiên Đông", "Phước Mỹ", "Thọ Quang"]
            }
        };

        function switchTab(tabId) {
            document.getElementById('tab-scrape').classList.toggle('hidden', tabId !== 'tab-scrape');
            document.getElementById('tab-view').classList.toggle('hidden', tabId !== 'tab-view');
            
            document.getElementById('btn-tab-scrape').className = tabId === 'tab-scrape' 
                ? 'px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5'
                : 'px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5';
            
            document.getElementById('btn-tab-view').className = tabId === 'tab-view'
                ? 'px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-900 text-white shadow flex items-center gap-1.5'
                : 'px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5';

            if (tabId === 'tab-view') {
                if (!resultsMap) initResultsMap();
                refreshFileList();
            } else {
                if (pickerMap) setTimeout(() => pickerMap.invalidateSize(), 200);
            }
        }

        function setScrapeMode(mode) {
            currentMode = mode;
            document.getElementById('panel-mode-radius').classList.toggle('hidden', mode !== 'radius');
            document.getElementById('panel-mode-admin').classList.toggle('hidden', mode !== 'admin');
            
            document.getElementById('btn-mode-radius').className = mode === 'radius'
                ? 'py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1'
                : 'py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1';
            
            document.getElementById('btn-mode-admin').className = mode === 'admin'
                ? 'py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1'
                : 'py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1';

            if (mode === 'admin') onProvinceChange();
        }

        function initPickerMap() {
            const defaultLat = 21.028511;
            const defaultLon = 105.854167;
            
            pickerMap = L.map('picker-map').setView([defaultLat, defaultLon], 14);
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                maxZoom: 19,
                attribution: '© OpenStreetMap'
            }).addTo(pickerMap);

            pickerMarker = L.marker([defaultLat, defaultLon], { draggable: true }).addTo(pickerMap);
            pickerCircle = L.circle([defaultLat, defaultLon], {
                radius: 2000,
                color: '#4f46e5',
                fillColor: '#6366f1',
                fillOpacity: 0.18
            }).addTo(pickerMap);

            updatePickerInputs(defaultLat, defaultLon);

            pickerMarker.on('drag', function(e) {
                const pos = e.target.getLatLng();
                pickerCircle.setLatLng(pos);
                updatePickerInputs(pos.lat, pos.lng);
            });

            pickerMap.on('click', function(e) {
                pickerMarker.setLatLng(e.latlng);
                pickerCircle.setLatLng(e.latlng);
                updatePickerInputs(e.latlng.lat, e.latlng.lng);
            });
        }

        function updatePickerInputs(lat, lon) {
            document.getElementById('scrape-lat').value = lat.toFixed(6);
            document.getElementById('scrape-lon').value = lon.toFixed(6);
            document.getElementById('picked-location-text').textContent = `${lat.toFixed(4)}, ${lon.toFixed(4)}`;
        }

        function updateRadiusFromSlider() {
            const rKm = parseFloat(document.getElementById('scrape-radius-slider').value);
            const area = (Math.PI * rKm * rKm).toFixed(1);
            document.getElementById('radius-val-text').textContent = `${rKm.toFixed(1)} Km (~${area} km²)`;
            document.getElementById('radius-desc-text').textContent = `Bán kính quét: ${rKm.toFixed(1)} km`;
            if (pickerCircle) pickerCircle.setRadius(rKm * 1000);
        }

        async function searchAddressGeo() {
            const q = document.getElementById('geo-search-input').value.trim();
            if (!q) return;
            try {
                const res = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(q)}&countrycodes=vn&limit=1`);
                const data = await res.json();
                if (data && data.length > 0) {
                    const lat = parseFloat(data[0].lat);
                    const lon = parseFloat(data[0].lon);
                    pickerMap.setView([lat, lon], 14);
                    pickerMarker.setLatLng([lat, lon]);
                    pickerCircle.setLatLng([lat, lon]);
                    updatePickerInputs(lat, lon);
                } else {
                    alert("Không tìm thấy địa chỉ. Vui lòng thử lại với từ khóa khác!");
                }
            } catch (e) {
                alert("Lỗi tìm kiếm vị trí: " + e.message);
            }
        }

        function locateUser() {
            if ("geolocation" in navigator) {
                navigator.geolocation.getCurrentPosition((pos) => {
                    const lat = pos.coords.latitude;
                    const lon = pos.coords.longitude;
                    pickerMap.setView([lat, lon], 15);
                    pickerMarker.setLatLng([lat, lon]);
                    pickerCircle.setLatLng([lat, lon]);
                    updatePickerInputs(lat, lon);
                }, (err) => {
                    alert("Không thể lấy vị trí GPS: " + err.message);
                });
            } else {
                alert("Trình duyệt không hỗ trợ Geolocation.");
            }
        }

        function onProvinceChange() {
            const p = document.getElementById('select-province').value;
            const distSelect = document.getElementById('select-district');
            distSelect.innerHTML = '';
            
            const districts = VIETNAM_ADMIN_DATA[p] ? Object.keys(VIETNAM_ADMIN_DATA[p]) : [];
            if (districts.length === 0) {
                distSelect.innerHTML = `<option value="">(Nhập khu vực theo ${p})</option>`;
                updateWardsList([]);
                return;
            }
            districts.forEach(d => {
                const opt = document.createElement('option');
                opt.value = d;
                opt.textContent = d;
                distSelect.appendChild(opt);
            });
            onDistrictChange();
        }

        function onDistrictChange() {
            const p = document.getElementById('select-province').value;
            const d = document.getElementById('select-district').value;
            const wards = (VIETNAM_ADMIN_DATA[p] && VIETNAM_ADMIN_DATA[p][d]) ? VIETNAM_ADMIN_DATA[p][d] : [];
            updateWardsList(wards);
        }

        function updateWardsList(wards) {
            const box = document.getElementById('wards-checkbox-list');
            box.innerHTML = '';
            if (wards.length === 0) {
                box.innerHTML = `<span class="text-slate-400 text-[11px]">Không có danh sách phường/xã chi tiết.</span>`;
                return;
            }
            wards.forEach(w => {
                const lbl = document.createElement('label');
                lbl.className = 'inline-flex items-center gap-1 bg-white border border-slate-200 px-2 py-0.5 rounded text-[11px] cursor-pointer hover:border-indigo-400';
                lbl.innerHTML = `<input type="checkbox" name="ward-item" value="${w}" checked class="rounded text-indigo-600"> ${w}`;
                box.appendChild(lbl);
            });
        }

        function selectAllWards(check) {
            document.querySelectorAll('input[name="ward-item"]').forEach(el => el.checked = check);
        }

        async function triggerScrape() {
            const keywords = document.getElementById('scrape-keywords').value.trim();
            if (!keywords) {
                alert("Vui lòng nhập từ khóa tìm kiếm!");
                return;
            }

            const payload = {
                keywords: keywords,
                mode: currentMode,
                depth: parseInt(document.getElementById('scrape-depth').value) || 30,
                zoom: parseInt(document.getElementById('scrape-zoom').value) || 15,
                email: document.getElementById('scrape-email').checked
            };

            if (currentMode === 'radius') {
                const lat = parseFloat(document.getElementById('scrape-lat').value);
                const lon = parseFloat(document.getElementById('scrape-lon').value);
                const rKm = parseFloat(document.getElementById('scrape-radius-slider').value);
                payload.lat = lat;
                payload.lon = lon;
                payload.radius_km = rKm;

                if (document.getElementById('scrape-grid-auto').checked && pickerCircle) {
                    const bounds = pickerCircle.getBounds();
                    payload.use_grid = true;
                    payload.bbox = `${bounds.getSouth().toFixed(6)},${bounds.getWest().toFixed(6)},${bounds.getNorth().toFixed(6)},${bounds.getEast().toFixed(6)}`;
                }
            } else {
                const province = document.getElementById('select-province').value;
                const district = document.getElementById('select-district').value;
                const checkedWards = Array.from(document.querySelectorAll('input[name="ward-item"]:checked')).map(el => el.value);
                
                let areas = [];
                if (checkedWards.length > 0) {
                    areas = checkedWards.map(w => `Phường ${w}, ${district}, ${province}`);
                } else if (district) {
                    areas = [`${district}, ${province}`];
                } else {
                    areas = [province];
                }
                payload.admin_areas = areas;
            }

            const btn = document.getElementById('btn-start-scrape');
            btn.disabled = true;
            btn.classList.add('opacity-50', 'cursor-not-allowed');

            document.getElementById('progress-box').classList.remove('hidden');
            document.getElementById('scrape-logs').textContent = "Đang gửi yêu cầu cào dữ liệu...";

            try {
                const res = await fetch(apiUrl('/api/start-scrape'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const data = await res.json();
                if (res.ok) {
                    pollScrapeStatus();
                } else {
                    alert("Lỗi: " + (data.error || "Không thể khởi động cào"));
                    btn.disabled = false;
                    btn.classList.remove('opacity-50', 'cursor-not-allowed');
                }
            } catch (e) {
                alert("Lỗi kết nối máy chủ: " + e.message);
                btn.disabled = false;
                btn.classList.remove('opacity-50', 'cursor-not-allowed');
            }
        }

        let pollInterval = null;
        function pollScrapeStatus() {
            if (pollInterval) clearInterval(pollInterval);
            pollInterval = setInterval(async () => {
                try {
                    const res = await fetch(apiUrl('/api/status'));
                    const status = await res.json();
                    
                    const logBox = document.getElementById('scrape-logs');
                    logBox.textContent = status.logs.join('\n');
                    logBox.scrollTop = logBox.scrollHeight;

                    const statusTag = document.getElementById('scrape-status-tag');
                    if (status.status === 'running') {
                        statusTag.className = 'px-2 py-0.5 rounded text-[11px] font-bold bg-amber-100 text-amber-800';
                        statusTag.textContent = 'Đang chạy...';
                    } else if (status.status === 'completed') {
                        statusTag.className = 'px-2 py-0.5 rounded text-[11px] font-bold bg-emerald-100 text-emerald-800';
                        statusTag.textContent = 'Hoàn tất!';
                        clearInterval(pollInterval);
                        
                        const btn = document.getElementById('btn-start-scrape');
                        btn.disabled = false;
                        btn.classList.remove('opacity-50', 'cursor-not-allowed');
                        
                        // Tự động chuyển qua tab xem dữ liệu
                        setTimeout(() => {
                            switchTab('tab-view');
                            refreshFileList();
                        }, 1200);
                    } else if (status.status === 'error') {
                        statusTag.className = 'px-2 py-0.5 rounded text-[11px] font-bold bg-rose-100 text-rose-800';
                        statusTag.textContent = 'Có lỗi xảy ra';
                        clearInterval(pollInterval);
                        const btn = document.getElementById('btn-start-scrape');
                        btn.disabled = false;
                        btn.classList.remove('opacity-50', 'cursor-not-allowed');
                    }
                } catch (e) {}
            }, 1500);
        }

        // ==================== TAB 2: XEM KẾT QUẢ & BẢN ĐỒ ====================
        function initResultsMap() {
            resultsMap = L.map('results-map').setView([21.028511, 105.854167], 13);
            L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                maxZoom: 19,
                attribution: '© OpenStreetMap'
            }).addTo(resultsMap);
            resultsLayerGroup = L.layerGroup().addTo(resultsMap);
        }

        async function refreshFileList() {
            try {
                const res = await fetch(apiUrl('/api/files'));
                const files = await res.json();
                const sel = document.getElementById('file-select');
                sel.innerHTML = '';
                if (!files || files.length === 0) {
                    sel.innerHTML = '<option value="">Không có file dữ liệu</option>';
                    return;
                }
                files.forEach(f => {
                    const opt = document.createElement('option');
                    opt.value = f.name;
                    opt.textContent = `${f.name} (${f.size_kb} KB - ${f.modified})`;
                    sel.appendChild(opt);
                });
                loadFileData();
            } catch (e) {
                console.error("Lỗi lấy danh sách file:", e);
            }
        }

        async function loadFileData() {
            const filename = document.getElementById('file-select').value;
            if (!filename) return;

            try {
                const res = await fetch(apiUrl(`/api/data?filename=${encodeURIComponent(filename)}`));
                const json = await res.json();
                if (res.ok) {
                    currentData = json.data || [];
                    renderStats(currentData);
                    renderResultsMap(currentData);
                    renderTable(currentData);
                } else {
                    alert("Lỗi tải dữ liệu: " + (json.error || "Không thể nạp file"));
                }
            } catch (e) {
                console.error("Lỗi tải nội dung file:", e);
            }
        }

        function renderStats(data) {
            document.getElementById('stat-total').textContent = data.length;
            const withPhone = data.filter(d => d.phone && d.phone.trim()).length;
            document.getElementById('stat-phone').textContent = withPhone;
            const withZalo = data.filter(d => d.phone_info && d.phone_info.is_mobile).length;
            document.getElementById('stat-zalo').textContent = withZalo;
        }

        function renderResultsMap(data) {
            if (!resultsMap || !resultsLayerGroup) return;
            resultsLayerGroup.clearLayers();

            const bounds = [];
            data.forEach(item => {
                const lat = parseFloat(item.latitude);
                const lon = parseFloat(item.longitude);
                if (!isNaN(lat) && !isNaN(lon) && lat !== 0 && lon !== 0) {
                    bounds.push([lat, lon]);
                    
                    const pInfo = item.phone_info || {};
                    const phone = pInfo.raw || item.phone || '';
                    const cleanPhone = pInfo.clean || '';
                    const isMobile = pInfo.is_mobile;
                    const zaloUrl = pInfo.zalo_url;

                    let actionButtons = '';
                    if (cleanPhone) {
                        actionButtons += `
                            <button onclick="makeCallToPhone('${cleanPhone}')" class="px-2 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-bold flex items-center gap-1 shadow-sm">
                                📞 Gọi V9 PBX
                            </button>
                        `;
                    }
                    if (isMobile && zaloUrl) {
                        actionButtons += `
                            <a href="${zaloUrl}" target="_blank" class="px-2 py-1 zalo-badge text-white rounded text-[11px] font-bold flex items-center gap-1 shadow-sm">
                                💬 Zalo
                            </a>
                        `;
                    }

                    const popupHtml = `
                        <div class="text-xs max-w-xs leading-relaxed">
                            <h4 class="font-bold text-sm text-indigo-700">${escapeHtml(item.title || 'Địa điểm')}</h4>
                            <p class="text-slate-500 text-[11px]">${escapeHtml(item.category || '')}</p>
                            <p class="text-slate-700 mt-1">${escapeHtml(item.address || '')}</p>
                            ${phone ? `<p class="font-bold text-slate-900 mt-1">ĐT: ${escapeHtml(phone)}</p>` : ''}
                            <div class="flex items-center gap-2 mt-2 pt-2 border-t border-slate-200">
                                ${actionButtons}
                                ${item.link ? `<a href="${item.link}" target="_blank" class="text-indigo-600 hover:underline text-[11px] ml-auto">Xem Maps</a>` : ''}
                            </div>
                        </div>
                    `;

                    L.marker([lat, lon]).bindPopup(popupHtml).addTo(resultsLayerGroup);
                }
            });

            if (bounds.length > 0) {
                resultsMap.fitBounds(bounds, { padding: [30, 30] });
            }
        }

        function renderTable(data) {
            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';

            const showingInfo = document.getElementById('table-showing-info');
            showingInfo.textContent = `Đang hiển thị ${data.length} địa điểm`;

            if (data.length === 0) {
                tbody.innerHTML = `<tr><td colspan="9" class="p-6 text-center text-slate-400">Không tìm thấy địa điểm phù hợp</td></tr>`;
                return;
            }

            data.forEach((item, idx) => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

                const pInfo = item.phone_info || {};
                const rawPhone = pInfo.raw || item.phone || '';
                const cleanPhone = pInfo.clean || '';
                const isMobile = pInfo.is_mobile;
                const zaloUrl = pInfo.zalo_url;

                let contactCol = '<span class="text-slate-400 text-[11px]">Không có</span>';
                if (cleanPhone) {
                    contactCol = `
                        <div class="flex flex-col gap-1">
                            <span class="font-bold text-slate-900 text-xs">${escapeHtml(rawPhone)}</span>
                            <div class="flex items-center gap-1.5 flex-wrap">
                                <button onclick="makeCallToPhone('${cleanPhone}')" 
                                    class="px-2 py-0.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded font-bold text-[10px] flex items-center gap-1 shadow-sm"
                                    title="Gọi trực tiếp qua tổng đài V9 Tech">
                                    <i data-lucide="phone" class="w-3 h-3"></i> Gọi V9
                                </button>
                                ${isMobile ? `
                                    <a href="${zaloUrl}" target="_blank" 
                                        class="px-2 py-0.5 zalo-badge text-white rounded font-bold text-[10px] flex items-center gap-1 shadow-sm"
                                        title="Mở Zalo nhắn tin với số ${cleanPhone}">
                                        <span>Zalo</span>
                                    </a>
                                ` : ''}
                            </div>
                        </div>
                    `;
                }

                tr.innerHTML = `
                    <td class="p-3 text-slate-400 font-mono text-[11px]">${idx + 1}</td>
                    <td class="p-3 font-semibold text-indigo-950">${escapeHtml(item.title || '')}</td>
                    <td class="p-3 text-slate-600"><span class="bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[11px]">${escapeHtml(item.category || 'N/A')}</span></td>
                    <td class="p-3">${contactCol}</td>
                    <td class="p-3 text-slate-600 max-w-xs truncate" title="${escapeHtml(item.address || '')}">${escapeHtml(item.address || '')}</td>
                    <td class="p-3 text-center font-bold text-amber-600">${item.review_rating || '-'}</td>
                    <td class="p-3 text-center text-slate-500">${item.review_count || 0}</td>
                    <td class="p-3 text-slate-500">
                        ${item.website ? `<a href="${item.website}" target="_blank" class="text-indigo-600 hover:underline max-w-[120px] truncate block text-[11px]" title="${item.website}">${item.website.replace(/^https?:\\/\\//, '')}</a>` : '-'}
                    </td>
                    <td class="p-3 text-center">
                        ${item.link ? `<a href="${item.link}" target="_blank" class="text-indigo-600 hover:text-indigo-800 p-1 inline-block" title="Mở trên Google Maps"><i data-lucide="external-link" class="w-4 h-4"></i></a>` : '-'}
                    </td>
                `;
                tbody.appendChild(tr);
            });

            lucide.createIcons();
        }

        function filterTable() {
            const query = document.getElementById('table-search').value.toLowerCase().trim();
            const zaloOnly = document.getElementById('filter-zalo-only').checked;

            const filtered = currentData.filter(item => {
                if (zaloOnly && (!item.phone_info || !item.phone_info.is_mobile)) {
                    return false;
                }
                if (!query) return true;
                const title = (item.title || '').toLowerCase();
                const addr = (item.address || '').toLowerCase();
                const phone = (item.phone || '').toLowerCase();
                const cat = (item.category || '').toLowerCase();
                return title.includes(query) || addr.includes(query) || phone.includes(query) || cat.includes(query);
            });

            renderTable(filtered);
        }

        // ==================== TÍCH HỢP TỔNG ĐÀI V9 TECH ====================

        async function initV9Integration() {
            try {
                const res = await fetch(apiUrl('/api/v9tech/config'));
                v9Config = await res.json();
                updateConfigFormFields(v9Config);

                if (v9Config.mode === 'sdk' && v9Config.token && v9Config.domain && v9Config.client_id) {
                    connectV9WebRTC();
                } else if (v9Config.mode === 'api' && v9Config.extension) {
                    updateSipBadge('registered', `PBX API [${v9Config.extension}]`);
                } else {
                    updateSipBadge('not_configured', 'Cấu hình V9 Tech');
                }
            } catch (e) {
                console.error("Lỗi lấy config V9:", e);
                updateSipBadge('error', 'Lỗi kết nối V9');
            }
        }

        function updateConfigFormFields(cfg) {
            document.getElementById('v9-sdk-token').value = cfg.token || '';
            document.getElementById('v9-sdk-domain').value = cfg.domain || 'demo.v9.com.vn';
            document.getElementById('v9-sdk-client-id').value = cfg.client_id || 'v9-web-client';
            document.getElementById('v9-sdk-server').value = cfg.sip_server || cfg.domain || 'demo.v9.com.vn';
            document.getElementById('v9-sdk-port').value = cfg.sip_port || 8089;
            document.getElementById('v9-sdk-username').value = cfg.sip_username || cfg.extension || '101';
            document.getElementById('v9-sdk-password').value = cfg.sip_password || '';

            document.getElementById('v9-cfg-ext').value = cfg.extension || '101';
            document.getElementById('v9-cfg-url').value = cfg.api_url || 'https://api.v9tech.vn/api/v1/call/click-to-call';
            document.getElementById('v9-cfg-key').value = cfg.api_key || '';

            switchV9ModeTab(cfg.mode || 'sdk');
        }

        function switchV9ModeTab(mode) {
            currentV9Mode = mode;
            document.getElementById('panel-v9-sdk').classList.toggle('hidden', mode !== 'sdk');
            document.getElementById('panel-v9-api').classList.toggle('hidden', mode !== 'api');

            document.getElementById('tab-v9-sdk').className = mode === 'sdk'
                ? 'py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1.5'
                : 'py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1.5';
            
            document.getElementById('tab-v9-api').className = mode === 'api'
                ? 'py-1.5 rounded-md bg-white text-indigo-700 shadow-sm flex items-center justify-center gap-1.5'
                : 'py-1.5 rounded-md text-slate-600 hover:text-slate-900 flex items-center justify-center gap-1.5';
        }

        function updateSipBadge(state, text) {
            const badge = document.getElementById('v9-sip-badge');
            const dot = document.getElementById('v9-sip-dot');
            const txt = document.getElementById('v9-sip-text');
            txt.textContent = text;

            if (state === 'registered') {
                dot.className = 'w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse';
                badge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-lg bg-emerald-950/80 hover:bg-emerald-900 text-xs font-semibold cursor-pointer border border-emerald-500/50 shadow-inner text-emerald-200';
            } else if (state === 'connecting') {
                dot.className = 'w-2.5 h-2.5 rounded-full bg-amber-400 animate-ping';
                badge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-lg bg-amber-950/80 hover:bg-amber-900 text-xs font-semibold cursor-pointer border border-amber-500/50 shadow-inner text-amber-200';
            } else if (state === 'in_call') {
                dot.className = 'w-2.5 h-2.5 rounded-full bg-rose-500 animate-ping';
                badge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-lg bg-rose-950/80 hover:bg-rose-900 text-xs font-bold cursor-pointer border border-rose-500/50 shadow-inner text-rose-200';
            } else {
                dot.className = 'w-2.5 h-2.5 rounded-full bg-slate-400';
                badge.className = 'flex items-center gap-2 px-3 py-1.5 rounded-lg bg-indigo-900/80 hover:bg-indigo-900 text-xs font-medium cursor-pointer border border-indigo-600/50 shadow-inner text-slate-300';
            }
        }

        function connectV9WebRTC() {
            if (!window.VoiceSDKDialer) {
                console.warn("VoiceSDKDialer chưa có trong window.");
                updateSipBadge('not_configured', 'Cần nạp SDK V9');
                return;
            }

            try {
                if (v9Dialer && typeof v9Dialer.destroy === 'function') {
                    try { v9Dialer.destroy(); } catch (e) {}
                }

                updateSipBadge('connecting', 'Đang kết nối SIP...');

                v9Dialer = new window.VoiceSDKDialer({
                    token: v9Config.token,
                    domain: v9Config.domain,
                    clientId: v9Config.client_id
                });

                v9Dialer.start({
                    containerId: 'sip-dialer-container',
                    defaultCountryCode: '+84',
                    enableUI: false,
                    enableDialerLauncher: false,
                    sipConfig: {
                        server: v9Config.sip_server || v9Config.domain,
                        port: parseInt(v9Config.sip_port) || 8089,
                        username: v9Config.sip_username || v9Config.extension || '101',
                        password: v9Config.sip_password || '',
                        displayName: v9Config.sip_display_name || 'NV Kinh Doanh'
                    },
                    callbacks: {
                        onRegistered: function(data, dialer) {
                            console.log('SIP Registered thành công:', data);
                            updateSipBadge('registered', `Sẵn sàng [${v9Config.sip_username || '101'}]`);
                        },
                        onUnregistered: function(data, dialer) {
                            console.log('SIP Unregistered:', data);
                            updateSipBadge('unregistered', 'Chưa đăng ký SIP');
                        },
                        onServerConnect: function(data, dialer) {
                            console.log('SIP Server Connect:', data);
                            updateSipBadge('connecting', 'Đã nối máy chủ SIP');
                        },
                        onServerDisconnect: function(data, dialer) {
                            console.log('SIP Server Disconnect:', data);
                            updateSipBadge('disconnected', 'Mất kết nối SIP');
                        },
                        onOutgoingCall: function(data, dialer) {
                            console.log('Outgoing call:', data);
                            openCallWidget(currentActivePhone || 'Đang gọi...', 'Đang đổ chuông...');
                            updateSipBadge('in_call', `Đang gọi ${currentActivePhone}`);
                        },
                        onCallAnswered: function(data, dialer) {
                            console.log('Call Answered:', data);
                            updateCallWidgetStatus('Đang đàm thoại (WebRTC)');
                            startCallTimer();
                            updateSipBadge('in_call', `Đàm thoại: ${currentActivePhone}`);
                        },
                        onCallEnded: function(data, dialer) {
                            console.log('Call Ended:', data);
                            updateCallWidgetStatus('Cuộc gọi đã kết thúc');
                            stopCallTimer();
                            updateSipBadge('registered', `Sẵn sàng [${v9Config.sip_username || '101'}]`);
                            setTimeout(() => closeCallWidget(), 2500);
                        },
                        onCallDeclined: function(data, dialer) {
                            console.log('Call Declined:', data);
                            updateCallWidgetStatus('Khách hàng từ chối cuộc gọi');
                            stopCallTimer();
                            updateSipBadge('registered', `Sẵn sàng [${v9Config.sip_username || '101'}]`);
                            setTimeout(() => closeCallWidget(), 3000);
                        },
                        onCallRejected: function(data, dialer) {
                            console.log('Call Rejected:', data);
                            updateCallWidgetStatus('Máy bận / Từ chối');
                            stopCallTimer();
                            updateSipBadge('registered', `Sẵn sàng [${v9Config.sip_username || '101'}]`);
                            setTimeout(() => closeCallWidget(), 3000);
                        },
                        onStatusChange: function(status, dialer) {
                            console.log('SIP Status Change:', status);
                        },
                        onError: function(err) {
                            console.error('SIP Error:', err);
                            showToast('Lỗi V9 Softphone', err || 'Lỗi kết nối');
                        }
                    }
                });
            } catch (err) {
                console.error("Lỗi khởi tạo VoiceSDKDialer:", err);
                updateSipBadge('error', 'Lỗi SIP Softphone');
            }
        }

        function makeCallToPhone(phone) {
            const clean = phone.replace(/[^0-9]/g, '');
            if (!clean) {
                alert("Số điện thoại không hợp lệ!");
                return;
            }

            currentActivePhone = clean;

            if (v9Config.mode === 'sdk' && v9Config.token) {
                if (!v9Dialer) {
                    connectV9WebRTC();
                }
                openCallWidget(clean, 'Đang thiết lập cuộc gọi...');
                try {
                    v9Dialer.makeCall(clean);
                } catch (err) {
                    console.error("Make call error:", err);
                    showToast('Lỗi gọi WebRTC', err.message || err);
                }
            } else if (v9Config.mode === 'api' || v9Config.extension) {
                callViaPbxApi(clean);
            } else {
                openV9Modal();
            }
        }

        async function callViaPbxApi(phone) {
            showToast('Đang gọi PBX V9 Tech', `Đang yêu cầu kết nối tới ${phone}...`);
            try {
                const res = await fetch(apiUrl('/api/v9tech/call'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ phone: phone })
                });
                const data = await res.json();
                if (data.success) {
                    showToast('Đang đổ chuông máy nhánh!', data.message);
                } else if (data.fallback) {
                    showToast('Chuyển sang ứng dụng Softphone', `Mở quay số tới ${phone}...`);
                    window.location.href = `tel:${phone}`;
                } else {
                    showToast('Thông báo V9 Tech', data.message);
                }
            } catch (e) {
                window.location.href = `tel:${phone}`;
            }
        }

        // ==================== CALL WIDGET CONTROLS ====================
        function openCallWidget(phone, status) {
            currentActivePhone = phone;
            document.getElementById('call-phone-number').textContent = phone;
            document.getElementById('call-status-label').textContent = status;
            document.getElementById('call-timer').textContent = '00:00';
            document.getElementById('v9-call-widget').classList.remove('hidden');
            lucide.createIcons();
        }

        function updateCallWidgetStatus(status) {
            document.getElementById('call-status-label').textContent = status;
        }

        function closeCallWidget() {
            document.getElementById('v9-call-widget').classList.add('hidden');
            stopCallTimer();
        }

        function startCallTimer() {
            stopCallTimer();
            callSeconds = 0;
            callTimerInterval = setInterval(() => {
                callSeconds++;
                const mins = String(Math.floor(callSeconds / 60)).padStart(2, '0');
                const secs = String(callSeconds % 60).padStart(2, '0');
                document.getElementById('call-timer').textContent = `${mins}:${secs}`;
            }, 1000);
        }

        function stopCallTimer() {
            if (callTimerInterval) {
                clearInterval(callTimerInterval);
                callTimerInterval = null;
            }
        }

        function hangupCurrentCall() {
            if (v9Dialer && typeof v9Dialer.hangupCall === 'function') {
                try { v9Dialer.hangupCall(); } catch (e) {}
            }
            updateCallWidgetStatus('Đã cúp máy');
            stopCallTimer();
            setTimeout(() => closeCallWidget(), 1000);
        }

        function toggleCallMute() {
            if (!v9Dialer) return;
            isMuted = !isMuted;
            try {
                if (isMuted) {
                    v9Dialer.mute();
                    document.getElementById('btn-call-mute').className = 'p-3 rounded-full bg-amber-500 text-white transition';
                } else {
                    v9Dialer.unmute();
                    document.getElementById('btn-call-mute').className = 'p-3 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 transition';
                }
            } catch (e) {
                console.error("Mute error:", e);
            }
        }

        function toggleCallHold() {
            if (!v9Dialer) return;
            isHeld = !isHeld;
            try {
                if (isHeld) {
                    v9Dialer.hold();
                    document.getElementById('btn-call-hold').className = 'p-3 rounded-full bg-amber-500 text-white transition';
                    updateCallWidgetStatus('Đang giữ cuộc gọi (Hold)');
                } else {
                    v9Dialer.unhold();
                    document.getElementById('btn-call-hold').className = 'p-3 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 transition';
                    updateCallWidgetStatus('Đang đàm thoại');
                }
            } catch (e) {
                console.error("Hold error:", e);
            }
        }

        function showToast(title, message) {
            const toast = document.getElementById('v9-toast');
            document.getElementById('toast-title').textContent = title;
            document.getElementById('toast-message').textContent = message;
            toast.classList.remove('hidden');
            lucide.createIcons();
            setTimeout(() => toast.classList.add('hidden'), 6000);
        }

        // ==================== CẤU HÌNH V9 MODAL ====================
        function openV9Modal() {
            document.getElementById('v9-modal').classList.remove('hidden');
            lucide.createIcons();
        }

        function closeV9Modal() {
            document.getElementById('v9-modal').classList.add('hidden');
        }

        async function saveV9ConfigAndConnect() {
            const payload = {
                mode: currentV9Mode,
                token: document.getElementById('v9-sdk-token').value.trim(),
                domain: document.getElementById('v9-sdk-domain').value.trim(),
                client_id: document.getElementById('v9-sdk-client-id').value.trim(),
                sip_server: document.getElementById('v9-sdk-server').value.trim(),
                sip_port: parseInt(document.getElementById('v9-sdk-port').value) || 8089,
                sip_username: document.getElementById('v9-sdk-username').value.trim(),
                sip_password: document.getElementById('v9-sdk-password').value.trim(),
                extension: document.getElementById('v9-cfg-ext').value.trim() || document.getElementById('v9-sdk-username').value.trim(),
                api_url: document.getElementById('v9-cfg-url').value.trim(),
                api_key: document.getElementById('v9-cfg-key').value.trim()
            };

            try {
                const res = await fetch(apiUrl('/api/v9tech/config'), {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(payload)
                });
                const data = await res.json();
                if (res.ok) {
                    v9Config = data.config;
                    alert("Đã lưu cấu hình tổng đài V9 Tech thành công!");
                    closeV9Modal();
                    if (currentV9Mode === 'sdk') {
                        connectV9WebRTC();
                    } else {
                        updateSipBadge('registered', `PBX API [${payload.extension}]`);
                    }
                } else {
                    alert("Lỗi lưu cấu hình: " + (data.error || ""));
                }
            } catch (e) {
                alert("Lỗi lưu cấu hình: " + e.message);
            }
        }

        // ==================== TẢI FILE & SỬA LỖI FONT ====================
        function downloadActiveFile(format) {
            const filename = document.getElementById('file-select').value;
            if (!filename) {
                alert("Vui lòng chọn một tệp dữ liệu!");
                return;
            }
            window.location.href = apiUrl(`/api/download?filename=${encodeURIComponent(filename)}&format=${format}`);
        }

        async function fixAllFilesVietnamese() {
            if (!confirm("Hệ thống sẽ tự động chuẩn hóa toàn bộ file CSV sang chuẩn UTF-8 BOM và tạo thêm bản Excel .xlsx để khi mở trực tiếp bằng Excel trên Windows hiển thị đúng 100% tiếng Việt có dấu.\n\nBạn có muốn tiếp tục?")) {
                return;
            }
            try {
                const res = await fetch(apiUrl('/api/fix-all-files'), { method: 'POST' });
                const data = await res.json();
                if (data.success) {
                    alert(`Đã chuẩn hóa thành công ${data.fixed_count} file sang chuẩn UTF-8 BOM và Excel .xlsx!\n\nBây giờ bạn có thể nhấp đúp mở file trực tiếp bằng Excel không còn bị lỗi font nữa.`);
                    refreshFileList();
                } else {
                    alert("Có lỗi khi chuẩn hóa file.");
                }
            } catch (e) {
                alert("Lỗi kết nối máy chủ: " + e.message);
            }
        }

        function escapeHtml(str) {
            if (!str) return '';
            const div = document.createElement('div');
            div.textContent = str;
            return div.innerHTML;
        }

        window.addEventListener('DOMContentLoaded', () => {
            initPickerMap();
            initV9Integration();
            lucide.createIcons();
        });
    </script>
</body>
</html>
"""
    return html

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"=== KHỞI ĐỘNG MÁY CHỦ GOOGLE MAPS PRO (VIETNAM + V9 TECH + ZALO) ===")
    print(f"URL: http://127.0.0.1:{port}")
    uvicorn.run("viewer_server:app", host="0.0.0.0", port=port, reload=False)
