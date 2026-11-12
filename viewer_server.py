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
import db_manager

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
TEMPLATES_DIR = WORKSPACE_DIR / "templates"
CONFIG_FILE = WORKSPACE_DIR / "v9tech_config.json"
INDEX_HTML_PATH = TEMPLATES_DIR / "index.html"

RESULTS_DIR.mkdir(parents=True, exist_ok=True)
STATIC_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

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
            try:
                kw = current_scrape_job.get("query_keywords", "")
                db_res = db_manager.import_csv_to_db(raw_csv_path, keyword=kw, filename=raw_csv_path.name)
                if db_res.get("success"):
                    current_scrape_job["logs"].append(
                        f">>> CSDL: Đã lưu vào CSDL & Lọc trùng thành công! Tổng: {db_res['total']} địa điểm "
                        f"(Mới: {db_res['new']}, Trùng lặp đã gộp & cập nhật: {db_res['duplicates']})"
                    )
            except Exception as dbe:
                current_scrape_job["logs"].append(f">>> CSDL: Cảnh báo lưu CSDL: {str(dbe)}")
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
    current_scrape_job["query_keywords"] = ", ".join(query_lines[:3])
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

# ================= KHO CSDL & LỌC TRÙNG LẶP =================
@app.get("/api/database/stats")
def get_database_stats():
    """Lấy thống kê tổng kho CSDL, lọc trùng và phân loại số điện thoại"""
    try:
        stats = db_manager.get_db_stats()
        return {"success": True, "stats": stats}
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

@app.get("/api/database/places")
def get_database_places(
    q: str = Query("", description="Từ khóa tìm kiếm"),
    filter_type: str = Query("all", alias="filter", description="all, mobile, email, website, duplicates"),
    page: int = Query(1, ge=1),
    page_size: int = Query(25, ge=1, le=200),
    sort_by: str = Query("last_scanned_at"),
    sort_dir: str = Query("DESC")
):
    """Lấy danh sách địa điểm trong kho CSDL có phân trang và lọc trùng"""
    try:
        res = db_manager.get_places(
            q=q,
            filter_type=filter_type,
            page=page,
            page_size=page_size,
            sort_by=sort_by,
            sort_dir=sort_dir
        )
        return {"success": True, **res}
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

@app.get("/api/database/export")
def export_database(
    filter_type: str = Query("all", alias="filter"),
    format: str = Query("excel")
):
    """Xuất file Excel hoặc CSV kho dữ liệu sạch không trùng lặp"""
    try:
        out_file = db_manager.export_places_to_file(filter_type=filter_type, export_format=format)
        if not out_file.exists():
            return JSONResponse(status_code=500, content={"error": "Không tạo được file xuất"})
        
        media_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if format == "excel" else "text/csv"
        )
        return FileResponse(out_file, filename=out_file.name, media_type=media_type)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.post("/api/database/sync-all")
def sync_all_database_files():
    """Đồng bộ toàn bộ các file CSV đã cào vào CSDL và tự động loại bỏ / gộp trùng lặp"""
    try:
        res = db_manager.sync_all_outputs(RESULTS_DIR)
        return res
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

@app.delete("/api/database/places/{place_id}")
def delete_database_place(place_id: int):
    """Xóa một địa điểm khỏi kho CSDL"""
    try:
        db_manager.delete_place(place_id)
        return {"success": True}
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})

# Giao diện HTML Tiếng Việt Chuyên Nghiệp
@app.get("/", response_class=HTMLResponse)
def index_html():
    if INDEX_HTML_PATH.exists():
        with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
            return f.read()
    return HTMLResponse("<h1>Chưa tìm thấy templates/index.html</h1>", status_code=500)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"=== KHỞI ĐỘNG MÁY CHỦ GOOGLE MAPS PRO (VIETNAM + V9 TECH + ZALO) ===")
    print(f"URL: http://127.0.0.1:{port}")
    uvicorn.run("viewer_server:app", host="0.0.0.0", port=port, reload=False)
