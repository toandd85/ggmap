import os
import re
import sqlite3
import datetime
import pandas as pd
from pathlib import Path
from typing import Optional, List, Dict, Any

WORKSPACE_DIR = Path(os.environ.get("WORKSPACE_DIR", Path(__file__).parent.resolve())).resolve()
DB_PATH = WORKSPACE_DIR / "leads.db"

def clean_phone(raw_phone: str) -> dict:
    if not raw_phone or pd.isna(raw_phone):
        return {"raw": "", "clean": "", "is_mobile": False, "zalo_url": "", "type": "none"}
    digits = "".join(c for c in str(raw_phone) if c.isdigit())
    if digits.startswith("84") and len(digits) >= 11:
        digits = "0" + digits[2:]
    elif not digits.startswith("0") and len(digits) == 9:
        digits = "0" + digits

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

def get_connection(db_file: Optional[Path] = None) -> sqlite3.Connection:
    target_path = db_file or DB_PATH
    conn = sqlite3.connect(str(target_path), timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn

def init_db(db_file: Optional[Path] = None):
    conn = get_connection(db_file)
    with conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS places (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            place_id TEXT UNIQUE,
            cid TEXT,
            data_id TEXT,
            title TEXT NOT NULL,
            category TEXT,
            phone TEXT,
            phone_clean TEXT,
            is_mobile INTEGER DEFAULT 0,
            zalo_link TEXT,
            address TEXT,
            complete_address TEXT,
            website TEXT,
            emails TEXT,
            review_count INTEGER DEFAULT 0,
            review_rating REAL DEFAULT 0,
            link TEXT,
            latitude REAL,
            longitude REAL,
            scan_count INTEGER DEFAULT 1,
            first_scanned_at DATETIME,
            last_scanned_at DATETIME,
            last_keyword TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_places_place_id ON places(place_id);
        CREATE INDEX IF NOT EXISTS idx_places_cid ON places(cid);
        CREATE INDEX IF NOT EXISTS idx_places_phone_clean ON places(phone_clean);
        CREATE INDEX IF NOT EXISTS idx_places_title ON places(title);
        CREATE INDEX IF NOT EXISTS idx_places_category ON places(category);
        CREATE INDEX IF NOT EXISTS idx_places_is_mobile ON places(is_mobile);

        CREATE TABLE IF NOT EXISTS scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_time DATETIME,
            keyword TEXT,
            filename TEXT,
            total_found INTEGER DEFAULT 0,
            new_places INTEGER DEFAULT 0,
            duplicate_places INTEGER DEFAULT 0
        );
        CREATE INDEX IF NOT EXISTS idx_scans_time ON scans(scan_time);
        """)
    conn.close()

def import_csv_to_db(csv_path: Path, keyword: str = "", filename: str = "", db_file: Optional[Path] = None) -> Dict[str, Any]:
    if not csv_path.exists():
        return {"total": 0, "new": 0, "duplicates": 0, "error": "File không tồn tại"}

    init_db(db_file)
    conn = get_connection(db_file)

    try:
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_path, encoding='utf-8-sig', errors='replace')
    except Exception as e:
        conn.close()
        return {"total": 0, "new": 0, "duplicates": 0, "error": f"Lỗi đọc CSV: {str(e)}"}

    df = df.fillna("")
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    target_filename = filename or csv_path.name

    total_records = len(df)
    new_count = 0
    duplicate_count = 0

    cursor = conn.cursor()

    for _, row in df.iterrows():
        title = str(row.get("title", "")).strip()
        if not title:
            continue

        place_id = str(row.get("place_id", "")).strip()
        cid = str(row.get("cid", "")).strip()
        data_id = str(row.get("data_id", "")).strip()
        category = str(row.get("category", "")).strip()
        raw_phone = str(row.get("phone", "")).strip()
        address = str(row.get("address", "")).strip()
        complete_address = str(row.get("complete_address", "")).strip() or address
        website = str(row.get("website", "")).strip()
        emails = str(row.get("emails", "")).strip()
        link = str(row.get("link", "")).strip()

        try:
            review_count = int(row.get("review_count", 0) or 0)
        except (ValueError, TypeError):
            review_count = 0

        try:
            review_rating = float(row.get("review_rating", 0) or 0)
        except (ValueError, TypeError):
            review_rating = 0.0

        try:
            latitude = float(row.get("latitude", 0) or 0)
            longitude = float(row.get("longitude", 0) or 0)
        except (ValueError, TypeError):
            latitude = 0.0
            longitude = 0.0

        phone_info = clean_phone(raw_phone)
        phone_clean = phone_info["clean"]
        is_mobile = 1 if phone_info["is_mobile"] else 0
        zalo_link = phone_info["zalo_url"]

        # 1. Deduplication checking logic
        existing = None
        if place_id:
            cursor.execute("SELECT id, scan_count, phone, phone_clean, website, emails, review_count, zalo_link FROM places WHERE place_id = ?", (place_id,))
            existing = cursor.fetchone()
        
        if not existing and cid:
            cursor.execute("SELECT id, scan_count, phone, phone_clean, website, emails, review_count, zalo_link FROM places WHERE cid = ?", (cid,))
            existing = cursor.fetchone()

        if not existing and phone_clean:
            cursor.execute("SELECT id, scan_count, phone, phone_clean, website, emails, review_count, zalo_link FROM places WHERE phone_clean = ?", (phone_clean,))
            existing = cursor.fetchone()

        if not existing and len(title) >= 4 and address:
            cursor.execute("SELECT id, scan_count, phone, phone_clean, website, emails, review_count, zalo_link FROM places WHERE title = ? AND address = ?", (title, address))
            existing = cursor.fetchone()

        if existing:
            # Duplicate found -> update fresh data and increment scan_count
            duplicate_count += 1
            existing_id = existing["id"]
            new_scan_count = (existing["scan_count"] or 1) + 1

            # Keep the best/latest phone, website, emails
            final_phone = raw_phone if raw_phone else existing["phone"]
            final_phone_clean = phone_clean if phone_clean else existing["phone_clean"]
            final_zalo_link = zalo_link if zalo_link else existing["zalo_link"]
            final_website = website if website else existing["website"]
            final_emails = emails if emails else existing["emails"]
            final_reviews = max(review_count, existing["review_count"] or 0)

            cursor.execute("""
                UPDATE places SET 
                    scan_count = ?,
                    last_scanned_at = ?,
                    last_keyword = ?,
                    phone = ?,
                    phone_clean = ?,
                    zalo_link = ?,
                    website = ?,
                    emails = ?,
                    review_count = ?,
                    review_rating = CASE WHEN ? > 0 THEN ? ELSE review_rating END,
                    category = CASE WHEN category = '' THEN ? ELSE category END,
                    address = CASE WHEN address = '' THEN ? ELSE address END
                WHERE id = ?
            """, (
                new_scan_count, now_str, keyword or target_filename,
                final_phone, final_phone_clean, final_zalo_link,
                final_website, final_emails, final_reviews,
                review_rating, review_rating,
                category, address,
                existing_id
            ))
        else:
            # New record -> insert into DB
            new_count += 1
            # Fallback place_id if blank to ensure uniqueness if possible
            gen_place_id = place_id if place_id else (f"gen_{cid}" if cid else None)
            cursor.execute("""
                INSERT INTO places (
                    place_id, cid, data_id, title, category,
                    phone, phone_clean, is_mobile, zalo_link,
                    address, complete_address, website, emails,
                    review_count, review_rating, link,
                    latitude, longitude, scan_count,
                    first_scanned_at, last_scanned_at, last_keyword
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
            """, (
                gen_place_id, cid, data_id, title, category,
                raw_phone, phone_clean, is_mobile, zalo_link,
                address, complete_address, website, emails,
                review_count, review_rating, link,
                latitude, longitude,
                now_str, now_str, keyword or target_filename
            ))

    # Log to scans table
    cursor.execute("""
        INSERT INTO scans (scan_time, keyword, filename, total_found, new_places, duplicate_places)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (now_str, keyword or target_filename, target_filename, total_records, new_count, duplicate_count))

    conn.commit()
    conn.close()

    return {
        "success": True,
        "total": total_records,
        "new": new_count,
        "duplicates": duplicate_count,
        "filename": target_filename
    }

def sync_all_outputs(results_dir: Optional[Path] = None, db_file: Optional[Path] = None) -> Dict[str, Any]:
    target_dir = results_dir or (WORKSPACE_DIR / "outputs")
    init_db(db_file)
    
    csv_files = []
    for p in target_dir.glob("*.csv"):
        if p.name.endswith("_excel.csv"):
            continue
        csv_files.append(p)

    total_files = len(csv_files)
    total_new = 0
    total_duplicates = 0
    total_records = 0

    for csv_file in csv_files:
        # Extract keyword from filename e.g. "quà_tặng_20261112_143439.csv"
        kw = csv_file.stem.rsplit("_", 2)[0].replace("_", " ") if "_" in csv_file.stem else csv_file.stem
        res = import_csv_to_db(csv_file, keyword=kw, filename=csv_file.name, db_file=db_file)
        if res.get("success"):
            total_records += res.get("total", 0)
            total_new += res.get("new", 0)
            total_duplicates += res.get("duplicates", 0)

    return {
        "success": True,
        "total_files": total_files,
        "total_records": total_records,
        "total_new": total_new,
        "total_duplicates": total_duplicates
    }

def get_db_stats(db_file: Optional[Path] = None) -> Dict[str, Any]:
    init_db(db_file)
    conn = get_connection(db_file)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM places")
    total_places = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM places WHERE is_mobile = 1")
    mobile_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM places WHERE emails != '' AND emails IS NOT NULL")
    email_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM places WHERE website != '' AND website IS NOT NULL")
    website_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM places WHERE scan_count > 1")
    duplicate_places_count = cursor.fetchone()[0]

    cursor.execute("SELECT COALESCE(SUM(scan_count - 1), 0) FROM places WHERE scan_count > 1")
    duplicates_prevented = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM scans")
    total_scans = cursor.fetchone()[0]

    conn.close()

    return {
        "total_places": total_places,
        "mobile_count": mobile_count,
        "email_count": email_count,
        "website_count": website_count,
        "duplicate_places_count": duplicate_places_count,
        "duplicates_prevented": duplicates_prevented,
        "total_scans": total_scans
    }

def get_places(
    q: str = "",
    filter_type: str = "all", # all, mobile, email, website, duplicates
    page: int = 1,
    page_size: int = 25,
    sort_by: str = "last_scanned_at",
    sort_dir: str = "DESC",
    db_file: Optional[Path] = None
) -> Dict[str, Any]:
    init_db(db_file)
    conn = get_connection(db_file)
    cursor = conn.cursor()

    where_clauses = []
    params = []

    if q:
        query_pattern = f"%{q.strip()}%"
        where_clauses.append("(title LIKE ? OR phone LIKE ? OR phone_clean LIKE ? OR address LIKE ? OR category LIKE ? OR last_keyword LIKE ?)")
        params.extend([query_pattern] * 6)

    if filter_type == "mobile":
        where_clauses.append("is_mobile = 1")
    elif filter_type == "email":
        where_clauses.append("emails != '' AND emails IS NOT NULL")
    elif filter_type == "website":
        where_clauses.append("website != '' AND website IS NOT NULL")
    elif filter_type == "duplicates":
        where_clauses.append("scan_count > 1")

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

    # Count total
    count_sql = f"SELECT COUNT(*) FROM places {where_sql}"
    cursor.execute(count_sql, params)
    total = cursor.fetchone()[0]

    # Valid sort fields
    allowed_sorts = {
        "id": "id",
        "title": "title",
        "review_rating": "review_rating",
        "review_count": "review_count",
        "scan_count": "scan_count",
        "last_scanned_at": "last_scanned_at",
        "first_scanned_at": "first_scanned_at"
    }
    safe_sort = allowed_sorts.get(sort_by, "last_scanned_at")
    safe_dir = "ASC" if sort_dir.upper() == "ASC" else "DESC"

    offset = max(0, (page - 1) * page_size)
    query_sql = f"""
        SELECT 
            id, place_id, cid, title, category,
            phone, phone_clean, is_mobile, zalo_link,
            address, website, emails, review_count, review_rating,
            link, scan_count, first_scanned_at, last_scanned_at, last_keyword
        FROM places 
        {where_sql}
        ORDER BY {safe_sort} {safe_dir}
        LIMIT ? OFFSET ?
    """
    cursor.execute(query_sql, params + [page_size, offset])
    rows = cursor.fetchall()

    places = []
    for r in rows:
        d = dict(r)
        d["phone_info"] = {
            "raw": d.get("phone", ""),
            "clean": d.get("phone_clean", ""),
            "is_mobile": bool(d.get("is_mobile")),
            "zalo_url": d.get("zalo_link", ""),
            "type": "mobile" if d.get("is_mobile") else ("landline" if d.get("phone_clean") else "none")
        }
        places.append(d)

    conn.close()

    total_pages = (total + page_size - 1) // page_size if total > 0 else 1

    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "places": places
    }

def delete_place(place_id: int, db_file: Optional[Path] = None) -> bool:
    init_db(db_file)
    conn = get_connection(db_file)
    with conn:
        conn.execute("DELETE FROM places WHERE id = ?", (place_id,))
    conn.close()
    return True

def export_places_to_file(filter_type: str = "all", export_format: str = "excel", db_file: Optional[Path] = None) -> Path:
    init_db(db_file)
    conn = get_connection(db_file)

    where_sql = ""
    if filter_type == "mobile":
        where_sql = "WHERE is_mobile = 1"
    elif filter_type == "email":
        where_sql = "WHERE emails != '' AND emails IS NOT NULL"
    elif filter_type == "website":
        where_sql = "WHERE website != '' AND website IS NOT NULL"
    elif filter_type == "duplicates":
        where_sql = "WHERE scan_count > 1"

    query_sql = f"""
        SELECT 
            id as ID,
            title as [Tên địa điểm],
            category as [Danh mục],
            phone as [Số điện thoại gốc],
            phone_clean as [SĐT Chuẩn],
            CASE WHEN is_mobile = 1 THEN 'Di động' ELSE 'Cố định' END as [Loại SĐT],
            zalo_link as [Link Zalo],
            emails as [Email],
            website as [Website],
            address as [Địa chỉ],
            review_count as [Số đánh giá],
            review_rating as [Điểm sao],
            scan_count as [Số lần quét (Gộp trùng)],
            last_keyword as [Từ khóa quét gần nhất],
            first_scanned_at as [Ngày thêm đầu tiên],
            last_scanned_at as [Cập nhật gần nhất],
            link as [Link Google Maps]
        FROM places
        {where_sql}
        ORDER BY last_scanned_at DESC
    """
    df = pd.read_sql_query(query_sql, conn)
    conn.close()

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    export_dir = WORKSPACE_DIR / "outputs"
    export_dir.mkdir(parents=True, exist_ok=True)

    if export_format == "excel":
        file_path = export_dir / f"CSDL_GoogleMaps_Master_{timestamp}.xlsx"
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="KhachHang_GoogleMaps")
        return file_path
    else:
        file_path = export_dir / f"CSDL_GoogleMaps_Master_{timestamp}.csv"
        df.to_csv(file_path, index=False, encoding="utf-8-sig")
        return file_path
