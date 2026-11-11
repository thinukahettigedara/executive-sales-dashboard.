import http.server
import io
import json
import os
import socketserver
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path
import pandas as pd

PORT = 8000
DIRECTORY = Path(__file__).parent.resolve()
DOWNLOADS_DIR = DIRECTORY.parent

COLUMNS = ["Date", "Order_ID", "Region", "Category", "Product", "Sales_Rep",
           "Segment", "Units", "Unit_Price", "Revenue", "Cost", "Profit"]
NUMERIC = ["Units", "Unit_Price", "Revenue", "Cost", "Profit"]
TEXT = ["Order_ID", "Region", "Category", "Product", "Sales_Rep", "Segment"]

cached_data = {}
cached_mtime = {}


def normalize_col_name(col):
    c = str(col).strip().replace(" ", "_").replace("%", "").strip("_")
    mapping = {
        "Order_ID": "Order_ID",
        "OrderID": "Order_ID",
        "Order_Id": "Order_ID",
        "Sales_Rep": "Sales_Rep",
        "SalesRep": "Sales_Rep",
        "Unit_Price": "Unit_Price",
        "UnitPrice": "Unit_Price",
        "Margin": "Margin",
    }
    return mapping.get(c, c)


def get_available_data_files():
    """Find all relevant .xlsx and .csv files in workspace and downloads directory."""
    files = []
    seen_paths = set()

    search_dirs = [DIRECTORY]
    if DOWNLOADS_DIR.exists():
        search_dirs.append(DOWNLOADS_DIR)

    for search_dir in search_dirs:
        for ext in ("*.xlsx", "*.csv"):
            for f in search_dir.glob(ext):
                if f.name.startswith("~$") or f.resolve() in seen_paths:
                    continue
                seen_paths.add(f.resolve())
                try:
                    mtime = os.path.getmtime(f)
                    files.append({
                        "name": f.name,
                        "path": f,
                        "folder": "files" if search_dir == DIRECTORY else "downloads",
                        "mtime": mtime,
                        "size": os.path.getsize(f),
                        "last_updated": datetime.fromtimestamp(mtime).strftime("%b %d, %Y %I:%M:%S %p")
                    })
                except OSError:
                    pass

    # Sort with newest modification time first
    files.sort(key=lambda x: x["mtime"], reverse=True)
    return files


def resolve_target_file(requested_file=None):
    available = get_available_data_files()
    if not available:
        return None, []

    if requested_file and requested_file != "AUTO":
        for f in available:
            if f["name"] == requested_file:
                return f, available

    # Default to newest
    return available[0], available


def parse_dataframe_to_records(df, filename, sheet_name=None):
    df.columns = [normalize_col_name(c) for c in df.columns]
    df = df.dropna(how="all")

    # Match required columns
    if "Revenue" in df.columns and "Cost" in df.columns and "Profit" not in df.columns:
        df["Profit"] = pd.to_numeric(df["Revenue"], errors="coerce") - pd.to_numeric(df["Cost"], errors="coerce")

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
    else:
        return None, f"Missing 'Date' column in {filename}"

    for col in NUMERIC:
        if col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    for col in TEXT:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = ""

    df = df.sort_values("Date")

    records = []
    for _, row in df.iterrows():
        rev = float(row["Revenue"])
        cost = float(row["Cost"])
        profit = float(row["Profit"])
        margin = round((profit / rev * 100) if rev > 0 else 0, 1)

        date_val = row["Date"]
        date_str = date_val.strftime("%Y-%m-%d") if hasattr(date_val, "strftime") else str(date_val)

        records.append({
            "date": date_str,
            "order_id": str(row["Order_ID"]),
            "region": str(row["Region"]),
            "category": str(row["Category"]),
            "product": str(row["Product"]),
            "sales_rep": str(row["Sales_Rep"]),
            "segment": str(row["Segment"]),
            "units": int(row["Units"]),
            "unit_price": round(float(row["Unit_Price"]), 2),
            "revenue": round(rev, 2),
            "cost": round(cost, 2),
            "profit": round(profit, 2),
            "margin": margin
        })

    return records, None


def safe_read_data_file(file_info):
    file_path = file_info["path"]
    filename = file_info["name"]
    current_mtime = file_info["mtime"]

    # Safe byte read with retry
    content = None
    for attempt in range(5):
        try:
            with open(file_path, "rb") as f:
                content = f.read()
            break
        except (PermissionError, OSError):
            time.sleep(0.2)

    if content is None:
        return {"status": "error", "error": f"Could not read {filename} (file may be open in Excel - please save it first)", "records": []}

    bio = io.BytesIO(content)
    sheet_name = None

    try:
        if filename.lower().endswith(".csv"):
            df = pd.read_csv(bio)
            sheet_name = "CSV"
        else:
            xl = pd.ExcelFile(bio)
            # Find the best sheet
            for candidate in ["Data", "Sales_Data", "sales_data", "Sheet1"]:
                if candidate in xl.sheet_names:
                    sheet_name = candidate
                    break
            if not sheet_name:
                sheet_name = xl.sheet_names[-1] if "Dashboard" in xl.sheet_names else xl.sheet_names[0]

            bio.seek(0)
            df = pd.read_excel(bio, sheet_name=sheet_name)
    except Exception as e:
        return {"status": "error", "error": f"Error parsing {filename}: {e}", "records": []}

    records, err = parse_dataframe_to_records(df, filename, sheet_name)
    if err:
        return {"status": "error", "error": err, "records": []}

    regions = sorted(list({r["region"] for r in records if r["region"]}))
    categories = sorted(list({r["category"] for r in records if r["category"]}))
    segments = sorted(list({r["segment"] for r in records if r["segment"]}))
    reps = sorted(list({r["sales_rep"] for r in records if r["sales_rep"]}))

    version_str = f"{filename}_{current_mtime}_{len(records)}"

    result = {
        "status": "success",
        "file": filename,
        "sheet": sheet_name,
        "path": str(file_path),
        "data_version": version_str,
        "mtime": current_mtime,
        "last_updated": file_info["last_updated"],
        "total_records": len(records),
        "filters": {
            "regions": regions,
            "categories": categories,
            "segments": segments,
            "sales_reps": reps
        },
        "records": records
    }

    cached_data[filename] = result
    cached_mtime[filename] = current_mtime
    return result


def save_records_to_excel(records, target_filename="Sales_Dashboard.xlsx"):
    """Saves records directly to Excel file on disk."""
    out_file = DIRECTORY / target_filename
    df = pd.DataFrame(records)

    # Format column names to standard COLUMNS
    col_map = {
        "date": "Date",
        "order_id": "Order_ID",
        "region": "Region",
        "category": "Category",
        "product": "Product",
        "sales_rep": "Sales_Rep",
        "segment": "Segment",
        "units": "Units",
        "unit_price": "Unit_Price",
        "revenue": "Revenue",
        "cost": "Cost",
        "profit": "Profit"
    }
    df = df.rename(columns=col_map)
    export_cols = [c for c in COLUMNS if c in df.columns]
    df = df[export_cols]

    # Save to Excel
    for attempt in range(5):
        try:
            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Data", index=False)
                df.to_excel(writer, sheet_name="Sales_Data", index=False)

            # Also update sales_data.xlsx for consistency
            sales_data_file = DIRECTORY / "sales_data.xlsx"
            if sales_data_file.exists():
                with pd.ExcelWriter(sales_data_file, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Sales_Data", index=False)
            return True, f"Saved {len(records)} records to {target_filename}"
        except PermissionError:
            time.sleep(0.3)
        except Exception as e:
            if attempt == 4:
                return False, str(e)
            time.sleep(0.3)
    return False, "File is locked by Excel. Please close or save in Excel."


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DIRECTORY), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        requested_file = params.get("file", [None])[0]

        if parsed.path == "/api/status":
            target, available = resolve_target_file(requested_file)
            if not target:
                payload = {"status": "error", "message": "No Excel or CSV files found"}
            else:
                payload = {
                    "status": "success",
                    "active_file": target["name"],
                    "mtime": target["mtime"],
                    "last_updated": target["last_updated"],
                    "data_version": f"{target['name']}_{target['mtime']}_{target['size']}",
                    "available_files": [{"name": f["name"], "last_updated": f["last_updated"], "mtime": f["mtime"], "folder": f["folder"]} for f in available]
                }
            content = json.dumps(payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        if parsed.path == "/api/data":
            target, available = resolve_target_file(requested_file)
            if not target:
                data = {"status": "error", "error": "No Excel file found", "records": []}
            else:
                data = safe_read_data_file(target)
                data["available_files"] = [{"name": f["name"], "last_updated": f["last_updated"], "mtime": f["mtime"], "folder": f["folder"]} for f in available]

            content = json.dumps(data).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        if parsed.path == "/" or parsed.path == "":
            self.path = "/index.html"

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)

        # POST /api/save: Save updated records directly to Excel
        if parsed.path == "/api/save":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body)
                records = payload.get("records", [])
                filename = payload.get("file", "Sales_Dashboard.xlsx")

                success, msg = save_records_to_excel(records, filename)
                res_payload = {"status": "success" if success else "error", "message": msg, "count": len(records)}
            except Exception as e:
                res_payload = {"status": "error", "message": str(e)}

            content = json.dumps(res_payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        # POST /api/upload: Upload new Excel/CSV file directly
        if parsed.path == "/api/upload":
            try:
                length = int(self.headers.get("Content-Length", 0))
                file_bytes = self.rfile.read(length)
                filename = self.headers.get("X-File-Name", "uploaded_sales_data.xlsx")
                dest = DIRECTORY / filename
                with open(dest, "wb") as f:
                    f.write(file_bytes)

                res_payload = {"status": "success", "file": filename, "message": f"Successfully uploaded {filename}"}
            except Exception as e:
                res_payload = {"status": "error", "message": str(e)}

            content = json.dumps(res_payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        # POST /api/delete_file: Delete an Excel or CSV file
        if parsed.path == "/api/delete_file":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(length).decode("utf-8")
                payload = json.loads(body)
                filename = payload.get("file", "").strip()

                if not filename or ".." in filename or "/" in filename or "\\" in filename:
                    res_payload = {"status": "error", "message": "Invalid file name"}
                elif not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".csv")):
                    res_payload = {"status": "error", "message": "Only .xlsx and .csv files can be deleted"}
                else:
                    target_path = None
                    if (DIRECTORY / filename).exists():
                        target_path = DIRECTORY / filename
                    elif (DOWNLOADS_DIR / filename).exists():
                        target_path = DOWNLOADS_DIR / filename

                    if not target_path:
                        res_payload = {"status": "error", "message": f"File '{filename}' not found"}
                    else:
                        try:
                            os.remove(target_path)
                            cached_data.pop(filename, None)
                            cached_mtime.pop(filename, None)
                            res_payload = {"status": "success", "message": f"Deleted '{filename}' successfully"}
                        except PermissionError:
                            res_payload = {"status": "error", "message": f"File '{filename}' is open in Excel. Please close it first."}
                        except Exception as e:
                            res_payload = {"status": "error", "message": str(e)}
            except Exception as e:
                res_payload = {"status": "error", "message": str(e)}

            content = json.dumps(res_payload).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(content)
            return

        self.send_response(404)
        self.end_headers()


def run():
    port = PORT
    for attempt in range(5):
        try:
            with socketserver.TCPServer(("", port), DashboardHandler) as httpd:
                print(f"Executive Sales Dashboard running at http://localhost:{port}")
                sys.stdout.flush()
                httpd.serve_forever()
        except OSError as e:
            if e.errno in (98, 10048):
                port += 1
                continue
            raise


if __name__ == "__main__":
    run()
