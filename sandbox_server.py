import os
import sys
import json
import re
import zipfile
import shutil
from datetime import datetime, timedelta
import yaml
import requests
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from msal import PublicClientApplication, SerializableTokenCache

sys.stdout.reconfigure(encoding="utf-8")

SANDBOX_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(os.path.dirname(SANDBOX_DIR))
_workspace_src = os.path.join(BASE_DIR, "PowerApp", "SEAZEN HR", "Src")
DEFAULT_SRC_DIR = _workspace_src if os.path.exists(_workspace_src) else os.path.join(SANDBOX_DIR, "samples")
_workspace_token = os.path.join(BASE_DIR, "SeaZen-AI-Checkin-Service", "token_cache.bin")
TOKEN_CACHE_PATH = _workspace_token if os.path.exists(_workspace_token) else os.path.join(SANDBOX_DIR, "token_cache.bin")
PUBLISHED_DIR = os.path.join(SANDBOX_DIR, "published")

TENANT_ID = "29101901-3778-4339-827c-7d2b50914db0"
SITE_ID = "seazenvn.sharepoint.com,d47e794f-870a-4d63-8173-4d0721d5d77d,69313196-298d-4be2-af46-f7893c08cccf"

# Known SharePoint Lists in ONESEAZEN mapped by Display Name
KNOWN_LISTS = {
    "APP CHẤM CÔNG - THÔNG BÁO": "35164ecb-9333-4c0f-8c68-2c18edcda7b3",
    "APP CHẤM CÔNG - LOG CHECKIN/CHECKOUT": "4dfb189e-9b8e-43df-9f72-d04bbf18f6d3",
    "APP CHẤM CÔNG - NGAN HANG CAU HOI": "3df4ac7f-5dbb-4df0-835e-9a003c596e01",
    "APP CHẤM CÔNG - KET QUA KIEM TRA": "4bfd6f2d-f92d-4f36-8ee2-f729c050fe49",
    "APP CHẤM CÔNG - CAM KẾT TUÂN THỦ COC": "48d19ce4-dfdd-45dd-880a-ba441fa9dae3",
    "APP CHẤM CÔNG - ĐĂNG KÝ NGHỈ VIỆC": "21ca89c7-753c-4b49-acff-00e3d1602eec",
    "APP CHẤM CÔNG - ĐĂNG KÍ NGHỈ VIỆC": "21ca89c7-753c-4b49-acff-00e3d1602eec",
    "APP CHẤM CÔNG - ĐĂNG KÝ NGHỈ PHÉP/NGHỈ OFF": "3a093630-3146-4319-b9db-2f65f1e1f71a",
    "APP CHẤM CÔNG - ĐĂNG KÝ TĂNG CA": "7246207b-83d2-4676-92fe-51dc8edb36fe",
    "APP CHẤM CÔNG - ĐĂNG KÝ CÔNG TÁC": "1c1a7658-6c81-4e57-9d0f-e905f56b411d",
    "APP CHẤM CÔNG - ĐĂNG KÝ ĐIỀU CHUYỂN": "ad31bd68-f86f-41da-a426-4192c90cceef",
    "APP CHẤM CÔNG - ĐĂNG KÝ VÀO TRỄ/RA SỚM": "195a595b-0fc1-4b6c-a144-d7ae6a20700c",
    "APP CHẤM CÔNG - ĐĂNG KÝ ĐIỂM DANH BỔ SUNG": "d6a645ef-7d12-424e-b8fc-c632cf42a20a",
    "APP CHẤM CÔNG - PHẢN ÁNH": "e8834a65-c3c9-4b74-813c-072e6cfa9d1e",
    "APP CHẤM CÔNG - PHIẾU LƯƠNG": "df7a3038-92b5-4658-be3d-0b6485586c3a",
    "APP CHẤM CÔNG - PHÂN CA TRỰC": "30eb7a52-24ef-4a42-a68c-d791f3074a26",
    "MASTERLOCATION": "36152949-e342-4a2e-ac8e-9897d15066fc",
    "DANH SÁCH NHÂN VIÊN SEAZEN": "36ce86a4-8f94-4410-9113-92340ce251d7",
}

WORKSPACE_STATE = {
    "active_dir": DEFAULT_SRC_DIR,
    "opened_files": {},  # filename -> full_path
}


class CustomSafeLoader(yaml.SafeLoader):
    pass


def default_ctor(loader, tag_suffix, node):
    if isinstance(node, yaml.ScalarNode):
        return loader.construct_scalar(node)
    elif isinstance(node, yaml.SequenceNode):
        return loader.construct_sequence(node)
    elif isinstance(node, yaml.MappingNode):
        return loader.construct_mapping(node)
    return None


CustomSafeLoader.add_multi_constructor("", default_ctor)


def get_auth_context():
    """Acquire OAuth2 Bearer Token silently from token_cache.bin (No-Cookie Standard)."""
    if not os.path.exists(TOKEN_CACHE_PATH):
        return {"authenticated": False, "token": None, "username": None, "client_id": None}

    cache = SerializableTokenCache()
    try:
        with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    except Exception:
        return {"authenticated": False, "token": None, "username": None, "client_id": None}

    for cid in ["0390280b-12b8-49c4-adbf-a9c2b8bccfb5", "67ab195e-f6ff-40ff-9169-73035faaf572"]:
        app = PublicClientApplication(
            cid,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=cache,
        )
        accounts = app.get_accounts()
        if accounts:
            res = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=accounts[0])
            if res and "access_token" in res:
                if cache.has_state_changed:
                    with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
                        f.write(cache.serialize())
                return {
                    "authenticated": True,
                    "token": res["access_token"],
                    "username": accounts[0].get("username"),
                    "client_id": cid,
                }
    return {"authenticated": False, "token": None, "username": None, "client_id": None}


def interactive_oauth_login():
    """Launch interactive Microsoft OAuth2 login using SeaZen App Registration."""
    cache = SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        try:
            with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception:
            pass

    cid = "0390280b-12b8-49c4-adbf-a9c2b8bccfb5"  # SEAZENHR-PORTAL
    app = PublicClientApplication(
        cid,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    res = app.acquire_token_interactive(scopes=["https://graph.microsoft.com/.default"])
    if res and "access_token" in res:
        os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())
        accounts = app.get_accounts()
        username = accounts[0].get("username") if accounts else "Authenticated User"
        return {"authenticated": True, "username": username, "client_id": cid}
    return {"authenticated": False, "error": res.get("error_description", "Login failed")}


def resolve_list_id(list_name_or_id, token):
    """Resolve a SharePoint list displayName to its GUID."""
    clean = str(list_name_or_id).strip().strip("'").strip('"')
    if clean in KNOWN_LISTS:
        return KNOWN_LISTS[clean]
    if re.match(r"^[0-9a-fA-F-]{36}$", clean):
        return clean
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists?$top=500", headers=headers, timeout=8)
    if r.status_code == 200:
        for lst in r.json().get("value", []):
            dn = lst.get("displayName", "")
            KNOWN_LISTS[dn] = lst.get("id")
            if dn.upper() == clean.upper():
                return lst.get("id")
    return None


def lint_yaml_and_powerfx(raw_yaml, parsed_data):
    diagnostics = []
    if re.search(r"LookUp\(\s*'APP CHẤM CÔNG - THÔNG BÁO'", raw_yaml):
        diagnostics.append({
            "level": "error",
            "rule": "SP-DELEGATION-5000",
            "message": "LookUp trực tiếp vào 'APP CHẤM CÔNG - THÔNG BÁO' bị giới hạn 500 dòng đầu (trả về tin cũ ID 424). Hãy lọc qua 'Created >= varNotiCutoffDate'.",
        })
    if re.search(r"Filter\(\s*'APP CHẤM CÔNG - THÔNG BÁO'\s*,\s*HienThiNoi", raw_yaml):
        diagnostics.append({
            "level": "error",
            "rule": "SP-UNINDEXED-COLUMN",
            "message": "Cột 'HienThiNoi' chưa được đánh Index trên SharePoint. Lọc trực tiếp sẽ bị chặn ở 500 dòng đầu.",
        })
    if "ThisItem.StatusBg" in raw_yaml and "StatusBg:" not in raw_yaml:
        diagnostics.append({
            "level": "error",
            "rule": "MISSING-COLLECTION-PROP",
            "message": "Tham chiếu 'ThisItem.StatusBg' không tồn tại trong bộ sưu tập -> Gây lỗi đỏ và làm nút Classic/Button bị đen.",
        })
    return diagnostics


LIVE_CACHE = {"data": None}


def fetch_live_sp_context(msnv="14703"):
    auth = get_auth_context()
    token = auth.get("token")
    context = {
        "msnv": msnv,
        "hoTen": "NGUYỄN HUỲNH VĂN SANG",
        "donVi": "SEAZEN TECH",
        "authenticated": auth.get("authenticated", False),
        "username": auth.get("username"),
        "notifications": [],
        "latestPopup": None,
        "examAlerts": [],
        "source": "Local Fallback",
    }

    if token:
        headers = {"Authorization": f"Bearer {token}"}
        try:
            list_tb = KNOWN_LISTS["APP CHẤM CÔNG - THÔNG BÁO"]
            cutoff_iso = (datetime.utcnow() - timedelta(days=3)).strftime("%Y-%m-%dT00:00:00Z")
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_tb}/items"
                f"?expand=fields&$filter=fields/Created ge '{cutoff_iso}'&$top=999"
            )
            r = requests.get(url, headers=headers, timeout=8)
            if r.status_code == 200:
                items = r.json().get("value", [])
                items.sort(key=lambda x: int(x.get("id", 0)), reverse=True)
                notis = []
                latest_popup = None
                for it in items:
                    f = it.get("fields", {})
                    nguoi_nhan = str(f.get("NguoiNhan", ""))
                    loai_tb = str(f.get("LoaiThongBao", "Hệ thống"))
                    if loai_tb == "Hệ thống" or nguoi_nhan == "ALL" or msnv in nguoi_nhan:
                        rec = {
                            "ID": int(it.get("id", 0)),
                            "Title": f.get("Title", "Thông báo"),
                            "NoiDung": f.get("NoiDung", ""),
                            "LoaiDon": {"Value": f.get("LoaiDon", "Khác")},
                            "LoaiThongBao": {"Value": loai_tb},
                            "KieuThongBao": {"Value": f.get("KieuThongBao", "Thông báo thường")},
                            "HienThiNoi": {"Value": f.get("HienThiNoi", "Ẩn")},
                            "NguoiNhan": nguoi_nhan,
                            "Created": f.get("Created", ""),
                            "CreatedBy": "Phòng Nhân Sự",
                        }
                        notis.append(rec)
                        if f.get("HienThiNoi") == "Hiện" and latest_popup is None:
                            latest_popup = rec
                context["notifications"] = notis[:30]
                context["latestPopup"] = latest_popup
                context["source"] = f"SharePoint Live ({auth.get('username')})"
        except Exception as e:
            print("Warning live SP fetch:", e)

    if not context["latestPopup"]:
        context["latestPopup"] = {
            "ID": 5307,
            "Title": "CHĂM LO TẾT ĐINH MÙI 2027",
            "NoiDung": "<div style='padding:12px;background:#fff8e9;border:1px solid #d9a441;border-radius:8px;'><h3 style='color:#a51f24;margin:0 0 8px 0;'>THÔNG BÁO ĐĂNG KÝ CHĂM LO TẾT ĐINH MÙI 2027</h3><p>Công đoàn SeaZen được phân bổ <b>91 suất chăm lo Tết</b>.</p><p><b>Hạn đăng ký:</b> 17h00 thứ Hai, 02/11/2026.</p></div>",
            "LoaiDon": {"Value": "Khác"},
            "LoaiThongBao": {"Value": "Hệ thống"},
            "KieuThongBao": {"Value": "Thông báo HTML"},
            "HienThiNoi": {"Value": "Hiện"},
            "Created": "2026-10-30T05:15:26Z",
            "CreatedBy": "Ban Chấp Hành Công Đoàn SeaZen",
        }

    context["examAlerts"] = [
        {
            "AlertId": "COC_COMMIT",
            "ChuDeCode": "COC",
            "Tag": "VĂN BẢN QUY CHẾ",
            "TimeTag": "Phụ lục 04",
            "Title": "1. Đọc & Ký Cam kết Tuân thủ COC",
            "Desc": "Nghiên cứu 5 chương quy tắc ứng xử, các lằn ranh đỏ (Red Lines) và thực hiện ký cam kết điện tử theo QĐ 382/2026/QĐ-SEAZEN.",
            "StatusText": "● Đã ký cam kết",
            "StatusColor": "rgba(22, 101, 52, 1)",
            "ActionText": "Xem lại",
            "ScreenTarget": "DOC_VA_KY",
            "Priority": 3,
        },
        {
            "AlertId": "EXAM_COC",
            "ChuDeCode": "COC",
            "Tag": "KHẢO SÁT HIỂU BIẾT (COC)",
            "TimeTag": "10 câu • 15 phút",
            "Title": "2. Bộ Quy Tắc Ứng Xử & Đạo Đức (COC)",
            "Desc": "Trắc nghiệm nhận thức & tình huống thực tế theo QĐ 382/2026/QĐ-SEAZEN (Điều 49).",
            "StatusText": "● Chưa đạt (8/10) - Làm lại",
            "StatusColor": "rgba(225, 29, 72, 1)",
            "ActionText": "Làm lại ngay",
            "ScreenTarget": "KIEM_TRA",
            "Priority": 2,
        },
    ]
    return context


class SandboxHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path in ["/", "/index.html", "/player"]:
            index_file = os.path.join(SANDBOX_DIR, "index.html")
            with open(index_file, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(content)
            return

        if path == "/api/auth/status":
            auth = get_auth_context()
            self._send_json({
                "authenticated": auth["authenticated"],
                "username": auth["username"],
                "client_id": auth["client_id"],
                "tenant_id": TENANT_ID,
            })
            return

        if path == "/api/screens":
            active_dir = WORKSPACE_STATE["active_dir"]
            files = list(WORKSPACE_STATE["opened_files"].keys())
            if os.path.exists(active_dir):
                for fn in sorted(os.listdir(active_dir)):
                    if fn.endswith(".pa.yaml") and not fn.startswith("App.") and fn not in files:
                        files.append(fn)
            self._send_json({"screens": files, "activeDir": active_dir})
            return

        if path == "/api/screen":
            name = qs.get("name", ["ThongBao_LichSu.pa.yaml"])[0]
            fpath = WORKSPACE_STATE["opened_files"].get(name) or os.path.join(WORKSPACE_STATE["active_dir"], name)
            if not os.path.exists(fpath):
                self._send_json({"error": f"File not found: {fpath}"}, 404)
                return
            with open(fpath, "r", encoding="utf-8") as f:
                raw_yaml = f.read()
            try:
                ast_data = yaml.load(raw_yaml, Loader=CustomSafeLoader)
                diags = lint_yaml_and_powerfx(raw_yaml, ast_data)
                self._send_json({
                    "name": name,
                    "path": fpath,
                    "rawYaml": raw_yaml,
                    "ast": ast_data,
                    "diagnostics": diags,
                    "valid": True,
                })
            except Exception as e:
                self._send_json({
                    "name": name,
                    "path": fpath,
                    "rawYaml": raw_yaml,
                    "ast": None,
                    "diagnostics": [{"level": "error", "rule": "YAML-SYNTAX", "message": str(e)}],
                    "valid": False,
                })
            return

        if path == "/api/sp-data":
            msnv = qs.get("msnv", ["14703"])[0]
            refresh = qs.get("refresh", ["0"])[0] == "1"
            if LIVE_CACHE["data"] is None or refresh:
                LIVE_CACHE["data"] = fetch_live_sp_context(msnv)
            self._send_json(LIVE_CACHE["data"])
            return

        if path == "/api/sp/lists":
            self._send_json({"lists": [{"name": k, "id": v} for k, v in KNOWN_LISTS.items()]})
            return

        if path == "/api/sp/items":
            list_name = qs.get("list", ["APP CHẤM CÔNG - THÔNG BÁO"])[0]
            top = int(qs.get("top", ["50"])[0])
            auth = get_auth_context()
            if not auth["authenticated"]:
                self._send_json({"error": "Chưa đăng nhập OAuth2 Bearer Token"}, 401)
                return
            token = auth["token"]
            list_id = resolve_list_id(list_name, token)
            if not list_id:
                self._send_json({"error": f"Không tìm thấy SharePoint List: {list_name}"}, 404)
                return
            headers = {"Authorization": f"Bearer {token}"}
            # Fetch recent items via Created filter if APP CHẤM CÔNG - THÔNG BÁO
            if list_id == KNOWN_LISTS["APP CHẤM CÔNG - THÔNG BÁO"]:
                cutoff_iso = (datetime.utcnow() - timedelta(days=3)).strftime("%Y-%m-%dT00:00:00Z")
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_id}/items?expand=fields&$filter=fields/Created ge '{cutoff_iso}'&$top=999"
            else:
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_id}/items?expand=fields&$top=200"
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code != 200:
                self._send_json({"error": r.text}, r.status_code)
                return
            raw_items = r.json().get("value", [])
            raw_items.sort(key=lambda x: int(x.get("id", 0)), reverse=True)
            records = []
            for it in raw_items[:top]:
                fld = it.get("fields", {})
                fld["ID"] = int(it.get("id", 0))
                records.append(fld)
            self._send_json({"list": list_name, "listId": list_id, "items": records})
            return

        self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        body_raw = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
        payload = json.loads(body_raw)

        # 1. Native Windows Open File Dialog
        if parsed.path == "/api/dialog/open-file":
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                file_paths = filedialog.askopenfilenames(
                    title="Chọn file màn hình Power Apps (.pa.yaml / .msapp)",
                    initialdir=WORKSPACE_STATE["active_dir"],
                    filetypes=[("Power Apps Files", "*.pa.yaml *.yaml *.yml *.msapp"), ("All Files", "*.*")],
                )
                root.destroy()
                if not file_paths:
                    self._send_json({"cancelled": True})
                    return

                loaded_names = []
                for fp in file_paths:
                    if fp.lower().endswith(".msapp"):
                        extract_dir = os.path.join(SANDBOX_DIR, "unpacked_msapp")
                        os.makedirs(extract_dir, exist_ok=True)
                        with zipfile.ZipFile(fp, "r") as zf:
                            zf.extractall(extract_dir)
                        src_sub = os.path.join(extract_dir, "Src")
                        target_dir = src_sub if os.path.exists(src_sub) else extract_dir
                        WORKSPACE_STATE["active_dir"] = target_dir
                        for fn in os.listdir(target_dir):
                            if fn.endswith(".pa.yaml"):
                                loaded_names.append(fn)
                    else:
                        fn = os.path.basename(fp)
                        WORKSPACE_STATE["opened_files"][fn] = fp
                        WORKSPACE_STATE["active_dir"] = os.path.dirname(fp)
                        loaded_names.append(fn)

                self._send_json({
                    "cancelled": False,
                    "loaded": loaded_names,
                    "firstFile": loaded_names[0] if loaded_names else None,
                    "activeDir": WORKSPACE_STATE["active_dir"],
                })
            except Exception as e:
                self._send_json({"error": str(e)}, 500)
            return

        # 2. Native Windows Open Folder Dialog
        if parsed.path == "/api/dialog/open-folder":
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                folder_path = filedialog.askdirectory(
                    title="Chọn thư mục chứa các file màn hình (.pa.yaml)",
                    initialdir=WORKSPACE_STATE["active_dir"],
                )
                root.destroy()
                if not folder_path:
                    self._send_json({"cancelled": True})
                    return
                WORKSPACE_STATE["active_dir"] = folder_path
                files = [fn for fn in sorted(os.listdir(folder_path)) if fn.endswith(".pa.yaml")]
                self._send_json({
                    "cancelled": False,
                    "activeDir": folder_path,
                    "screens": files,
                    "firstFile": files[0] if files else None,
                })
            except Exception as e:
                self._send_json({"error": str(e)}, 500)
            return

        # 3. Interactive OAuth2 Login
        if parsed.path == "/api/auth/login":
            res = interactive_oauth_login()
            LIVE_CACHE["data"] = None
            self._send_json(res)
            return

        # 4. Live SharePoint CRUD (Create, Update, Delete via OAuth2 Bearer Token)
        if parsed.path == "/api/sp/crud":
            auth = get_auth_context()
            if not auth["authenticated"]:
                self._send_json({"error": "Chưa xác thực OAuth2 Bearer Token"}, 401)
                return
            token = auth["token"]
            action = payload.get("action", "create").lower()
            list_name = payload.get("list", "")
            item_id = payload.get("id")
            fields_data = payload.get("fields", {})

            list_id = resolve_list_id(list_name, token)
            if not list_id:
                self._send_json({"error": f"Không tìm thấy danh sách SharePoint: {list_name}"}, 404)
                return

            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            }

            # Normalize Power Apps Choice objects {Value: "X"} -> "X" for Graph API
            clean_fields = {}
            for k, v in fields_data.items():
                if k in ["ID", "id", "Created", "Modified", "Author", "Editor", "@odata.etag"]:
                    continue
                if isinstance(v, dict) and "Value" in v:
                    clean_fields[k] = v["Value"]
                else:
                    clean_fields[k] = v

            if action == "create":
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_id}/items"
                r = requests.post(url, headers=headers, json={"fields": clean_fields}, timeout=10)
                if r.status_code in [200, 201]:
                    created_item = r.json()
                    LIVE_CACHE["data"] = None
                    self._send_json({
                        "success": True,
                        "action": "create",
                        "id": created_item.get("id"),
                        "item": created_item.get("fields", {}),
                    })
                else:
                    self._send_json({"success": False, "error": r.text}, r.status_code)
                return

            elif action == "update":
                if not item_id:
                    self._send_json({"error": "Thiếu ID bản ghi để cập nhật (Patch)"}, 400)
                    return
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_id}/items/{item_id}/fields"
                r = requests.patch(url, headers=headers, json=clean_fields, timeout=10)
                if r.status_code in [200, 204]:
                    LIVE_CACHE["data"] = None
                    self._send_json({
                        "success": True,
                        "action": "update",
                        "id": item_id,
                        "item": r.json() if r.text else clean_fields,
                    })
                else:
                    self._send_json({"success": False, "error": r.text}, r.status_code)
                return

            elif action == "delete":
                if not item_id:
                    self._send_json({"error": "Thiếu ID bản ghi để xóa (Remove)"}, 400)
                    return
                url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{list_id}/items/{item_id}"
                r = requests.delete(url, headers=headers, timeout=10)
                if r.status_code in [200, 204]:
                    LIVE_CACHE["data"] = None
                    self._send_json({"success": True, "action": "delete", "id": item_id})
                else:
                    self._send_json({"success": False, "error": r.text}, r.status_code)
                return

            self._send_json({"error": f"Unknown CRUD action: {action}"}, 400)
            return

        # 5. Parse YAML
        if parsed.path == "/api/parse":
            raw_yaml = payload.get("yaml", "")
            try:
                ast_data = yaml.load(raw_yaml, Loader=CustomSafeLoader)
                diags = lint_yaml_and_powerfx(raw_yaml, ast_data)
                self._send_json({
                    "valid": True,
                    "ast": ast_data,
                    "diagnostics": diags,
                })
            except Exception as e:
                self._send_json({
                    "valid": False,
                    "ast": None,
                    "diagnostics": [{"level": "error", "rule": "YAML-SYNTAX", "message": str(e)}],
                })
            return

        # 6. Save Screen
        if parsed.path == "/api/save-screen":
            name = payload.get("name", "")
            raw_yaml = payload.get("yaml", "")
            if not name or not name.endswith(".pa.yaml"):
                self._send_json({"error": "Invalid screen file name"}, 400)
                return
            try:
                ast_data = yaml.load(raw_yaml, Loader=CustomSafeLoader)
                diags = lint_yaml_and_powerfx(raw_yaml, ast_data)
                fpath = WORKSPACE_STATE["opened_files"].get(name) or os.path.join(WORKSPACE_STATE["active_dir"], name)
                with open(fpath, "w", encoding="utf-8") as f:
                    f.write(raw_yaml)
                self._send_json({"saved": True, "name": name, "path": fpath, "diagnostics": diags})
            except Exception as e:
                self._send_json({"saved": False, "error": str(e)}, 400)
            return

        # 7. Publish App Launcher
        if parsed.path == "/api/publish":
            os.makedirs(PUBLISHED_DIR, exist_ok=True)
            studio_bat = os.path.join(BASE_DIR, "PowerApp", "Chay_SeaZen_PowerApp_Studio.bat")
            player_bat = os.path.join(BASE_DIR, "PowerApp", "Chay_SeaZen_App_Published.bat")
            with open(studio_bat, "w", encoding="utf-8") as f:
                f.write('@echo off\nchcp 65001 >nul\nstart "" "http://localhost:8095"\npython "%~dp0sandbox\\sandbox_server.py"\n')
            with open(player_bat, "w", encoding="utf-8") as f:
                f.write('@echo off\nchcp 65001 >nul\nstart "" "http://localhost:8095/player?mode=published"\npython "%~dp0sandbox\\sandbox_server.py"\n')
            self._send_json({
                "published": True,
                "playerUrl": "http://localhost:8095/player?mode=published",
                "studioLauncher": studio_bat,
                "playerLauncher": player_bat,
            })
            return

        self._send_json({"error": "Not found"}, 404)


if __name__ == "__main__":
    port = 8095
    print(f"Starting SeaZen Power Apps Local Studio & Live CRUD Engine on http://localhost:{port}")
    server = HTTPServer(("127.0.0.1", port), SandboxHandler)
    server.serve_forever()
