"""
Ecount ERP API 客戶端

申請 API 步驟：
1. 登入 Ecount ERP → 環境設定 → API 認證金鑰管理
2. 建立 API 金鑰後填入 .env 的 ECOUNT_* 欄位
3. API 文件：https://oapi.ecounterp.com/OAPI/V2/Swagger/Index

若 ECOUNT_COMPANY_NO 未設定，自動使用 mock 資料方便開發測試。
"""

import json
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path

import httpx

from config import settings

# ERP 可售庫存快取（庫存情況報表同步後的可售庫存）
_AVAILABLE_PATH = Path(__file__).parent.parent / "data" / "available.json"
_SYNC_SCRIPT = Path(__file__).parent.parent / "scripts" / "auto_sync_unfulfilled.py"
_PYTHON = r"C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe"

# 超過此秒數就觸發同步（2 分鐘）
_STALE_SECONDS = 2 * 60
_sync_lock = threading.Lock()


def _sync_and_wait():
    """
    執行同步並等待完成（最多 60 秒）。
    若同步正在執行中，等待它結束即可。
    """
    with _sync_lock:
        try:
            print("[Ecount] 同步庫存中，請稍候...")
            subprocess.run(
                [_PYTHON, str(_SYNC_SCRIPT)],
                creationflags=subprocess.CREATE_NO_WINDOW,
                cwd=str(_SYNC_SCRIPT.parent.parent),
                timeout=60,
            )
            print("[Ecount] 同步完成")
        except subprocess.TimeoutExpired:
            print("[Ecount] 同步逾時（60 秒），使用既有資料")
        except Exception as e:
            print(f"[Ecount] 同步失敗: {e}")


class EcountClient:
    def __init__(self):
        self._session_id: str | None = None
        self._session_expires: float = 0.0
        self._session_lock = threading.Lock()
        self._avail_lock = threading.Lock()
        # 品項快取：list of {"code": str, "name": str}
        self._product_cache: list[dict] = []
        self._cache_expires: float = 0.0
        # O(1) 查找索引（由 _ensure_product_cache 建立）
        self._product_by_code: dict[str, dict] = {}  # uppercase code -> item
        self._product_by_name: dict[str, dict] = {}  # uppercase name -> item
        # 持久 HTTP 連線池
        self._http = httpx.Client(timeout=15)

    # ------------------------------------------------------------------
    # 公開方法
    # ------------------------------------------------------------------

    def lookup(self, keyword: str) -> dict | None:
        """
        以產品編號或名稱（模糊）查詢庫存。

        優先使用 data/available.json（庫存情況報表的可售庫存），
        若無同步資料則 fallback 到 OAPI 取 BAL_QTY。

        Returns:
            dict  — {"code": str, "name": str, "qty": int}
            None  — 查無此產品
        """
        if not self._is_configured():
            return self._mock_lookup(keyword)

        # 先找出對應的 PROD_CD
        prod_cd = self._resolve_product_code(keyword)
        if not prod_cd:
            return None

        # 從快取取得中文名稱
        name = self._get_product_name(prod_cd) or prod_cd

        # 優先：使用 available.json 的可售庫存（已含 ERP 未出貨計算）
        erp_data = self._get_erp_available(prod_cd)
        if erp_data is not None:
            available = erp_data.get("available") or 0
            return {
                "code":     prod_cd,
                "name":     name,
                "qty":      available,
                "balance":  erp_data.get("balance"),   # 庫存數量
                "unfilled": erp_data.get("unfilled"),  # 未出貨
                "incoming": erp_data.get("incoming"),  # 未進貨（總公司未到）
                "preorder": erp_data.get("preorder"),  # 可預購數量
                "source":   "available.json",
            }

        # Fallback：OAPI 取 BAL_QTY（不扣 ERP 未出貨，但即時）
        qty = self._fetch_inventory(prod_cd)
        if qty is None:
            return None

        return {
            "code":   prod_cd,
            "name":   name,
            "qty":    qty,
            "source": "oapi",
        }

    def get_product_detail(self, prod_cd: str) -> dict | None:
        """
        查詢單一品項的詳細資料（單價、規格等）。

        Returns:
            dict  — {"code", "name", "spec", "price", "unit"}
            None  — 查無此品項
        """
        if not self._is_configured():
            return None

        session_id = self._ensure_session()
        if not session_id:
            return None

        try:
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBasic/ViewBasicProduct",
                params={"SESSION_ID": session_id},
                json={"PROD_CD": prod_cd, "PROD_TYPE": "0"},
                timeout=10,
            )
            if str(data.get("Status")) == "200":
                r = data.get("Data", {}).get("Result", {})
                if r:
                    return {
                        "code":  r.get("PROD_CD", ""),
                        "name":  r.get("PROD_DES", ""),
                        "spec":  r.get("SIZE_DES", ""),
                        "price": r.get("PRICE", ""),
                        "unit":  r.get("UNIT", ""),
                    }
            return None
        except Exception as e:
            print(f"[Ecount] get_product_detail 錯誤: {e}")
            return None

    def get_price(self, keyword: str) -> dict | None:
        """
        以產品編號或名稱查詢售價。

        Returns:
            dict — {"code": str, "name": str, "price": float, "unit": str}
            None — 查無此產品
        """
        if not self._is_configured():
            return None

        prod_cd = self._resolve_product_code(keyword)
        if not prod_cd:
            return None

        _uc = prod_cd.upper()
        item = self._product_by_code.get(_uc)
        if item:
            return {
                "code":  item["code"],
                "name":  item["name"],
                "price": item["price"],
                "unit":  item["unit"],
            }
        return None

    def get_customers_list(self) -> list[dict]:
        """
        抓取 Ecount 所有客戶/供應商清單。

        Returns:
            list of {
                "code":  BUSINESS_NO,
                "name":  CUST_NAME,
                "phone": HP_NO（手機），
                "tel":   TEL（電話）,
            }
        """
        if not self._is_configured():
            return []

        session_id = self._ensure_session()
        if not session_id:
            return []

        try:
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/AccountBasic/GetBasicCustList",
                params={"SESSION_ID": session_id},
                json={},
                timeout=20,
            )
            if str(data.get("Status")) == "200":
                results = data.get("Data", {}).get("Result", [])
                customers = []
                for r in results:
                    code = (r.get("BUSINESS_NO") or r.get("CUST_CD") or "").strip()
                    if not code:
                        continue
                    customers.append({
                        "code":  code,
                        "name":  (r.get("CUST_NAME") or "").strip(),
                        "phone": (r.get("HP_NO") or "").strip(),
                        "tel":   (r.get("TEL") or "").strip(),
                    })
                print(f"[Ecount] 客戶清單抓取成功，共 {len(customers)} 筆")
                return customers
            print(f"[Ecount] get_customers_list 失敗: {data.get('Status')} {data.get('Message','')}")
            return []
        except Exception as e:
            print(f"[Ecount] get_customers_list 錯誤: {e}")
            return []

    def save_customer(self, business_no: str, cust_name: str, **kwargs) -> bool:
        """
        在 Ecount 新增客戶/供應商。

        Args:
            business_no: 統一編號（客戶代碼）
            cust_name:   公司/客戶名稱
            kwargs:      選填欄位：tel, email, addr, hp_no, remarks

        Returns:
            True  — 新增成功
            False — 失敗
        """
        if not self._is_configured():
            print(f"[Ecount Mock] save_customer: {business_no} {cust_name}")
            return True

        session_id = self._ensure_session()
        if not session_id:
            return False

        try:
            bulk = {
                "BUSINESS_NO": business_no,
                "CUST_NAME":   cust_name,
                "CUST_GROUP1": "01",
                "HP_NO":       kwargs.get("hp_no", ""),
                "TEL":         kwargs.get("tel", ""),
                "EMAIL":       kwargs.get("email", ""),
                "ADDR":        kwargs.get("addr", ""),
                "REMARKS":     kwargs.get("remarks", ""),
            }
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/AccountBasic/SaveBasicCust",
                params={"SESSION_ID": session_id},
                json={"CustList": [{"BulkDatas": bulk}]},
                timeout=10,
            )
            if str(data.get("Status")) == "200":
                success = int(data.get("Data", {}).get("SuccessCnt", 0))
                if success > 0:
                    return True
                # SuccessCnt=0 可能是 session 過期，重試一次
                self._session_id = None
                session_id = self._ensure_session()
                if not session_id:
                    return False
                data2 = self._checked_post(
                    f"{settings.ECOUNT_BASE_URL}/OAPI/V2/AccountBasic/SaveBasicCust",
                    params={"SESSION_ID": session_id},
                    json={"CustList": [{"BulkDatas": bulk}]},
                    timeout=10,
                )
                return int(data2.get("Data", {}).get("SuccessCnt", 0)) > 0
            # session 失效 → 重新取得並重試
            self._session_id = None
            session_id = self._ensure_session()
            if not session_id:
                return False
            data2 = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/AccountBasic/SaveBasicCust",
                params={"SESSION_ID": session_id},
                json={"CustList": [{"BulkDatas": bulk}]},
                timeout=10,
            )
            return int(data2.get("Data", {}).get("SuccessCnt", 0)) > 0
        except Exception as e:
            print(f"[Ecount] save_customer 錯誤: {e}")
            return False

    def save_product(
        self,
        prod_cd:   str,
        prod_name: str,
        unit:      str = "個",
        size_flag: str = "",   # 規格類型：1=規格名稱 2=規格組合 3=規格計算 4=規格計算組合
        size_des:  str = "",   # 規格描述
        extra:     dict | None = None,  # 其他欄位（直接塞入 BulkDatas）
    ) -> str | None:
        """
        在 Ecount 新增品項。

        Returns:
            str  — 品項編碼（如 "TESTBOT001-"）
            None — 失敗
        """
        if not self._is_configured():
            print(f"[Ecount Mock] save_product: {prod_cd} {prod_name}")
            return prod_cd

        session_id = self._ensure_session()
        if not session_id:
            return None

        bulk = {"PROD_CD": prod_cd.upper(), "PROD_DES": prod_name}
        if unit:      bulk["UNIT"]      = unit
        if size_flag: bulk["SIZE_FLAG"] = size_flag
        if size_des:  bulk["SIZE_DES"]  = size_des
        if extra:     bulk.update(extra)

        try:
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBasic/SaveBasicProduct",
                params={"SESSION_ID": session_id},
                json={"ProductList": [{"BulkDatas": bulk}]},
                timeout=15,
            )
            if str(data.get("Status")) == "200" and not data.get("Errors"):
                slip = (data.get("Data", {}).get("SlipNos") or [""])[0]
                print(f"[Ecount] save_product 成功: {prod_cd} → {slip}")
                return slip or prod_cd
            errors = data.get("Errors") or []
            msgs = " | ".join(e.get("Message", "") for e in errors)
            print(f"[Ecount] save_product 失敗: {msgs}")
            return None
        except Exception as e:
            print(f"[Ecount] save_product 錯誤: {e}")
            return None

    def save_order(self, cust_code: str, items: list[dict], phone: str = "") -> str | None:
        """
        在 Ecount 新增訂貨單。

        Args:
            cust_code: 客戶編碼（BUSINESS_NO）
            items:     list of {"prod_cd": str, "qty": int, "price": str}
            phone:     客戶手機號碼（HP_NO），可選

        Returns:
            str   — 訂貨單號（如 "20260310-1"）
            None  — 失敗
        """
        if not self._is_configured():
            print(f"[Ecount Mock] save_order: cust={cust_code} items={items}")
            return "MOCK-ORDER-001"

        session_id = self._ensure_session()
        if not session_id:
            return None

        try:
            self._ensure_product_cache()
            today = datetime.now().strftime("%Y%m%d")
            bulk_list = []
            for item in items:
                # item 自帶 price 優先；否則從品項快取補 OUT_PRICE
                price = str(item.get("price", "")).strip()
                if not price:
                    _pcd = item["prod_cd"].upper()
                    p = self._product_by_code.get(_pcd)
                    if p and p.get("price"):
                        price = str(round(float(p["price"])))
                # SUPPLY_AMT = 單價 × 數量
                try:
                    supply_amt = str(round(float(price) * int(item["qty"]))) if price else ""
                except (ValueError, TypeError):
                    supply_amt = ""
                bulk_list.append({"BulkDatas": {
                    "UPLOAD_SER_NO":  "1",
                    "TIME_DATE":      today,
                    "CUST":           cust_code,
                    "U_MEMO1":        phone or "",
                    "WH_CD":          "101",
                    "PROD_CD":        item["prod_cd"],
                    "QTY":            str(item["qty"]),
                    "PRICE":          price,
                    "SUPPLY_AMT":     supply_amt,
                    "ITEM_TIME_DATE": today,
                    "REMARKS":        item.get("note", ""),   # 對應 Ecount 訂貨單行項「摘要」欄
                }})

            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/SaleOrder/SaveSaleOrder",
                params={"SESSION_ID": session_id},
                json={"SaleOrderList": bulk_list},
                timeout=10,
            )
            if str(data.get("Status")) == "200":
                slip_nos = data.get("Data", {}).get("SlipNos", [])
                slip_no = slip_nos[0] if slip_nos else ""
                success = int(data.get("Data", {}).get("SuccessCnt", 0))
                if success > 0 or slip_no:
                    return slip_no or "OK"
            details = data.get("Data", {}).get("ResultDetails", "")
            print(f"[Ecount] save_order 失敗: {details}")
            return None
        except Exception as e:
            print(f"[Ecount] save_order 錯誤: {e}")
            return None

    def _get_erp_available(self, prod_cd: str) -> dict | None:
        """
        從 data/available.json 讀取庫存明細。
        回傳 dict：{incoming, unfilled, balance, available, preorder}
        若檔案超過 2 分鐘未更新，先同步再回答。
        """
        try:
            if _AVAILABLE_PATH.exists():
                age = time.time() - _AVAILABLE_PATH.stat().st_mtime
                if age > _STALE_SECONDS:
                    print(f"[Ecount] available.json 已 {int(age/60)} 分鐘未更新，嘗試同步...")
                    mtime_before = _AVAILABLE_PATH.stat().st_mtime
                    _sync_and_wait()
                    # 同步後檔案若仍未更新 → 改用 OAPI 即時查詢
                    mtime_after = _AVAILABLE_PATH.stat().st_mtime if _AVAILABLE_PATH.exists() else 0
                    if mtime_after <= mtime_before:
                        print("[Ecount] 同步未能更新 available.json，改用 OAPI 即時查詢")
                        return None

            if _AVAILABLE_PATH.exists():
                with self._avail_lock:
                    data = json.loads(_AVAILABLE_PATH.read_text(encoding="utf-8"))
                if prod_cd in data:
                    entry = data[prod_cd]
                    # 新格式（dict）
                    if isinstance(entry, dict):
                        return entry
                    # 舊格式（int）→ 僅有可售庫存，其餘為 None
                    return {
                        "incoming": None,
                        "unfilled": None,
                        "balance":  None,
                        "available": int(entry),
                        "preorder":  None,
                    }
        except Exception as e:
            print(f"[Ecount] 讀取 available.json 失敗: {e}")
        return None

    def get_order(self, order_id: str) -> dict | None:
        """
        查詢訂單狀態（目前 Ecount OAPI 無查詢訂貨單端點，回傳 None）。
        """
        return None

    # ------------------------------------------------------------------
    # 內部：產品快取
    # ------------------------------------------------------------------

    def _ensure_product_cache(self):
        """
        品項快取刷新策略：
        - 上班時間（11:00–21:00）：每 2 小時更新一次
        - 下班時間：有快取就直接用，不更新
        - 無快取時（初次啟動）：無論何時都先抓一次
        """
        now  = datetime.now()
        hour = now.hour

        is_working = 11 <= hour < 21

        # 下班時間且已有快取 → 直接用
        if not is_working and self._product_cache:
            return

        # 快取未過期 → 直接用
        if self._product_cache and time.time() < self._cache_expires:
            return

        session_id = self._ensure_session()
        if not session_id:
            return

        try:
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBasic/GetBasicProductsList",
                params={"SESSION_ID": session_id},
                json={},
                timeout=20,
            )
            if str(data.get("Status")) == "200":
                results = data.get("Data", {}).get("Result", [])
                import re as _re
                _ZX_RE = _re.compile(r'^Z[A-Za-z]', _re.IGNORECASE)
                self._product_cache = [
                    {
                        "code":    r.get("PROD_CD", "").strip(),
                        "name":    r.get("PROD_DES", "").strip(),
                        "price":   float(r.get("OUT_PRICE") or 0),
                        "unit":    r.get("UNIT", "").strip(),
                        "box_qty": int(float(r.get("EXCH_RATE") or 0)),  # 裝箱數
                    }
                    for r in results
                    if r.get("PROD_CD") and not _ZX_RE.match(r.get("PROD_CD", ""))
                ]
                # 建立 O(1) 查找索引
                self._product_by_code = {
                    item["code"].upper(): item for item in self._product_cache
                }
                self._product_by_name = {
                    item["name"].upper(): item for item in self._product_cache
                }
                self._cache_expires = time.time() + 2 * 3600  # 上班時間 2 小時 TTL
                print(f"[Ecount] 品項快取刷新，共 {len(self._product_cache)} 筆（已排除 Z+英文 開頭貨號）")
        except Exception as e:
            print(f"[Ecount] 品項快取刷新失敗: {e}")

    def _resolve_product_code(self, keyword: str) -> str | None:
        """
        以關鍵字找出 PROD_CD（完全符合編號 > 名稱包含）。

        防誤判規則：
        - 長度 < 2：直接跳過
        - 純英文（無數字）且長度 < 4：不做名稱子字串比對
          （避免 "IN"、"LL" 等 OCR 雜訊命中含 "MINI" 的產品名稱）
        - 回傳「第一個」符合的結果，不再覆寫（避免最後一筆覆蓋最佳命中）
        """
        self._ensure_product_cache()

        kw = keyword.strip().upper()
        if len(kw) < 2:
            return None

        # O(1) 精確比對：code 完全符合
        exact = self._product_by_code.get(kw)
        if exact:
            return exact["code"]

        # O(1) 精確比對：name 完全符合
        exact_name = self._product_by_name.get(kw)
        if exact_name:
            return exact_name["code"]

        # 純英文字母（無數字）且太短 → 不做名稱子字串比對，只允許完全符合
        import re as _re
        _is_short_alpha = bool(_re.fullmatch(r'[A-Z]{1,3}', kw))

        # Fallback：線性掃描做部分/模糊匹配
        if not _is_short_alpha:
            for item in self._product_cache:
                code = item["code"].upper()
                name = item["name"].upper()
                if kw in name or kw in code:
                    return item["code"]  # 第一個名稱符合即回傳，不繼續覆寫

        return None

    def _get_product_name(self, prod_cd: str) -> str | None:
        """從快取中取得產品中文名稱（O(1) 索引查找）"""
        item = self._product_by_code.get(prod_cd.upper())
        return item["name"] if item else None

    def search_products_by_name(self, keyword: str) -> list[str]:
        """以關鍵字模糊搜尋所有符合品名的產品編號清單（最多 20 筆）"""
        self._ensure_product_cache()
        kw = keyword.strip().upper()
        if not kw:
            return []
        matched = [
            item["code"] for item in self._product_cache
            if kw in item["name"].upper() or kw in item["code"].upper()
        ]
        return matched[:20]

    def get_product_cache_item(self, prod_cd: str) -> dict | None:
        """從快取取得完整產品資料（含 unit, box_qty），O(1) 索引查找"""
        self._ensure_product_cache()
        return self._product_by_code.get(prod_cd.strip().upper())

    # ------------------------------------------------------------------
    # 內部：庫存查詢
    # ------------------------------------------------------------------

    def _fetch_inventory(self, prod_cd: str) -> int | None:
        """查詢特定 PROD_CD 在倉庫 101 的庫存數量"""
        session_id = self._ensure_session()
        if not session_id:
            print("[Ecount] 無法取得 Session")
            return None

        try:
            today = datetime.now().strftime("%Y%m%d")
            data = self._checked_post(
                f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus",
                params={"SESSION_ID": session_id},
                json={"PROD_CD": prod_cd, "WH_CD": "101", "BASE_DATE": today, "ZERO_FLAG": "Y"},
                timeout=10,
            )
            if str(data.get("Status")) == "200":
                results = data.get("Data", {}).get("Result", [])
                if results:
                    return int(float(results[0].get("BAL_QTY", 0)))
                return 0  # 有產品但庫存 0
            return None
        except Exception as e:
            print(f"[Ecount] _fetch_inventory 錯誤: {e}")
            return None

    def get_all_stock_products(self) -> list[dict]:
        """
        取得倉庫 101 所有有庫存（qty > 0）的品項。
        回傳 [{"code": ..., "name": ..., "qty": ...}, ...]

        策略：
        1. 優先讀取 data/available.json（已同步，最快）
        2. 若無檔案，嘗試不帶 PROD_CD 的批量 OAPI 查詢
        3. 若仍失敗，改用 product_cache 逐筆並發查詢
        """
        # ── 優先：讀取 available.json ──────────────────────────────────
        if _AVAILABLE_PATH.exists():
            try:
                with self._avail_lock:
                    raw = json.loads(_AVAILABLE_PATH.read_text(encoding="utf-8"))
                self._ensure_product_cache()
                code_to_name = {
                    item["code"].upper(): item["name"]
                    for item in (self._product_cache or [])
                }
                out = []
                for prod_cd, entry in raw.items():
                    qty = entry.get("available", 0) if isinstance(entry, dict) else int(entry or 0)
                    if qty and qty > 0:
                        out.append({
                            "code": prod_cd.upper(),
                            "name": code_to_name.get(prod_cd.upper(), ""),
                            "qty":  qty,
                        })
                print(f"[Ecount] get_all_stock_products（available.json）→ {len(out)} 筆有庫存品項")
                return out
            except Exception as e:
                print(f"[Ecount] 讀取 available.json 失敗（{e}），改用 OAPI")

        # ── fallback 1：批量 OAPI 查詢 ────────────────────────────────
        session_id = self._ensure_session()
        if session_id:
            try:
                today = datetime.now().strftime("%Y%m%d")
                data = self._checked_post(
                    f"{settings.ECOUNT_BASE_URL}/OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus",
                    params={"SESSION_ID": session_id},
                    json={"WH_CD": "101", "BASE_DATE": today, "ZERO_FLAG": "N"},
                    timeout=30,
                )
                if str(data.get("Status")) == "200":
                    results = data.get("Data", {}).get("Result", [])
                    if results:
                        out = []
                        for r in results:
                            qty = int(float(r.get("BAL_QTY", 0)))
                            if qty > 0:
                                out.append({
                                    "code": r.get("PROD_CD", "").strip().upper(),
                                    "name": r.get("PROD_DES", "").strip(),
                                    "qty":  qty,
                                })
                        print(f"[Ecount] get_all_stock_products（批量 OAPI）→ {len(out)} 筆有庫存品項")
                        return out
                    print("[Ecount] 批量庫存查詢回傳空，改用逐筆查詢")
            except Exception as e:
                print(f"[Ecount] 批量庫存查詢失敗（{e}），改用逐筆查詢")

        # ── fallback 2：逐筆查詢（並發 8 條）─────────────────────────
        if not session_id:
            session_id = self._ensure_session()
        self._ensure_product_cache()
        if not self._product_cache:
            print("[Ecount] product_cache 為空，無法逐筆查詢")
            return []

        from concurrent.futures import ThreadPoolExecutor, as_completed

        def _query_one(item: dict) -> dict | None:
            qty = self._fetch_inventory(item["code"])
            if qty and qty > 0:
                return {"code": item["code"], "name": item["name"], "qty": qty}
            return None

        out = []
        total = len(self._product_cache)
        print(f"[Ecount] 逐筆查詢 {total} 筆品項庫存…")
        with ThreadPoolExecutor(max_workers=8) as pool:
            futures = {pool.submit(_query_one, item): item for item in self._product_cache}
            for fut in as_completed(futures):
                result = fut.result()
                if result:
                    out.append(result)
        print(f"[Ecount] 逐筆查詢完成，{len(out)} 筆有庫存")
        return out

    # ------------------------------------------------------------------
    # 內部：Session 管理
    # ------------------------------------------------------------------

    def _is_configured(self) -> bool:
        return bool(settings.ECOUNT_COMPANY_NO)

    def _ensure_session(self) -> str | None:
        """取得（或重用）Ecount Session ID"""
        with self._session_lock:
            if self._session_id and time.time() < self._session_expires:
                return self._session_id

            try:
                resp = self._http.post(
                    f"{settings.ECOUNT_BASE_URL}/OAPI/V2/OAPILogin",
                    json={
                        "COM_CODE":     settings.ECOUNT_COMPANY_NO,
                        "USER_ID":      settings.ECOUNT_USER_ID,
                        "API_CERT_KEY": settings.ECOUNT_API_CERT_KEY,
                        "LAN_TYPE":     "zh-TW",
                        "ZONE":         settings.ECOUNT_ZONE,
                    },
                    timeout=10,
                )
                data = self._safe_json(resp)
                if str(data.get("Status")) == "200" and str(data.get("Data", {}).get("Code")) == "00":
                    self._session_id = data["Data"]["Datas"]["SESSION_ID"]
                    self._session_expires = time.time() + 3600  # 1 小時有效
                    return self._session_id
            except Exception as e:
                print(f"[Ecount] 取得 Session 失敗: {e}")
            return None

    # ------------------------------------------------------------------
    # 內部工具
    # ------------------------------------------------------------------

    def close(self):
        """關閉 HTTP 連線池"""
        self._http.close()

    def _checked_post(self, url, **kwargs):
        """POST with HTTP status code validation"""
        resp = self._http.post(url, **kwargs)
        if resp.status_code >= 400:
            print(f"[ecount] HTTP {resp.status_code} for {url}")
            return {}
        return self._safe_json(resp)

    @staticmethod
    def _safe_json(resp) -> dict:
        """
        安全解析 httpx Response 為 dict。
        Ecount 有時在 session 過期後回傳 Big5/空 body，需要容錯處理。
        """
        content = resp.content
        if not content:
            return {}
        for enc in ("utf-8", "big5", "gbk", "gb18030"):
            try:
                import json as _json
                return _json.loads(content.decode(enc))
            except Exception:
                continue
        return {}

    # ------------------------------------------------------------------
    # Mock 資料（開發用，正式設定 Ecount API 後自動停用）
    # ------------------------------------------------------------------

    def _mock_lookup(self, keyword: str) -> dict | None:
        mock = {
            "A001": {"name": "測試商品A", "qty": 50},
            "A002": {"name": "測試商品B", "qty": 0},
            "B001": {"name": "測試商品C", "qty": 12},
        }
        kw = keyword.strip().upper()
        for code, info in mock.items():
            if kw == code or kw in info["name"]:
                print(f"[Ecount Mock] 庫存查詢: {kw} → qty={info['qty']}")
                return {"code": code, "name": info["name"], "qty": info["qty"]}
        print(f"[Ecount Mock] 庫存查詢: {kw} → 查無資料")
        return None

    def _mock_order(self, order_id: str) -> dict | None:
        upper = order_id.upper()
        if upper.startswith("ORD"):
            return {
                "status": "confirmed",
                "items": "測試商品 x2",
                "eta": "2026-03-10",
            }
        print(f"[Ecount Mock] 訂單查詢: {order_id} → 查無資料")
        return None


def _map_ecount_status(flag: str) -> str:
    return {
        "1": "pending",
        "2": "confirmed",
        "3": "shipped",
        "4": "delivered",
        "9": "cancelled",
    }.get(flag, "pending")


ecount_client = EcountClient()
