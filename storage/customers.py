"""
客戶資料庫（SQLite）

欄位：
- line_user_id   : LINE userId（bot 互動時自動更新）
- display_name   : LINE 顯示名稱（from API / CSV，為 LINE 暱稱）
- real_name      : 客戶提供的真實姓名（透過 awaiting_contact_info 流程填寫）
- chat_label     : LINE OA 對話標籤（CSV 檔名）
- phone          : 電話（從對話萃取）
- address        : 收貨地址
- note           : 備註（VIP 等級、習慣等）
- first_seen     : 第一次互動時間
- last_seen      : 最後互動時間

來源：
1. 匯入腳本：scripts/import_customers.py（從 CSV 批次匯入）
2. Bot 執行時：客戶傳訊息自動 upsert line_user_id + display_name
3. awaiting_contact_info 流程：客戶提供姓名時寫入 real_name
"""

import re
import sqlite3
from datetime import datetime
from pathlib import Path

_HELPER_RE = re.compile(r'helper:(\d+)')

DB_PATH = Path(__file__).parent.parent / "data" / "customers.db"


class CustomerStore:
    def __init__(self):
        DB_PATH.parent.mkdir(exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS customers (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    line_user_id  TEXT    UNIQUE,
                    display_name  TEXT,
                    chat_label    TEXT,
                    phone         TEXT,
                    address       TEXT,
                    note          TEXT,
                    first_seen    TEXT,
                    last_seen     TEXT
                )
            """)
            # 舊 DB 升級：補欄位
            cols = [r[1] for r in conn.execute("PRAGMA table_info(customers)").fetchall()]
            if "address" not in cols:
                conn.execute("ALTER TABLE customers ADD COLUMN address TEXT")
            if "ecount_cust_cd" not in cols:
                conn.execute("ALTER TABLE customers ADD COLUMN ecount_cust_cd TEXT")
            if "real_name" not in cols:
                conn.execute("ALTER TABLE customers ADD COLUMN real_name TEXT")
            # 額外電話表（一人多支）
            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_phones (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_id INTEGER NOT NULL,
                    phone       TEXT    NOT NULL,
                    UNIQUE(customer_id, phone)
                )
            """)
            # 多地址 / 多 Ecount 代碼表（一人多址）
            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_ecount_codes (
                    id             INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_id    INTEGER NOT NULL,
                    ecount_cust_cd TEXT    NOT NULL,
                    address_label  TEXT    DEFAULT '',
                    UNIQUE(customer_id, ecount_cust_cd)
                )
            """)
            # 升級：補 cust_name 欄
            ec_cols = [r[1] for r in conn.execute("PRAGMA table_info(customer_ecount_codes)").fetchall()]
            if "cust_name" not in ec_cols:
                conn.execute("ALTER TABLE customer_ecount_codes ADD COLUMN cust_name TEXT DEFAULT ''")
            # 客戶群組預設地址（LINE 群組 → 對應 Ecount 代碼）
            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_group_address (
                    group_id       TEXT PRIMARY KEY,
                    ecount_cust_cd TEXT NOT NULL,
                    label          TEXT DEFAULT ''
                )
            """)
            # 升級：個人預設地址（Du→饒河、Rachel→文山 等）
            if "preferred_ecount_cust_cd" not in cols:
                conn.execute(
                    "ALTER TABLE customers ADD COLUMN preferred_ecount_cust_cd TEXT DEFAULT NULL"
                )
            # 升級：客戶分類標籤（VIP/野獸國/標準/中句/K霸，JSON 陣列）
            if "tags" not in cols:
                conn.execute(
                    "ALTER TABLE customers ADD COLUMN tags TEXT DEFAULT NULL"
                )
            # 常用查詢欄位索引
            conn.execute("CREATE INDEX IF NOT EXISTS idx_customers_display_name ON customers(display_name)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_customers_real_name ON customers(real_name)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_customers_ecount_cust_cd ON customers(ecount_cust_cd)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_customer_phones_phone ON customer_phones(phone)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_customer_ecount_codes_cust_cd ON customer_ecount_codes(ecount_cust_cd)")
            conn.commit()

    # ── 內部工具 ──────────────────────────────────────────
    def _resolve_master(self, conn, customer_id: int) -> int:
        """若 note 含 'helper:XX'，回傳 master id；否則回傳自身 id"""
        row = conn.execute(
            "SELECT note FROM customers WHERE id=?", (customer_id,)
        ).fetchone()
        if row and row[0]:
            m = _HELPER_RE.search(row[0])
            if m:
                master_id = int(m.group(1))
                print(f"[customer] ✓ 幫手代理: id={customer_id} → master id={master_id}")
                return master_id
        return customer_id

    # ── 從 Bot 執行時更新 ─────────────────────────────────
    def upsert_from_line(self, line_user_id: str, display_name: str) -> int:
        """Bot 收到訊息時呼叫，更新 LINE ID 及顯示名稱。

        自動合併策略（依序）：
        1. line_user_id 已存在 → 更新 last_seen
        2. display_name 完全符合 + line_user_id=NULL → 自動合併（寫入 line_user_id）
        3. real_name 完全符合 + line_user_id=NULL → 自動合併
        4. 全部找不到 → 新增
        """
        now = datetime.now().isoformat()
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("BEGIN IMMEDIATE")
            # Step 1：用 line_user_id 查
            row = conn.execute(
                "SELECT id FROM customers WHERE line_user_id = ?",
                (line_user_id,)
            ).fetchone()
            if row:
                conn.execute(
                    "UPDATE customers SET display_name=?, last_seen=? WHERE line_user_id=?",
                    (display_name, now, line_user_id)
                )
                return self._resolve_master(conn, row[0])

            # Step 2：用 display_name 比對 CSV 匯入（line_user_id=NULL）
            row2 = conn.execute(
                "SELECT id FROM customers WHERE display_name=? AND line_user_id IS NULL LIMIT 1",
                (display_name,)
            ).fetchone()
            if row2:
                conn.execute(
                    """UPDATE customers
                       SET line_user_id=?, last_seen=?, first_seen=COALESCE(first_seen,?)
                       WHERE id=?""",
                    (line_user_id, now, now, row2[0])
                )
                print(f"[customer] ✓ 自動合併(display_name={display_name!r}) → id={row2[0]}")
                return self._resolve_master(conn, row2[0])

            # Step 3：用 real_name 比對（LINE 暱稱改過但我們知道真實姓名）
            row3 = conn.execute(
                "SELECT id FROM customers WHERE real_name=? AND line_user_id IS NULL LIMIT 1",
                (display_name,)
            ).fetchone()
            if row3:
                conn.execute(
                    """UPDATE customers
                       SET line_user_id=?, display_name=?, last_seen=?,
                           first_seen=COALESCE(first_seen,?)
                       WHERE id=?""",
                    (line_user_id, display_name, now, now, row3[0])
                )
                print(f"[customer] ✓ 自動合併(real_name={display_name!r}) → id={row3[0]}")
                return self._resolve_master(conn, row3[0])

            # Step 4：全找不到 → 新增
            cur = conn.execute(
                """INSERT INTO customers
                   (line_user_id, display_name, first_seen, last_seen)
                   VALUES (?, ?, ?, ?)""",
                (line_user_id, display_name, now, now)
            )
            return cur.lastrowid

    # ── 查詢 ──────────────────────────────────────────────
    def get_by_line_id(self, line_user_id: str) -> dict | None:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT * FROM customers WHERE line_user_id=?",
                (line_user_id,)
            ).fetchone()
        return dict(row) if row else None

    def update_phone(self, line_user_id: str, phone: str) -> bool:
        """更新客戶主電話（若該客戶無電話才寫入），同時補到 customer_phones"""
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                "SELECT id, phone FROM customers WHERE line_user_id=?",
                (line_user_id,)
            ).fetchone()
            if not row:
                return False
            cid = row[0]
            if not row[1]:
                conn.execute(
                    "UPDATE customers SET phone=? WHERE id=?", (phone, cid)
                )
            conn.execute(
                "INSERT OR IGNORE INTO customer_phones (customer_id, phone) VALUES (?,?)",
                (cid, phone)
            )
            conn.commit()
        return True

    def update_address(self, line_user_id: str, address: str) -> bool:
        """更新客戶住址，回傳是否找到該客戶"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET address=? WHERE line_user_id=?",
                (address, line_user_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def get_by_phone(self, phone: str) -> dict | None:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                """SELECT c.* FROM customers c
                   JOIN customer_phones p ON p.customer_id = c.id
                   WHERE p.phone=? LIMIT 1""",
                (phone,)
            ).fetchone()
        return dict(row) if row else None

    def search(self, keyword: str) -> list[dict]:
        """搜尋名稱/電話/備註/分類標籤"""
        q = f"%{keyword}%"
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """SELECT DISTINCT c.* FROM customers c
                   LEFT JOIN customer_phones p ON p.customer_id = c.id
                   WHERE c.display_name LIKE ? OR c.real_name LIKE ?
                      OR c.chat_label LIKE ? OR c.phone LIKE ?
                      OR p.phone LIKE ? OR c.note LIKE ? OR c.address LIKE ?
                      OR c.tags LIKE ?
                   ORDER BY c.last_seen DESC LIMIT 50""",
                (q, q, q, q, q, q, q, q)
            ).fetchall()
        return [dict(r) for r in rows]

    def search_by_name(self, name: str, real_name_only: bool = False) -> list[dict]:
        """
        用姓名找客戶。
        real_name_only=True：只精確比對 real_name / chat_label，不做模糊搜尋（內部群組用）
        real_name_only=False：real_name / chat_label / display_name 三欄都查，完全比對優先再模糊比對
        回傳 list（可能多筆）。
        """
        # 去除括號（如「陳怡如(彥鈞)」→「陳怡如」）後再搜尋
        name_clean = re.sub(r'[\(（][^\)）]*[\)）]', '', name).strip() or name

        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            if real_name_only:
                # 完全比對 real_name / chat_label（不含 display_name LINE 暱稱）；不做模糊搜尋
                for n in dict.fromkeys([name, name_clean]):  # 原始名優先，再試去括號版
                    rows = conn.execute(
                        """SELECT * FROM customers
                           WHERE real_name=? OR chat_label=?
                           ORDER BY CASE WHEN real_name=? THEN 0 ELSE 1 END,
                                    last_seen DESC""",
                        (n, n, n)
                    ).fetchall()
                    if rows:
                        return [dict(r) for r in rows]
                rows = []  # 找不到
            else:
                # 完全比對（real_name 優先，再 chat_label，再 display_name）
                rows = conn.execute(
                    """SELECT * FROM customers
                       WHERE real_name=? OR chat_label=? OR display_name=?
                       ORDER BY
                         CASE WHEN real_name=? THEN 0
                              WHEN chat_label=? THEN 1
                              ELSE 2 END,
                         last_seen DESC""",
                    (name, name, name, name, name)
                ).fetchall()
                if rows:
                    return [dict(r) for r in rows]
                # 模糊比對
                q = f"%{name}%"
                rows = conn.execute(
                    """SELECT * FROM customers
                       WHERE real_name LIKE ? OR chat_label LIKE ? OR display_name LIKE ?
                       ORDER BY last_seen DESC LIMIT 5""",
                    (q, q, q)
                ).fetchall()
        return [dict(r) for r in rows]

    def all(self, limit: int = 500) -> list[dict]:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM customers ORDER BY last_seen DESC LIMIT ?",
                (limit,)
            ).fetchall()
        return [dict(r) for r in rows]

    # ── 匯入用 ────────────────────────────────────────────
    def import_from_csv_data(
        self,
        display_name: str,
        chat_label: str,
        phones: list[str],
        note: str = "",
        address: str = "",
    ) -> int:
        """從 CSV 分析結果批次匯入（不覆蓋已有 LINE ID）"""
        with sqlite3.connect(DB_PATH) as conn:
            # 先找有沒有同名
            row = conn.execute(
                "SELECT id FROM customers WHERE display_name=? LIMIT 1",
                (display_name,)
            ).fetchone()
            if row:
                cid = row[0]
                conn.execute(
                    "UPDATE customers SET chat_label=?, note=? WHERE id=?",
                    (chat_label, note or None, cid)
                )
                if phones:
                    conn.execute(
                        "UPDATE customers SET phone=? WHERE id=? AND phone IS NULL",
                        (phones[0], cid)
                    )
                if address:
                    conn.execute(
                        "UPDATE customers SET address=? WHERE id=? AND address IS NULL",
                        (address, cid)
                    )
            else:
                cur = conn.execute(
                    """INSERT INTO customers
                       (display_name, chat_label, phone, address, note)
                       VALUES (?, ?, ?, ?, ?)""",
                    (display_name, chat_label, phones[0] if phones else None, address or None, note or None)
                )
                cid = cur.lastrowid

            # 所有電話存到 customer_phones
            for ph in phones:
                try:
                    conn.execute(
                        "INSERT OR IGNORE INTO customer_phones (customer_id, phone) VALUES (?,?)",
                        (cid, ph)
                    )
                except Exception:
                    pass
            conn.commit()
        return cid

    def update_ecount_cust_cd(self, line_user_id: str, ecount_cd: str) -> bool:
        """手動綁定客戶的 Ecount 客戶代碼（by LINE user_id）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET ecount_cust_cd=? WHERE line_user_id=?",
                (ecount_cd, line_user_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def update_ecount_cust_cd_by_db_id(self, db_id: int, ecount_cd: str) -> bool:
        """手動綁定客戶的 Ecount 客戶代碼（by DB primary key，用於 CSV 匯入客戶）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET ecount_cust_cd=? WHERE id=?",
                (ecount_cd, db_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def upsert_ecount_code(
        self, customer_id: int, ecount_cust_cd: str, address_label: str = ""
    ) -> None:
        """新增或更新客戶的一個 Ecount 代碼（同一客戶可多筆，對應多個送貨地址）"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("""
                INSERT INTO customer_ecount_codes (customer_id, ecount_cust_cd, address_label)
                VALUES (?, ?, ?)
                ON CONFLICT(customer_id, ecount_cust_cd)
                DO UPDATE SET address_label = excluded.address_label
            """, (customer_id, ecount_cust_cd, address_label or ""))
            conn.commit()

    def get_ecount_codes_by_line_id(self, line_user_id: str) -> list[dict]:
        """取得 LINE 用戶所有 Ecount 代碼（含地址標籤），按 id 排序"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT id FROM customers WHERE line_user_id=?", (line_user_id,)
            ).fetchone()
            if not row:
                return []
            rows = conn.execute(
                """SELECT ecount_cust_cd, address_label
                   FROM customer_ecount_codes
                   WHERE customer_id=? ORDER BY id""",
                (row[0],)
            ).fetchall()
        return [dict(r) for r in rows]

    def get_ecount_codes_by_db_id(self, db_id: int) -> list[dict]:
        """取得客戶（DB id）所有 Ecount 代碼（含地址標籤）"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                """SELECT ecount_cust_cd, address_label
                   FROM customer_ecount_codes
                   WHERE customer_id=? ORDER BY id""",
                (db_id,)
            ).fetchall()
        return [dict(r) for r in rows]

    def update_real_name(self, line_user_id: str, real_name: str) -> bool:
        """更新客戶提供的真實姓名（不覆蓋 LINE 暱稱 display_name）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET real_name=? WHERE line_user_id=?",
                (real_name.strip(), line_user_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def update_chat_label(self, line_user_id: str, label: str) -> bool:
        """手動設定客戶標籤（用於在待處理清單顯示的名稱）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET chat_label=? WHERE line_user_id=?",
                (label.strip() or None, line_user_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def update_chat_label_by_db_id(self, db_id: int, label: str) -> bool:
        """手動設定客戶標籤（by DB primary key，適用於無 LINE ID 的 CSV 匯入客戶）"""
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET chat_label=? WHERE id=?",
                (label.strip() or None, db_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def get_by_db_id(self, db_id: int) -> dict | None:
        """用 DB primary key 查詢客戶"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT * FROM customers WHERE id=?", (db_id,)
            ).fetchone()
        return dict(row) if row else None

    def update_tags_by_db_id(self, db_id: int, tags: list[str]) -> bool:
        """更新客戶分類標籤（VIP/野獸國/標準/中句/K霸），以 JSON 陣列儲存"""
        import json
        tag_json = json.dumps(tags, ensure_ascii=False) if tags else None
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute(
                "UPDATE customers SET tags=? WHERE id=?",
                (tag_json, db_id)
            )
            conn.commit()
        return cur.rowcount > 0

    def get_customers_by_tag(self, tag: str) -> list[dict]:
        """取得含有指定分類標籤的所有客戶（有 LINE ID 才能推送）"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            # 用 LIKE 搜尋 JSON 陣列中的標籤（簡單、效能尚可）
            rows = conn.execute(
                "SELECT * FROM customers WHERE tags LIKE ? AND line_user_id IS NOT NULL",
                (f'%"{tag}"%',)
            ).fetchall()
        return [dict(r) for r in rows]

    def get_ecount_cust_code(self, line_user_id: str, default: str = "LINECUST") -> str:
        """
        取得該 LINE 用戶對應的 Ecount 客戶代碼。
        優先使用手動綁定的 ecount_cust_cd；
        若無則回傳 default（通常來自 settings.ECOUNT_DEFAULT_CUST_CD）。
        """
        cust_info = self.get_by_line_id(line_user_id)
        if cust_info and cust_info.get("ecount_cust_cd"):
            return cust_info["ecount_cust_cd"]
        return default

    def sync_real_names_from_ecount(self, ecount_list: list[dict]) -> int:
        """
        用 Ecount 客戶名稱回填 real_name（只填空白欄位，不覆蓋客戶已提供的姓名）。

        Args:
            ecount_list: ecount_client.get_customers_list() 的回傳值
                         每筆格式：{"code": "CUST001", "name": "王小明", ...}

        Returns:
            int — 成功更新的筆數
        """
        updated = 0
        with sqlite3.connect(DB_PATH) as conn:
            for ec in ecount_list:
                code = (ec.get("code") or "").strip()
                name = (ec.get("name") or "").strip()
                if not code or not name:
                    continue

                # 找出對應的客戶 id（主表 ecount_cust_cd 或多地址表）
                row = conn.execute(
                    """SELECT id FROM customers WHERE ecount_cust_cd = ?
                       AND (real_name IS NULL OR real_name = '')""",
                    (code,)
                ).fetchone()

                if not row:
                    # 從多地址表找
                    row = conn.execute(
                        """SELECT c.id FROM customers c
                           JOIN customer_ecount_codes ec2 ON ec2.customer_id = c.id
                           WHERE ec2.ecount_cust_cd = ?
                             AND (c.real_name IS NULL OR c.real_name = '')
                           LIMIT 1""",
                        (code,)
                    ).fetchone()

                if row:
                    conn.execute(
                        "UPDATE customers SET real_name = ? WHERE id = ?",
                        (name, row[0])
                    )
                    updated += 1

            conn.commit()
        return updated

    def sync_ecount_names_full(self, ecount_list: list[dict]) -> dict:
        """
        完整版 Ecount 姓名同步：
        除了用 ecount_cust_cd 比對，還嘗試用電話號碼比對。
        把 Ecount 姓名存入 real_name（讓內部群組用 Ecount 姓名搜到客戶）。

        同時也把 ecount_cust_cd 回填到 customers 主表（讓後續下單用）。

        Args:
            ecount_list: [{"code": "M2509260001", "name": "王小明", "phone": "0912345678"}, ...]

        Returns:
            dict — {"by_code": N, "by_phone": N, "skipped": N}
        """
        by_code = by_phone = skipped = 0
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            for ec in ecount_list:
                code  = (ec.get("code") or "").strip()
                name  = (ec.get("name") or "").strip()
                phone = (ec.get("phone") or "").strip().replace("-", "").replace(" ", "")
                if not name:
                    skipped += 1
                    continue

                matched_id = None

                # ── 方法 1：ecount_cust_cd 直接比對 ─────────────
                if code:
                    row = conn.execute(
                        "SELECT id FROM customers WHERE ecount_cust_cd=? LIMIT 1", (code,)
                    ).fetchone()
                    if not row:
                        row = conn.execute(
                            """SELECT c.id FROM customers c
                               JOIN customer_ecount_codes ec2 ON ec2.customer_id = c.id
                               WHERE ec2.ecount_cust_cd=? LIMIT 1""",
                            (code,)
                        ).fetchone()
                    if row:
                        matched_id = row[0]
                        conn.execute(
                            "UPDATE customers SET real_name=?, ecount_cust_cd=? WHERE id=? AND (real_name IS NULL OR real_name='')",
                            (name, code, matched_id)
                        )
                        # 即使 real_name 已存在，也確保 ecount_cust_cd 有填
                        conn.execute(
                            "UPDATE customers SET ecount_cust_cd=? WHERE id=? AND (ecount_cust_cd IS NULL OR ecount_cust_cd='')",
                            (code, matched_id)
                        )
                        by_code += 1
                        continue

                # ── 方法 2：電話號碼比對 ─────────────────────────
                if phone and len(phone) >= 8:
                    # 找 customer_phones 表
                    row = conn.execute(
                        """SELECT c.id FROM customers c
                           JOIN customer_phones p ON p.customer_id = c.id
                           WHERE p.phone LIKE ?
                           LIMIT 1""",
                        (f"%{phone[-8:]}",)  # 後 8 碼比對（容錯區碼）
                    ).fetchone()
                    if not row:
                        # 找主表 phone 欄位
                        row = conn.execute(
                            "SELECT id FROM customers WHERE phone LIKE ? LIMIT 1",
                            (f"%{phone[-8:]}",)
                        ).fetchone()
                    if row:
                        matched_id = row[0]
                        conn.execute(
                            "UPDATE customers SET real_name=? WHERE id=? AND (real_name IS NULL OR real_name='')",
                            (name, matched_id)
                        )
                        if code:
                            conn.execute(
                                "UPDATE customers SET ecount_cust_cd=? WHERE id=? AND (ecount_cust_cd IS NULL OR ecount_cust_cd='')",
                                (code, matched_id)
                            )
                        by_phone += 1
                        continue

                skipped += 1

            conn.commit()
        return {"by_code": by_code, "by_phone": by_phone, "skipped": skipped}

    # ── 個人預設地址（Du→饒河、Rachel→文山 等）────────────────
    def set_preferred_address(self, customer_id: int, ecount_cust_cd: str | None) -> None:
        """
        設定某客戶的個人預設訂單地址。
        叫貨結帳時會直接問「是否送到 XX？」而非列出全部地址。
        傳 None 或空字串表示清除。
        """
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                "UPDATE customers SET preferred_ecount_cust_cd=? WHERE id=?",
                (ecount_cust_cd or None, customer_id),
            )
            conn.commit()

    def get_preferred_address(self, customer_id: int) -> str | None:
        """取得客戶的個人預設 Ecount 代碼（無則回 None）"""
        with sqlite3.connect(DB_PATH) as conn:
            row = conn.execute(
                "SELECT preferred_ecount_cust_cd FROM customers WHERE id=?",
                (customer_id,),
            ).fetchone()
        return row[0] if row else None

    # ── 客戶群組預設地址 ──────────────────────────────────
    def get_group_default(self, group_id: str) -> dict | None:
        """
        查詢 LINE 群組對應的預設 Ecount 代碼。
        回傳 {"ecount_cust_cd": ..., "label": ...} 或 None。
        """
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT ecount_cust_cd, label FROM customer_group_address WHERE group_id=?",
                (group_id,)
            ).fetchone()
        return dict(row) if row else None

    def set_group_address(
        self, group_id: str, ecount_cust_cd: str, label: str = ""
    ) -> None:
        """
        登記（或更新）LINE 群組對應的預設 Ecount 代碼。
        用法：bot 收到群組訊息後，透過 /admin/set-group-address 登記。
        """
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("""
                INSERT INTO customer_group_address (group_id, ecount_cust_cd, label)
                VALUES (?, ?, ?)
                ON CONFLICT(group_id)
                DO UPDATE SET ecount_cust_cd = excluded.ecount_cust_cd,
                              label = excluded.label
            """, (group_id, ecount_cust_cd, label))
            conn.commit()

    def list_group_addresses(self) -> list[dict]:
        """回傳所有已登記的群組預設地址"""
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT group_id, ecount_cust_cd, label FROM customer_group_address ORDER BY group_id"
            ).fetchall()
        return [dict(r) for r in rows]

    def count(self) -> int:
        with sqlite3.connect(DB_PATH) as conn:
            return conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]


customer_store = CustomerStore()
