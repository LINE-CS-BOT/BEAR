# 小蠻牛 LINE 客服機器人 — 除錯紀錄

> 最後更新：2026-03-18

---

## 2026-03-18 修正記錄

### BUG-01｜Ecount API 回應 Big5 編碼解析失敗
- **問題**：`lookup()` 回傳 null，內部群庫存查詢無回應
- **原因**：Ecount OAPI v2 部分回應為 Big5/GBK 編碼，直接 `.json()` 解析失敗
- **修正**：`services/ecount.py` 新增 `_safe_json()`，依序嘗試 UTF-8 → Big5 → GBK → GB18030
- **檔案**：`services/ecount.py`

---

### BUG-02｜新建品項在 Ecount API 失敗時仍存入 DB
- **問題**：Ecount API 回傳失敗（result=None）時，品項仍被存入 `new_products.db`
- **修正**：`_build_one_product()` 改為 `if result:` 才呼叫 `new_products_store.add()`
- **檔案**：`handlers/internal.py`

---

### BUG-03｜`_PROD_CODE_RE` 正則 `\b` 邊界在中文環境失效
- **問題**：`庫存Z3207` 無法正確擷取貨號（`\b` 不支援中文/英文邊界）
- **原因**：Python regex `\b` 僅判斷 `\w/\W` 邊界，中文字與英文字之間無效
- **修正**：改為 `(?<![A-Za-z])` 負向回顧斷言，並確保內部群庫存查詢不論庫存數量均顯示完整格式
- **檔案**：`handlers/internal.py`

---

### BUG-04｜新建品項完成訊息誤啟動 30 秒 session
- **問題**：「新建品項 Z3579...」成功完成後，bot 仍啟動 30 秒確認 session 等待回覆
- **原因**：`_build_new_product()` 流程結束後未提早 return，仍走入一般指令處理
- **修正**：新品建立成功後立即回覆並 return，不進入等待狀態
- **檔案**：`handlers/internal.py`

---

### BUG-05｜`(大)` CLASS_CD 標籤顯示錯誤
- **問題**：CLASS_CD `00002` 對應標籤顯示為「盒」，應為「改裝」
- **修正**：
  - `handlers/internal.py`：`_CLASS_LABEL_NP['00002']` 改為 `'改裝'`
  - `static/admin.html`：`CLASS_LABEL['00002']` 改為 `'改裝'`
- **檔案**：`handlers/internal.py`, `static/admin.html`

---

### BUG-06｜Format B 代訂備註未傳入 `_do_order()`
- **問題**：多行格式「鄭鉅耀 訂\nZ3340 10個\n備註 送松山」建立的訂單沒有備註
- **原因**：Format B 解析時擷取了 `note_b` 但未傳入 `_do_order(note=note_b)`
- **修正**：`_do_order(cust_name, items_raw, note=note_b)` 補上 note 參數
- **檔案**：`handlers/internal.py`

---

### BUG-07｜Format B2 純姓名首行支援 + 多品項備註
- **問題**：格式「鄭鉅耀\nZ3340 10個\nZ3338 20個\n備註 送松山」無回應
- **原因**：Format B2 未實作（只有 Format B 含「訂」關鍵字）
- **修正**：新增 Format B2 判斷：第一行無貨號、無「訂」關鍵字時，視後續行為品項清單
- **檔案**：`handlers/internal.py`

---

### BUG-08｜新建品項全形括號未識別（Z3579）
- **問題**：`（原定）多色麥克風音響` 未識別 CLASS_CD，品名解析錯誤
- **原因**：`_CLASS_CD_MAP` 僅支援半形括號 `(原定)`
- **修正**：`_CLASS_CD_MAP` 正則改為 `[（(]原定[)）]` 支援全形/半形括號
- **檔案**：`handlers/internal.py`

---

### BUG-09｜售價含 `` 前綴未解析（Z3579）
- **問題**：售價欄位格式「售價:$299」中的 `$` 導致數字無法擷取
- **修正**：售價/入庫價正則加入 `[$＄]?` 允許貨幣符號前綴
- **檔案**：`handlers/internal.py`

---

### BUG-10｜`"明天看有沒有空過去載"` 意圖誤判
- **問題**：該句應為一般閒聊/未知，但被誤判為：
  1. `INVENTORY`：因含「有沒有」
  2. `VISIT_STORE`：因含「過去」
- **修正一**：Intent 偵測新增 `_INV_EXCLUDE` 排除清單，包含「有沒有空」「有沒有時間」等非庫存語境
- **修正二**：`handlers/visit.py` 的 `VISIT_KEYWORDS` 移除過於寬泛的 `"過去"`，保留 `"過去拿"` `"過去看"` 等複合詞
- **檔案**：`handlers/intent.py`, `handlers/visit.py`

---

### BUG-11｜tray.py 使用 `DEVNULL` 丟棄所有 uvicorn log
- **問題**：uvicorn 啟動後所有輸出（包括錯誤）被丟棄，無法排查問題
- **修正**：改用 shell `>>` 重定向將 stdout/stderr 追加到 `server.log`
  ```python
  cmd_str = " ".join(f'"{c}"' if " " in c else c for c in UVICORN_CMD)
  p = subprocess.Popen(
      f'{cmd_str} >> "{log_path}" 2>&1',
      cwd=BASE_DIR, shell=True, creationflags=CREATE_NO_WINDOW,
  )
  ```
- **檔案**：`tray.py`

---

### BUG-12｜tray.py `BASE_DIR / "server.log"` 型別錯誤
- **問題**：`BASE_DIR` 為 `str`，不支援 `/` 運算子
- **修正**：改用 `os.path.join(BASE_DIR, "server.log")`
- **檔案**：`tray.py`

---

### BUG-13｜tray.py 持有 log 檔案句柄導致 uvicorn 重啟失敗
- **問題**：tray 進程開啟 log 檔案後持有 file handle，uvicorn 重啟時無法寫入同一檔案而崩潰
- **修正**：改用 shell `>>` 重定向（由 OS shell 管理 fd），tray 進程不持有任何 log 句柄
- **檔案**：`tray.py`

---

### BUG-14｜空白客戶記錄（display_name/real_name/phone 全空）
- **問題**：群組訊息觸發時，即使無法取得使用者資料，仍呼叫 `customer_store.upsert_from_line()` 建立空記錄
- **修正**：在兩處呼叫點加入 `if display_name:` 防衛判斷
- **清除**：一次性 SQL 刪除 4 筆空記錄（id=287, 298, 300, 301）
- **檔案**：`main.py`

---

### BUG-15｜群組訊息 log 缺少使用者識別資訊
- **問題**：`data/group_ids.txt` 只記錄群組 ID，無法判斷說話的人是誰
- **修正**：改為記錄格式 `MM-DD HH:MM | group_id | user_id | display_name`
  - 每次收到群組訊息時，呼叫 `get_group_member_profile()` 取得 display_name
- **檔案**：`main.py`

---

## 功能新增記錄

### FEAT-01｜7 類新意圖支援（計劃中）
- **意圖**：`BARGAINING`, `SPEC`, `RETURN`, `MULTI_PRODUCT`, `ADDRESS_CHANGE`, `COMPLAINT`, `URGENT_ORDER`
- **狀態**：Intent enum 與關鍵字已加入 `handlers/intent.py`，handler 邏輯在 `handlers/service.py`
- **DB**：`storage/issues.py` → `issues.db`（退換貨/投訴/地址更改）

### FEAT-02｜Format B2 多行下單（無「訂」關鍵字）
- 格式：`鄭鉅耀\nZ3340 10個\nZ3338 20個\n備註 送松山`
- **狀態**：✅ 已實作

### FEAT-03｜箱/件換算（Format A/B/B2/C）
- unit=箱/件 → `resolve_order_qty(prod_cd, qty)` 乘以 `EXCH_RATE`
- unit=個/其他 → 不換算
- **狀態**：✅ 已實作

### FEAT-04｜貨架標籤 PDF（佇列式，3個一張）
- 指令：`貨架標籤 Z2095`（內部群）
- 命名：`架上標_{YYYYMMDD}_{CODE1}_{CODE2}_{CODE3}.pdf`
- **狀態**：✅ 已實作

### FEAT-05｜總公司群組調貨完整流程
- 客戶缺貨 → 問數量 → 通知 HQ → HQ 回「有貨/叫貨N週」→ 建單/問等
- **狀態**：✅ 已實作

### FEAT-06｜Admin 後台面板順序
1. 📅 預計到店客人
2. 👤 人工接手面板
3. 處理紀錄
4. 🆕 待審核新品項
- **狀態**：✅ 已確認

### FEAT-07｜HTTPS（DuckDNS + Caddy，取代 ngrok）
- 域名：`xmnline.duckdns.org`
- Webhook：`https://xmnline.duckdns.org/webhook`
- 路由：Router 80/443 → 192.168.10.170 → Caddy → localhost:8000
- **狀態**：✅ 已上線

### FEAT-08｜系統架構文件
- 新增 `ARCHITECTURE.md`：完整架構明細、群組 ID、Ecount API 端點、CLASS_CD 對應
- 新增 `BUGFIX_LOG.md`（本檔）：除錯記錄
- **狀態**：✅ 已建立

---

## 待修正清單

| ID | 項目 | 優先度 |
|----|------|--------|
| P-01 | `_followup_loop` 中 `line_api` 未在 scope（定時跟進可能無法推訊息）| 高 |
| P-02 | 開機自動啟動仍指向 `start.bat`，應改為 `start_tray.bat` | 中 |
| P-03 | `customer_group_address` 表已建但資料空（客戶多地址功能待完成）| 低 |
| P-04 | Google Calendar 整合未完成（`services/google_cal.py`）| 低 |
| P-05 | 模糊選品後 `pending_ambiguous_resolve` state 直接建單（目前多一次確認）| 低 |
