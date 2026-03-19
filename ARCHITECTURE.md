# 小蠻牛 LINE 客服機器人 — 架構明細

> 最後更新：2026-03-18

---

## 系統概覽

| 項目 | 規格 |
|------|------|
| 框架 | FastAPI + uvicorn |
| 入口 | `POST /webhook`（LINE Signature 驗證）|
| 常駐 | `tray.py`（系統匣圖示 + watchdog 5秒自動重啟）|
| 資料庫 | SQLite × 多個（data/ 目錄）|
| ERP | Ecount OAPI v2（正式區 IB，`oapiIB.ecount.com`）|
| HTTPS | Caddy 反向代理（port 80/443 → localhost:8000）|
| 域名 | DuckDNS（`xmnline.duckdns.org`，動態 IP 同步）|
| Webhook | `https://xmnline.duckdns.org/webhook` |

---

## 對話群組

| 環境變數 | Group ID | 說明 |
|---------|---------|------|
| `LINE_GROUP_ID` | `Cbe854062792d1177f44446b0835ea4dc` | 內部群組（員工操作）|
| `LINE_GROUP_ID_HQ` | `Cdef8491598e17fa9fb2fa96a3d9bbff7` | 總公司群（調貨/新品）|
| `LINE_GROUP_ID_SHOWCASE` | `C6494c5b991b61f05a7309f87fe8702dd` | 新品看貨群 |
| 未設定 | `Ce4433a39d70eea301135c0a6357d8320` | 待確認 |
| 未設定 | `C4c6032278ec52776fc77b17fd40c433e` | 待確認 |

---

## 檔案結構

### 核心

| 檔案 | 行數 | 說明 |
|------|------|------|
| `main.py` | 3003 | FastAPI webhook + lifespan scheduler + admin API |
| `config.py` | 45 | pydantic-settings 環境變數 |
| `tray.py` | 237 | 系統匣常駐 + watchdog 自動重啟 uvicorn + Caddy |

### handlers/（業務邏輯）

| 檔案 | 行數 | 說明 |
|------|------|------|
| `intent.py` | 241 | 意圖偵測（22種意圖）|
| `internal.py` | 2658 | 內部群/總部群完整邏輯（代訂/庫存/新品/上架）|
| `inventory.py` | 207 | 客戶端庫存查詢 |
| `ordering.py` | 447 | 客戶端下單流程 + 箱件換算 |
| `restock.py` | 151 | HQ 群組調貨回覆處理 |
| `summary.py` | 389 | 待處理清單（內部群「清單」指令）|
| `hours.py` | 106 | 營業時間（支援日期查詢）|
| `delivery.py` | 40 | 配送時程 |
| `escalate.py` | 33 | 轉真人客服 |
| `tone.py` | 661 | 真人語氣模擬（269則對話分析）|
| `service.py` | 358 | 砍價/規格/退換貨/地址更改/投訴 |
| `visit.py` | 184 | 到店預約（客戶/內部群查詢）|
| `followup.py` | 56 | 定時跟進（24h提醒/48h清除）|
| `price.py` | 51 | 價格查詢 |
| `payment.py` | 28 | 付款處理 |
| `orders.py` | 18 | 訂單追蹤記錄 |

### services/（外部服務）

| 檔案 | 行數 | 說明 |
|------|------|------|
| `ecount.py` | 797 | Ecount OAPI v2 client（含 Big5 編碼處理）|
| `vision.py` | 373 | 圖片識別（pHash + OCR）|
| `refresh.py` | 188 | 資料庫定時刷新 |
| `google_cal.py` | 124 | Google Calendar 整合（待完成）|
| `inventory_csv.py` | 54 | 庫存 CSV 匯入 |

### storage/（資料存取層）

| 檔案 | 行數 | SQLite DB | 說明 |
|------|------|-----------|------|
| `customers.py` | 680 | `customers.db` | 客戶資料（264筆）|
| `state.py` | 128 | in-memory | 對話狀態（暫存）|
| `persistent_state.py` | 126 | `persistent_state.db` | 對話狀態持久化（重啟後恢復）|
| `restock.py` | 104 | `restock_requests.db` | 調貨請求狀態 |
| `new_products.py` | 92 | `new_products.db` | 待審核新品項 |
| `pending.py` | 67 | `pending_queries.db` | 待確認查詢 |
| `issues.py` | 103 | `issues.db` | 退換貨/投訴/地址變更 |
| `visits.py` | 78 | `visits.db` | 到店預約記錄 |
| `specs.py` | 99 | `specs.json` | 產品規格 |
| `reserved.py` | 99 | in-memory | 保留庫存 |
| `payments.py` | 79 | `payments.db` | 付款記錄 |
| `queue.py` | 77 | `queue.db` | 離峰訊息佇列 |
| `cart.py` | 41 | in-memory | 購物車 |
| `notify.py` | 100 | `notify.db` | 通知記錄 |
| `tags_config.py` | 55 | `tags_config.json` | 標籤設定 |

### scripts/（工具腳本）

| 檔案 | 行數 | 說明 |
|------|------|------|
| `generate_shelf_label.py` | 285 | 貨架標籤 PDF（每3個一張）|
| `auto_sync_unfulfilled.py` | 485 | 庫存同步（Playwright + Chrome 爬 Ecount）|
| `sync_cust_from_web.py` | 892 | 從 Ecount 網頁同步客戶資料 |
| `sync_cust_ecount.py` | 230 | Ecount 客戶 API 同步 |
| `import_customers.py` | 153 | 從 LINE OA CSV/ZIP 批次匯入客戶 |
| `import_specs.py` | 174 | 匯入產品規格 |
| `build_image_hashes.py` | 101 | 建立圖片 pHash 索引 |

### 靜態資源

| 路徑 | 說明 |
|------|------|
| `static/admin.html` | Admin 後台單頁應用 |
| `static/products/{CODE}/` | 產品圖片 |
| `data/available.json` | 庫存快取（定時同步）|
| `data/ecount_customers.json` | Ecount 客戶清單快取 |
| `data/image_hashes.json` | 圖片 pHash 索引 |
| `data/label_queue.json` | 貨架標籤佇列（未滿3個暫存）|
| `data/specs.json` | 產品規格快取 |
| `data/group_ids.txt` | 群組訊息記錄（含 user_id + 姓名）|

---

## Ecount API 端點

| 功能 | 端點 |
|------|------|
| 登入 | `/OAPI/V2/OAPILogin` |
| 品項清單 | `/OAPI/V2/InventoryBasic/GetBasicProductsList` |
| 新建品項 | `/OAPI/V2/InventoryBasic/SaveBasicProduct` |
| 庫存查詢 | `/OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus` |
| 建立訂單 | `/OAPI/V2/SaleOrder/SaveSaleOrder` |

- SESSION_ID 放 query string
- 回應編碼：自動偵測 UTF-8 / Big5 / GBK / GB18030
- 品項快取：6小時更新，排除 Z+英文開頭貨號

---

## Admin 後台面板順序

1. 📅 預計到店客人
2. 👤 人工接手面板
3. 處理紀錄
4. 🆕 待審核新品項

---

## 啟動方式

```
start_tray.bat  →  pythonw tray.py
                     ├── uvicorn main:app --reload (port 8000)
                     └── caddy.exe (port 80/443 → 8000)
```

開機自動啟動：排程器執行 `start_tray.bat`（應確認已設定）

---

## 新建品項 CLASS_CD 對應

| 前綴（半形/全形）| CLASS_CD | 標籤 | 入庫計算 |
|----------------|---------|------|---------|
| `(原)` / `（原）` | 00001 | 原裝 | 售價 × 0.95 |
| `(大)` / `（大）` | 00002 | 改裝 | 售價 × 0.85 |
| `(原定)` / `（原定）` | 00004 | 定裝 | 加盟商價 |
| `(定)` / `（定）` | 00004 | 定裝 | 加盟商價 |
| 無前綴 | — | — | 加盟商價；無則空 |

---

## 內部群代訂格式

| 格式 | 範例 |
|------|------|
| A 單行 | `楊庭瑋 訂 Z2095 30個` |
| B 多行有「訂」 | `鄭鉅耀 訂` + `Z3340 10個` + `備註 送松山` |
| B2 多行純姓名 | `鄭鉅耀` + `Z3340 10個` + `備註 送松山` |
| C 直接 | `方力緯 Z3562 5個` |
| D 品名單品 | `曹竣智 要 洗衣球 5` |
| E 品名多品 | `楊庭瑋 衛生紙30箱 泡澡球10件` |

備註支援：`備註` / `備誌` / `備记` + 空格或冒號
