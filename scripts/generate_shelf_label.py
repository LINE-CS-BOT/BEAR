"""
上架標籤 PDF 生成腳本

規則：
  - 所有上架產品都納入（不限台型）
  - 每次上架後將產品加入持久佇列（data/label_queue.json）
  - 佇列滿 3 個即自動生成一張 A4 PDF，剩餘繼續等候下次上架
  - 存到 H:\\其他電腦\\我的電腦\\小蠻牛\\架上標\\架上標_YYYYMMDD_HHMMSS.pdf

欄位對應：
  設計名稱 → Ecount 品名
  商品編號 → 貨號
  設計尺寸 → specs.json size
  倉儲編號 → P + 售價數字（例：P109）
  備註     → specs.json weight

用法：
  python scripts/generate_shelf_label.py T1102 T1201 Q0314
  （也可由 main.py 在上架完成後自動呼叫）
"""

import json
import re
import sys
import io
from datetime import datetime
from pathlib import Path

# ── 路徑設定 ──────────────────────────────────────────────────
_ROOT        = Path(__file__).parent.parent
TEMPLATE_PDF = Path(r"H:\其他電腦\我的電腦\小蠻牛\空白文件及標示\小蠻牛架上合併標.pdf")
OUTPUT_DIR   = Path(r"H:\其他電腦\我的電腦\小蠻牛\架上標")
QUEUE_FILE   = _ROOT / "data" / "label_queue.json"

# ── A4 PDF 座標系（pt，左下角為原點）──────────────────────────
#   Block 1（上）: y_bottom=545
#   Block 2（中）: y_bottom=273
#   Block 3（下）: y_bottom=1
_BLOCK_Y_BOTTOMS = [545.0, 273.0, 1.0]

# 各欄位相對 block 底部的文字基線 y（從下往上量，實測微調確認）
_ROW_BASELINE_REL = {
    "設計名稱": 250,
    "商品編號": 203,
    "設計尺寸": 160,
    "倉儲編號": 117,
    "備註":      62,
}

# 文字起始 x 與最大寬度
_TEXT_X      = 190.0
_TEXT_MAX_W  = 260   # 超過自動縮小字型


# ── 佇列管理（檔案鎖保護，防止並發覆蓋）─────────────────────
import threading
_queue_lock = threading.Lock()


def _load_queue() -> list[dict]:
    if QUEUE_FILE.exists():
        try:
            return json.loads(QUEUE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _save_queue(queue: list[dict]):
    QUEUE_FILE.parent.mkdir(parents=True, exist_ok=True)
    QUEUE_FILE.write_text(
        json.dumps(queue, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ── 資料處理 ──────────────────────────────────────────────────

def _load_specs() -> dict:
    sp = _ROOT / "data" / "specs.json"
    if sp.exists():
        return json.loads(sp.read_text(encoding="utf-8"))
    return {}


def _ecount_name(code: str) -> str:
    """從 Ecount 取品名（同步）"""
    try:
        sys.path.insert(0, str(_ROOT))
        from services.ecount import ecount_client
        info = ecount_client.lookup(code)
        return (info.get("name") or "").strip() if info else ""
    except Exception:
        return ""


def _parse_price(price_str: str) -> str:
    """'109元' → 'P109'"""
    m = re.search(r"\d+", price_str or "")
    return f"P{m.group()}" if m else ""


def _strip_unit(s: str) -> str:
    """去掉「約：」前綴"""
    return re.sub(r"^約[：:]?\s*", "", (s or "").strip())


def _notify_missing_specs(codes: list[str]):
    """規格缺失時推通知到內部群"""
    try:
        sys.path.insert(0, str(_ROOT))
        from config import settings
        from linebot.v3.messaging import (
            ApiClient, Configuration, MessagingApi, PushMessageRequest, TextMessage
        )
        group_id = settings.line_group_id
        if not group_id:
            return
        text = (
            "⚠️ 架上標籤規格缺失\n"
            f"以下貨號找不到規格資料，標籤未生成：\n"
            + "\n".join(f"• {c}" for c in codes)
            + "\n\n請補上 PO文（含尺寸/重量/價格）後重新上架。"
        )
        cfg = Configuration(access_token=settings.line_channel_access_token)
        with ApiClient(cfg) as client:
            MessagingApi(client).push_message(
                PushMessageRequest(to=group_id, messages=[TextMessage(text=text)])
            )
        print(f"[label] 規格缺失通知已送出：{codes}")
    except Exception as e:
        print(f"[label] 規格缺失通知失敗：{e}")


def _build_product_data(code: str, specs: dict) -> dict | None:
    """組合一個產品的欄位資料；找不到規格或品名則回傳 None"""
    spec = specs.get(code.upper())
    if not spec:
        print(f"[label] 找不到 {code} 的規格資料，跳過")
        return None

    # 從 Ecount 取品名和出庫單價
    try:
        sys.path.insert(0, str(_ROOT))
        from services.ecount import ecount_client
        ecount_client._ensure_product_cache()
        cache = ecount_client.get_product_cache_item(code.upper())
    except Exception:
        cache = None

    name = (cache.get("name") or "").strip() if cache else ""
    if not name:
        print(f"[label] {code} 在 Ecount 無品名，跳過（請先新增品項）")
        return None

    # 價格優先用 Ecount 出庫單價，沒有才用 specs.json
    ecount_price = cache.get("price") if cache else None
    if ecount_price and float(ecount_price) > 0:
        price = f"P{int(float(ecount_price))}"
    else:
        price = _parse_price(spec.get("price", ""))

    size   = _strip_unit(spec.get("size", ""))
    weight = _strip_unit(spec.get("weight", ""))

    return {
        "設計名稱": name,
        "商品編號": code,
        "設計尺寸": size,
        "倉儲編號": price,
        "備註":     weight,
    }


# ── PDF 生成 ───────────────────────────────────────────────────

def _generate_one_pdf(products: list[dict], output_path: Path):
    """將 3 個產品資料填入模板，生成一張 A4 PDF。"""
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from pypdf import PdfReader, PdfWriter

    # 載入中文字型
    font_name = "MsYaHei"
    for fp in [
        Path(r"C:\Windows\Fonts\msjh.ttc"),
        Path(r"C:\Windows\Fonts\mingliu.ttc"),
        Path(r"C:\Windows\Fonts\kaiu.ttf"),
    ]:
        if fp.exists():
            try:
                pdfmetrics.registerFont(TTFont(font_name, str(fp)))
                break
            except Exception:
                continue
    else:
        font_name = "Helvetica"

    # 建立 overlay（透明背景，只有文字）
    overlay_buf = io.BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=A4)
    c.setFillColorRGB(0, 0, 0)

    for i, prod in enumerate(products[:3]):
        y_bot = _BLOCK_Y_BOTTOMS[i]
        for field, rel_y in _ROW_BASELINE_REL.items():
            val = prod.get(field, "")
            if not val:
                continue
            font_size = 20.0
            while font_size > 8:
                c.setFont(font_name, font_size)
                if c.stringWidth(val, font_name, font_size) <= _TEXT_MAX_W:
                    break
                font_size -= 0.5
            c.setFont(font_name, font_size)
            c.drawString(_TEXT_X, y_bot + rel_y, val)

    c.save()
    overlay_buf.seek(0)

    # 疊加到模板
    template = PdfReader(str(TEMPLATE_PDF))
    overlay  = PdfReader(overlay_buf)
    writer   = PdfWriter()
    page = template.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(str(output_path), "wb") as f:
        writer.write(f)

    print(f"[label] 已生成：{output_path}")


# ── 主入口 ─────────────────────────────────────────────────────

def generate_labels(codes: list[str]) -> list[Path]:
    """
    傳入本次上架的貨號清單，加入持久佇列，
    佇列每滿 3 個自動生成一張 PDF，剩餘保留至下次。
    回傳本次生成的檔案路徑清單。
    """
    specs = _load_specs()

    # 組合本次上架產品資料
    new_products = []
    missing_codes = []
    for code in codes:
        d = _build_product_data(code.upper(), specs)
        if d:
            new_products.append(d)
        else:
            missing_codes.append(code.upper())

    # 有規格缺失 → 記錄（由呼叫方在回覆訊息中顯示）
    if missing_codes:
        print(f"[label] 規格缺失，跳過：{missing_codes}")

    if not new_products:
        print("[label] 本次上架無法取得規格資料，不更新佇列")
        return []

    with _queue_lock:
        # 合併至佇列（佇列內同貨號 → 取代為最新資料；已印過的不影響）
        queue = _load_queue()
        existing_idx = {p["商品編號"]: i for i, p in enumerate(queue)}
        replaced = 0
        for p in new_products:
            code = p["商品編號"]
            if code in existing_idx:
                queue[existing_idx[code]] = p
                replaced += 1
            else:
                queue.append(p)
                existing_idx[code] = len(queue) - 1
        added = len(new_products) - replaced
        print(f"[label] 佇列現有 {len(queue)} 個（新增 {added} 個，更新 {replaced} 個）")

        if len(queue) < 3:
            _save_queue(queue)
            print(f"[label] 尚不足 3 個，暫存等候下次上架")
            return []

        # 每 3 個生成一張 PDF
        ts = datetime.now().strftime("%Y%m%d")
        generated = []
        while len(queue) >= 3:
            batch = queue[:3]
            queue = queue[3:]
            codes_str = "_".join(p["商品編號"] for p in batch)
            out_path = OUTPUT_DIR / f"架上標_{ts}_{codes_str}.pdf"
            _generate_one_pdf(batch, out_path)
            generated.append(out_path)

        # 剩餘存回佇列
        _save_queue(queue)
        if queue:
            print(f"[label] 剩餘 {len(queue)} 個暫存至佇列，下次上架再湊")

    return generated


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python generate_shelf_label.py T1102 T1201 Q0314 ...")
        sys.exit(1)
    paths = generate_labels(sys.argv[1:])
    if paths:
        print(f"✅ 共生成 {len(paths)} 張 PDF：")
        for p in paths:
            print(f"  {p}")
    else:
        print("未生成任何 PDF（規格缺失或佇列未滿 3 個）")
