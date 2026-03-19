"""
廣告圖生成模組

觸發：內部群組 `廣告圖更新` 或 Admin 介面按鈕
流程：
  1. 掃描 AD_IMAGE_DIR（H:\...\廣告圖\產品圖）→ 從檔名抓貨號
  2. 每個貨號：Ecount 品名 + PO文 + 聯繫資料 → 組提示詞
  3. Playwright 開 Gemini → 上傳產品圖 + 貼提示詞 → 生成圖
  4. 儲存 Z3432_line.png / Z3432_fb.png 到 AD_OUTPUT_DIR
  5. 通知內部群
"""

import re
import subprocess
import sys
from pathlib import Path

from storage.state import state_manager
from services.ecount import ecount_client
import storage.specs as spec_store
from config import settings

# ── 固定路徑 ─────────────────────────────────────────────────────────────────
AD_IMAGE_DIR  = Path(r"H:\其他電腦\我的電腦\小蠻牛\廣告圖\產品圖")
AD_OUTPUT_DIR = Path(r"H:\其他電腦\我的電腦\小蠻牛\廣告圖\廣告圖")

# ── 廣告品牌聯繫資訊 ────────────────────────────────────────────────────────
STORE_NAME   = "小蠻牛新北旗艦店"
CONTACT_LINE = "LINE ID: @774jucyh"

# ── 支援的圖片副檔名 ─────────────────────────────────────────────────────────
_IMG_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}

# ── 從檔名抓貨號 ─────────────────────────────────────────────────────────────
_FILENAME_CODE_RE = re.compile(r'([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)', re.IGNORECASE)

# ── 觸發詞 ────────────────────────────────────────────────────────────────────
_AD_TRIGGER_WORDS = {"廣告圖更新", "廣告圖 更新", "廣告更新"}
_AD_TEST_RE       = re.compile(r'^廣告圖測試\s+([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)$')


# ── 工具函式 ─────────────────────────────────────────────────────────────────

def _extract_code_from_filename(filename: str) -> str | None:
    """從檔名抓產品代碼，例 Z3432A.jpg → Z3432"""
    stem = Path(filename).stem
    m = _FILENAME_CODE_RE.search(stem)
    return m.group(1).upper() if m else None


def _get_product_name(prod_code: str) -> str:
    """取得產品名稱（Ecount 快取 > specs > 原代碼）"""
    item = ecount_client.get_product_cache_item(prod_code)
    if item and item.get("PROD_NM"):
        return item["PROD_NM"]
    spec = spec_store.get_by_code(prod_code)
    if spec and spec.get("name"):
        return spec["name"]
    return prod_code


def _get_po_summary(prod_code: str) -> str:
    """從 PO文.txt 或 specs 取簡短描述（≤ 150 字）"""
    try:
        from handlers.internal import _get_raw_po_block
        raw = _get_raw_po_block(prod_code)
        if raw:
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
            lines = [l for l in lines if not re.match(r'^[A-Z]{1,3}\d{3,6}$', l)]
            return "\n".join(lines)[:150]
    except Exception:
        pass

    spec = spec_store.get_by_code(prod_code)
    if spec:
        parts = [str(spec[k]) for k in ("name", "desc", "spec", "size") if spec.get(k)]
        return "\n".join(parts)[:150]

    item = ecount_client.get_product_cache_item(prod_code)
    if item:
        return item.get("PROD_NM", prod_code)

    return prod_code


def build_gemini_prompt_with_claude(
    prod_code: str,
    platform: str,
    image_paths: list[str],
) -> str:
    """
    用本機 Claude Code CLI (claude --print) 當子代理：
      - 把 H:\ 產品圖複製到 temp，傳路徑給 CLI
      - Claude 用 Read tool 讀圖（有圖片視覺）
      - system-prompt 內嵌 ad-design + ad-style-optimizer 核心邏輯
      - stdout 就是 Gemini 提示詞
    失敗時 fallback 到 build_gemini_prompt()。
    """
    import subprocess, shutil as _sh, tempfile

    prod_name = _get_product_name(prod_code)
    po_text   = _get_po_summary(prod_code)

    if platform == "line":
        size_desc   = "1040×1040px，1:1 正方形，LINE 官方帳號貼文廣告"
        layout_hint = "中央聚焦型：品牌名（上12%）→ 產品主視覺（中52%）→ 商品名+貨號（中下26%）→ LINE ID（下10%）"
    else:
        size_desc   = "1200×630px，1.91:1 橫式，Facebook 動態貼文廣告"
        layout_hint = "左右分割型：左側產品圖（46%），右側垂直排列 品牌名→橘線→商品名→貨號→賣點1-2條→LINE ID"

    # ── 把 H:\ 圖片複製到本機 temp ────────────────────────────────────────────
    local_paths: list[str] = []
    for img_path in image_paths[:3]:
        p = Path(img_path)
        if not p.exists():
            continue
        if p.drive.upper() not in ("C:", "D:", "E:"):
            try:
                tmp = Path(tempfile.gettempdir()) / p.name
                _sh.copy2(str(p), str(tmp))
                local_paths.append(str(tmp))
                print(f"[claude-prompt] 複製到本機：{tmp.name}")
            except Exception as e:
                print(f"[claude-prompt] 複製失敗（{p.name}）：{e}")
        else:
            local_paths.append(str(p))

    if not local_paths:
        print("[claude-prompt] 無可用圖片，使用預設模板")
        return build_gemini_prompt(prod_code, platform)

    paths_str = "\n".join(local_paths)

    _SYSTEM = (
        "你是一位資深廣告視覺設計師，同時精通廣告圖片設計與廣告風格優化。\n\n"
        "【廣告圖片設計原則】\n"
        "- 每張廣告在固定尺寸內完成一個任務，前 0.5 秒抓注意力\n"
        "- 資訊極度精簡：品牌名 + 商品名 + 貨號 + 聯繫方式\n"
        "- 視覺動線明確：品牌 → 產品 → 商品名 → 聯繫\n\n"
        "【廣告風格：極簡高端電商風】\n"
        "- 配色：主色 #FFFFFF 白、輔色 #F0F0F0 淺灰、強調色 #FF6B00 橘、LINE綠 #06C755\n"
        "- 背景：白色到極淺灰漸層，乾淨無雜訊，四邊 5% 安全邊距\n"
        "- 字體：無襯線黑體，主標極粗，輔助文字細且間距寬\n"
        "- 產品圖：精確去背，底部橢圓投影（#CCCCCC，模糊16px，橫向壓扁），懸浮感\n"
        "- 裝飾：橘色品牌細線（36px×2.5px，圓角），其他地方不放裝飾\n"
        "- 整體：乾淨、有呼吸感、精品電商質感\n\n"
        "【輸出規則】\n"
        "只輸出 Gemini 製圖提示詞，繁體中文，從「極簡高端電商風格，」開頭，"
        "400-800字，不要前言或標題。"
    )

    _USER = (
        f"請用 Read 工具讀取以下產品照片，觀察產品外觀後生成廣告製圖提示詞。\n\n"
        f"產品照片路徑：\n{paths_str}\n\n"
        f"產品資訊：\n"
        f"品牌：{STORE_NAME}\n"
        f"商品名稱：{prod_name}\n"
        f"貨號：{prod_code}\n"
        f"聯繫：{CONTACT_LINE}\n"
        f"商品描述：{po_text}\n\n"
        f"廣告規格：{size_desc}\n"
        f"版面配置：{layout_hint}\n\n"
        f"必須顯示文字：品牌名「{STORE_NAME}」、商品名「{prod_name}」、"
        f"貨號「{prod_code}」、「{CONTACT_LINE}」（附LINE綠色圖示）\n\n"
        f"Read 圖片後直接輸出提示詞。"
    )

    # ── 找 claude CLI 執行檔 ──────────────────────────────────────────────────
    claude_bin = _sh.which("claude") or _sh.which("claude.exe")
    if not claude_bin:
        for candidate in [
            Path.home() / ".claude" / "local" / "claude.exe",
            Path(r"C:\Users\bear\AppData\Local\AnthropicClaude\claude.exe"),
            Path(r"C:\Program Files\Claude\claude.exe"),
        ]:
            if candidate.exists():
                claude_bin = str(candidate)
                break

    if not claude_bin:
        print("[claude-prompt] ⚠️  找不到 claude CLI，使用預設模板")
        return build_gemini_prompt(prod_code, platform)

    # ── 呼叫 claude --print（prompt 透過 stdin 傳入）─────────────────────────
    try:
        print(f"[claude-prompt] 呼叫 claude CLI → {prod_code}/{platform}", flush=True)
        # system 邏輯寫進 prompt 開頭，一起透過 stdin 送入
        full_prompt = _SYSTEM + "\n\n---\n\n" + _USER
        proc = subprocess.run(
            [
                claude_bin,
                "--print",
                "--allowedTools", "Read",
                "--model", "claude-opus-4-6",
            ],
            input=full_prompt,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
            cwd=str(Path(__file__).parent.parent),
        )

        out = proc.stdout.strip()
        if proc.returncode != 0:
            err = proc.stderr.strip() or out
            print(f"[claude-prompt] CLI 錯誤碼 {proc.returncode}：{err[:200]}")
            print("[claude-prompt] 使用預設模板")
            return build_gemini_prompt(prod_code, platform)
        if len(out) > 100:
            print(f"[claude-prompt] ✅ 生成提示詞（{len(out)} 字）")
            return out
        print(f"[claude-prompt] 內容不足（{len(out)} 字），使用預設模板")

    except subprocess.TimeoutExpired:
        print("[claude-prompt] ⚠️  逾時（120s），使用預設模板")
    except Exception as e:
        print(f"[claude-prompt] ⚠️  失敗（{e}），使用預設模板")

    return build_gemini_prompt(prod_code, platform)


def build_gemini_prompt(prod_code: str, platform: str) -> str:
    """組合 Gemini 廣告圖生成提示詞（由 ad-design skill 設計）"""
    prod_name = _get_product_name(prod_code)
    po_text   = _get_po_summary(prod_code)

    if platform == "line":
        return f"""極簡高端電商風格，1:1 正方形構圖，1040×1040px，LINE 官方帳號貼文廣告。

畫面採「中央聚焦型」版面，大量留白，讓商品自己說話。整體質感如精品電商主圖，乾淨、有呼吸感、值得信賴。

【背景】
#FFFFFF 白色為主，底部向 #F0F0F0 極淺灰過渡，漸層極為細膩，幾乎感覺不到。四邊保留 5% 安全邊距，絕對不放任何元素在邊緣。

【頂部區塊（畫面上方 12%）】
品牌名稱「{STORE_NAME}」，字色 #1A1A1A，字重極粗，字母間距略寬，置中。
品牌名下方有一條橘色細線（#FF6B00，寬 36px，高 2.5px，圓角），像精品品牌的 signature 線條。

【主視覺（畫面中央 52%）】
⚠️ 使用我上傳的產品照片作為商品主圖，不要替換成其他圖片。
處理方式：精確去除原始背景，保留商品本體，邊緣乾淨平滑。
去背後的商品置於畫面正中央，寬度佔畫面 68%。
適當調亮色調、提升對比，讓商品在白色背景上更鮮明好看。
底部加橢圓投影（顏色 #CCCCCC，模糊 16px，偏移 8px，橫向壓扁），呈現懸浮貼地感。

【商品資訊區塊（畫面中下 26%）】
- 商品名稱「{prod_name}」：字色 #1A1A1A，字重極粗，字級「大」，置中，行距 1.3
- 貨號「{prod_code}」：商品名稱正下方，字色 #999999，字級「小」，字母間距略寬，置中
- 兩者之間有 8px 的呼吸空間

【底部區塊（畫面下方 10%）】
「{CONTACT_LINE}」，字色 #555555，字級「小」，置中。
LINE 圖示（#06C755 綠色）緊鄰文字左側，圖示大小與文字等高。
底部區塊背景可加極淺的 #FAFAFA 色帶作為視覺分隔。

【整體設計原則】
- 主色白 #FFFFFF，輔色淺灰 #F0F0F0，強調色橘 #FF6B00
- 橘色僅用於品牌線條這一處，其他地方不出現橘色
- 字體統一無襯線黑體，乾淨俐落
- 無多餘裝飾，無漸層色塊，無花俏邊框
- 整體視覺動線：品牌名 → 商品 → 商品名 → 貨號 → LINE 聯繫

【商品資訊參考（幫助理解商品，不需全部顯示在圖上）】
{po_text}

⚠️ 必須清晰顯示貨號「{prod_code}」在商品名稱正下方，這是硬性要求。
請直接生成這張廣告圖片。"""

    else:  # fb
        return f"""極簡高端電商風格，1.91:1 橫式構圖，1200×630px，Facebook 動態貼文廣告。

畫面採「左右分割型」版面，左圖右文，視覺乾淨、資訊層次清晰。整體質感如精品電商廣告，讓消費者在滑動 FB 動態時自然停下來。

【背景】
全版 #FFFFFF 白色，右側資訊區底部向 #FAFAFA 極淺灰過渡，幾乎察覺不到漸層。左右分界用一條 #EBEBEB 細線（1px）輕柔分隔，不搶眼。

【左側商品展示區（畫面寬度 46%）】
⚠️ 使用我上傳的產品照片作為商品主圖，不要替換成其他圖片。
處理方式：精確去除原始背景，保留商品本體，邊緣乾淨平滑。
去背後的商品居中展示於左側區塊，高度佔左側 78%。
適當調亮色調、提升對比，讓商品更鮮明。
底部加橢圓投影（#C8C8C8，模糊 18px），自然貼地感。
商品輕微右傾 2°，增加動感但不失穩重。

【右側品牌資訊區（畫面寬度 54%）】
區塊內容垂直居中，靠左對齊，左側內邊距 48px：

第一層｜品牌標識：
「{STORE_NAME}」，字色 #FF6B00，字重粗，字母間距略寬，字級「小」。
緊接一條橘色線（#FF6B00，寬 28px，高 2px，圓角），線條在文字下方 6px。

第二層｜商品主標（最大視覺重量）：
「{prod_name}」，字色 #1A1A1A，字重極粗（900），字級「大」，行距 1.25。
商品名下方 10px 留白。

第三層｜貨號：
「{prod_code}」，字色 #AAAAAA，字母間距略寬，字級「小」。
貨號下方 16px 留白。

第四層｜產品重點（從商品資訊取 1-2 個最核心賣點）：
每點前加 · 符號（#FF6B00 橘色），字色 #555555，字級「小」，行距 1.6。
最多 2 行，超過不顯示。

第五層｜聯繫資訊（貼近底部）：
「{CONTACT_LINE}」，LINE 圖示（#06C755）+ 文字，字色 #444444，字級「小」。

【整體設計原則】
- 主色白 #FFFFFF，輔色 #FAFAFA，強調色橘 #FF6B00，LINE 綠 #06C755
- 視覺動線：商品圖 → 品牌名 → 商品名 → 貨號 → 賣點 → LINE
- 乾淨、專業、有信任感，橘色只出現在品牌線條和 bullet 點這兩處
- 無多餘裝飾，無漸層色塊，字體統一無襯線

【商品資訊參考（幫助理解商品，不需全部顯示在圖上）】
{po_text}

⚠️ 必須清晰顯示貨號「{prod_code}」在商品名稱正下方，這是硬性要求。
請直接生成這張廣告圖片。"""


# ── 主入口 ────────────────────────────────────────────────────────────────────

def handle_ad_update_trigger(text: str, group_id: str,
                              line_api=None) -> str | None:
    """偵測 `廣告圖更新` 或 `廣告圖測試 Z3432`，啟動生成"""
    stripped = text.strip()

    # ── 測試模式：廣告圖測試 Z3432（只生成一個貨號）────────────────────────
    m = _AD_TEST_RE.match(stripped)
    if m:
        code = m.group(1).upper()

        # 在產品圖資料夾找該貨號的圖片
        try:
            if not AD_IMAGE_DIR.is_dir():
                return f"❌ 找不到產品圖片資料夾：\n{AD_IMAGE_DIR}"
        except OSError:
            return f"❌ 無法存取產品圖片資料夾（磁碟機未連線？）"

        imgs = [str(f) for f in sorted(AD_IMAGE_DIR.iterdir())
                if f.is_file() and f.suffix.lower() in _IMG_EXTS
                and _extract_code_from_filename(f.name) == code]

        try:
            AD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            return f"❌ 無法建立輸出資料夾：{e}"

        prod_name = _get_product_name(code)
        _start_ad_generation_bg({code: imgs}, str(AD_OUTPUT_DIR), group_id, line_api)

        img_note = f"（找到 {len(imgs)} 張圖）" if imgs else "（⚠️ 未找到對應圖片，僅用文字提示詞）"
        return (
            f"🧪 廣告圖測試：{code}  {prod_name}  {img_note}\n"
            f"生成 LINE + FB 共 2 張\n"
            f"⏳ Gemini 即將開啟，完成後通知"
        )

    # ── 全量更新模式 ─────────────────────────────────────────────────────────
    if stripped not in _AD_TRIGGER_WORDS:
        return None

    # 檢查來源資料夾
    try:
        if not AD_IMAGE_DIR.is_dir():
            return f"❌ 找不到產品圖片資料夾：\n{AD_IMAGE_DIR}"
    except OSError:
        return f"❌ 無法存取產品圖片資料夾（磁碟機未連線？）：\n{AD_IMAGE_DIR}"

    # 確保輸出資料夾存在
    try:
        AD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return f"❌ 無法建立廣告圖輸出資料夾：{e}"

    # 掃描圖片
    images = [f for f in AD_IMAGE_DIR.iterdir()
              if f.is_file() and f.suffix.lower() in _IMG_EXTS]
    if not images:
        return f"❌ 產品圖片資料夾內沒有圖片（jpg/png/webp）：\n{AD_IMAGE_DIR}"

    # 從檔名抓貨號
    all_found: dict[str, list[str]] = {}
    for img in sorted(images):
        code = _extract_code_from_filename(img.name)
        if code:
            all_found.setdefault(code, []).append(str(img))

    if not all_found:
        return (
            "❌ 圖片檔名中找不到貨號\n"
            "請確認檔名格式，例如：Z3432.jpg 或 Z3432A.jpg"
        )

    # ── 過濾：廣告圖已存在且比產品圖新 → 跳過 ───────────────────────────────
    def _needs_update(code: str, img_paths: list[str]) -> bool:
        """任一平台廣告圖不存在，或比產品圖舊 → 需要重新生成"""
        src_mtime = max(Path(p).stat().st_mtime for p in img_paths)
        for platform in ("line", "fb"):
            ad_file = AD_OUTPUT_DIR / f"{code}_{platform}.png"
            if not ad_file.exists():
                return True
            if ad_file.stat().st_mtime < src_mtime:
                return True
        return False

    found: dict[str, list[str]] = {
        code: imgs for code, imgs in all_found.items()
        if _needs_update(code, imgs)
    }
    skipped = len(all_found) - len(found)

    if not found:
        return (
            f"✅ 廣告圖已是最新，共 {len(all_found)} 個產品\n"
            f"（產品圖未更新，無需重新生成）"
        )

    # 列出待生成的貨號
    code_list = "\n".join(
        f"  • {c}  {_get_product_name(c)}"
        for c in list(found.keys())[:10]
    )
    extra = f"\n  ...等共 {len(found)} 個" if len(found) > 10 else ""
    skip_note = f"（{skipped} 個已是最新，跳過）\n" if skipped else ""

    # 啟動背景生成
    _start_ad_generation_bg(found, str(AD_OUTPUT_DIR), group_id, line_api)

    return (
        f"📸 廣告圖更新啟動\n"
        f"{skip_note}"
        f"待生成 {len(found)} 個產品：\n{code_list}{extra}\n\n"
        f"⏳ Gemini 瀏覽器即將開啟\n"
        f"每個產品生成 LINE + FB 共 2 張\n"
        f"完成後會在此通知"
    )


# ── 背景執行 Playwright 腳本 ─────────────────────────────────────────────────

def _start_ad_generation_bg(
    images: dict[str, list[str]],
    output_folder: str,
    group_id: str,
    line_api=None,
) -> None:
    """以子程序執行 scripts/generate_ad_gemini.py"""
    import json
    import threading

    def _run():
        script = Path(__file__).parent.parent / "scripts" / "generate_ad_gemini.py"
        payload = json.dumps({
            "codes":         list(images.keys()),
            "images":        images,
            "output_folder": output_folder,
            "notify_group":  group_id,
        }, ensure_ascii=False)

        python = sys.executable
        flags  = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        proc = subprocess.Popen(
            [python, str(script), "--payload", payload],
            cwd=str(Path(__file__).parent.parent),
            creationflags=flags,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
        )
        stdout, _ = proc.communicate()
        if stdout:
            print(stdout.decode("utf-8", errors="replace"), flush=True)

    threading.Thread(target=_run, daemon=True).start()
