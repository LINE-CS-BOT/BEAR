"""
在 Claude Code (cowork) 環境中執行，生成廣告提示詞並存檔。

流程：
  STEP 1 → claude --print (ad-design skill)：看照片 + 產品資訊 → 完整設計規格 + 初步提示詞
  STEP 2 → claude --print (ad-style-optimizer skill)：依產品類別選風格 → 風格化提示詞
  最後從輸出中解析出 LINE / FB 提示詞分別存檔

用法：
  python scripts/generate_prompts.py           # 全部產品
  python scripts/generate_prompts.py P0137     # 單一產品
"""

import re
import subprocess
import sys
import shutil
import tempfile
from pathlib import Path

# Windows stdout 強制 utf-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

# 載入 .env（必須在 import handlers 之前）
import os
os.chdir(str(ROOT))
try:
    from dotenv import load_dotenv
    load_dotenv(ROOT / ".env")
except ImportError:
    pass

from handlers.ad_maker import (
    AD_IMAGE_DIR, _IMG_EXTS, _extract_code_from_filename,
    _get_product_name, _get_po_summary,
    STORE_NAME, CONTACT_LINE,
)

PROMPT_DIR  = ROOT / "data" / "ad_prompts"
SKILLS_DIR  = ROOT / "data" / "skills"
PROMPT_DIR.mkdir(parents=True, exist_ok=True)

# ── 讀取 Skill 內容 ────────────────────────────────────────────────────────────

def _load_skill(filename: str) -> str:
    path = SKILLS_DIR / filename
    if path.exists():
        return path.read_text(encoding="utf-8")
    return ""

SKILL_AD_DESIGN   = _load_skill("ad_design_skill.txt")
SKILL_AD_STYLE    = _load_skill("ad_style_optimizer_skill.txt")

# ── 廣告規格定義 ───────────────────────────────────────────────────────────────

PLATFORMS = {
    "line": {
        "size":    "1080×1080px，1:1 正方形，LINE 官方帳號貼文廣告",
        "layout":  "中央聚焦型：品牌名（上12%）→ 產品主視覺（中52%）→ 商品名+貨號（中下26%）→ LINE ID（下10%）",
        "label":   "LINE 1040×1040",
    },
    "fb": {
        "size":    "1200×630px，1.91:1 橫式，Facebook 動態貼文廣告",
        "layout":  "左右分割型：左側產品圖（46%），右側垂直排列 品牌名→橘線→商品名→貨號→賣點1-2條→LINE ID",
        "label":   "FB 1200×630",
    },
}

# ── 取圖路徑 ───────────────────────────────────────────────────────────────────

def get_image_paths(code: str) -> list[str]:
    """找產品圖，H:\\ 複製到本機 temp 回傳路徑"""
    imgs = [
        f for f in sorted(AD_IMAGE_DIR.iterdir())
        if f.is_file()
        and f.suffix.lower() in _IMG_EXTS
        and _extract_code_from_filename(f.name) == code
    ]
    local = []
    for p in imgs[:3]:
        if p.drive.upper() not in ("C:", "D:", "E:"):
            try:
                tmp = Path(tempfile.gettempdir()) / p.name
                shutil.copy2(str(p), str(tmp))
                local.append(str(tmp))
            except Exception:
                pass
        else:
            local.append(str(p))
    return local

# ── 呼叫 claude --print ────────────────────────────────────────────────────────

def _run_claude(system: str, user: str, timeout: int = 180) -> str | None:
    claude_bin = shutil.which("claude") or shutil.which("claude.exe")
    if not claude_bin:
        print("  ❌ 找不到 claude CLI")
        return None
    try:
        proc = subprocess.run(
            [claude_bin, "--print", "--allowedTools", "Read",
             "--model", "claude-opus-4-6"],
            input=system + "\n\n---\n\n" + user,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            cwd=str(ROOT),
        )
        out = proc.stdout.strip()
        if proc.returncode != 0:
            err = proc.stderr.strip() or out[:120]
            print(f"  ❌ CLI 錯誤（{proc.returncode}）：{err[:200]}")
            return None
        if len(out) > 100:
            return out
        print(f"  ⚠️  輸出太短（{len(out)} 字）")
        return None
    except subprocess.TimeoutExpired:
        print(f"  ❌ 逾時（{timeout}s）")
        return None
    except Exception as e:
        print(f"  ❌ 失敗：{e}")
        return None

# ── STEP 1：ad-design skill ────────────────────────────────────────────────────

def step1_design(code: str, local_paths: list[str]) -> str | None:
    """呼叫 ad-design skill，讓 Claude 看照片後產出完整設計規格 + 初步提示詞"""
    prod_name = _get_product_name(code)
    po_text   = _get_po_summary(code)
    paths_str = "\n".join(local_paths)

    line_spec = PLATFORMS["line"]
    fb_spec   = PLATFORMS["fb"]

    from handlers.ad_maker import _get_product_specs, _get_product_price
    specs = _get_product_specs(code)
    size_text = specs.get("size", "")
    weight_text = specs.get("weight", "")
    price = _get_product_price(code)

    user_msg = (
        f"請用 Read 工具讀取以下產品照片，依照廣告圖片設計規範設計以下兩張廣告並輸出完整設計規格與 AI 製圖提示詞。\n\n"
        f"產品照片路徑：\n{paths_str}\n\n"
        f"【重要：產品外觀描述要求】\n"
        f"- 仔細觀察照片中的所有元素：產品本體、包裝盒、配件、標籤、印刷圖案\n"
        f"- 必須在提示詞中用至少 300 字詳細描述產品外觀（盒型、顏色、文字、圖案位置等）\n"
        f"- 描述顏色時請具體（如「深藍底印金色花紋」），避免模糊說法\n"
        f"- 商品盒型必須完整呈現，四個角清楚可見，不可裁切、變形\n\n"
        f"產品資訊：\n"
        f"品牌：{STORE_NAME}\n"
        f"商品名稱：{prod_name}\n"
        f"貨號：{code}\n"
        f"價格：{price}\n"
        f"尺寸：{size_text}\n"
        f"重量：{weight_text}\n"
        f"聯繫：{CONTACT_LINE}\n"
        f"商品描述：{po_text}\n\n"
        f"廣告 1 規格：{line_spec['size']}\n"
        f"廣告 1 版面：{line_spec['layout']}\n\n"
        f"廣告 2 規格：{fb_spec['size']}\n"
        f"廣告 2 版面：{fb_spec['layout']}\n\n"
        f"【廣告圖必須顯示的文字】\n"
        f"- 品牌名「{STORE_NAME}」\n"
        f"- 商品名「{prod_name}」\n"
        f"- 貨號「{code}」\n"
        f"- 價格「{price}」\n"
        f"- 尺寸「{size_text}」（有值才顯示）\n"
        f"- 重量「{weight_text}」（有值才顯示）\n"
        f"- 「{CONTACT_LINE}」（附LINE綠色圖示）\n\n"
        f"【斟酌顯示】\n"
        f"- 產品特色描述（從商品描述中挑 1-2 個賣點）\n\n"
        f"請輸出兩張廣告的完整設計規格與 AI 製圖提示詞（繁體中文）。"
    )

    system = SKILL_AD_DESIGN if SKILL_AD_DESIGN else (
        "你是一位資深廣告視覺設計師。請根據產品照片設計廣告並輸出 AI 製圖提示詞（繁體中文）。"
    )

    print("  [STEP 1] ad-design skill...", flush=True)
    return _run_claude(system, user_msg, timeout=180)

# ── STEP 2：ad-style-optimizer skill ──────────────────────────────────────────

def step2_optimize(code: str, design_output: str) -> str | None:
    """呼叫 ad-style-optimizer skill，依產品類別選風格優化提示詞"""
    prod_name = _get_product_name(code)

    user_msg = (
        f"以下是貨號 {code}「{prod_name}」兩張廣告的設計規格與初步提示詞，"
        f"請根據產品類別選擇最適合的廣告風格，對 LINE 和 FB 兩張廣告的提示詞進行風格優化。\n\n"
        f"品牌：{STORE_NAME}\n"
        f"目標：在 LINE 和 Facebook 上吸引點擊，導購\n\n"
        f"【STEP 1 設計規格輸入】\n"
        f"{design_output}\n\n"
        f"請輸出風格優化後的 AI 製圖提示詞（繁體中文），"
        f"LINE 和 FB 各一段，每段 400-800 字。"
    )

    system = SKILL_AD_STYLE if SKILL_AD_STYLE else (
        "你是一位廣告風格優化專家。請優化廣告提示詞，使其更有質感（繁體中文）。"
    )

    print("  [STEP 2] ad-style-optimizer skill...", flush=True)
    return _run_claude(system, user_msg, timeout=180)

# ── 解析風格化輸出，取出 LINE / FB 提示詞 ─────────────────────────────────────

def parse_optimized_prompts(text: str) -> dict[str, str]:
    """
    解析 LINE / FB 提示詞。
    支援兩種格式：
      1. \"\"\"...\"\"\"  （三引號）
      2. ```...```      （反引號 Markdown 程式碼區塊）
    取出順序：第1個=LINE，第2個=FB
    """
    # 先嘗試三引號
    blocks = re.findall(r'"""(.*?)"""', text, re.DOTALL)
    # 若不足，再嘗試反引號區塊（含語言標記如 ```text）
    if len(blocks) < 2:
        backtick_blocks = re.findall(r'```(?:\w*\n)?(.*?)```', text, re.DOTALL)
        # 過濾掉太短的（可能是程式碼片段非提示詞）
        backtick_blocks = [b.strip() for b in backtick_blocks if len(b.strip()) > 100]
        blocks = blocks + backtick_blocks

    result = {}
    if len(blocks) >= 1:
        result["line"] = blocks[0].strip()
    if len(blocks) >= 2:
        result["fb"]   = blocks[1].strip()
    return result

# ── 處理單一貨號 ───────────────────────────────────────────────────────────────

def process_code(code: str, step1_only: bool = False) -> bool:
    prod_name   = _get_product_name(code)
    local_paths = get_image_paths(code)

    print(f"\n{'='*55}")
    print(f"貨號：{code}  {prod_name}")
    print(f"圖片：{[Path(p).name for p in local_paths]}")
    print(f"模式：{'僅 STEP 1' if step1_only else 'STEP 1 + STEP 2'}")

    if not local_paths:
        print("  ⚠️  找不到圖片，跳過")
        return False

    # ── STEP 1：設計規格 ──
    design_output = step1_design(code, local_paths)
    if not design_output:
        print("  ❌ STEP 1 失敗，跳過")
        return False
    print(f"  ✅ STEP 1 完成（{len(design_output)} 字）")

    step1_file = PROMPT_DIR / f"{code}_step1_design.txt"
    step1_file.write_text(design_output, encoding="utf-8")

    if step1_only:
        # 直接從 STEP 1 解析提示詞
        prompts = parse_optimized_prompts(design_output)
    else:
        # ── STEP 2：風格優化 ──
        optimized = step2_optimize(code, design_output)
        if not optimized:
            print("  ❌ STEP 2 失敗，改用 STEP 1 原始提示詞")
            prompts = parse_optimized_prompts(design_output)
        else:
            print(f"  ✅ STEP 2 完成（{len(optimized)} 字）")
            step2_file = PROMPT_DIR / f"{code}_step2_style.txt"
            step2_file.write_text(optimized, encoding="utf-8")
            prompts = parse_optimized_prompts(optimized)

    # ── 存最終提示詞 ──
    ok = True
    for platform in ("line", "fb"):
        out_file = PROMPT_DIR / f"{code}_{platform}.txt"
        prompt   = prompts.get(platform, "")
        if prompt and len(prompt) > 100:
            out_file.write_text(prompt, encoding="utf-8")
            print(f"  ✅ {platform}: {len(prompt)} 字 → {out_file.name}")
        else:
            print(f"  ⚠️  {platform}: 解析失敗，未存檔（解析到 {len(prompt)} 字）")
            ok = False

    return ok

# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main():
    # 預設跑完整兩步流程（STEP 1 + STEP 2），加 --step1-only 才只跑 STEP 1
    step1_only = "--step1-only" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]

    if args:
        code = args[0].upper()
        process_code(code, step1_only=step1_only)
        return

    try:
        if not AD_IMAGE_DIR.is_dir():
            print(f"❌ 找不到產品圖片資料夾：{AD_IMAGE_DIR}")
            return
    except OSError:
        print(f"❌ 無法存取磁碟機（H:\\ 未連線？）")
        return

    codes: dict[str, list] = {}
    for f in sorted(AD_IMAGE_DIR.iterdir()):
        if f.is_file() and f.suffix.lower() in _IMG_EXTS:
            c = _extract_code_from_filename(f.name)
            if c:
                codes.setdefault(c, []).append(f)

    if not codes:
        print("❌ 沒有找到任何產品圖片")
        return

    print(f"找到 {len(codes)} 個產品")
    success = 0
    for code in codes:
        if process_code(code, step1_only=step1_only):
            success += 1

    print(f"\n完成：{success}/{len(codes)} 個產品生成提示詞")
    print(f"存放位置：{PROMPT_DIR}")


if __name__ == "__main__":
    main()
