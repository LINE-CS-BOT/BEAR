"""
將 產品PO文.txt 解析並匯入 data/specs.json

格式範例：
  編號：T1202
  建議：標準台用
  品名：杜卡迪合金回力摩托車
  尺寸-約18X9X9公分
  重量：約：237公克
  價格：109元
  ...

執行方式（每次更新PO文.txt後重跑）：
  python scripts/import_specs.py
"""

import io
import json
import re
import sys
from pathlib import Path

# Windows cmd 編碼修正
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# 壓制 Windows 磁碟機未連線彈窗（H: 雲端磁碟離線時靜默失敗）
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
    except Exception:
        pass

SOURCE = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt")
OUTPUT = Path(__file__).parent.parent / "data" / "specs.json"


# 台型正規化對照表
_MACHINE_MAP = {
    "k霸":   "K霸台",
    "k霸台": "K霸台",
    "小k霸": "小K霸台",
    "小k霸台": "小K霸台",
    "中巨":   "中巨台",
    "中巨台": "中巨台",
    "標準":   "標準台",
    "標準台": "標準台",
    "迷你":   "迷你台",
    "迷你台": "迷你台",
    "超k":    "超K台",
    "超k台":  "超K台",
    "巨無霸": "巨無霸台",
    "巨無霸台": "巨無霸台",
}

def _normalize_machine(raw: str) -> str:
    key = raw.strip().lower()
    return _MACHINE_MAP.get(key, raw.strip())


def parse_specs(text: str) -> dict:
    specs = {}
    # 按空行切割產品區塊
    blocks = re.split(r"\n{2,}", text.strip())

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        code = name = size = weight = machine = price = ""
        first_untagged = ""   # 第一行無標籤文字（備用品名）

        block_lines = [l.strip() for l in block.splitlines()]
        for i, line in enumerate(block_lines):
            if not line:
                continue

            # 去除行首 emoji/符號，方便匹配
            _clean = re.sub(r'^[\U0001f300-\U0001faff\u2600-\u27bf\u2702-\u27b0‼️⁉️*✨⭐️🔥💥⚠️🎉❤️]+\s*', '', line)

            # 編號（支援「編號：」「產品編號：」「商品編號：」「貨號：」「新編號：」）
            m = re.match(r"(?:產品|商品|新)?(?:編號|貨號)[：:](.+)", _clean)
            if m:
                code = m.group(1).strip()
                continue

            # 建議台型
            m = re.match(r"建議[：:](.+)", _clean)
            if m:
                machine = m.group(1).strip()
                continue

            # 品名（多種格式：「品名：」或「品名」不含冒號）
            m = re.match(r"品名[：:]?(.+)", _clean)
            if m:
                name = m.group(1).strip()
                continue

            # 尺寸（支援「尺寸」「尺吋」「包裝尺寸」「產品包裝尺寸」「外盒尺寸」）
            m = re.match(r"(?:產品)?(?:包裝)?(?:外盒)?(?:尺寸|尺吋)[-：: 約]*(.*)", _clean)
            if m:
                val = m.group(1).strip()
                if val:
                    size = val
                elif i + 1 < len(block_lines) and block_lines[i + 1].strip():
                    # 尺寸標籤後面沒值，取下一行
                    size = block_lines[i + 1].strip()
                continue
            # 尺寸 fallback：行內含「N*N*N公分」或「N×N×N cm」格式
            if not size:
                m_size_fb = re.search(r'(\d+(?:\.\d+)?\s*[*×xX]\s*\d+(?:\.\d+)?(?:\s*[*×xX]\s*\d+(?:\.\d+)?)?)\s*(?:公分|cm|CM)', line)
                if m_size_fb:
                    size = m_size_fb.group(0).strip()
                    continue

            # 重量（支援「重量：」「產品重量：」「單盒重量：」「單顆重量：」）
            m = re.match(r"(?:產品|單盒|單顆)?重量[：: ]*(.+)", _clean)
            if m:
                weight = m.group(1).strip()
                # 去除「約：」「約」前綴
                weight = re.sub(r"^約[：:]?\s*", "", weight).strip()
                # 去除重量後面的備註文字（如「260公克不易脫爪」→「260公克」）
                weight = re.sub(r"(\d+(?:\.\d+)?(?:公克|克|g|kg|公斤)).*", r"\1", weight)
                continue
            # 重量 fallback：行內含「約N克」「約Ng」格式
            if not weight:
                m_wt_fb = re.search(r'約?\s*(\d+(?:\.\d+)?)\s*(?:公克|克|g|kg|公斤)', line)
                if m_wt_fb:
                    weight = m_wt_fb.group(0).strip()
                    weight = re.sub(r"^約\s*", "", weight).strip()
                    continue

            # 價格（格式變化多：「價格：」「售價：」「單盒特價：」「批價：」「零售價：」）
            m = re.match(r"(?:價格|售價|單盒特價|特價|批價|零售價)[：:](.+)", _clean)
            if m:
                price = m.group(1).strip()
                continue

            # 無標籤行 → 記錄第一行作為備用品名（排除純數字/英文短詞）
            if (not first_untagged
                    and len(line) >= 4
                    and not re.match(r'^[\d\s/＋+×x*]+$', line)):
                first_untagged = line

        # 品名為空時，用第一行無標籤文字補上
        if not name and first_untagged:
            name = first_untagged

        if not code:
            continue

        code = code.upper()   # 統一大寫，避免 k0216 vs K0216 查不到

        # 台型清單：「標準/迷你台用」 → ["標準台", "迷你台"]
        machine_clean = machine.replace("用", "").replace("專", "").strip()
        raw_list = [m.strip() for m in re.split(r"[/／、]", machine_clean) if m.strip()]
        machine_list = [_normalize_machine(m) for m in raw_list]

        specs[code] = {
            "code":    code,
            "name":    name,
            "size":    size,
            "weight":  weight,
            "machine": machine_list,
            "price":   price,
        }

    return specs


def main():
    try:
        exists = SOURCE.exists()
    except OSError:
        exists = False
    if not exists:
        print(f"找不到來源檔案：{SOURCE}")
        sys.exit(1)

    text = SOURCE.read_text(encoding="utf-8")
    specs = parse_specs(text)

    OUTPUT.parent.mkdir(exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(specs, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 成功匯入 {len(specs)} 筆規格資料 → {OUTPUT}")
    for code, s in specs.items():
        machines = "、".join(s["machine"]) if s["machine"] else "通用"
        print(f"  {code}：{s['name']}｜{s['size']}｜{s['weight']}｜{machines}｜{s['price']}")


if __name__ == "__main__":
    main()
