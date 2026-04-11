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
    _last_code = ""
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

            # 去除行首 emoji/符號，方便匹配（涵蓋所有 Unicode emoji 區段）
            _clean = re.sub(r'^[\U0001f300-\U0001faff\U0001f600-\U0001f64f\U0001f680-\U0001f6ff\u2600-\u27bf\u2702-\u27b0\ufe0f\u200d‼️⁉️*✨⭐️🔥💥⚠️🎉❤️❇️🐚🌈☀️]+\s*', '', line)

            # 編號（支援「編號：」「產品編號：」「商品編號：」「貨號：」「新編號：」）
            _new_code = ""
            m = re.match(r"(?:產品|商品|新)?(?:編號|貨號)[：:](.+)", _clean)
            if m:
                raw_code = m.group(1).strip()
                m_code_only = re.match(r'([A-Za-z]{1,3}-?\d{3,6}(?:-[A-Za-z0-9]+)*)', raw_code)
                _new_code = m_code_only.group(1) if m_code_only else raw_code
            # 編號 fallback：純貨號行（如 U0380、T1202、Z3300）
            if not _new_code and not code:
                m = re.match(r'^([A-Za-z]\d{3,5}(?:-[A-Za-z0-9]+)*)$', _clean.strip())
                if m:
                    _new_code = m.group(1).strip()
            if _new_code:
                # block 內尚無貨號但有累積 specs → 屬於前一個 block 的貨號
                if not code and _last_code and _last_code in specs and (size or weight or price):
                    if size and not specs[_last_code]["size"]:
                        specs[_last_code]["size"] = size
                    if weight and not specs[_last_code]["weight"]:
                        specs[_last_code]["weight"] = weight
                    if price and not specs[_last_code]["price"]:
                        specs[_last_code]["price"] = price
                    name = size = weight = machine = price = ""
                    first_untagged = ""
                # 同一 block 出現第二個貨號 → 先存前一個貨號的 specs
                elif code and code != _new_code.upper() and (size or weight or price):
                    _save_code = code.upper()
                    if _save_code not in specs:
                        specs[_save_code] = {"code": _save_code, "name": name, "size": size, "weight": weight, "machine": [], "price": price}
                    else:
                        if size and not specs[_save_code]["size"]:
                            specs[_save_code]["size"] = size
                        if weight and not specs[_save_code]["weight"]:
                            specs[_save_code]["weight"] = weight
                        if price and not specs[_save_code]["price"]:
                            specs[_save_code]["price"] = price
                    _last_code = _save_code
                    name = size = weight = machine = price = ""
                    first_untagged = ""
                code = _new_code
                continue

            # 建議台型
            m = re.match(r"建議[：:](.+)", _clean)
            if m:
                machine = m.group(1).strip()
                continue

            # 品名（多種格式：「品名：」「名稱：」「產品名稱：」）
            m = re.match(r"(?:產品名稱|品名|名稱)[：:]?(.+)", _clean)
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

            # 重量（支援「重量：」「產品重量：」「單盒重量：」「單顆重量：」「每盒重量：」）
            m = re.match(r"(?:產品|單盒|單顆|每盒)?重量[：: ]*(.+)", _clean)
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
            # 無貨號但有規格 → 補到前一個貨號
            if _last_code and _last_code in specs:
                if size and not specs[_last_code]["size"]:
                    specs[_last_code]["size"] = size
                if weight and not specs[_last_code]["weight"]:
                    specs[_last_code]["weight"] = weight
                if price and not specs[_last_code]["price"]:
                    specs[_last_code]["price"] = price
            continue

        code = code.upper()   # 統一大寫，避免 k0216 vs K0216 查不到
        _last_code = code

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


def _strip_paren_prefix(name: str) -> str:
    """去除品名開頭的括號前綴，如 '(大)潮牌跑酷' → '潮牌跑酷'、'（原）三麗鷗' → '三麗鷗'"""
    return re.sub(r'^[\(（][^)）]*[\)）]\s*', '', name).strip()


def _format_price(price_val) -> str:
    """價格統一為 '數字元' 格式，如 109.0 → '109元'、'109元' → '109元'"""
    if isinstance(price_val, (int, float)) and price_val > 0:
        n = int(price_val) if price_val == int(price_val) else price_val
        return f"{n}元"
    if isinstance(price_val, str):
        m = re.search(r'(\d+(?:\.\d+)?)', price_val)
        if m:
            n = float(m.group(1))
            n = int(n) if n == int(n) else n
            return f"{n}元"
    return ""


def _enrich_from_ecount(specs: dict) -> dict:
    """用 Ecount 品項資料覆蓋品名和價格"""
    sys.path.insert(0, str(Path(__file__).parent.parent))
    try:
        from services.ecount import ecount_client
        ecount_client._ensure_product_cache()
    except Exception as e:
        print(f"⚠️ 無法載入 Ecount 品項快取：{e}")
        return specs

    matched = 0
    for code, s in specs.items():
        item = ecount_client.get_product_cache_item(code)
        if not item:
            # 品名/價格保留 PO 文原始值，但統一格式
            s["name"] = _strip_paren_prefix(s["name"]) if s["name"] else ""
            s["price"] = _format_price(s["price"])
            continue
        matched += 1
        # 品名：用 Ecount 的，去掉括號前綴
        ec_name = _strip_paren_prefix(item.get("name", ""))
        if ec_name:
            s["name"] = ec_name
        # 價格：用 Ecount 的 OUT_PRICE
        ec_price = item.get("price", 0)
        if ec_price and ec_price > 0:
            s["price"] = _format_price(ec_price)
        else:
            s["price"] = _format_price(s["price"])

    print(f"[enrich] Ecount 匹配 {matched}/{len(specs)} 筆")
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

    # 用 Ecount 覆蓋品名和價格
    specs = _enrich_from_ecount(specs)

    OUTPUT.parent.mkdir(exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(specs, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 成功匯入 {len(specs)} 筆規格資料 → {OUTPUT}")
    for code, s in specs.items():
        machines = "、".join(s["machine"]) if s["machine"] else "通用"
        print(f"  {code}：{s['name']}｜{s['size']}｜{s['weight']}｜{machines}｜{s['price']}")


if __name__ == "__main__":
    main()
