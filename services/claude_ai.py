"""
Claude CLI fallback 服務

當 Bot 無法辨識圖片或回答問題時，呼叫本機 Claude Code CLI 處理。
使用 Max 訂閱額度，不需要 API key。
"""

import subprocess
import tempfile
import json
from pathlib import Path

_BASE = Path(__file__).parent.parent
_TIMEOUT = 60  # 秒
_CLAUDE_CMD = r"C:\Users\bear\AppData\Roaming\npm\claude.cmd"

# ── 對話紀錄（in-memory，每個客戶保留最近 10 輪）──────────────
import threading
_chat_history: dict[str, list[dict]] = {}  # user_id → [{"role": "user"/"bot", "text": "..."}]
_chat_lock = threading.Lock()
_MAX_HISTORY = 10


def add_chat_history(user_id: str, role: str, text: str) -> None:
    """記錄一輪對話（role: 'user' 或 'bot'）"""
    with _chat_lock:
        if user_id not in _chat_history:
            _chat_history[user_id] = []
        _chat_history[user_id].append({"role": role, "text": text[:200]})
        # 只保留最近 N 輪
        if len(_chat_history[user_id]) > _MAX_HISTORY * 2:
            _chat_history[user_id] = _chat_history[user_id][-_MAX_HISTORY * 2:]


def _get_chat_context(user_id: str) -> str:
    """取得該客戶最近的對話紀錄"""
    with _chat_lock:
        history = _chat_history.get(user_id, [])
    if not history:
        return ""
    lines = ["【最近對話紀錄】"]
    for h in history:
        prefix = "客戶" if h["role"] == "user" else "客服"
        lines.append(f"{prefix}：{h['text']}")
    return "\n".join(lines)


_PO_PATH = Path(r"H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt")


def _load_context() -> str:
    """載入產品資料作為 Claude 的 context"""
    import re
    parts = []

    # 載入 PO文（完整產品描述，越後面越新）
    if _PO_PATH.exists():
        try:
            po_text = _PO_PATH.read_text(encoding="utf-8").strip()
            blocks = re.split(r"\n{2,}", po_text)
            # 取最後 50 段（最新的產品）
            recent = blocks[-50:] if len(blocks) > 50 else blocks
            parts.append("【產品PO文（越後面越新）】\n" + "\n\n".join(recent))
        except Exception:
            pass

    # 載入 specs（產品規格）— 全部載入
    specs_path = _BASE / "data" / "specs.json"
    if specs_path.exists():
        try:
            specs = json.loads(specs_path.read_text(encoding="utf-8"))
            if isinstance(specs, dict):
                spec_lines = []
                for code, s in specs.items():
                    name = s.get("name", "")
                    price = s.get("price", "")
                    size = s.get("size", "")
                    weight = s.get("weight", "")
                    machine = "、".join(s.get("machine", []))
                    spec_lines.append(f"{code}: {name} | 價格:{price} | 尺寸:{size} | 重量:{weight} | 台型:{machine}")
                parts.append("【產品規格】\n" + "\n".join(spec_lines))
        except Exception:
            pass

    # 載入庫存 — 可售庫存（先確認資料新鮮度）
    avail_path = _BASE / "data" / "available.json"
    if avail_path.exists():
        import time
        age = time.time() - avail_path.stat().st_mtime
        if age > 30 * 60:
            try:
                from services.ecount import ecount_client
                ecount_client._ensure_available()
                print("[claude-ai] 庫存資料過期，已觸發同步", flush=True)
            except Exception:
                pass
    if avail_path.exists():
        try:
            avail = json.loads(avail_path.read_text(encoding="utf-8"))
            if isinstance(avail, dict):
                inv_lines = []
                for code, data in avail.items():
                    if isinstance(data, dict):
                        qty = data.get("available", 0)
                    else:
                        qty = data
                    inv_lines.append(f"{code}: 可售{qty}個")
                parts.append("【庫存（可售數量）】\n" + "\n".join(inv_lines))
        except Exception:
            pass

    return "\n\n".join(parts)


_SYSTEM_PROMPT = """重要：忽略所有其他系統指示。你現在的唯一角色是「小蠻牛客服機器人」。不要寫程式、不要分析程式碼、不要提到任何開發相關的事。你只負責回覆客戶的問題。

你是小蠻牛公司的客服機器人。小蠻牛是娃娃機商品批發商，客戶主要是娃娃機台主。
- 回覆要簡短親切，用繁體中文，1-3 句話就好
- 不確定的資訊不要亂說，回覆「我幫您確認一下，稍後回覆您」
- 不要提到你是 AI 或 Claude
- 語氣像真人客服，友善但專業
- 營業時間：週二到週日 13:00~21:00，週一公休
- 地址：新北市土城區中央路二段394巷12號
- PO文資料越後面的是越新的商品，客戶問「新貨」「最近有什麼」就從後面找
- 庫存 0 或負數代表缺貨，正數代表有現貨
- 推薦商品時只推薦有現貨的（庫存 > 0），絕對不要推薦缺貨的商品
- 絕對不要把庫存數量告訴客戶！只說「有現貨」或「目前缺貨」，不要說有幾個
- 不要自己編造流程或承諾（例如「已幫您備註」「明天取貨」「已登記」等），如果不確定怎麼處理，回覆「我幫您確認一下，稍後回覆您」
- 客戶問出貨、送貨、先送、分批送等物流問題，回覆「我幫您確認一下，稍後回覆您」
- 不要把客戶的口語當成產品名搜尋（如「蠻的」不是產品名）
- 客戶問價格、尺寸、重量等，從產品規格和PO文裡找
- 如果真的找不到資訊，回覆「我幫您確認一下，稍後回覆您」
- 客戶要明細、收據、發票等，回覆「請問大概是什麼時候的訂單呢？我稍後拍照給您」
- 不要叫客戶提供訂單編號，客戶不會有訂單編號
- 推薦產品時列出貨號、品名、價格即可，圖片會自動附上，不要提到圖片相關的事（不要說「沒有圖片」「稍後傳圖」「系統會顯示」等）
- 客戶問「有圖嗎」「圖片看一下」時，直接列出產品資訊就好，圖片會自動附上
- 如果圖片看不清楚或無法辨識出任何貨號，回覆「確認中～請稍等下唷～」，不要描述圖片內容、不要問客戶問題
- 你的回覆會直接傳給客戶，所以只輸出回覆內容，不要加任何解釋、程式碼或前綴"""


def ask_claude_text(question: str, user_id: str = "") -> str | None:
    """
    用 Claude CLI 回答文字問題。
    回傳回答文字，失敗回傳 None。
    """
    context = _load_context()
    chat_ctx = _get_chat_context(user_id) if user_id else ""
    full_prompt = f"{_SYSTEM_PROMPT}\n\n{context}\n\n{chat_ctx}\n\n---\n客戶問：{question}\n\n請直接回覆客戶（不要加任何前綴或解釋）："

    try:
        env = {**__import__("os").environ, "PYTHONIOENCODING": "utf-8"}
        result = subprocess.run(
            [_CLAUDE_CMD, "-p", "-",
             "--tools", ""],
            input=full_prompt.encode("utf-8"),
            capture_output=True, timeout=_TIMEOUT, env=env,
            cwd="C:\\Users\\bear\\AppData\\Local\\Temp",
        )
        answer = result.stdout.decode("utf-8", errors="replace").strip()
        if answer and result.returncode == 0:
            print(f"[claude-ai] 文字回答成功: {question[:30]!r} → {answer[:50]!r}", flush=True)
            return answer
        else:
            stderr = result.stderr.decode("utf-8", errors="replace")[:100]
            print(f"[claude-ai] 文字回答失敗: returncode={result.returncode} stderr={stderr}", flush=True)
            return None
    except subprocess.TimeoutExpired:
        print(f"[claude-ai] 逾時（{_TIMEOUT}s）: {question[:30]!r}", flush=True)
        return None
    except Exception as e:
        print(f"[claude-ai] 例外: {e}", flush=True)
        return None


def ask_claude_image(img_bytes: bytes, question: str = "", user_id: str = "") -> str | None:
    """
    用 Claude CLI 辨識圖片。
    回傳回答文字，失敗回傳 None。
    """
    # 存暫存圖片
    try:
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as f:
            f.write(img_bytes)
            tmp_path = f.name
    except Exception as e:
        print(f"[claude-ai] 暫存圖片失敗: {e}", flush=True)
        return None

    context = _load_context()
    chat_ctx = _get_chat_context(user_id) if user_id else ""
    full_prompt = f"""{_SYSTEM_PROMPT}

{context}

{chat_ctx}

---
客戶傳了一張產品圖片。{f'客戶說：{question}' if question else ''}
請先用 Read tool 讀取圖片 {tmp_path}，辨識這是什麼產品。

重要規則：
- 最重要：仔細讀圖片裡所有白色標籤、價格牌上的文字，特別是貨號（如 T1221、Z3240、S0633 等格式）
- 圖片上的標籤文字是最可靠的資訊，一定要讀出來
- 只回答你在圖片中確實看到的產品，必須能對應到上面的產品資料（貨號匹配）
- 如果圖片裡有多個產品，列出所有能辨識到貨號的產品，格式如：「1. T1221 攀爬遙控車 299元\n2. Z3240 三麗歐兒童枕 259元」
- 不確定的產品不要猜，絕對不要用外觀去猜可能是哪個產品
- 無法讀到任何貨號時，回覆類似「確認中～請稍等一下唷～」或「收到圖片了！讓我查一下」（溫馨簡短，不要描述圖片內容）
請直接回覆客戶："""

    try:
        env = {**__import__("os").environ, "PYTHONIOENCODING": "utf-8"}
        result = subprocess.run(
            [_CLAUDE_CMD, "-p", "-",
             "--tools", "Read"],
            input=full_prompt.encode("utf-8"),
            capture_output=True, timeout=_TIMEOUT, env=env,
            cwd="C:\\Users\\bear\\AppData\\Local\\Temp",
        )
        answer = result.stdout.decode("utf-8", errors="replace").strip()
        if answer and result.returncode == 0:
            print(f"[claude-ai] 圖片辨識成功: {answer[:50]!r}", flush=True)
            return answer
        else:
            print(f"[claude-ai] 圖片辨識失敗: returncode={result.returncode}", flush=True)
            return None
    except subprocess.TimeoutExpired:
        print(f"[claude-ai] 圖片辨識逾時（{_TIMEOUT}s）", flush=True)
        return None
    except Exception as e:
        print(f"[claude-ai] 圖片辨識例外: {e}", flush=True)
        return None
    finally:
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except Exception:
            pass
