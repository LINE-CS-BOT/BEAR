# ── 最優先：process 啟動時就壓制 Windows 磁碟機未連線彈窗 ────────────
import sys as _sys
if _sys.platform == "win32":
    try:
        import ctypes as _ctypes
        # SEM_FAILCRITICALERRORS(0x0001) + SEM_NOOPENFILEERRORBOX(0x8000)
        # 壓制 H:/B: 磁碟機未就緒時的系統彈窗，讓 Python 直接收到 OSError
        _ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
    except Exception:
        pass
    # 強制 stdout/stderr 使用 UTF-8，避免 emoji 在 cp950 終端機 crash
    try:
        _sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        _sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import asyncio
import base64
import re
import secrets
import sqlite3
import uvicorn
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from pathlib import Path
from fastapi import FastAPI, Request, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, Response
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, PushMessageRequest, TextMessage,
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent, ImageMessageContent, VideoMessageContent
from linebot.v3.exceptions import InvalidSignatureError

from config import settings
from handlers.intent import detect_intent, Intent
from handlers.inventory import handle_inventory
from handlers.orders import handle_order_tracking
from handlers.delivery import handle_delivery
from handlers.hours import handle_business_hours
from handlers.ordering import handle_order_quantity, handle_checkout, extract_quantity
from handlers.inventory import notify_hq_restock
from handlers.restock import handle_hq_reply
from handlers.internal import (
    handle_internal_arrival, handle_internal_order,
    handle_internal_image, handle_internal_order_from_state,
    handle_internal_notify_register, handle_internal_inventory,
    handle_internal_product_info, handle_internal_spec_query,
    handle_internal_tag_push,
    handle_internal_product_upload, handle_internal_save_images,
    handle_internal_add_images, handle_internal_save_text,
    handle_internal_upload_start, handle_internal_upload_add_media,
    handle_internal_upload_text, handle_internal_upload_finish,
    handle_ambiguous_resolve, handle_name_order_confirm,
    handle_internal_new_product,
    _split_new_product_entries,
    handle_internal_spec_inquiry, handle_spec_inquiry_reply, handle_spec_inquiry_qty,
    handle_internal_price_query,
    handle_internal_add_customer,
    handle_internal_product_info_by_name,
    handle_internal_consumable,
    handle_internal_rebate,
    handle_internal_unfulfilled,
    handle_internal_unclaimed,
    handle_internal_showcase_push,
    handle_internal_label_queue,
    _NEW_PROD_TRIGGER_RE,
    _SAVE_IMG_RE as _INTERNAL_SAVE_IMG_RE,
    _ADD_IMG_RE  as _INTERNAL_ADD_IMG_RE,
    _UPLOAD_TRIGGERS as _INTERNAL_UPLOAD_TRIGGERS,
    _UPLOAD_FINISH_RE as _INTERNAL_UPLOAD_FINISH_RE,
)
from handlers.ad_maker import handle_ad_update_trigger
from handlers.payment import is_payment_message, handle_payment
from handlers.price import handle_price
from handlers.summary import send_pending_summary, send_full_report, build_full_report, build_pending_text
from services.refresh import check_and_refresh
from handlers.escalate import handle_unknown
from handlers.service import (
    handle_bargaining, handle_spec, handle_return,
    handle_address_change, handle_complaint, handle_multi_product,
    handle_urgent_order, handle_image_product, handle_notify_request,
    detect_machine_query, handle_machine_query,
)
from handlers import tone
from handlers.visit import handle_visit, handle_visit_query, is_visit_query
from storage.state import state_manager
from storage.customers import customer_store
from storage.payments import payment_store
from storage.restock import restock_store
from storage.delivery import delivery_store
from storage.issues import issue_store
from storage.pending import pending_store
from storage.notify import notify_store
import storage.visits as visit_store
import storage.queue as queue_store

# 標記完成指令：支援單個、連續多個、範圍、全部
# 格式：✅ I2 I3 P1 / I1-I6已處理 / 全部已處理
_RESOLVE_TRIGGER_RE = re.compile(r"[✅☑️√v]|已處理|已完成|全部已處理")
_RESOLVE_ITEM_RE    = re.compile(r"([PRDIQprdiq])\s*(\d+)")
_RESOLVE_RANGE_RE   = re.compile(r"([PRDIQprdiq])\s*(\d+)\s*[-~–]\s*(?:[PRDIQprdiq]\s*)?(\d+)", re.IGNORECASE)
_RESOLVE_ALL_RE     = re.compile(r"全部已處理|全部\s*(?:已處理|完成|標記)")

# 內部群組呼叫 bot 名字的指令（到貨通知代客登記）
_BOT_NAME_RE = re.compile(r"^(新北小蠻牛|小蠻牛)\s*")
_PROD_CODE_RE_STAFF = re.compile(r"[A-Za-z][A-Za-z0-9\-]{2,}")
_STAFF_NOTIFY_KW = ["需要到貨通知", "要到貨通知", "到貨通知", "需要通知", "要通知"]

# ── 路徑常數 ──────────────────────────────────────────
_BASE_DIR   = Path(__file__).parent          # 專案根目錄（絕對路徑）
_ADMIN_HTML = _BASE_DIR / "static" / "admin.html"

# ── 機器人開關（admin 介面控制）──────────────────────
_bot_active: bool = True

# ── Server 啟動時間 ────────────────────────────────────
import time as _time_module
_server_start_time: float = _time_module.time()

# ── 最近看到的群組 ID（in-memory，重啟後清空）──────────
_seen_group_ids: set[str] = set()

# ── 未知群組追蹤（group_id -> last_seen timestamp）──────
_unknown_groups: dict[str, str] = {}

# ── Webhook 訊息去重（防止重複處理同一則訊息）──────────
_processed_msg_ids: dict[str, float] = {}  # message_id -> timestamp

# ── LINE profile 快取 ─────────────────────────────────
_profile_cache: dict[str, tuple[object, float]] = {}  # user_id -> (profile, timestamp)
_PROFILE_CACHE_TTL = 3600  # 1 hour

def _get_profile_cached(line_api, user_id: str):
    now = _time_module.time()
    cached = _profile_cache.get(user_id)
    if cached and now - cached[1] < _PROFILE_CACHE_TTL:
        return cached[0]
    try:
        profile = line_api.get_profile(user_id)
        _profile_cache[user_id] = (profile, now)
        return profile
    except Exception:
        if cached:
            return cached[0]
        return None

# ── 圖片訊息合併緩衝（等待後續文字，最多 N 秒）──────────
import threading as _threading

_IMG_COALESCE_SECS = 6.0          # 等幾秒看有沒有連續文字（必須 > _TXT_COALESCE_SECS=5.0，讓文字先觸發帶走圖片）

# key: user_id
# value: {"msg_id": str, "context": "user"|"group", "group_id": str|None, "timer": Timer}
_img_buffer: dict[str, dict] = {}
_img_buffer_lock = _threading.Lock()


def _send_reply(reply_token: str | None, to: str, text: str, line_api) -> None:
    """
    優先用 reply_message（免費，不佔月額度），
    token 不存在或已過期才 fallback 到 push_message。
    """
    if reply_token:
        try:
            line_api.reply_message(ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=text)],
            ))
            return
        except Exception as _re:
            print(f"[send_reply] reply_message 失敗: {_re}", flush=True)
            # token 過期或已用過，fallback
    else:
        print(f"[send_reply] 無 reply_token，直接 push to={to[:10]}...", flush=True)
    try:
        line_api.push_message(PushMessageRequest(
            to=to, messages=[TextMessage(text=text)]))
    except Exception as _pe:
        print(f"[send_reply] push_message 也失敗: {_pe}", flush=True)


def _img_buf_flush(user_id: str) -> None:
    """Timer callback：N 秒後沒有文字跟上，單獨處理圖片"""
    with _img_buffer_lock:
        entry = _img_buffer.pop(user_id, None)
    if not entry:
        return   # 已被文字訊息消費

    line_api = _line_api
    reply_token = entry.get("reply_token")

    if entry["context"] == "group":
        # flush 時再次確認是否有 upload session（可能在 3 秒後才建立）
        _up_state = state_manager.get(user_id)
        if _up_state and _up_state.get("action") == "uploading":
            # 補進上架 session，不走圖片識別
            handle_internal_upload_add_media(user_id, entry["msg_id"], "image")
            _upload_timer_reset(user_id, entry["group_id"], reply_token)
            return
        # 內部群 → handle_internal_image
        reply_text = handle_internal_image(entry["group_id"], entry["msg_id"], line_api)
        _send_reply(reply_token, entry["group_id"], reply_text, line_api)
    else:
        # 1:1 客戶
        if _in_quiet_hours():
            queue_store.add(user_id, "image", msg_id=entry["msg_id"])
            return
        reply_text = handle_image_product(user_id, entry["msg_id"], line_api)
        _send_reply(reply_token, user_id, reply_text, line_api)


def _img_buf_set(user_id: str, msg_id: str, context: str, group_id: str | None,
                 reply_token: str | None = None) -> None:
    """存入緩衝並啟動 timer（支援多張圖片累積）"""
    timer = _threading.Timer(_IMG_COALESCE_SECS, _img_buf_flush, args=(user_id,))
    timer.daemon = True
    with _img_buffer_lock:
        old = _img_buffer.get(user_id)
        if old:
            old["timer"].cancel()   # 取消舊 timer（連續圖片更新）
            # 累積多張圖片
            old_ids = old.get("msg_ids", [old["msg_id"]] if old.get("msg_id") else [])
            old_ids.append(msg_id)
            _img_buffer[user_id] = {
                "msg_id": old_ids[0], "msg_ids": old_ids, "context": context,
                "group_id": group_id, "timer": timer,
                "reply_token": reply_token,
            }
        else:
            _img_buffer[user_id] = {
                "msg_id": msg_id, "msg_ids": [msg_id], "context": context,
                "group_id": group_id, "timer": timer,
                "reply_token": reply_token,
            }
    timer.start()


def _img_buf_pop(user_id: str) -> dict | None:
    """文字到達時取出圖片緩衝並取消 timer（若存在）"""
    with _img_buffer_lock:
        entry = _img_buffer.pop(user_id, None)
    if entry:
        entry["timer"].cancel()
    return entry


# ── 內部群組多媒體緩衝（累積圖片/影片，等 PO 文字跟上）────────────
_MEDIA_COALESCE_SECS = 15.0

# key: user_id
# value: {"media": [{"msg_id": str, "type": "image"|"video"}], "group_id": str, "timer": Timer}
_media_buf: dict[str, dict] = {}
_media_buf_lock = _threading.Lock()

# ── 上架 Session 自動完成 timer（圖片傳完 30 秒無後續自動 finish）────
_UPLOAD_AUTO_FINISH_SECS = 30.0
_upload_finish_timers: dict[str, _threading.Timer] = {}
_upload_finish_timers_lock = _threading.Lock()

# ── 新增品項 Session timer（30 秒無後續自動處理）───────────────────────
_NP_SESSION_SECS = 30.0
_new_prod_timers: dict[str, _threading.Timer] = {}
_new_prod_timers_lock = _threading.Lock()


def _new_prod_auto_finish(user_id: str, group_id: str, reply_token: str | None) -> None:
    """30 秒無後續 → 自動處理累積的新增品項訊息"""
    with _new_prod_timers_lock:
        _new_prod_timers.pop(user_id, None)
    _state = state_manager.get(user_id)
    if not _state or _state.get("action") != "new_product_session":
        return
    lines = _state.get("lines", [])
    state_manager.clear(user_id)
    if not lines:
        return
    ack = handle_internal_new_product("\n".join(lines))
    if ack:
        _send_reply(reply_token, group_id, ack, _line_api)


def _new_prod_timer_reset(user_id: str, group_id: str, reply_token: str | None) -> None:
    """每次新訊息加入 session 就重置 30 秒 timer"""
    with _new_prod_timers_lock:
        old = _new_prod_timers.pop(user_id, None)
        if old:
            old.cancel()
        t = _threading.Timer(
            _NP_SESSION_SECS, _new_prod_auto_finish,
            args=(user_id, group_id, reply_token),
        )
        t.daemon = True
        _new_prod_timers[user_id] = t
    t.start()


def _new_prod_timer_cancel(user_id: str) -> None:
    """手動「完成」時取消 auto-finish timer"""
    with _new_prod_timers_lock:
        t = _new_prod_timers.pop(user_id, None)
        if t:
            t.cancel()


def _upload_auto_finish(user_id: str, group_id: str, reply_token: str | None) -> None:
    """30 秒無後續 → 自動執行 upload finish 並通知群組（需有至少 1 張圖片才執行）"""
    with _upload_finish_timers_lock:
        _upload_finish_timers.pop(user_id, None)
    # 確認有圖片，避免只有 PO 文就觸發
    _state = state_manager.get(user_id)
    if not _state:
        return
    _has_media = (
        len(_state.get("current_media", [])) > 0
        or any(len(g.get("media", [])) > 0 for g in _state.get("groups", []))
    )
    if not _has_media:
        print(f"[upload-auto] {user_id[:10]}... 無圖片，跳過 auto-finish", flush=True)
        return
    ack = handle_internal_upload_finish(user_id)
    if ack:
        _send_reply(reply_token, group_id, ack, _line_api)


def _upload_timer_reset(user_id: str, group_id: str, reply_token: str | None) -> None:
    """每次收到新媒體就重置 30 秒 timer"""
    with _upload_finish_timers_lock:
        old = _upload_finish_timers.pop(user_id, None)
        if old:
            old.cancel()
        t = _threading.Timer(
            _UPLOAD_AUTO_FINISH_SECS,
            _upload_auto_finish,
            args=(user_id, group_id, reply_token),
        )
        t.daemon = True
        _upload_finish_timers[user_id] = t
    t.start()


def _upload_timer_cancel(user_id: str) -> None:
    """手動說「完成」時取消 auto-finish timer"""
    with _upload_finish_timers_lock:
        t = _upload_finish_timers.pop(user_id, None)
        if t:
            t.cancel()


def _media_buf_add(user_id: str, msg_id: str, media_type: str, group_id: str,
                   reply_token: str | None = None) -> None:
    """內部群組：累積圖片/影片到緩衝，重置 15 秒 timer"""
    with _media_buf_lock:
        if user_id in _media_buf:
            _media_buf[user_id]["timer"].cancel()
            _media_buf[user_id]["media"].append({"msg_id": msg_id, "type": media_type})
            # 保留最新的 reply_token
            if reply_token:
                _media_buf[user_id]["reply_token"] = reply_token
        else:
            _media_buf[user_id] = {
                "media":       [{"msg_id": msg_id, "type": media_type}],
                "group_id":    group_id,
                "reply_token": reply_token,
            }
        timer = _threading.Timer(_MEDIA_COALESCE_SECS, _media_buf_flush, args=(user_id,))
        timer.daemon = True
        _media_buf[user_id]["timer"] = timer
    timer.start()


def _media_buf_pop(user_id: str) -> dict | None:
    """文字到達時取出媒體緩衝並取消 timer"""
    with _media_buf_lock:
        entry = _media_buf.pop(user_id, None)
    if entry:
        entry["timer"].cancel()
    return entry


def _media_buf_flush(user_id: str) -> None:
    """15 秒後沒有文字跟上：單張圖片 → 商品識別；多張/含影片 → 提示補文字"""
    with _media_buf_lock:
        entry = _media_buf.pop(user_id, None)
    if not entry:
        return
    media    = entry["media"]
    group_id = entry["group_id"]

    line_api = _line_api
    reply_token = entry.get("reply_token")
    if len(media) == 1 and media[0]["type"] == "image":
        # 單張圖片，無文字 → 既有商品識別
        reply_text = handle_internal_image(user_id, media[0]["msg_id"], line_api)
    else:
        n = len(media)
        reply_text = (
            f"收到 {n} 個檔案，請補上指令：\n"
            f"• 上架（圖片 + PO文一起存）\n"
            f"• 存圖 Z3432（只存圖片）"
        )
    _send_reply(reply_token, group_id, reply_text, line_api)


# ── 文字訊息合併緩衝（等待連續訊息，5 秒無新訊息才統一處理）────────
_TXT_COALESCE_SECS = 5.0

# key: user_id
# value: {"lines": [str], "context": "user"|"group", "group_id": str|None, "timer": Timer}
_txt_buffer: dict[str, dict] = {}
_txt_buffer_lock = _threading.Lock()


def _txt_buf_add(user_id: str, text: str, context: str, group_id: str | None,
                 reply_token: str | None = None) -> None:
    """
    新增一行文字到緩衝，並重置 timer（debounce）。
    - 一般訊息：5 秒後觸發
    - 含上架指令（上架/存圖）：延長到 15 秒，等圖片/影片跟上
    reply_token：最後一則訊息的 reply_token，flush 時優先用 reply_message（不佔月額度）
    """
    _is_upload_cmd = ("上架" in text or "存圖" in text or "加圖" in text or "存文" in text)
    wait_secs = _MEDIA_COALESCE_SECS if _is_upload_cmd else _TXT_COALESCE_SECS

    # 取消圖片 timer（圖片不要單獨處理，等文字一起）
    with _img_buffer_lock:
        img_e = _img_buffer.get(user_id)
        if img_e:
            img_e["timer"].cancel()

    # 取消媒體 timer（多媒體等文字一起處理）
    with _media_buf_lock:
        med_e = _media_buf.get(user_id)
        if med_e:
            med_e["timer"].cancel()

    with _txt_buffer_lock:
        if user_id in _txt_buffer:
            _txt_buffer[user_id]["timer"].cancel()
            _txt_buffer[user_id]["lines"].append(text)
            # 每次更新 reply_token（保留最新的）
            if reply_token:
                _txt_buffer[user_id]["reply_token"] = reply_token
            all_text = "\n".join(_txt_buffer[user_id]["lines"])
            if "上架" in all_text or "存圖" in all_text or "加圖" in all_text or "存文" in all_text:
                wait_secs = _MEDIA_COALESCE_SECS
        else:
            _txt_buffer[user_id] = {
                "lines":       [text],
                "context":     context,
                "group_id":    group_id,
                "reply_token": reply_token,
            }
        timer = _threading.Timer(wait_secs, _txt_buf_flush, args=(user_id,))
        timer.daemon = True
        _txt_buffer[user_id]["timer"] = timer

    timer.start()


def _dispatch_internal_fallback(combined: str, group_id: str, line_api) -> str | None:
    """內部群組 fallback dispatch chain"""
    return (
        handle_spec_inquiry_reply(group_id, combined, line_api)
        or handle_spec_inquiry_qty(group_id, combined, line_api)
        or handle_ad_update_trigger(combined, group_id, line_api)
        or _handle_pending_list_command(combined)
        or _handle_staff_resolve(combined)
        or _handle_visit_resolve(combined)
        or _handle_visit_query_command(combined)
        or _handle_spec_rebuild_command(combined)
        or _handle_bot_notify_command(combined)
        or handle_internal_showcase_push(combined, line_api)
        or handle_internal_label_queue(combined)
        or handle_internal_tag_push(combined, line_api)
        or handle_internal_add_customer(combined)
        or handle_internal_notify_register(combined, line_api)
        or handle_internal_arrival(combined, line_api)
        or handle_ambiguous_resolve(group_id, combined)
        or handle_name_order_confirm(group_id, combined)
        or handle_internal_order(combined, line_api, group_id=group_id)
        or handle_internal_rebate(combined, group_id)
        or handle_internal_unfulfilled(combined, group_id)
        or handle_internal_unclaimed(combined, group_id)
        or handle_internal_consumable(combined, group_id)
        or handle_internal_spec_query(combined)
        or handle_internal_product_info(combined, group_id)
        or handle_internal_price_query(combined)
        or handle_internal_inventory(combined, group_id)
        or handle_internal_spec_inquiry(combined, group_id)
        or handle_internal_product_info_by_name(combined, group_id)
    )


def _txt_buf_flush(user_id: str) -> None:
    """
    5 秒後觸發：合併所有緩衝文字（+ 圖片若有），統一處理並 push_message 回覆。
    """
    with _txt_buffer_lock:
        entry = _txt_buffer.pop(user_id, None)
    if not entry:
        return

    combined     = "\n".join(entry["lines"])
    context      = entry["context"]
    group_id     = entry.get("group_id")
    reply_token  = entry.get("reply_token")

    # ── 圖片緩衝（1:1）/ 媒體緩衝（內部群）：一起處理 ──────────────────
    img_e   = _img_buf_pop(user_id)
    media_e = _media_buf_pop(user_id)

    line_api = _line_api

    if True:  # preserve indentation block
        # ── 回覆輔助：優先 reply_message，失敗才 push（共用 _send_reply）──
        _reply_token_used = [False]

        def _send_group_ack(text: str) -> None:
            nonlocal reply_token
            if reply_token and not _reply_token_used[0]:
                try:
                    line_api.reply_message(ReplyMessageRequest(
                        reply_token=reply_token,
                        messages=[TextMessage(text=text)],
                    ))
                    _reply_token_used[0] = True
                    return
                except Exception as _reply_err:
                    print(f"[txt-buf] reply_message 失敗: {_reply_err}", flush=True)
                    reply_token = None  # 標記 token 已失效
            try:
                line_api.push_message(PushMessageRequest(
                    to=group_id, messages=[TextMessage(text=text)]))
            except Exception as _push_err:
                print(f"[txt-buf] push_message 失敗（可能月額度用完）: {_push_err}", flush=True)

        # ════ 內部群組 ════
        if context == "group":
            # 上架 session：per-user（各人獨立上架）
            upload_state = state_manager.get(user_id)
            # 訂單 state：per-group（任何人都能接下「客戶名 N個」）
            group_order_state = state_manager.get(group_id)

            # ── 新增品項：完整訊息 → 優先立即處理（不論是否在 session 中）──
            if _NEW_PROD_TRIGGER_RE.match(combined.strip()) and _split_new_product_entries(combined):
                _new_prod_timer_cancel(user_id)
                state_manager.clear(user_id)
                ack = handle_internal_new_product(combined)
                if ack:
                    _send_group_ack(ack)
                return

            # ── 新增品項 Session 進行中 ──────────────────────────────────
            if upload_state and upload_state.get("action") == "new_product_session":
                if _INTERNAL_UPLOAD_FINISH_RE.match(combined.strip()):   # 「完成」
                    _new_prod_timer_cancel(user_id)
                    lines = upload_state.get("lines", [])
                    state_manager.clear(user_id)
                    ack = handle_internal_new_product("\n".join(lines)) if lines else None
                else:
                    upload_state["lines"].append(combined)
                    state_manager.set(user_id, upload_state)
                    _new_prod_timer_reset(user_id, group_id, reply_token)
                    ack = None  # 靜默等待
                if ack:
                    _send_group_ack(ack)
                return

            # ── 批次上架 Session 進行中 ──────────────────────────────────
            if upload_state and upload_state.get("action") == "uploading":
                if _INTERNAL_UPLOAD_FINISH_RE.match(combined.strip()):
                    _upload_timer_cancel(user_id)
                    # 修 Bug1：完成前先把 img_buf 中還沒處理的圖片補進 session
                    _pending_img = _img_buf_pop(user_id)
                    if _pending_img:
                        handle_internal_upload_add_media(user_id, _pending_img["msg_id"], "image")
                    ack = handle_internal_upload_finish(user_id)
                else:
                    # 收到文字也重置 timer，避免文字和 auto-finish 同時觸發兩個回覆
                    _upload_timer_reset(user_id, group_id, reply_token)
                    ack = handle_internal_upload_text(user_id, combined)
                if ack:
                    _send_group_ack(ack)
                return

            # ── 觸發批次上架 Session ─────────────────────────────────────
            _first_line = combined.split('\n')[0].strip()
            if _first_line in _INTERNAL_UPLOAD_TRIGGERS:
                ack = handle_internal_upload_start(user_id)
                if ack:
                    _send_group_ack(ack)
                _remaining = combined[len(_first_line):].strip()
                if _remaining:
                    # 修 Bug2：_remaining 最後一行若是「完成」→ 拆開，先存 PO 文再 finish
                    _rem_lines = _remaining.splitlines()
                    _finish_idx = next(
                        (i for i, l in enumerate(_rem_lines)
                         if _INTERNAL_UPLOAD_FINISH_RE.match(l.strip())),
                        -1
                    )
                    if _finish_idx >= 0:
                        _po_part = "\n".join(_rem_lines[:_finish_idx]).strip()
                        if _po_part:
                            handle_internal_upload_text(user_id, _po_part)
                        # 補進 img_buf 中還沒處理的圖片
                        _pending_img = _img_buf_pop(user_id)
                        if _pending_img:
                            handle_internal_upload_add_media(user_id, _pending_img["msg_id"], "image")
                        # 補進 media_buf（圖片/影片在 txt_buf 15s 內到達的情況）
                        if media_e:
                            for _mi in media_e["media"]:
                                handle_internal_upload_add_media(user_id, _mi["msg_id"], _mi["type"])
                        ack2 = handle_internal_upload_finish(user_id)
                    else:
                        ack2 = handle_internal_upload_text(user_id, _remaining)
                        # 補進 media_buf（圖片/影片在 txt_buf 15s 內到達的情況）
                        if media_e:
                            for _mi in media_e["media"]:
                                handle_internal_upload_add_media(user_id, _mi["msg_id"], _mi["type"])
                            _upload_timer_reset(user_id, group_id, reply_token)
                    if ack2:
                        _send_group_ack(ack2)
                elif media_e:
                    # 只有 "上架" 沒有 PO文，但有圖片（快速傳送情境）
                    for _mi in media_e["media"]:
                        handle_internal_upload_add_media(user_id, _mi["msg_id"], _mi["type"])
                    _upload_timer_reset(user_id, group_id, reply_token)
                return

            # ── 存圖 Z3432（替換舊圖 + 圖片/影片）──────────────────────
            _save_img_m = _INTERNAL_SAVE_IMG_RE.search(combined)
            if _save_img_m and media_e:
                ack = handle_internal_save_images(
                    _save_img_m.group(1).upper(), media_e["media"])
                if ack:
                    _send_group_ack(ack)
                return

            # ── 加圖 Z3432（保留舊圖，追加新圖片/影片）──────────────────
            _add_img_m = _INTERNAL_ADD_IMG_RE.search(combined)
            if _add_img_m and media_e:
                ack = handle_internal_add_images(
                    _add_img_m.group(1).upper(), media_e["media"])
                if ack:
                    _send_group_ack(ack)
                return

            # ── 存文（純文字 → PO文.txt）────────────────────────────────
            if "存文" in combined:
                ack = handle_internal_save_text(combined)
                if ack:
                    _send_group_ack(ack)
                return

            # ── 既有：圖片識別（從 media_e 取第一張 或 img_e）──────────
            img_pc = None
            source_msg_id = None
            if media_e and media_e["media"] and media_e["media"][0]["type"] == "image":
                source_msg_id = media_e["media"][0]["msg_id"]
            elif img_e:
                source_msg_id = img_e["msg_id"]
            if source_msg_id:
                img_pc = _img_identify_from_buf(source_msg_id)

            if img_pc:
                print(f"[txt-buf] 圖片+文字 → 內部群 {img_pc}")
            elif source_msg_id:
                print("[txt-buf] 媒體+文字 → 內部群，圖片識別失敗")

            # 圖片辨識成功 → 查庫存資訊回覆（不設 state，要下單打完整格式）
            ack = None
            if img_pc:
                from services.ecount import ecount_client as _ec
                _gi = _ec.lookup(img_pc)
                if _gi:
                    from handlers.internal import _format_po, _fmt_stock_lines
                    _po = _format_po(img_pc)
                    _stock = _fmt_stock_lines(_gi)
                    ack = f"{_po}\n{_stock}"

            if not ack:
                    # ── 新增品項觸發 ──
                    if _NEW_PROD_TRIGGER_RE.match(combined.strip()):
                        # 訊息已含品項資料 → 直接處理，不開 session
                        if _split_new_product_entries(combined):
                            ack = handle_internal_new_product(combined)
                        else:
                            state_manager.set(user_id, {
                                "action": "new_product_session",
                                "lines":  [combined],
                            })
                            _new_prod_timer_reset(user_id, group_id, reply_token)
                            ack = "📝 品項建立模式，請依序輸入各品項資料\n完成後傳「完成」，或等待 30 秒自動處理"
                    else:
                        ack = _dispatch_internal_fallback(combined, group_id, line_api)
            else:
                try:
                    # ── 新增品項觸發 ──
                    if _NEW_PROD_TRIGGER_RE.match(combined.strip()):
                        # 訊息已含品項資料 → 直接處理，不開 session
                        if _split_new_product_entries(combined):
                            ack = handle_internal_new_product(combined)
                        else:
                            state_manager.set(user_id, {
                                "action": "new_product_session",
                                "lines":  [combined],
                            })
                            _new_prod_timer_reset(user_id, group_id, reply_token)
                            ack = "📝 品項建立模式，請依序輸入各品項資料\n完成後傳「完成」，或等待 30 秒自動處理"
                    else:
                        print(f"[txt-buf] dispatch_internal_fallback 開始: {combined!r}", flush=True)
                        ack = _dispatch_internal_fallback(combined, group_id, line_api)
                        print(f"[txt-buf] dispatch_internal_fallback 結果: {ack!r}", flush=True)
                except Exception as _ge:
                    print(f"[txt-buf] 內部群處理例外: {type(_ge).__name__}: {_ge}", flush=True)
                    import traceback; traceback.print_exc()
                    ack = f"❌ 處理時發生錯誤：{_ge}"
            if ack:
                _send_group_ack(ack)

        # ════ 1:1 客戶 ════
        else:
            try:
                current_state = state_manager.get(user_id)

                # 1:1 圖片識別：優先從文字提取貨號，沒有才辨識圖片
                img_pcs = []
                _text_codes = _PROD_CODE_RE.findall(combined) if img_e else []
                if _text_codes:
                    # 文字裡有貨號，直接用
                    for _tc in _text_codes:
                        _tc_upper = _tc.upper()
                        if _tc_upper not in img_pcs:
                            img_pcs.append(_tc_upper)
                elif img_e:
                    # 文字沒貨號，才用圖片辨識
                    msg_ids = img_e.get("msg_ids", [img_e["msg_id"]] if img_e.get("msg_id") else [])
                    for _mid in msg_ids:
                        _pc = _img_identify_from_buf(_mid)
                        if _pc and _pc not in img_pcs:
                            img_pcs.append(_pc)
                img_pc = img_pcs[0] if img_pcs else None

                # ── 補圖指令：「補圖 Z3278」+ 圖片 → 存到產品資料夾 ──────
                _add_img_m = re.match(r'補圖\s+([A-Za-z]{1,3}-?\d{3,6}(?:-\d+)?)', combined.strip(), re.IGNORECASE)
                if _add_img_m and img_e:
                    _add_code = _add_img_m.group(1).upper()
                    msg_ids = img_e.get("msg_ids", [img_e["msg_id"]] if img_e.get("msg_id") else [])
                    from handlers.internal import handle_internal_add_images
                    media_items = [{"type": "image", "msg_id": mid} for mid in msg_ids]
                    ack = handle_internal_add_images(_add_code, media_items)
                    _send_reply(reply_token, user_id, ack, line_api)
                    reply_text = None
                    img_pc = None
                    img_pcs = []

                # ── 多張圖片 + 「各X」→ 批次加入購物車 ──────────────
                _each_m = re.search(r'各\s*(\d+)\s*(?:個|箱|件|盒|套|組)?', combined) if img_pcs else None
                if len(img_pcs) > 1 and _each_m and not current_state:
                    _each_qty = int(_each_m.group(1))
                    from services.ecount import ecount_client as _ec
                    from storage import cart as cart_store
                    _added = []
                    for _pc in img_pcs:
                        _info = _ec.lookup(_pc)
                        _pn = (_info.get("name") if _info else None) or _pc
                        cart_store.add_to_cart(user_id, _pc, _pn, _each_qty)
                        _added.append(f"• {_pn} × {_each_qty}")
                    state_manager.set(user_id, {"action": "awaiting_multi_img_confirm"})
                    reply_text = (
                        "收到！已加入購物車：\n"
                        + "\n".join(_added)
                        + "\n\n如果沒有問題我就送出訂單囉"
                    )
                    _send_reply(reply_token, user_id, reply_text, line_api)
                    reply_text = None
                    img_pc = None  # 已處理，不走單張流程

                # ── 圖片 + 文字：規則判斷意圖，再決定處理方式 ──────────────
                if img_pc and not current_state:
                    from services.ecount import ecount_client as _ec
                    _ui   = _ec.lookup(img_pc)
                    _uqty = _ui.get("qty") if _ui else None
                    _un   = (_ui.get("name") if _ui else None) or img_pc

                    # 從文字判斷客戶意圖
                    _txt_lower = combined.lower()
                    _delivery_kw = any(k in combined for k in [
                        "什麼時候到", "何時到", "到貨了嗎", "到了嗎", "到貨時間",
                        "什麼時候會到", "幾時到", "到貨沒", "出貨了嗎", "出貨沒",
                        "寄了嗎", "送到了嗎", "等很久", "到貨", "催貨",
                    ])
                    _inv_kw  = any(k in combined for k in ["有嗎", "有沒有", "庫存", "缺貨", "還有", "有貨", "剩幾"])
                    _price_kw = any(k in combined for k in ["多少", "幾元", "幾錢", "價格", "價錢", "多少錢", "售價"])
                    _qty_m   = re.search(r'(\d+)\s*(?:個|箱|件|盒|套|組)', combined)

                    if _delivery_kw:
                        # 問到貨/出貨 → 轉真人處理（Bot 無法查訂單進度）
                        print(f"[txt-buf] 圖片+文字 → 問到貨時間 {img_pc}", flush=True)
                        issue_store.add(user_id, "delivery_query",
                                        f"（傳圖詢問到貨）{_un}（{img_pc}）：{combined}")
                        reply_text = tone.urgent_order_ack()
                        _send_reply(reply_token, user_id, reply_text, line_api)
                        reply_text = None
                    elif _inv_kw:
                        # 問庫存 → 直接回答，不進購物車流程
                        print(f"[txt-buf] 圖片+文字 → 問庫存 {img_pc}", flush=True)
                        _inv_reply = None
                        if _uqty and _uqty > 0:
                            _inv_reply = handle_inventory(user_id, img_pc, line_api)
                        else:
                            # 沒貨 → 進缺貨詢問流程
                            state_manager.set(user_id, {
                                "action":    "awaiting_restock_qty",
                                "prod_cd":   img_pc,
                                "prod_name": _un,
                            })
                            current_state = state_manager.get(user_id)
                        if _inv_reply:
                            _send_reply(reply_token, user_id, _inv_reply, line_api)
                        reply_text = None  # 已處理，後面跳過
                    elif _price_kw:
                        # 問價格 → 走 price handler
                        print(f"[txt-buf] 圖片+文字 → 問價格 {img_pc}", flush=True)
                        reply_text = handle_price(user_id, img_pc)
                        if reply_text:
                            _send_reply(reply_token, user_id, reply_text, line_api)
                        reply_text = None
                    elif _qty_m and (_uqty and _uqty > 0):
                        # 直接說要幾個 + 有貨 → 設 awaiting_quantity，讓 combined text 觸發結帳
                        _direct_qty = int(_qty_m.group(1))
                        print(f"[txt-buf] 圖片+文字 → 直接下單 {img_pc} x{_direct_qty}", flush=True)
                        state_manager.set(user_id, {
                            "action":    "awaiting_quantity",
                            "prod_cd":   img_pc,
                            "prod_name": _un,
                        })
                        current_state = state_manager.get(user_id)
                    else:
                        # 意圖不明 → 依庫存決定走購物車或缺貨流程
                        if _uqty and _uqty > 0:
                            state_manager.set(user_id, {
                                "action":    "awaiting_quantity",
                                "prod_cd":   img_pc,
                                "prod_name": _un,
                            })
                            print(f"[txt-buf] 圖片+文字 → 有貨 {img_pc}，等待數量", flush=True)
                        else:
                            state_manager.set(user_id, {
                                "action":    "awaiting_restock_qty",
                                "prod_cd":   img_pc,
                                "prod_name": _un,
                            })
                            print(f"[txt-buf] 圖片+文字 → 缺貨 {img_pc}", flush=True)
                        current_state = state_manager.get(user_id)

                elif img_e and not current_state:
                    issue_store.add(user_id, "image_query", "（圖片+文字，圖片無法辨識）")
                    print("[txt-buf] 圖片+文字 → 1:1，圖片識別失敗", flush=True)

                # ── 凍結判斷：有待處理問題 → 完全靜默，等真人標記完成 ──────
                if not current_state and issue_store.has_pending_issue(user_id):
                    print(f"[frozen] {user_id[:10]}... 有待處理問題，靜默", flush=True)
                    return

                if is_payment_message(combined):
                    reply_text = handle_payment(user_id, combined)
                elif current_state:
                    reply_text = _handle_stateful(user_id, combined, current_state, line_api)
                else:
                    intent     = detect_intent(combined)
                    reply_text = _dispatch(user_id, combined, intent, line_api)

                if reply_text:
                    _send_reply(reply_token, user_id, reply_text, line_api)
            except Exception as _ce:
                print(f"[txt-buf] 1:1 客戶處理例外: {_ce}", flush=True)
                import traceback; traceback.print_exc()


def _img_identify_from_buf(msg_id: str) -> str | None:
    """
    從緩衝的 msg_id 下載圖片 → pHash → OCR，回傳 prod_code（或 None）。
    供文字 handler 在「圖片+文字合併」時使用。
    """
    from services.vision import download_image, identify_product, ocr_extract_candidates
    from services.ecount import ecount_client as _ec

    img_bytes = download_image(msg_id)
    if not img_bytes:
        return None

    prod_code = identify_product(img_bytes)
    if not prod_code:
        # 只嘗試「貨號格式」(字母+數字) 或中文詞，跳過純英文短詞（OCR 雜訊）
        _CODE_OR_ZH = re.compile(r'(?:[A-Za-z]\d{2,}|[\u4e00-\u9fff]{2,})')
        for candidate in ocr_extract_candidates(img_bytes):
            if not _CODE_OR_ZH.search(candidate):
                continue
            matched = _ec._resolve_product_code(candidate)
            if matched:
                prod_code = matched
                break
    return prod_code


# ── 離峰時段判斷（00:00 ~ 10:00）──────────────────────
def _in_quiet_hours() -> bool:
    """是否在離峰時段（00:00 ~ 10:00，靜默收集訊息，10:00 自動補處理）"""
    now = datetime.now()
    return now.hour < 10   # 00:00 ~ 09:59


# 離峰時段仍直接回覆的意圖（只有營業時間查詢，其餘全部靜默入佇列）
_QUIET_HOURS_DIRECT_INTENTS = {
    Intent.BUSINESS_HOURS,   # 營業時間（客戶問今天有開嗎，隨時可回）
}


async def _refresh_data_loop():
    """每 2 小時檢查一次資料庫是否需要刷新"""
    while True:
        await asyncio.sleep(2 * 3600)
        try:
            await asyncio.to_thread(check_and_refresh)
        except Exception as e:
            print(f"[scheduler] 資料庫刷新失敗: {e}")


async def _process_queued_messages():
    """處理離峰佇列中所有未處理的訊息（使用 push_message 補發回覆）"""
    msgs = queue_store.get_unprocessed()
    if not msgs:
        print("[queue] 無待處理的離峰訊息")
        return

    print(f"[queue] 開始補處理 {len(msgs)} 則離峰訊息...")

    line_api = _line_api

    if True:  # preserve indentation block
        for msg in msgs:
            try:
                uid = msg["user_id"]
                if msg["msg_type"] == "text":
                    intent = detect_intent(msg["content"])
                    reply_text = _dispatch(uid, msg["content"], intent, line_api)
                elif msg["msg_type"] == "image":
                    reply_text = handle_image_product(uid, msg["msg_id"], line_api)
                else:
                    reply_text = None

                if reply_text:
                    try:
                        line_api.push_message(
                            PushMessageRequest(
                                to=uid,
                                messages=[TextMessage(text=reply_text)],
                            )
                        )
                        print(f"[queue] OK {uid[:10]}... ({msg['msg_type']})", flush=True)
                    except Exception as _push_e:
                        # 429 月額度用完時跳過，不讓 crash 影響整個佇列
                        print(f"[queue] push_message 失敗 (跳過): {_push_e}", flush=True)
            except Exception as e:
                print(f"[queue] 處理失敗 uid={msg['user_id'][:10]}...: {e}", flush=True)
            finally:
                queue_store.mark_processed(msg["id"])

    print(f"[queue] 補處理完成，共 {len(msgs)} 則")


async def _rebate_sync_loop():
    """每日凌晨 1:00 自動同步回饋金資料"""
    while True:
        now = datetime.now()
        target = now.replace(hour=1, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        await asyncio.sleep((target - now).total_seconds())
        try:
            import subprocess as _sp
            _python = sys.executable
            _root = str(Path(__file__).parent)
            _flags = _sp.CREATE_NO_WINDOW if sys.platform == "win32" else 0

            # 同步回饋金
            print("[rebate] 凌晨自動同步本月資料...")
            proc = await asyncio.to_thread(
                _sp.run, [_python, "-m", "scripts.sync_rebate"],
                cwd=_root, capture_output=True, timeout=180, creationflags=_flags,
            )
            if proc.stdout:
                print(proc.stdout.decode("utf-8", errors="replace"), flush=True)
            if proc.returncode != 0 and proc.stderr:
                print(proc.stderr.decode("utf-8", errors="replace"), flush=True)

            # 每月 1~3 日額外同步上月資料
            if now.day <= 3:
                print("[rebate] 月初，額外同步上月資料...")
                proc = await asyncio.to_thread(
                    _sp.run, [_python, "-m", "scripts.sync_rebate", "--last-month"],
                    cwd=_root, capture_output=True, timeout=180, creationflags=_flags,
                )
                if proc.stdout:
                    print(proc.stdout.decode("utf-8", errors="replace"), flush=True)

            # 同步未處理訂單 + 未取訂單
            print("[unfulfilled] 凌晨自動同步未處理+未取訂單...")
            proc = await asyncio.to_thread(
                _sp.run, [_python, "-m", "scripts.sync_unfulfilled"],
                cwd=_root, capture_output=True, timeout=180, creationflags=_flags,
            )
            if proc.stdout:
                print(proc.stdout.decode("utf-8", errors="replace"), flush=True)

        except Exception as e:
            print(f"[rebate/unfulfilled] 自動同步失敗: {e}")


async def _queue_processor_loop():
    """每日 10:00 處理離峰佇列"""
    while True:
        now = datetime.now()
        target = now.replace(hour=10, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        await asyncio.sleep((target - now).total_seconds())
        try:
            await _process_queued_messages()
        except Exception as e:
            print(f"[queue] 處理失敗: {e}")


async def _check_restock_notifications():
    """
    檢查所有等待到貨通知的記錄，查詢 Ecount 庫存：
    有貨 → push 通知客戶 + 標記 notified。
    """
    from services.ecount import ecount_client

    pending = notify_store.get_pending(source=None)  # 全部（客戶+內部群登記）
    if not pending:
        print("[notify] 無等待通知記錄")
        return

    print(f"[notify] 開始檢查 {len(pending)} 筆到貨通知...")
    notified_count = 0

    line_api = _line_api

    if True:  # preserve indentation block
        for record in pending:
            try:
                result = ecount_client.lookup(record["prod_code"])
                qty = result.get("qty") if result else None
                if qty and qty > 0:
                    source = record.get("source", "customer")
                    qty_wanted = record.get("qty_wanted", 1)

                    if source == "staff":
                        # 內部群登記：用訂購格式通知
                        # 換算箱數顯示
                        item = ecount_client.get_product_cache_item(record["prod_code"])
                        box_qty = (item.get("box_qty") or 0) if item else 0
                        if box_qty > 1 and qty_wanted >= box_qty and qty_wanted % box_qty == 0:
                            qty_display = f"{qty_wanted // box_qty}箱"
                        else:
                            qty_display = f"{qty_wanted}個"
                        msg = (
                            f"老闆您好，您之前訂的貨已經到了\n"
                            f"{record['prod_name']}（{record['prod_code']}）× {qty_display}"
                        )
                    else:
                        # 客戶自己登記：用原本的到貨通知格式
                        msg = tone.restock_back_in_stock(
                            name=record["prod_name"],
                            code=record["prod_code"],
                        )

                    line_api.push_message(
                        PushMessageRequest(
                            to=record["user_id"],
                            messages=[TextMessage(text=msg)],
                        )
                    )
                    notify_store.mark_notified(record["id"])
                    notified_count += 1
                    print(
                        f"[notify] OK 已通知 {record['user_id'][:10]}... "
                        f"-> {record['prod_name']} 庫存={qty} (source={source})"
                    )
                else:
                    print(
                        f"[notify] 仍無貨：{record['prod_name']}（{record['prod_code']}）"
                    )
            except Exception as e:
                print(f"[notify] FAIL 通知失敗 id={record['id']}: {e}", flush=True)

    print(f"[notify] 完成，共通知 {notified_count} 筆")


async def _restock_notify_loop():
    """每日 21:00 執行到貨通知"""
    while True:
        now = datetime.now()
        target = now.replace(hour=21, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)
        await asyncio.sleep((target - now).total_seconds())
        try:
            await _check_restock_notifications()
        except Exception as e:
            print(f"[notify] 排程執行失敗: {e}")


async def _followup_loop():
    """每小時檢查對話狀態：24h 提醒 / 48h 清除"""
    await asyncio.sleep(60)   # 啟動後 1 分鐘才第一次跑（避免啟動擠塞）
    while True:
        try:
            from handlers.followup import check_and_followup
            result = await asyncio.to_thread(check_and_followup, _line_api)
            if result["reminded"] or result["expired"]:
                print(f"[followup] 提醒 {result['reminded']} 人，清除 {result['expired']} 筆過期狀態")
        except Exception as e:
            print(f"[followup] 排程執行失敗: {e}")
        await asyncio.sleep(3600)   # 每小時一次


def _startup_verify():
    """啟動時確認各模組功能是否正確載入，結果印至 log"""
    import inspect
    checks = []
    try:
        from handlers.internal import handle_ambiguous_resolve
        checks.append("✅ handle_ambiguous_resolve")
    except Exception as e:
        checks.append(f"❌ handle_ambiguous_resolve ({e})")
    try:
        src = inspect.getsource(
            __import__("services.ecount", fromlist=["EcountClient"]).EcountClient._ensure_product_cache
        )
        if "_ZX_RE" in src:
            checks.append("✅ Ecount Z+英文過濾")
        else:
            checks.append("❌ Ecount Z+英文過濾（未套用）")
    except Exception as e:
        checks.append(f"❌ Ecount Z+英文過濾 ({e})")
    try:
        src2 = inspect.getsource(_txt_buf_flush)
        checks.append("✅ 內部群 try/except" if "內部群處理例外" in src2 else "❌ 內部群 try/except（未套用）")
        checks.append("✅ ambiguous dispatch" if "handle_ambiguous_resolve" in src2 else "❌ ambiguous dispatch（未套用）")
    except Exception as e:
        checks.append(f"❌ dispatch 檢查 ({e})")
    print("[startup] 功能確認：", flush=True)
    for c in checks:
        print(f"  {c}", flush=True)


@asynccontextmanager
async def lifespan(app: FastAPI):
    _startup_verify()               # 確認各模組是否正確載入
    queue_store.init()              # 建立離峰佇列 table
    notify_store._init_db()         # 確保到貨通知 table 存在
    visit_store.init()              # 建立客戶到店記錄 table
    check_and_refresh()             # 啟動時檢查並刷新規格/圖片資料庫
    state_manager.restore_from_db() # 從 SQLite 恢復對話狀態
    _restore_txt_buffer()           # 恢復 reload 前未處理的文字 buffer
    asyncio.create_task(_refresh_data_loop())
    asyncio.create_task(_queue_processor_loop())
    asyncio.create_task(_restock_notify_loop())
    asyncio.create_task(_followup_loop())
    asyncio.create_task(_midnight_inventory_check_loop())
    asyncio.create_task(_rebate_sync_loop())
    # 啟動時若已過 10:00 且佇列有待處理訊息，立刻補發（防止 server 重啟後錯過 10:00 觸發）
    if datetime.now().hour >= 10 and queue_store.count_unprocessed() > 0:
        print(f"[queue] 啟動補處理：發現 {queue_store.count_unprocessed()} 則未處理的離峰訊息")
        asyncio.create_task(_process_queued_messages())
    yield
    # ── shutdown：持久化未處理的文字 buffer ──
    _persist_txt_buffer()


# ── 文字 buffer 持久化（防 reload 丟訊息）────────────────────────
_TXT_BUF_PERSIST_PATH = _BASE_DIR / "data" / "txt_buffer_pending.json"


def _persist_txt_buffer():
    """shutdown 時把未處理的 _txt_buffer 存到 JSON"""
    import json as _json
    with _txt_buffer_lock:
        pending = {}
        for uid, entry in _txt_buffer.items():
            try:
                entry["timer"].cancel()
            except Exception:
                pass
            pending[uid] = {
                "lines":       entry["lines"],
                "context":     entry["context"],
                "group_id":    entry.get("group_id"),
            }
    if pending:
        try:
            _TXT_BUF_PERSIST_PATH.write_text(
                _json.dumps(pending, ensure_ascii=False), encoding="utf-8")
            print(f"[shutdown] 已保存 {len(pending)} 筆未處理的文字 buffer", flush=True)
        except Exception as e:
            print(f"[shutdown] buffer 保存失敗: {e}", flush=True)


def _restore_txt_buffer():
    """startup 時恢復未處理的 _txt_buffer 並立即 flush"""
    import json as _json
    if not _TXT_BUF_PERSIST_PATH.exists():
        return
    try:
        pending = _json.loads(_TXT_BUF_PERSIST_PATH.read_text(encoding="utf-8"))
        _TXT_BUF_PERSIST_PATH.unlink()  # 讀完就刪
        if not pending:
            return
        print(f"[startup] 恢復 {len(pending)} 筆未處理的文字 buffer", flush=True)
        for uid, entry in pending.items():
            # 直接 flush，不再等 5 秒
            combined = "\n".join(entry["lines"])
            context = entry["context"]
            group_id = entry.get("group_id")
            # 塞回 buffer 然後立即 flush
            with _txt_buffer_lock:
                _txt_buffer[uid] = {
                    "lines":       entry["lines"],
                    "context":     context,
                    "group_id":    group_id,
                    "reply_token": None,  # reload 後 token 已過期
                }
                # 不設 timer，直接啟動 flush
                _txt_buffer[uid]["timer"] = _threading.Timer(0, lambda: None)
            _threading.Thread(target=_txt_buf_flush, args=(uid,), daemon=True).start()
    except Exception as e:
        print(f"[startup] buffer 恢復失敗: {e}", flush=True)


app = FastAPI(title="LINE Customer Service Bot", lifespan=lifespan)

_SERVER_START_TIME = datetime.now()

@app.get("/health")
def health():
    """Server 啟動時間確認用，更新後呼叫此端點驗證是否為新版"""
    import os
    return {
        "status": "ok",
        "started_at": _SERVER_START_TIME.strftime("%Y-%m-%d %H:%M:%S"),
        "pid": os.getpid(),
    }

# ── 靜態檔案（產品圖片等）────────────────────────────────────────────
_STATIC_DIR = _BASE_DIR / "static"
_PRODUCTS_IMG_DIR = _STATIC_DIR / "products"
_PRODUCTS_IMG_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")


# ── Admin 介面密碼保護（HTTP Basic Auth）─────────────
@app.middleware("http")
async def admin_auth_middleware(request: Request, call_next):
    if request.url.path.startswith("/admin"):
        auth = request.headers.get("Authorization", "")
        authorized = False
        if auth.startswith("Basic "):
            try:
                decoded  = base64.b64decode(auth[6:]).decode("utf-8")
                username, _, password = decoded.partition(":")
                ok_user = secrets.compare_digest(username, settings.ADMIN_USER)
                ok_pass = secrets.compare_digest(password, settings.ADMIN_PASS)
                authorized = ok_user and ok_pass
            except Exception:
                pass
        if not authorized:
            return Response(
                content="需要登入",
                status_code=401,
                headers={"WWW-Authenticate": 'Basic realm="Admin"'},
            )
    return await call_next(request)


_PHONE_RE = re.compile(r"09\d{8}")


def _user_phone(user_id: str) -> str:
    """從 customer_store 取得客戶手機號碼，供 save_order HP_NO 使用。"""
    info = customer_store.get_by_line_id(user_id)
    return (info or {}).get("phone", "") or ""
_ADDR_RE  = re.compile(
    r"[\u4e00-\u9fff]{2,5}[市縣][\u4e00-\u9fff]{1,5}[區鄉鎮市]"
    r"[\u4e00-\u9fff\d\-]+[路街][^\n,，。]{0,40}"
)


def _auto_save_contact_info(user_id: str, text: str) -> None:
    """從訊息自動擷取電話/住址並儲存到客戶資料"""
    phones = _PHONE_RE.findall(text)
    if phones:
        customer_store.update_phone(user_id, phones[0])
    addr = _ADDR_RE.search(text)
    if addr:
        customer_store.update_address(user_id, addr.group(0).strip())

_YES_KW = {"好", "好了", "是", "對", "ok", "OK", "yes", "YES", "好的", "是的", "對的", "確認", "沒問題", "沒有問題", "可以", "沒錯", "正確", "確定", "下單"}
_NO_KW = {"不", "否", "no", "NO", "不要", "不用", "取消", "算了", "不訂", "不對", "錯了", "不是", "不行", "等不了", "太久", "換一個", "其他"}

_configuration = Configuration(access_token=settings.LINE_CHANNEL_ACCESS_TOKEN)
_api_client = ApiClient(_configuration)
_line_api = MessagingApi(_api_client)
_webhook_handler = WebhookHandler(settings.LINE_CHANNEL_SECRET)


def _resolve_one(kind: str, item_id: int) -> str:
    """標記單一項目並回傳結果文字"""
    if kind == "P":
        ok = payment_store.resolve(item_id)
        return f"✅ #P{item_id} 轉帳確認完成" if ok else f"⚠️ #P{item_id} 找不到或已標記過"
    elif kind == "R":
        ok = restock_store.update_status(item_id, "confirmed")
        return f"✅ #R{item_id} 調貨處理完成" if ok else f"⚠️ #R{item_id} 找不到（可能已處理過）"
    elif kind == "D":
        ok = delivery_store.resolve(item_id)
        return f"✅ #D{item_id} 配送詢問處理完成" if ok else f"⚠️ #D{item_id} 找不到（可能已處理過）"
    elif kind == "I":
        ok = issue_store.resolve(item_id)
        return f"✅ #I{item_id} 問題處理完成" if ok else f"⚠️ #I{item_id} 找不到（可能已處理過）"
    elif kind == "Q":
        ok = pending_store.mark_answered(item_id)
        return f"✅ #Q{item_id} 商品查詢已回覆" if ok else f"⚠️ #Q{item_id} 找不到（可能已處理過）"
    return ""


_PENDING_LIST_RE = re.compile(r'^(清單|待確認|待處理|清單查詢|待確認清單|待處理清單)$')

def _get_ngrok_url_sync() -> str | None:
    """同步取得 ngrok 目前公開網址（daemon thread 可用）"""
    try:
        import httpx
        r = httpx.get("http://localhost:4040/api/tunnels", timeout=2.0)
        tunnels = r.json().get("tunnels", [])
        for t in tunnels:
            if t.get("proto") == "https":
                return t["public_url"]
        if tunnels:
            return tunnels[0]["public_url"]
    except Exception:
        pass
    return None


def _handle_visit_query_command(text: str) -> str | None:
    """內部群「誰要來 / 哪些客人要來」→ 回傳到店預約清單"""
    if not is_visit_query(text):
        return None
    return handle_visit_query()


def _handle_visit_resolve(text: str) -> str | None:
    """內部群「✅ V1」→ 標記客人已到店"""
    if not re.search(r"[✅☑️√v]", text) and "已到" not in text:
        return None
    m = re.search(r"[Vv]\s*(\d+)", text)
    if not m:
        return None
    vid = int(m.group(1))
    ok = visit_store.mark_visited(vid)
    return f"✅ #V{vid} 已到店" if ok else f"找不到 #V{vid}"


def _handle_pending_list_command(text: str) -> str | None:
    """
    偵測「清單」/「待確認」等指令 → 回傳待處理清單文字 + admin 網址。
    """
    if not _PENDING_LIST_RE.match(text.strip()):
        return None
    body = build_pending_text()
    if not body:
        return "目前沒有待處理項目"
    url = _get_ngrok_url_sync()
    suffix = f"\n\n🌐 管理介面：{url}/admin" if url else ""
    return body + suffix


def _handle_staff_resolve(text: str) -> str | None:
    """
    內部群組人工標記完成指令。

    支援格式：
      ✅ I2              → 單筆
      ✅ I2 I3 P1        → 多筆
      I1-I6已處理        → 範圍（同類型）
      全部已處理          → 全部未處理項目一次標記
    """
    # 沒有標記觸發詞（✅/已處理/已完成）→ 直接略過，避免誤判貨號連字號為範圍
    if not _RESOLVE_TRIGGER_RE.search(text):
        return None

    results = []

    # ── 全部已處理 ────────────────────────────────────
    if _RESOLVE_ALL_RE.search(text):
        count = 0
        for p in payment_store.get_pending():
            payment_store.resolve(p["id"]); count += 1
        for r in restock_store.get_unresolved():
            restock_store.update_status(r["id"], "confirmed"); count += 1
        for d in delivery_store.get_pending():
            delivery_store.resolve(d["id"]); count += 1
        for i in issue_store.get_pending():
            issue_store.resolve(i["id"]); count += 1
        for q in pending_store.get_pending():
            pending_store.mark_answered(q["id"]); count += 1
        return f"✅ 全部 {count} 筆待處理項目已標記完成" if count else "目前無待處理項目"

    # ── 範圍標記：I1-I6 / I1~I6 ─────────────────────
    for m in _RESOLVE_RANGE_RE.finditer(text):
        kind  = m.group(1).upper()
        start = int(m.group(2))
        end   = int(m.group(3))
        if start > end:
            start, end = end, start
        if end - start > 50:   # 防止誤操作，最多 50 筆
            results.append(f"⚠️ 範圍太大（{start}~{end}），最多一次 50 筆")
            continue
        for i in range(start, end + 1):
            results.append(_resolve_one(kind, i))

    # ── 範圍標記後不再處理單筆（避免重複）────────────
    if results:
        return "\n".join(results)

    # ── 單筆 / 多筆 ──────────────────────────────────
    has_trigger = bool(_RESOLVE_TRIGGER_RE.search(text))
    if not has_trigger:
        return None

    items = _RESOLVE_ITEM_RE.findall(text)
    if not items:
        return None

    for kind_raw, id_raw in items:
        results.append(_resolve_one(kind_raw.upper(), int(id_raw)))

    return "\n".join(results) if results else None


def _handle_spec_rebuild_command(text: str) -> str | None:
    """「規格更新」→ 立即觸發 import_specs + image_hashes 重建（背景非同步）"""
    if text.strip() not in ("規格更新", "更新規格", "重建規格", "specs rebuild"):
        return None
    from services.refresh import trigger_rebuild
    trigger_rebuild()
    return "🔄 規格DB 更新中（背景執行），完成後自動生效，約需 10~30 秒"


def _handle_bot_notify_command(text: str) -> str | None:
    """
    偵測內部群組呼叫 bot 名字的到貨通知代客登記指令。
    格式：「新北小蠻牛 王小明 AA001 需要到貨通知」

    1. 確認前綴含「新北小蠻牛」或「小蠻牛」
    2. 確認含到貨通知關鍵字
    3. 解析客戶姓名 + 貨號 → 查 DB + Ecount → 登記 notify_store
    回傳回覆文字，或 None（不是此指令）。
    """
    m_bot = _BOT_NAME_RE.match(text)
    if not m_bot:
        return None

    remainder = text[m_bot.end():].strip()
    if not any(kw in remainder for kw in _STAFF_NOTIFY_KW):
        return None

    # ── 提取貨號 ──────────────────────────────────────
    m_prod = _PROD_CODE_RE_STAFF.search(remainder)
    if not m_prod:
        return "請加上貨號哦\n格式：新北小蠻牛 王小明 AA001 需要到貨通知"

    prod_code_raw = m_prod.group(0).upper()

    # ── 提取客戶姓名（移除貨號 + 關鍵字後剩下的文字）──
    name_part = _PROD_CODE_RE_STAFF.sub("", remainder)
    for kw in _STAFF_NOTIFY_KW:
        name_part = name_part.replace(kw, "")
    cust_name = name_part.strip()

    if not cust_name:
        return "請加上客戶姓名哦\n格式：新北小蠻牛 王小明 AA001 需要到貨通知"

    # ── 查 Ecount 確認產品 ─────────────────────────────
    from services.ecount import ecount_client
    result = ecount_client.lookup(prod_code_raw)
    if not result:
        return f"找不到貨號「{prod_code_raw}」，請確認一下哦"

    prod_name = result.get("name") or prod_code_raw

    # ── 查客戶 DB ─────────────────────────────────────
    matches = customer_store.search_by_name(cust_name, real_name_only=True)
    if not matches:
        return (
            f"找不到「{cust_name}」的資料\n"
            f"（客戶需要先傳訊息給我才能登記到貨通知唷）"
        )

    # 取第一個有 LINE user_id 的客戶
    matched = next((m for m in matches if m.get("line_user_id")), None)
    if not matched:
        return (
            f"「{cust_name}」尚未和我互動過，無法自動推送\n"
            f"（請客戶先傳一則訊息給我唷）"
        )

    # ── 登記到貨通知 ───────────────────────────────────
    from storage.notify import notify_store
    notify_store.add(matched["line_user_id"], prod_code_raw, prod_name, 1)

    display = (matched.get("real_name") or matched.get("display_name") or cust_name).strip()
    print(
        f"[notify] 人工登記 → {display}（{matched['line_user_id'][:10]}...）"
        f"：{prod_name}（{prod_code_raw}）"
    )
    return (
        f"✅ 已登記！\n"
        f"👤 {display}\n"
        f"📦 「{prod_name}」（{prod_code_raw}）\n"
        f"到貨後會自動推播通知哦"
    )


@app.post("/admin/push-summary")
async def push_summary():
    """手動觸發待處理清單推送"""
    await asyncio.to_thread(send_pending_summary)
    return {"status": "ok"}


@app.post("/admin/generate-ad")
async def admin_generate_ad():
    """觸發廣告圖更新（同內部群 `廣告圖更新` 指令）"""
    from handlers.ad_maker import handle_ad_update_trigger
    result = await asyncio.to_thread(handle_ad_update_trigger, "廣告圖更新", settings.LINE_GROUP_ID)
    return {"status": "ok", "message": result or "已啟動"}


@app.post("/admin/process-queue")
async def process_queue():
    """手動觸發離峰佇列補處理"""
    pending = queue_store.count_unprocessed()
    await _process_queued_messages()
    return {"status": "ok", "processed": pending}


@app.post("/admin/full-report")
async def push_full_report(days: int = 3):
    """推送完整報表（已處理 + 未處理）到內部群組，預設顯示 3 天內已處理記錄"""
    await asyncio.to_thread(send_full_report, days=days)
    return {"status": "ok", "days": days}


@app.get("/admin/full-report")
async def get_full_report(days: int = 3):
    """直接回傳完整報表文字（不推送到 LINE，方便瀏覽器查看）"""
    report = await asyncio.to_thread(build_full_report, days=days)
    return {"report": report, "days": days}


@app.post("/admin/check-restock-notify")
async def check_restock_notify():
    """手動觸發到貨通知檢查（測試用）"""
    pending = notify_store.count_pending()
    await _check_restock_notifications()
    return {"status": "ok", "checked": pending}


@app.post("/admin/update-customer-label")
async def update_customer_label(
    line_user_id: str = "", db_id: int = 0, label: str = ""
):
    """
    手動設定客戶標籤（顯示在待處理清單的名稱）。
    優先用 line_user_id；若空則用 db_id（適用於 CSV 匯入無 LINE ID 的客戶）。
    label 傳空字串表示清除。

    用法：POST /admin/update-customer-label?line_user_id=Uxxxxxxx&label=王小明
          POST /admin/update-customer-label?db_id=42&label=王小明
    """
    if line_user_id:
        ok = customer_store.update_chat_label(line_user_id, label)
        cust = customer_store.get_by_line_id(line_user_id)
    elif db_id:
        ok = customer_store.update_chat_label_by_db_id(db_id, label)
        cust = customer_store.get_by_db_id(db_id)
    else:
        raise HTTPException(status_code=400, detail="需要 line_user_id 或 db_id")
    if not ok:
        raise HTTPException(status_code=404, detail=f"找不到客戶 (line_user_id={line_user_id!r} db_id={db_id})")
    display = (cust or {}).get("display_name", "")
    print(f"[admin] 客戶標籤更新: uid={line_user_id or 'N/A'} db_id={db_id} → label={label!r} (display_name={display!r})")
    return {"status": "ok", "chat_label": label or None, "display_name": display}


@app.post("/admin/update-customer-tags")
async def update_customer_tags(db_id: int, tags: str = ""):
    """
    更新客戶分類標籤（VIP/野獸國/標準/中句/K霸）。
    tags 為逗號分隔的標籤字串，例如 "VIP,野獸國"。
    傳空字串表示清除所有標籤。

    用法：POST /admin/update-customer-tags?db_id=42&tags=VIP,野獸國
    """
    tag_list = [t.strip() for t in tags.split(",") if t.strip()] if tags.strip() else []
    ok = customer_store.update_tags_by_db_id(db_id, tag_list)
    if not ok:
        raise HTTPException(status_code=404, detail=f"找不到 db_id={db_id}")
    cust = customer_store.get_by_db_id(db_id)
    display = (cust or {}).get("display_name", "")
    print(f"[admin] 客戶分類標籤更新: db_id={db_id} → tags={tag_list!r} (display_name={display!r})")
    return {"status": "ok", "db_id": db_id, "tags": tag_list, "display_name": display}


@app.post("/admin/update-customer-real-name")
async def update_customer_real_name(line_user_id: str, real_name: str = ""):
    """
    手動更新客戶真實姓名（real_name）。
    real_name 傳空字串表示清除。

    用法：POST /admin/update-customer-real-name?line_user_id=Uxxxxxxx&real_name=王家文
    """
    import sqlite3 as _sq
    from storage.customers import DB_PATH as _CUST_DB
    with _sq.connect(str(_CUST_DB)) as _conn:
        cur = _conn.execute(
            "UPDATE customers SET real_name=? WHERE line_user_id=?",
            (real_name.strip() or None, line_user_id)
        )
        _conn.commit()
    if not cur.rowcount:
        raise HTTPException(status_code=404, detail=f"找不到 line_user_id={line_user_id}")
    print(f"[admin] 真實姓名更新: {line_user_id} → real_name={real_name!r}")
    return {"status": "ok", "line_user_id": line_user_id, "real_name": real_name.strip() or None}


# ── 客戶分類標籤設定 ──────────────────────────────────────────────────────

@app.get("/admin/tags-config")
async def get_tags_config():
    """取得目前所有分類標籤清單"""
    from storage.tags_config import load_tags
    return {"tags": load_tags()}


@app.post("/admin/tags-config/add")
async def add_tags_config(tag: str):
    """
    新增一個分類標籤。
    用法：POST /admin/tags-config/add?tag=新標籤
    """
    from storage.tags_config import add_tag
    tag = tag.strip()
    if not tag:
        raise HTTPException(status_code=400, detail="標籤名稱不可為空")
    if len(tag) > 20:
        raise HTTPException(status_code=400, detail="標籤名稱不可超過 20 字")
    tags = add_tag(tag)
    print(f"[admin] 新增分類標籤: {tag!r}")
    return {"status": "ok", "tags": tags}


@app.post("/admin/tags-config/remove")
async def remove_tags_config(tag: str):
    """
    移除一個分類標籤。
    用法：POST /admin/tags-config/remove?tag=舊標籤
    """
    from storage.tags_config import remove_tag
    tags = remove_tag(tag)
    print(f"[admin] 移除分類標籤: {tag!r}")
    return {"status": "ok", "tags": tags}


@app.get("/admin/product-images")
async def list_product_images(code: str = ""):
    """
    列出指定產品代碼的圖片清單（static/products/{CODE}/）。
    不傳 code 則列出全部產品目錄。
    """
    base = _PRODUCTS_IMG_DIR
    if code:
        folder = base / code.upper()
        if not folder.is_dir():
            return {"code": code.upper(), "images": []}
        imgs = sorted(
            p.name for p in folder.iterdir()
            if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif")
        )
        return {"code": code.upper(), "images": imgs}
    else:
        result = {}
        if base.is_dir():
            for d in sorted(base.iterdir()):
                if d.is_dir():
                    imgs = sorted(
                        p.name for p in d.iterdir()
                        if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif")
                    )
                    if imgs:
                        result[d.name] = imgs
        return {"products": result}


@app.get("/product-media/{file_path:path}")
async def serve_product_media(file_path: str):
    """
    轉發產品照片資料夾的檔案（圖片 / 影片），供 LINE push 使用的公開 URL。
    實際路徑：settings.PRODUCT_MEDIA_PATH / file_path
    例：GET /product-media/T1202A.jpg
    """
    import sys as _sys
    if _sys.platform == "win32":
        try:
            import ctypes as _ct
            _ct.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass
    media_file = Path(settings.PRODUCT_MEDIA_PATH) / file_path
    media_file = media_file.resolve()
    media_root = Path(settings.PRODUCT_MEDIA_PATH).resolve()
    if not str(media_file).startswith(str(media_root)):
        raise HTTPException(status_code=403, detail="Forbidden")
    try:
        if not media_file.exists() or not media_file.is_file():
            raise HTTPException(status_code=404, detail=f"找不到檔案: {file_path}")
    except OSError:
        raise HTTPException(status_code=503, detail="產品照片磁碟機未連線")
    # 判斷 Content-Type
    ext = media_file.suffix.lower()
    media_types = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".gif": "image/gif",
        ".mp4": "video/mp4", ".mov": "video/quicktime",
    }
    mt = media_types.get(ext, "application/octet-stream")
    return FileResponse(str(media_file), media_type=mt)


@app.get("/admin/product-media-list")
async def list_product_media(code: str = ""):
    """
    列出產品照片資料夾中匹配指定代碼的所有檔案。
    不傳 code 則列出全部（前 200 個）。
    """
    import sys as _sys
    if _sys.platform == "win32":
        try:
            import ctypes as _ct
            _ct.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
        except Exception:
            pass
    media_dir = Path(settings.PRODUCT_MEDIA_PATH)
    try:
        if not media_dir.is_dir():
            return {"ok": False, "detail": "產品照片磁碟機未連線或路徑不存在"}
    except OSError:
        return {"ok": False, "detail": "磁碟機未連線"}

    exts = {".jpg", ".jpeg", ".png", ".gif", ".mp4", ".mov"}
    if code:
        from handlers.internal import _match_product_media_files
        files = _match_product_media_files(code.upper(), media_dir)
        return {"code": code.upper(), "files": [f.name for f in files]}
    else:
        all_files = sorted(
            f.name for f in media_dir.iterdir()
            if f.is_file() and f.suffix.lower() in exts
        )[:200]
        return {"total": len(all_files), "files": all_files}


@app.post("/admin/update-customer-phone")
async def update_customer_phone(line_user_id: str, phone: str = ""):
    """
    手動更新客戶主電話（強制覆蓋），同時補到 customer_phones 多號表。
    phone 傳空字串表示清除主電話。

    用法：POST /admin/update-customer-phone?line_user_id=Uxxxxxxx&phone=0912345678
    """
    import sqlite3 as _sq
    from storage.customers import DB_PATH as _CUST_DB
    phone_val = phone.strip() or None
    with _sq.connect(str(_CUST_DB)) as _conn:
        cur = _conn.execute(
            "UPDATE customers SET phone=? WHERE line_user_id=?",
            (phone_val, line_user_id)
        )
        _conn.commit()
        if not cur.rowcount:
            raise HTTPException(status_code=404, detail=f"找不到 line_user_id={line_user_id}")
        if phone_val:
            row = _conn.execute(
                "SELECT id FROM customers WHERE line_user_id=?", (line_user_id,)
            ).fetchone()
            if row:
                _conn.execute(
                    "INSERT OR IGNORE INTO customer_phones (customer_id, phone) VALUES (?,?)",
                    (row[0], phone_val)
                )
                _conn.commit()
    print(f"[admin] 電話更新: {line_user_id} → phone={phone_val!r}")
    return {"status": "ok", "line_user_id": line_user_id, "phone": phone_val}


@app.post("/admin/update-customer-ecount-code")
async def update_customer_ecount_code(line_user_id: str, ecount_cust_cd: str = ""):
    """
    手動更新客戶 Ecount 客戶代碼（ecount_cust_cd）。
    傳空字串表示清除。

    用法：POST /admin/update-customer-ecount-code?line_user_id=Uxxxxxxx&ecount_cust_cd=M2509260001
    """
    import sqlite3 as _sq
    from storage.customers import DB_PATH as _CUST_DB
    val = ecount_cust_cd.strip() or None
    with _sq.connect(str(_CUST_DB)) as _conn:
        cur = _conn.execute(
            "UPDATE customers SET ecount_cust_cd=? WHERE line_user_id=?",
            (val, line_user_id)
        )
        _conn.commit()
    if not cur.rowcount:
        raise HTTPException(status_code=404, detail=f"找不到 line_user_id={line_user_id}")
    print(f"[admin] Ecount 代碼更新: {line_user_id} → ecount_cust_cd={val!r}")
    return {"status": "ok", "line_user_id": line_user_id, "ecount_cust_cd": val}


@app.get("/admin/customers")
async def admin_customers(q: str = ""):
    """
    查詢客戶資料（用於取得 line_user_id 再更新標籤）。
    q 可傳姓名或電話，不傳則回傳最近 50 筆。
    """
    if q:
        rows = customer_store.search(q)
    else:
        from storage.customers import DB_PATH as _CUST_DB
        import sqlite3 as _sq
        with _sq.connect(str(_CUST_DB)) as _conn:
            _conn.row_factory = _sq.Row
            rows = [dict(r) for r in _conn.execute(
                "SELECT id, line_user_id, display_name, chat_label, real_name, phone, ecount_cust_cd, tags "
                "FROM customers ORDER BY last_seen DESC LIMIT 50"
            ).fetchall()]
    return {"customers": rows, "count": len(rows)}


@app.post("/admin/merge-customers")
async def merge_customers(keep_id: int, remove_id: int):
    """
    將 remove_id 的客戶資料合併到 keep_id，合併後刪除 remove_id。

    規則：keep 的欄位優先（有值就不蓋掉），remove 有而 keep 沒有的才補入。
    用法：POST /admin/merge-customers?keep_id=268&remove_id=22
    """
    import sqlite3 as _sq
    from storage.customers import DB_PATH as _CUST_DB

    with _sq.connect(str(_CUST_DB)) as conn:
        conn.row_factory = _sq.Row

        keep = conn.execute("SELECT * FROM customers WHERE id=?", (keep_id,)).fetchone()
        remove = conn.execute("SELECT * FROM customers WHERE id=?", (remove_id,)).fetchone()

        if not keep:
            raise HTTPException(status_code=404, detail=f"找不到 keep_id={keep_id}")
        if not remove:
            raise HTTPException(status_code=404, detail=f"找不到 remove_id={remove_id}")

        keep = dict(keep)
        remove = dict(remove)

        # 各欄位：keep 有值就保留，沒有才從 remove 補
        fields = ["phone", "real_name", "address", "note", "ecount_cust_cd",
                  "chat_label", "preferred_ecount_cust_cd"]
        updates = {}
        for f in fields:
            if not keep.get(f) and remove.get(f):
                updates[f] = remove[f]

        if updates:
            set_clause = ", ".join(f"{k}=?" for k in updates)
            conn.execute(
                f"UPDATE customers SET {set_clause} WHERE id=?",
                list(updates.values()) + [keep_id]
            )

        # 搬移 customer_phones
        phones = conn.execute(
            "SELECT phone FROM customer_phones WHERE customer_id=?", (remove_id,)
        ).fetchall()
        for row in phones:
            conn.execute(
                "INSERT OR IGNORE INTO customer_phones (customer_id, phone) VALUES (?,?)",
                (keep_id, row["phone"])
            )

        # 刪除 remove 記錄
        conn.execute("DELETE FROM customers WHERE id=?", (remove_id,))
        conn.commit()

    merged_fields = list(updates.keys()) if updates else []
    print(f"[admin] 合併客戶: keep={keep_id} remove={remove_id} 補入欄位={merged_fields}")
    return {
        "status": "ok",
        "keep_id": keep_id,
        "remove_id": remove_id,
        "merged_fields": merged_fields,
    }


@app.post("/admin/sync-customer-names")
async def sync_customer_names():
    """
    完整版 Ecount 姓名同步（推薦使用這個）：
    用 ecount_cust_cd + 電話號碼 雙重比對，把 Ecount 姓名存入 real_name。
    讓內部群組可以用 Ecount 姓名搜尋到 LINE 客戶（通知登記、代訂單等）。

    需要先有 data/ecount_customers.json（由 scripts/sync_cust_from_web.py 爬取）。
    """
    import json
    from pathlib import Path
    json_path = Path(__file__).parent / "data" / "ecount_customers.json"
    if not json_path.exists():
        return {"status": "error", "message": f"找不到 {json_path.name}，請先執行 scripts/sync_cust_from_web.py"}
    ecount_list = json.loads(json_path.read_text(encoding="utf-8"))
    if not ecount_list:
        return {"status": "error", "message": "ecount_customers.json 內容為空"}
    result = await asyncio.to_thread(customer_store.sync_ecount_names_full, ecount_list)
    total = result["by_code"] + result["by_phone"]
    print(f"[admin] Ecount 姓名同步：{len(ecount_list)} 筆 → 更新 {total} 筆 "
          f"(ecount_cust_cd:{result['by_code']}, 電話:{result['by_phone']}, 跳過:{result['skipped']})")
    return {
        "status": "ok",
        "ecount_total": len(ecount_list),
        "matched_by_code": result["by_code"],
        "matched_by_phone": result["by_phone"],
        "skipped": result["skipped"],
        "total_updated": total,
    }


@app.post("/admin/set-group-address")
async def set_group_address(group_id: str, ecount_cust_cd: str, label: str = ""):
    """
    登記客戶群組預設 Ecount 地址。
    當 bot 收到來自該群組的訊息時，會自動帶入此地址作為訂單預設地址。

    用法：POST /admin/set-group-address?group_id=C1234xxxxx&ecount_cust_cd=M2509260001-2&label=樹林
    """
    customer_store.set_group_address(group_id, ecount_cust_cd, label)
    print(f"[admin] 群組地址設定: {group_id} → {ecount_cust_cd} ({label})")
    return {
        "status": "ok",
        "group_id": group_id,
        "ecount_cust_cd": ecount_cust_cd,
        "label": label,
    }


@app.get("/admin/group-addresses")
async def list_group_addresses():
    """列出所有已登記的客戶群組預設地址"""
    rows = customer_store.list_group_addresses()
    return {"status": "ok", "count": len(rows), "data": rows}


@app.post("/admin/set-customer-preferred-address")
async def set_customer_preferred_address(
    customer_id: int,
    ecount_cust_cd: str,
    label: str = "",
):
    """
    設定某位客戶的個人預設訂單地址。
    叫貨結帳時直接問「是否送到 {label}？」而非列出全部地址選單。
    個人設定優先於群組設定。傳 ecount_cust_cd=（空字串）可清除。

    例：Du 管饒河店
        POST /admin/set-customer-preferred-address
             ?customer_id=XXX&ecount_cust_cd=M2510010009-1&label=饒河

    例：Rachel 管文山店
        POST /admin/set-customer-preferred-address
             ?customer_id=XXX&ecount_cust_cd=M2510010009-4&label=文山
    """
    customer_store.set_preferred_address(customer_id, ecount_cust_cd or None)
    action = "清除" if not ecount_cust_cd else f"設為 {ecount_cust_cd} ({label})"
    print(f"[admin] 客戶 {customer_id} 個人預設地址{action}")
    return {
        "status": "ok",
        "customer_id": customer_id,
        "ecount_cust_cd": ecount_cust_cd or None,
        "label": label,
    }


# ── 根路徑導向 admin ────────────────────────────────────
from fastapi.responses import RedirectResponse

@app.get("/")
async def root():
    return RedirectResponse(url="/admin")

# ── 管理介面 HTML ────────────────────────────────────
@app.get("/admin")
async def admin_ui():
    """管理介面首頁"""
    return FileResponse(_ADMIN_HTML)


@app.get("/admin/ngrok-url")
async def admin_ngrok_url():
    """從 ngrok 本地 API 取得目前的公開網址"""
    import httpx
    try:
        async with httpx.AsyncClient(timeout=2.0) as client:
            r = await client.get("http://localhost:4040/api/tunnels")
            tunnels = r.json().get("tunnels", [])
            for t in tunnels:
                if t.get("proto") == "https":
                    return {"url": t["public_url"]}
            # 沒有 https 就回第一個
            if tunnels:
                return {"url": tunnels[0]["public_url"]}
    except Exception:
        pass
    return {"url": None}


@app.get("/admin/visits")
async def admin_visits():
    """取得預計到店客人清單"""
    return {"visits": visit_store.get_pending()}


@app.post("/admin/visits/{visit_id}/resolve")
async def admin_visit_resolve(visit_id: int):
    """標記客人已到店"""
    ok = visit_store.mark_visited(visit_id)
    return {"ok": ok}


@app.get("/admin/seen-groups")
async def admin_seen_groups():
    """列出這次伺服器啟動後收到訊息的所有群組 ID"""
    return {
        "groups": list(_seen_group_ids),
        "configured": {
            "LINE_GROUP_ID": settings.LINE_GROUP_ID or "(未設定)",
            "LINE_GROUP_ID_HQ": settings.LINE_GROUP_ID_HQ or "(未設定)",
            "LINE_GROUP_ID_SHOWCASE": settings.LINE_GROUP_ID_SHOWCASE or "(未設定)",
        }
    }


# ── 群組管理 ─────────────────────────────────────────
@app.get("/admin/groups")
async def admin_groups():
    """列出所有群組（已知+未知）"""
    known = {
        settings.LINE_GROUP_ID: "內部群",
        settings.LINE_GROUP_ID_HQ: "總公司群",
        settings.LINE_GROUP_ID_SHOWCASE: "看貨群",
    }
    registered = customer_store.list_group_addresses()
    reg_map = {r["group_id"]: r for r in registered}

    groups = []
    # Known groups
    for gid, name in known.items():
        if gid:
            groups.append({
                "group_id": gid,
                "type": "system",
                "label": name,
                "ecount_cust_cd": reg_map.get(gid, {}).get("ecount_cust_cd", ""),
            })
    # Registered customer groups
    for r in registered:
        if r["group_id"] not in known:
            groups.append({
                "group_id": r["group_id"],
                "type": "customer",
                "label": r.get("label", ""),
                "ecount_cust_cd": r.get("ecount_cust_cd", ""),
            })
    # Unknown groups (seen but not registered)
    for gid, last_seen in _unknown_groups.items():
        if gid not in known and gid not in reg_map:
            groups.append({
                "group_id": gid,
                "type": "unknown",
                "label": "",
                "ecount_cust_cd": "",
                "last_seen": last_seen,
            })
    return groups


@app.post("/admin/set-group")
async def admin_set_group(group_id: str, ecount_cust_cd: str = "", label: str = ""):
    """設定群組的客戶代碼和標籤"""
    customer_store.set_group_address(group_id, ecount_cust_cd, label)
    # Remove from unknown if it was there
    _unknown_groups.pop(group_id, None)
    return {"ok": True}


# ── Server 重啟 ──────────────────────────────────────
@app.post("/admin/restart")
async def admin_restart():
    """重啟 server（背景執行，不顯示視窗）
    - tray.py 環境：殺掉整個 process tree，watchdog 會自動重啟
    - 獨立執行：用 VBS 先啟動新 server 再退出
    """
    import threading, os, subprocess, signal

    _vbs = _BASE_DIR / "start_server_bg.vbs"
    _lock = _BASE_DIR / "data" / "tray.lock"
    _under_tray = _lock.exists()

    def _do_restart():
        import time
        time.sleep(0.8)  # 等本次 HTTP 回應送出

        if not _under_tray:
            # 獨立模式：先啟動新 server 再退出
            subprocess.Popen(
                ["wscript.exe", str(_vbs)],
                creationflags=0x08000000,
            )
            time.sleep(1.5)

        # 用 taskkill /T 殺掉整個 process tree（包括 reloader + worker + shell）
        # 這樣 port 才會被正確釋放
        my_pid = os.getpid()
        try:
            # 先找到 reloader 的 PID（父進程）
            import ctypes
            kernel32 = ctypes.windll.kernel32
            # 嘗試殺掉佔用 8000 port 的所有 process
            result = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True, text=True,
                creationflags=0x08000000,
            )
            pids_to_kill = set()
            for line in result.stdout.splitlines():
                if ":8000 " in line and "LISTENING" in line:
                    pid = line.strip().split()[-1]
                    if pid.isdigit():
                        pids_to_kill.add(int(pid))

            for pid in pids_to_kill:
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(pid)],
                        creationflags=0x08000000,
                        capture_output=True,
                        timeout=5,
                    )
                except Exception:
                    pass
        except Exception:
            pass

        # 確保自己也退出
        os._exit(0)

    threading.Thread(target=_do_restart, daemon=True).start()
    mode = "tray watchdog" if _under_tray else "VBS"
    print(f"[admin] 重啟指令收到，模式：{mode}，即將重啟...", flush=True)
    return {"ok": True}


# ── Bot 開關 ─────────────────────────────────────────
@app.post("/admin/bot/on")
async def bot_on():
    """啟動機器人（開始處理 LINE 訊息）"""
    global _bot_active
    _bot_active = True
    print("[admin] 機器人已啟動")
    return {"bot_active": True}


@app.post("/admin/bot/off")
async def bot_off():
    """關閉機器人（暫停處理 LINE 訊息）"""
    global _bot_active
    _bot_active = False
    print("[admin] 機器人已關閉")
    return {"bot_active": False}


# ── API 連線狀態 ─────────────────────────────────────
@app.get("/admin/api-status")
async def admin_api_status():
    """
    回傳各服務連線狀態：
      bot_active   — 機器人是否啟用
      line_api     — LINE API 是否正常
      ecount_api   — Ecount ERP 是否可登入
      db_customers — 客戶資料庫是否可讀
      db_orders    — 訂單/調貨/配送資料庫是否可讀
    """
    # LINE API
    try:
        line_ok = await asyncio.to_thread(lambda: bool(_line_api.get_bot_info()))
    except Exception:
        line_ok = False

    # Ecount API
    try:
        from services.ecount import ecount_client as _ec
        sid = await asyncio.to_thread(_ec._ensure_session)
        ecount_ok = bool(sid)
    except Exception:
        ecount_ok = False

    # SQLite DB 健康檢查（使用絕對路徑避免工作目錄問題）
    def _db_ok(rel: str) -> bool:
        p = _BASE_DIR / rel
        if not p.exists():
            return False
        try:
            with sqlite3.connect(str(p)) as conn:
                conn.execute("SELECT 1")
            return True
        except Exception:
            return False

    db_cust_ok = _db_ok("data/customers.db")
    db_orders_ok = all([
        _db_ok("data/payment_confirmations.db"),
        _db_ok("data/restock_requests.db"),
        _db_ok("data/delivery_inquiries.db"),
    ])

    # Server 狀態
    import os as _os
    uptime_sec = int(_time_module.time() - _server_start_time)
    server_info = {"ok": True, "pid": _os.getpid(), "uptime": uptime_sec}

    # ngrok 狀態（查詢 localhost:4040 API）
    ngrok_ok  = False
    ngrok_url = None
    try:
        import httpx as _httpx
        _r = _httpx.get("http://localhost:4040/api/tunnels", timeout=2)
        if _r.status_code == 200:
            tunnels = _r.json().get("tunnels", [])
            for t in tunnels:
                if t.get("proto") == "https":
                    ngrok_url = t.get("public_url")
                    ngrok_ok  = True
                    break
            if not ngrok_ok and tunnels:
                ngrok_ok  = True
                ngrok_url = tunnels[0].get("public_url")
    except Exception:
        pass

    return {
        "bot_active":    _bot_active,
        "line_api":      line_ok,
        "ecount_api":    ecount_ok,
        "db_customers":  db_cust_ok,
        "db_orders":     db_orders_ok,
        "server":        server_info,
        "ngrok":         {"ok": ngrok_ok, "url": ngrok_url},
    }


# ── 庫存情況表 ────────────────────────────────────────
_inventory_check_cache: dict = {"updated_at": None, "items": []}


def _build_inventory_check() -> list[dict]:
    """
    從 Ecount 取所有有庫存（qty > 0）的品項，
    排除 HH 開頭貨號，逐一比對 specs.json（PO文）與產品照片資料夾。
    只回傳「缺規格 or 缺照片」的品項，按庫存量由大到小排序。
    """
    from services.ecount import ecount_client as _ec
    from handlers.internal import _get_media_dir, _match_product_media_files

    # 1. 讀 specs.json
    specs_path = _BASE_DIR / "data" / "specs.json"
    try:
        import json as _j
        specs: dict = _j.loads(specs_path.read_text(encoding="utf-8")) if specs_path.exists() else {}
    except Exception:
        specs = {}

    # 2. 取有庫存品項（Ecount），排除 HH 開頭貨號
    stock_items = _ec.get_all_stock_products()
    if not stock_items:
        return []

    # 3. 取產品照片目錄
    media_dir = _get_media_dir()

    # 4. 只保留「缺規格 or 缺照片」的品項
    rows = []
    for item in stock_items:
        code = item["code"].upper()
        if code.startswith("HH"):          # 排除 HH 開頭
            continue
        has_spec  = code in specs
        has_photo = bool(media_dir and _match_product_media_files(code, media_dir))
        if has_spec and has_photo:         # 都有 → 跳過
            continue
        rows.append({
            "code":      code,
            "name":      item["name"] or specs.get(code, {}).get("name", ""),
            "qty":       item["qty"],
            "has_spec":  has_spec,
            "has_photo": has_photo,
        })

    # 庫存量由大到小排序
    rows.sort(key=lambda r: -r["qty"])
    return rows


async def _midnight_inventory_check_loop():
    """每天 00:00 自動更新庫存情況表"""
    while True:
        now = datetime.now()
        # 計算距明天 00:00 的秒數
        nxt = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        wait_sec = (nxt - now).total_seconds()
        await asyncio.sleep(wait_sec)
        try:
            items = await asyncio.to_thread(_build_inventory_check)
            _inventory_check_cache["items"]      = items
            _inventory_check_cache["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            print(f"[inventory-check] 每日更新完成，共 {len(items)} 筆有庫存品項")
        except Exception as e:
            print(f"[inventory-check] 每日更新失敗: {e}")


@app.get("/admin/inventory-check")
async def admin_inventory_check(refresh: bool = False):
    """
    回傳有庫存的品項清單，並標記是否有規格表、是否有產品照片。
    ?refresh=true 強制重新查詢 Ecount。
    """
    if refresh or not _inventory_check_cache["updated_at"]:
        try:
            items = await asyncio.to_thread(_build_inventory_check)
            _inventory_check_cache["items"]      = items
            _inventory_check_cache["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        except Exception as e:
            return {"error": str(e), "items": [], "updated_at": None}
    return {
        "updated_at": _inventory_check_cache["updated_at"],
        "items":      _inventory_check_cache["items"],
    }


@app.get("/admin/new-products")
async def admin_new_products():
    """回傳待審核品項清單"""
    from storage.new_products import new_products_store
    return {"items": new_products_store.get_pending()}

@app.post("/admin/new-products/{item_id}/confirm")
async def admin_confirm_new_product(item_id: int):
    """確認品項（從待審核清單移除）"""
    from storage.new_products import new_products_store
    ok = new_products_store.confirm(item_id)
    return {"ok": ok}

@app.post("/admin/new-products/{item_id}/delete")
async def admin_delete_new_product(item_id: int):
    """棄用品項（直接刪除）"""
    from storage.new_products import new_products_store
    ok = new_products_store.delete(item_id)
    return {"ok": ok}


# ── 接手面板 API ──────────────────────────────────────────────────────────
_TAKEOVER_EXPIRE_MINUTES = 60  # 60 分鐘後自動恢復

@app.post("/admin/takeover")
async def admin_takeover(user_id: str, display_name: str = ""):
    """員工接手客戶對話，bot 靜默 60 分鐘"""
    import time as _time
    state_manager.set(user_id, {
        "action":       "human_takeover",
        "taken_at":     _time.time(),
        "display_name": display_name,
    })
    print(f"[takeover] 接手: {display_name}（{user_id[:10]}...）")
    return {"status": "ok", "user_id": user_id}

@app.post("/admin/release")
async def admin_release(user_id: str):
    """釋放客戶，bot 恢復回覆"""
    st = state_manager.get(user_id) or {}
    if st.get("action") == "human_takeover":
        state_manager.clear(user_id)
        print(f"[takeover] 釋放: {user_id[:10]}...")
    return {"status": "ok", "user_id": user_id}

@app.get("/admin/takeovers")
async def admin_list_takeovers():
    """列出目前所有接手中的客戶"""
    import time as _time
    now = _time.time()
    results = []
    for uid, st in state_manager.all_states().items():
        if st.get("action") == "human_takeover":
            taken_at  = st.get("taken_at", now)
            elapsed   = int((now - taken_at) / 60)
            remaining = max(0, _TAKEOVER_EXPIRE_MINUTES - elapsed)
            if remaining == 0:
                state_manager.clear(uid)
                continue
            cust = customer_store.get_by_line_id(uid)
            results.append({
                "user_id":       uid,
                "display_name":  st.get("display_name") or (cust.get("real_name") or cust.get("display_name") if cust else uid[:10]),
                "elapsed_min":   elapsed,
                "remaining_min": remaining,
            })
    return results

@app.get("/admin/recent-customers")
async def admin_recent_customers():
    """取得最近 30 位互動的客戶（供接手面板使用）"""
    import sqlite3 as _sq, os as _os
    db_path = _os.path.join(_os.path.dirname(__file__), "data", "customers.db")
    try:
        with _sq.connect(db_path) as conn:
            rows = conn.execute("""
                SELECT line_user_id, real_name, display_name, chat_label
                FROM customers
                WHERE line_user_id IS NOT NULL
                ORDER BY last_seen DESC
                LIMIT 30
            """).fetchall()
        takeover_uids = {uid for uid, st in state_manager.all_states().items()
                         if st.get("action") == "human_takeover"}
        return [
            {
                "user_id":      r[0],
                "name":         r[1] or r[3] or r[2] or r[0][:10],
                "is_takeover":  r[0] in takeover_uids,
            }
            for r in rows
        ]
    except Exception as e:
        return []


# ── 到貨通知管理 ──────────────────────────────────────────
@app.get("/admin/notify")
async def admin_notify_list():
    """取得全部到貨通知記錄"""
    from services.ecount import ecount_client
    records = notify_store.get_all()
    for r in records:
        # 附加客戶名稱
        cust = customer_store.get_by_line_id(r["user_id"])
        r["customer_name"] = (cust.get("real_name") or cust.get("display_name") or r["user_id"][:10]) if cust else r["user_id"][:10]
        # 附加單位顯示
        item = ecount_client.get_product_cache_item(r["prod_code"])
        box_qty = (item.get("box_qty") or 0) if item else 0
        prod_unit = (item.get("unit") or "") if item else ""
        qty = r["qty_wanted"]
        if prod_unit == "箱":
            # 產品本身以箱計
            r["qty_display"] = f"{qty}箱"
            r["unit"] = "箱"
            r["box_qty"] = box_qty or 1
        elif box_qty > 1 and qty >= box_qty and qty % box_qty == 0:
            r["qty_display"] = f"{qty // box_qty}箱"
            r["unit"] = "箱"
            r["box_qty"] = box_qty
        else:
            r["qty_display"] = f"{qty}個"
            r["unit"] = "個"
            r["box_qty"] = box_qty
    return records

@app.put("/admin/notify/{notify_id}")
async def admin_notify_update(notify_id: int, request: Request):
    """更新到貨通知"""
    body = await request.json()
    notify_store.update(notify_id, **body)
    return {"ok": True}

@app.delete("/admin/notify/{notify_id}")
async def admin_notify_delete(notify_id: int):
    """刪除到貨通知"""
    notify_store.delete(notify_id)
    return {"ok": True}

# ── 回饋金 ──────────────────────────────────────────────
@app.get("/admin/rebate")
async def admin_rebate(sync: bool = False):
    """取得當月回饋金計算結果，sync=true 時先從 Ecount 同步"""
    if sync:
        try:
            import subprocess as _sp
            _python = sys.executable
            _root = str(Path(__file__).parent)
            _flags = _sp.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            proc = await asyncio.to_thread(
                _sp.run, [_python, "-m", "scripts.sync_rebate"],
                cwd=_root, capture_output=True, timeout=180, creationflags=_flags,
            )
            if proc.stdout:
                print(proc.stdout.decode("utf-8", errors="replace"), flush=True)
        except Exception as e:
            print(f"[rebate] 同步失敗: {e}")
    from services.rebate import calculate_rebates
    return await asyncio.to_thread(calculate_rebates)

@app.get("/admin/rebate/approaching")
async def admin_rebate_approaching():
    """1~14日：上月達標客戶；15日起：當月快接近達成"""
    from datetime import datetime as _dt
    day = _dt.now().day
    if day < 15:
        from services.rebate import get_last_month_achievers
        return await asyncio.to_thread(get_last_month_achievers)
    else:
        from services.rebate import get_approaching_customers
        return await asyncio.to_thread(get_approaching_customers)


@app.post("/webhook")
async def webhook(request: Request):
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()

    # ── chat_mode_changed 事件（員工點「使用手動聊天」）────────────────────
    try:
        import json as _json
        _payload = _json.loads(body)
        for _ev in _payload.get("events", []):
            _ev_type = _ev.get("type", "")
            # 記錄所有非 message 事件，方便除錯
            if _ev_type != "message":
                print(f"[webhook] 非message事件: type={_ev_type} ev={_json.dumps(_ev, ensure_ascii=False)[:300]}", flush=True)
            if _ev_type == "chat_mode_changed":
                _src  = _ev.get("source") or {}
                _uid  = _src.get("userId", "") or _src.get("groupId", "") or _src.get("roomId", "")
                _mode = _ev.get("mode", "")
                if _uid:
                    if _mode == "standby":
                        state_manager.set(_uid, {"action": "human_takeover"})
                        print(f"[chat_mode] 員工接手 {_uid[:10]}... → bot 靜默")
                    elif _mode == "active":
                        if (state_manager.get(_uid) or {}).get("action") == "human_takeover":
                            state_manager.clear(_uid)
                        print(f"[chat_mode] bot 恢復 {_uid[:10]}...")
    except Exception as _e:
        print(f"[webhook] chat_mode parse 錯誤: {_e}")

    try:
        await asyncio.to_thread(_webhook_handler.handle, body.decode("utf-8"), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")
    return "OK"


@_webhook_handler.add(MessageEvent, message=TextMessageContent)
def on_message(event: MessageEvent):
    # ── 訊息去重 ─────────────────────────────────────────
    msg_id = event.message.id
    now = _time_module.time()
    if msg_id in _processed_msg_ids:
        return  # Already processed
    _processed_msg_ids[msg_id] = now
    # Clean entries older than 60s periodically
    if len(_processed_msg_ids) > 500:
        cutoff = now - 60
        for k in [k for k, v in _processed_msg_ids.items() if v < cutoff]:
            _processed_msg_ids.pop(k, None)

    # 印出來源 ID，方便取得兩個 Group ID
    source_type = event.source.type
    if source_type == "group":
        group_id = event.source.group_id
        # in-memory 記錄所有見過的群組 ID
        if group_id not in _seen_group_ids:
            _seen_group_ids.add(group_id)
        print(f"[GROUP] {group_id}", flush=True)
        # 寫檔備份（含 user_id + display_name）
        try:
            _uid = event.source.user_id
            # 嘗試取得顯示名稱
            _display = ""
            try:
                with ApiClient(_configuration) as _api:
                    _profile = MessagingApi(_api).get_group_member_profile(group_id, _uid)
                    _display = _profile.display_name or ""
            except Exception:
                pass
            log_path = _BASE_DIR / "data" / "group_ids.txt"
            with open(str(log_path), "a", encoding="utf-8") as f:
                from datetime import datetime as _dt
                f.write(f"{_dt.now().strftime('%m-%d %H:%M')} | {group_id} | {_uid} | {_display}\n")
        except Exception as e:
            print(f"[GROUP] 寫檔失敗: {e}", flush=True)

    user_id = event.source.user_id
    text = event.message.text.strip()

    with ApiClient(_configuration) as api_client:
        line_api = MessagingApi(api_client)

        # 自動記錄客戶 LINE ID + 顯示名稱
        profile = _get_profile_cached(line_api, user_id)
        display_name = profile.display_name if profile else ""
        try:
            if display_name:  # display_name 空白時不建立空記錄
                customer_store.upsert_from_line(user_id, display_name)
        except Exception as e:
            print(f"[customers] upsert 失敗: {e}")

        # 自動從訊息擷取電話/住址
        try:
            _auto_save_contact_info(user_id, text)
        except Exception as e:
            print(f"[customers] 聯絡資訊擷取失敗: {e}")

        # ── 總公司群組回覆 → 特殊處理，不走一般客戶邏輯 ──────
        if (source_type == "group"
                and settings.LINE_GROUP_ID_HQ
                and event.source.group_id == settings.LINE_GROUP_ID_HQ):
            hq_group_id = event.source.group_id
            # 新增品項 Session 進行中
            _hq_np_state = state_manager.get(user_id)
            if _hq_np_state and _hq_np_state.get("action") == "new_product_session":
                if _INTERNAL_UPLOAD_FINISH_RE.match(text.strip()):
                    _new_prod_timer_cancel(user_id)
                    _lines = _hq_np_state.get("lines", [])
                    state_manager.clear(user_id)
                    ack = handle_internal_new_product("\n".join(_lines)) if _lines else None
                else:
                    _hq_np_state["lines"].append(text)
                    state_manager.set(user_id, _hq_np_state)
                    _new_prod_timer_reset(user_id, hq_group_id, event.reply_token)
                    ack = None
                if ack:
                    line_api.reply_message(ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=ack)],
                    ))
                return
            # 新增品項觸發
            if _NEW_PROD_TRIGGER_RE.match(text.strip()):
                # 訊息已含品項資料 → 直接處理，不開 session
                if _split_new_product_entries(text):
                    ack = handle_internal_new_product(text)
                    if ack:
                        line_api.reply_message(ReplyMessageRequest(
                            reply_token=event.reply_token,
                            messages=[TextMessage(text=ack)],
                        ))
                    return
                state_manager.set(user_id, {
                    "action": "new_product_session",
                    "lines":  [text],
                })
                _new_prod_timer_reset(user_id, hq_group_id, event.reply_token)
                ack = "📝 品項建立模式，請依序輸入各品項資料\n完成後傳「完成」，或等待 30 秒自動處理"
                line_api.reply_message(ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=ack)],
                ))
                return
            # ── 批次上架 Session 進行中 ──
            _hq_upload_state = state_manager.get(user_id)
            if _hq_upload_state and _hq_upload_state.get("action") == "uploading":
                if _INTERNAL_UPLOAD_FINISH_RE.match(text.strip()):
                    _upload_timer_cancel(user_id)
                    ack = handle_internal_upload_finish(user_id)
                else:
                    _upload_timer_reset(user_id, hq_group_id, event.reply_token)
                    ack = handle_internal_upload_text(user_id, text)
                if ack:
                    line_api.reply_message(ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=ack)],
                    ))
                return

            # ── 標籤指令 ──
            if text.strip().splitlines()[0].strip().startswith("標籤"):
                ack = handle_internal_label_queue(text)
                if ack:
                    line_api.reply_message(ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=ack)],
                    ))
                return

            # ── 觸發批次上架 Session ──
            if text.strip() in _INTERNAL_UPLOAD_TRIGGERS:
                ack = handle_internal_upload_start(user_id)
                if ack:
                    line_api.reply_message(ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=ack)],
                    ))
                _upload_timer_reset(user_id, hq_group_id, event.reply_token)
                return

            # 其他 HQ 訊息
            ack = handle_hq_reply(text, line_api)
            if ack:
                line_api.reply_message(
                    ReplyMessageRequest(
                        reply_token=event.reply_token,
                        messages=[TextMessage(text=ack)],
                    )
                )
            return

        # ── 看貨群 → 只回覆營業時間查詢，其餘全部靜默 ──────────────────
        if (source_type == "group"
                and settings.LINE_GROUP_ID_SHOWCASE
                and event.source.group_id == settings.LINE_GROUP_ID_SHOWCASE):
            if detect_intent(text) == Intent.BUSINESS_HOURS:
                reply = handle_business_hours(text)
                if reply:
                    line_api.reply_message(
                        ReplyMessageRequest(
                            reply_token=event.reply_token,
                            messages=[TextMessage(text=reply)],
                        )
                    )
            return  # 非營業時間查詢 → 靜默不回

        # ── 內部群組 ──────────────────────────────────────────
        if (source_type == "group"
                and settings.LINE_GROUP_ID
                and event.source.group_id == settings.LINE_GROUP_ID):
            # 簡單查詢指令：直接 reply 不走 buffer（避免 token 過期）
            from handlers.internal import handle_internal_spec_query as _spec_q
            _quick_reply = (
                handle_internal_rebate(text)
                or handle_internal_unfulfilled(text)
                or handle_internal_unclaimed(text)
                or _spec_q(text)
            )
            if _quick_reply:
                _send_reply(event.reply_token, event.source.group_id,
                            _quick_reply, line_api)
                return
            # 其餘走文字緩衝，5 秒後統一處理
            _txt_buf_add(user_id, text, "group", event.source.group_id,
                         reply_token=event.reply_token)
            return

        # ── 機器人關閉中 → 只擋客戶訊息，內部群/總公司群指令不受影響 ──
        if not _bot_active:
            return

        # ── 預設地址偵測（立即寫入 state，flush 時就能用）─────────────
        _cust_for_pref   = customer_store.get_by_line_id(user_id)
        _individual_pref = (
            _cust_for_pref.get("preferred_ecount_cust_cd") if _cust_for_pref else None
        )
        if _individual_pref:
            state_manager.set_group_cust_cd(user_id, _individual_pref)
        elif source_type == "group" and event.source.group_id:
            gid = event.source.group_id
            grp = customer_store.get_group_default(gid)
            if grp:
                state_manager.set_group_cust_cd(user_id, grp["ecount_cust_cd"])
                print(f"[group] 群組預設地址: {grp['ecount_cust_cd']} ({grp.get('label', '')})")
            else:
                print(f"[group] 未知客戶群組: {gid}")
                _unknown_groups[gid] = datetime.now().strftime("%Y-%m-%d %H:%M")

        # ── 離峰時段（00:00 ~ 10:00）→ 靜默收集，不進緩衝 ────────────
        if _in_quiet_hours() and source_type == "user":
            current_state = state_manager.get(user_id)
            if not current_state:
                _q_intent = detect_intent(text)
                if _q_intent not in _QUIET_HOURS_DIRECT_INTENTS:
                    queue_store.add(user_id, "text", content=text)
                    return

        # ── 真人介入中 → 凍結 bot，不進緩衝 ────────────────────────
        _check_id = group_id if source_type == "group" else user_id
        if (state_manager.get(_check_id) or {}).get("action") == "human_takeover":
            print(f"[escalate] 真人介入中（群組），靜默 | {_check_id[:10]}...: {text!r}")
            return
        if source_type == "user" and (
            issue_store.has_pending_issue(user_id)
            or delivery_store.has_pending(user_id)
            or pending_store.has_pending(user_id)
        ):
            print(f"[escalate] 真人介入中，靜默 | {user_id[:10]}...: {text!r}")
            return

        # ── 1:1 客戶 → 存文字緩衝，5 秒後統一處理（含圖片合併）────────
        _txt_buf_add(user_id, text, "user", None, reply_token=event.reply_token)


@_webhook_handler.add(MessageEvent, message=ImageMessageContent)
def on_image_message(event: MessageEvent):
    """
    處理傳來的產品圖片（客戶 or 內部群組）

    收到圖片後先存入緩衝，等待 _IMG_COALESCE_SECS 秒：
    - 若後續有文字（如「5個」「幫XX訂5個」）→ 由 on_message 合併處理
    - 若無後續文字 → timer 到期自動單獨處理（push_message）
    """
    message_id  = event.message.id
    source_type = event.source.type
    user_id     = event.source.user_id

    # ── 內部群組圖片 ─────────────────────────────────────────────────────
    if (source_type == "group"
            and settings.LINE_GROUP_ID
            and event.source.group_id == settings.LINE_GROUP_ID):
        # 批次上架 Session 進行中 → 直接追加到 state（不走緩衝）
        _st = state_manager.get(user_id)
        if _st and _st.get("action") == "uploading":
            handle_internal_upload_add_media(user_id, message_id, "image")
            _upload_timer_reset(user_id, event.source.group_id, event.reply_token)
            return  # 靜默，上架作業全程不回覆，直到完成才通知
        # 一般模式 → 存媒體緩衝，等文字跟上；15s timer 到期再單獨處理
        _media_buf_add(user_id, message_id, "image", event.source.group_id,
                       reply_token=event.reply_token)
        return

    # ── 總公司群圖片：上架 Session 進行中 → 追加到 state ──
    if (source_type == "group"
            and settings.LINE_GROUP_ID_HQ
            and event.source.group_id == settings.LINE_GROUP_ID_HQ):
        _st = state_manager.get(user_id)
        if _st and _st.get("action") == "uploading":
            handle_internal_upload_add_media(user_id, message_id, "image")
            _upload_timer_reset(user_id, event.source.group_id, event.reply_token)
            return
        # 非上架 session → 靜默
        return

    # 其他群組靜默
    if source_type == "group":
        return

    # 機器人關閉中 → 靜默不回應（客戶圖片）
    if not _bot_active:
        return

    # 自動記錄客戶資料（不影響緩衝邏輯）
    with ApiClient(_configuration) as api_client:
        line_api = MessagingApi(api_client)
        profile = _get_profile_cached(line_api, user_id)
        if profile and profile.display_name:  # display_name 空白時不建立空記錄
            customer_store.upsert_from_line(user_id, profile.display_name)

    # ── 離峰時段 → 直接收集，不走緩衝 ──
    if _in_quiet_hours():
        queue_store.add(user_id, "image", msg_id=message_id)
        return

    # 存入緩衝，等待後續文字（如「5個」「有貨嗎」）
    _img_buf_set(user_id, message_id, "user", None, reply_token=event.reply_token)


@_webhook_handler.add(MessageEvent, message=VideoMessageContent)
def on_video_message(event: MessageEvent):
    """
    處理傳來的影片（僅內部群組）
    存入媒體緩衝，等待 PO 文字跟上；15 秒後無文字則提示補充。
    """
    source_type = event.source.type
    user_id     = event.source.user_id

    if (source_type == "group"
            and settings.LINE_GROUP_ID
            and event.source.group_id == settings.LINE_GROUP_ID):
        # 批次上架 Session 進行中 → 直接追加到 state
        _st = state_manager.get(user_id)
        if _st and _st.get("action") == "uploading":
            handle_internal_upload_add_media(user_id, event.message.id, "video")
            _upload_timer_reset(user_id, event.source.group_id, event.reply_token)
            return  # 靜默，上架作業全程不回覆，直到完成才通知
        _media_buf_add(user_id, event.message.id, "video", event.source.group_id,
                       reply_token=event.reply_token)
        return

    # ── 總公司群影片：上架 Session 進行中 → 追加到 state ──
    if (source_type == "group"
            and settings.LINE_GROUP_ID_HQ
            and event.source.group_id == settings.LINE_GROUP_ID_HQ):
        _st = state_manager.get(user_id)
        if _st and _st.get("action") == "uploading":
            handle_internal_upload_add_media(user_id, event.message.id, "video")
            _upload_timer_reset(user_id, event.source.group_id, event.reply_token)
            return


def _handle_stateful(
    user_id: str, text: str, state: dict, line_api: MessagingApi
) -> str:
    """處理多輪對話狀態"""
    action = state.get("action")

    # ── 等待客戶選擇款式（多筆匹配）──────────────────
    if action == "awaiting_product_clarify":
        from handlers.inventory import _query_single_product
        candidates: list = state.get("candidates", [])   # [(code, name), ...]

        chosen_code = None

        # 用數字選（「1」「2」「3」）
        qty = extract_quantity(text)
        if qty and 1 <= qty <= len(candidates):
            chosen_code = candidates[qty - 1][0]
        else:
            # 用品名子字串選（說「大顆」「A款」等）
            t = text.strip()
            for code, name in candidates:
                if t in name or name in t:
                    chosen_code = code
                    break

        if chosen_code:
            state_manager.clear(user_id)
            return _query_single_product(user_id, chosen_code, line_api)
        else:
            # 選不到 → 再問一次
            return tone.ask_product_clarify(state.get("keyword", ""), candidates)

    # ── 等待數量：客戶確認要購買幾個 ──────────────
    if action == "awaiting_quantity":
        qty = extract_quantity(text)
        if qty:
            if state.get("from_image"):
                # 圖片識別下單 → 先顯示確認框，等待「確認」才建單
                prod_cd   = state.get("prod_cd", "")
                prod_name = state.get("prod_name", "此商品")
                state_manager.set(user_id, {
                    "action":    "awaiting_image_order_confirm",
                    "prod_cd":   prod_cd,
                    "prod_name": prod_name,
                    "qty":       qty,
                })
                return (
                    f"確認是這款對吧～\n"
                    f"{prod_name}（{prod_cd}）× {qty} 個"
                )
            else:
                # 一般文字下單 → 直接加購物車
                state_manager.clear(user_id)
                return handle_order_quantity(user_id, text, state, line_api)
        elif any(kw in text for kw in ["不要", "算了", "取消", "不訂", "不用"]):
            state_manager.clear(user_id)
            return f"好的{tone.suffix_light()} 已取消，{tone.boss()}有需要再找我哦"
        else:
            # 數量不明，保留狀態再問一次
            prod_name = state.get("prod_name", "這款")
            return tone.ask_quantity(prod_name)

    # ── 多圖+各X 確認：等待確認送出或其他回覆靜默 ──────
    elif action == "awaiting_multi_img_confirm":
        if any(kw == text.strip() for kw in _YES_KW):
            state_manager.clear(user_id)
            from handlers.ordering import handle_checkout
            return handle_checkout(user_id, line_api)
        else:
            # 非確認 → 靜默，加入待處理
            state_manager.clear(user_id)
            issue_store.add(user_id, "multi_img_order", f"客戶回覆：{text}")
            return None

    # ── 圖片下單確認框：等待「確認」或「取消」──────
    elif action == "awaiting_image_order_confirm":
        prod_cd   = state.get("prod_cd", "")
        prod_name = state.get("prod_name", "此商品")
        qty       = state.get("qty", 1)
        if any(kw in text for kw in _YES_KW):
            state_manager.clear(user_id)
            # 加入購物車後立即結帳
            from handlers.ordering import handle_checkout
            from storage import cart as cart_store
            cart_store.add_item(user_id, prod_cd, prod_name, qty)
            return handle_checkout(user_id, line_api)
        elif any(kw in text for kw in _NO_KW):
            state_manager.clear(user_id)
            return f"好的{tone.suffix_light()} 已取消，{tone.boss()}有需要再找我哦"
        else:
            # 非確認/取消 → 重新顯示確認框
            return (
                f"確認是這款對吧～\n"
                f"{prod_name}（{prod_cd}）× {qty} 個"
            )

    # ── 等待數量（缺貨調貨）：問客戶要幾個，再通知總公司 ─────
    elif action == "awaiting_restock_qty":
        qty = extract_quantity(text)
        if qty:
            prod_name = state.get("prod_name", "此商品")
            prod_cd = state.get("prod_cd", "")
            state_manager.clear(user_id)
            from storage.restock import restock_store
            restock_store.add(user_id, prod_name, prod_cd, qty)
            notify_hq_restock(prod_name, qty, line_api)
            return tone.restock_inquiry_sent(prod_name, qty)
        elif any(kw in text for kw in ["不要", "算了", "取消", "不訂", "不用",
                                       "收到", "好", "好的", "知道了", "了解",
                                       "知道", "ok", "OK", "謝謝", "感謝"]):
            state_manager.clear(user_id)
            return f"好的{tone.suffix_light()} 有需要再找我哦"
        else:
            prod_name = state.get("prod_name", "這款")
            return tone.ask_quantity(prod_name)

    # ── 等待客戶確認是否願意等叫貨 ────────────────────────
    elif action == "awaiting_wait_confirm":
        prod_name = state.get("prod_name", "此商品")
        prod_cd = state.get("prod_cd", "")
        qty = state.get("qty", 1)
        wait_time = state.get("wait_time", "")
        restock_id = state.get("restock_id")

        if any(kw in text for kw in _YES_KW):
            from storage.restock import restock_store
            from storage.customers import customer_store as _cs
            from services.ecount import ecount_client

            # 多地址 → 先問客戶選哪個地址（群組有預設則問確認）
            codes = _cs.get_ecount_codes_by_line_id(user_id)
            if len(codes) > 1:
                preferred = state_manager.get_group_cust_cd(user_id)
                if preferred:
                    pref_label = next(
                        (c.get("address_label") or c.get("cust_name") or preferred
                         for c in codes if c["ecount_cust_cd"] == preferred),
                        preferred
                    )
                    state_manager.set(user_id, {
                        "action":     "awaiting_address_selection",
                        "prod_cd":    prod_cd,
                        "prod_name":  prod_name,
                        "qty":        qty,
                        "restock_id": restock_id,
                        "preferred_cust_cd": preferred,
                    })
                    return tone.ask_group_address_confirm(pref_label)
                state_manager.set(user_id, {
                    "action":     "awaiting_address_selection",
                    "prod_cd":    prod_cd,
                    "prod_name":  prod_name,
                    "qty":        qty,
                    "restock_id": restock_id,
                })
                return tone.ask_address_selection(codes)

            state_manager.clear(user_id)
            cust_code = (
                codes[0]["ecount_cust_cd"] if codes
                else _cs.get_ecount_cust_code(
                    user_id, default=settings.ECOUNT_DEFAULT_CUST_CD
                )
            )
            slip_no = ecount_client.save_order(
                cust_code=cust_code,
                items=[{"prod_cd": prod_cd, "qty": qty}],
                phone=_user_phone(user_id),
            )
            if restock_id:
                restock_store.update_status(restock_id, "confirmed")
            if slip_no:
                return tone.restock_wait_confirmed(prod_name, qty, slip_no)
            else:
                print(f"[ordering] 訂單建立失敗（wait_confirm）: {cust_code} | {prod_name} x{qty}")
                issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個（等待確認後）")
                return None
        elif any(kw in text for kw in _NO_KW):
            state_manager.clear(user_id)
            from storage.restock import restock_store
            if restock_id:
                restock_store.update_status(restock_id, "cancelled")
            # 登記到貨通知：有現貨時自動 push 給客戶
            if prod_cd:
                notify_store.add(user_id, prod_cd, prod_name, qty)
                print(f"[notify] 登記到貨通知：{prod_name}（{prod_cd}）x{qty} for {user_id[:10]}...")
            return tone.restock_wait_declined(prod_name)
        else:
            return tone.restock_wait_ask(prod_name, qty, wait_time)

    # ── 等待地址選擇（多地址客戶下單）──────────────
    elif action == "awaiting_address_selection":
        prod_cd          = state.get("prod_cd", "")
        prod_name        = state.get("prod_name", "")
        qty              = state.get("qty", 1)
        restock_id       = state.get("restock_id")
        preferred_cust_cd = state.get("preferred_cust_cd")  # 群組預設地址

        codes = customer_store.get_ecount_codes_by_line_id(user_id)

        # 群組預設地址確認流程：接受「是/否」
        if preferred_cust_cd:
            if any(kw in text for kw in _YES_KW):
                cust_code = preferred_cust_cd
                addr_label = next(
                    (c.get("address_label") or "" for c in codes if c["ecount_cust_cd"] == preferred_cust_cd), ""
                )
                is_undecided = "未決定" in addr_label
                state_manager.clear(user_id)
                state_manager.clear_group_cust_cd(user_id)
                from services.ecount import ecount_client
                slip_no = ecount_client.save_order(
                    cust_code=cust_code,
                    items=[{"prod_cd": prod_cd, "qty": qty}],
                    phone=_user_phone(user_id),
                )
                if restock_id:
                    restock_store.update_status(restock_id, "confirmed")
                if slip_no:
                    print(f"[ordering] 群組預設地址訂單: {slip_no} | {cust_code} | {prod_name} x{qty}")
                    if is_undecided:
                        issue_store.add(user_id, "address_pending",
                            f"{prod_name} × {qty} 個（單號 {slip_no}，待確認配送地址）")
                    return tone.order_confirmed(prod_name, qty, slip_no)
                else:
                    issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個（群組地址後）")
                    return None
            elif any(kw in text for kw in _NO_KW):
                # 移除 preferred，改為一般地址選擇
                state_manager.set(user_id, {
                    "action":     "awaiting_address_selection",
                    "prod_cd":    prod_cd,
                    "prod_name":  prod_name,
                    "qty":        qty,
                    "restock_id": restock_id,
                })
                return tone.ask_address_selection(codes)
            else:
                # 重新確認
                pref_label = next(
                    (c.get("address_label") or c.get("cust_name") or preferred_cust_cd
                     for c in codes if c["ecount_cust_cd"] == preferred_cust_cd),
                    preferred_cust_cd
                )
                return tone.ask_group_address_confirm(pref_label)

        m = re.search(r"^\s*(\d+)\s*$", text)
        if m:
            idx = int(m.group(1)) - 1
            if 0 <= idx < len(codes):
                selected     = codes[idx]
                cust_code    = selected["ecount_cust_cd"]
                addr_label   = (selected.get("address_label") or "").strip()
                is_undecided = "未決定" in addr_label
                state_manager.clear(user_id)
                from services.ecount import ecount_client
                slip_no = ecount_client.save_order(
                    cust_code=cust_code,
                    items=[{"prod_cd": prod_cd, "qty": qty}],
                    phone=_user_phone(user_id),
                )
                if restock_id:
                    restock_store.update_status(restock_id, "confirmed")
                if slip_no:
                    print(f"[ordering] 訂單建立成功: {slip_no} | {cust_code} | {prod_name} x{qty}")
                    if is_undecided:
                        issue_store.add(user_id, "address_pending",
                            f"{prod_name} × {qty} 個（單號 {slip_no}，待確認配送地址）")
                        print(f"[ordering] 地址未決定，已記錄待處理: {slip_no}")
                    return tone.order_confirmed(prod_name, qty, slip_no)
                else:
                    print(f"[ordering] 訂單建立失敗（address_selection）: {cust_code} | {prod_name} x{qty}")
                    issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個（地址選擇後）")
                    return None

        # 無效輸入 → 重新顯示選項
        return tone.ask_address_selection(codes)

    # ── 等待地址選擇（購物車結帳版）──────────────────
    elif action == "awaiting_address_selection_checkout":
        from storage import cart as cart_store
        cart = cart_store.get_cart(user_id)
        codes = customer_store.get_ecount_codes_by_line_id(user_id)
        m = re.search(r"^\s*(\d+)\s*$", text)
        if m:
            idx = int(m.group(1)) - 1
            if 0 <= idx < len(codes):
                selected     = codes[idx]
                cust_code    = selected["ecount_cust_cd"]
                addr_label   = (selected.get("address_label") or "").strip()
                is_undecided = "未決定" in addr_label
                state_manager.clear(user_id)
                from services.ecount import ecount_client as _ec
                items   = [{"prod_cd": i["prod_cd"], "qty": i["qty"]} for i in cart]
                slip_no = _ec.save_order(cust_code=cust_code, items=items, phone=_user_phone(user_id))
                if slip_no:
                    cart_store.clear_cart(user_id)
                    if is_undecided:
                        desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart)
                        issue_store.add(user_id, "address_pending",
                            f"{desc}（單號 {slip_no}，待確認配送地址）")
                        print(f"[ordering] 地址未決定，已記錄待處理: {slip_no}")
                    return tone.checkout_confirmed(cart)
                else:
                    desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart)
                    issue_store.add(user_id, "order_failed", desc)
                    return None
        return tone.ask_address_selection(codes)

    # ── 等待客戶提供聯絡資料（姓名 + 手機） ──────────
    elif action == "awaiting_contact_info":
        prod_cd   = state.get("prod_cd", "")
        prod_name = state.get("prod_name", "")
        qty       = state.get("qty", 1)

        phone_match = _PHONE_RE.search(text)
        if phone_match:
            phone_str = phone_match.group()
            try:
                customer_store.update_phone(user_id, phone_str)
            except Exception:
                pass
            name_text = _PHONE_RE.sub("", text).strip()
            if name_text:
                try:
                    customer_store.update_real_name(user_id, name_text)
                except Exception:
                    pass

            state_manager.clear(user_id)
            from handlers.ordering import _resolve_cust_code, _create_ecount_customer
            from services.ecount import ecount_client as _ec

            # 第二次 JSON 比對（不重複同步，已在 handle_checkout 階段同步過）
            cust_code = _resolve_cust_code(user_id, do_refresh=False)
            if not cust_code:
                # JSON 仍找不到 → Ecount API 建立新客戶
                cust_code = _create_ecount_customer(user_id)

            if not cust_code:
                # 全部失敗 → 寫入待處理，客戶端不回應
                print(f"[ordering] 客戶代碼解析全失敗: {user_id} | {prod_name} x{qty}")
                issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個")
                return None

            slip_no = _ec.save_order(
                cust_code=cust_code,
                items=[{"prod_cd": prod_cd, "qty": qty}],
                phone=phone_str,
            )
            if slip_no:
                print(f"[ordering] 訂單建立成功: {slip_no} | {cust_code} | {prod_name} x{qty}")
                return tone.order_confirmed(prod_name, qty, slip_no)
            else:
                print(f"[ordering] 訂單建立失敗: {cust_code} | {prod_name} x{qty}")
                issue_store.add(user_id, "order_failed", f"{prod_name} × {qty} 個")
                return None
        else:
            return tone.ask_contact_info()

    # ── 等待聯絡資料（購物車結帳版）──────────────────
    elif action == "awaiting_contact_info_checkout":
        from storage import cart as cart_store
        phone_match = _PHONE_RE.search(text)
        if phone_match:
            phone_str = phone_match.group()
            try:
                customer_store.update_phone(user_id, phone_str)
            except Exception:
                pass
            name_text = _PHONE_RE.sub("", text).strip()
            if name_text:
                try:
                    customer_store.update_real_name(user_id, name_text)
                except Exception:
                    pass
            state_manager.clear(user_id)
            from handlers.ordering import _resolve_cust_code, _create_ecount_customer
            from services.ecount import ecount_client as _ec
            cart = cart_store.get_cart(user_id)

            # 第二次 JSON 比對（不重複同步）
            cust_code = _resolve_cust_code(user_id, do_refresh=False)
            if not cust_code:
                # JSON 仍找不到 → Ecount API 建立新客戶
                cust_code = _create_ecount_customer(user_id)

            if not cust_code:
                # 全部失敗 → 寫入待處理，客戶端不回應
                desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart) if cart else "（無購物車資料）"
                print(f"[ordering] 客戶代碼解析全失敗（購物車）: {user_id}")
                issue_store.add(user_id, "order_failed", desc)
                return None

            items   = [{"prod_cd": i["prod_cd"], "qty": i["qty"]} for i in cart]
            slip_no = _ec.save_order(cust_code=cust_code, items=items, phone=phone_str)
            if slip_no:
                cart_store.clear_cart(user_id)
                return tone.checkout_confirmed(cart)
            else:
                desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart)
                issue_store.add(user_id, "order_failed", desc)
                return None
        else:
            return tone.ask_contact_info()

    # ── 等待產品名稱 ───────────────────────────────
    elif action == "awaiting_product":
        # 長訊息且不含貨號 → 使用者已換話題，清除狀態並轉真人
        _has_prod_code = bool(re.search(r'[A-Za-z]{1,3}-?\d{3,6}', text))
        if len(text) > 20 and not _has_prod_code:
            state_manager.clear(user_id)
            from handlers.escalate import handle_unknown
            return handle_unknown(user_id, text, line_api)
        state_manager.clear(user_id)
        from handlers.inventory import query_product
        return query_product(user_id, text, line_api)

    # ── 等待客戶說要登記通知的產品名稱 ────────────
    elif action == "awaiting_notify_product":
        state_manager.clear(user_id)
        from services.ecount import ecount_client as _ec
        from storage.notify import notify_store
        result_item = _ec.lookup(text)
        if not result_item:
            # 查不到 → pending_store 等真人確認
            from storage.pending import pending_store
            pending_store.add(user_id, text)
            return tone.product_not_found(text)
        prod_code = result_item["code"]
        prod_name = result_item["name"] or prod_code
        qty = result_item.get("qty", 0)
        if qty and qty > 0:
            # 有貨 → 進下單流程
            state_manager.set(user_id, {
                "action":    "awaiting_quantity",
                "prod_cd":   prod_code,
                "prod_name": prod_name,
            })
            return tone.notify_request_in_stock(prod_name)
        # 無貨 → 登記到貨通知
        notify_store.add(user_id, prod_code, prod_name, 1)
        print(f"[notify] 登記到貨通知：{prod_name}（{prod_code}）for {user_id[:10]}...")
        return tone.notify_request_ack(prod_name)

    # ── 等待訂單編號 ───────────────────────────────
    elif action == "awaiting_order_id":
        state_manager.clear(user_id)
        from handlers.orders import handle_order_tracking
        return handle_order_tracking(user_id, text)

    # ── 群組預設地址確認（是/否）──────────────────
    elif action == "awaiting_group_address_confirm":
        from storage import cart as cart_store
        from services.ecount import ecount_client

        preferred_cust_cd = state_manager.get_group_cust_cd(user_id)
        if not preferred_cust_cd:
            # 無預設代碼（不應發生），降級為一般地址選擇
            codes = customer_store.get_ecount_codes_by_line_id(user_id)
            state_manager.set(user_id, {"action": "awaiting_address_selection_checkout"})
            return tone.ask_address_selection(codes)

        if any(kw in text for kw in _YES_KW):
            state_manager.clear(user_id)
            state_manager.clear_group_cust_cd(user_id)
            cart = cart_store.get_cart(user_id)
            items = [{"prod_cd": i["prod_cd"], "qty": i["qty"]} for i in cart]
            slip_no = ecount_client.save_order(cust_code=preferred_cust_cd, items=items, phone=_user_phone(user_id))
            if slip_no:
                print(f"[ordering] 群組地址訂單建立成功: {slip_no} | {preferred_cust_cd}")
                cart_store.clear_cart(user_id)
                return tone.checkout_confirmed(cart)
            else:
                print(f"[ordering] 群組地址訂單建立失敗: {preferred_cust_cd}")
                desc = "、".join(f"{i['prod_name']}×{i['qty']}" for i in cart)
                issue_store.add(user_id, "order_failed", desc)
                return None
        elif any(kw in text for kw in _NO_KW):
            # 改選其他地址
            codes = customer_store.get_ecount_codes_by_line_id(user_id)
            state_manager.set(user_id, {"action": "awaiting_address_selection_checkout"})
            return tone.ask_address_selection(codes)
        else:
            # 重新詢問確認
            codes = customer_store.get_ecount_codes_by_line_id(user_id)
            pref_label = next(
                (c.get("address_label") or c.get("cust_name") or preferred_cust_cd
                 for c in codes if c["ecount_cust_cd"] == preferred_cust_cd),
                preferred_cust_cd
            )
            return tone.ask_group_address_confirm(pref_label)

    else:
        state_manager.clear(user_id)
        from handlers.escalate import handle_unknown
        return handle_unknown(user_id, text, line_api)


def _dispatch(
    user_id: str, text: str, intent: Intent, line_api: MessagingApi
) -> str:
    """根據意圖分發到對應 handler"""
    if intent == Intent.INVENTORY:
        return handle_inventory(user_id, text, line_api)
    elif intent == Intent.PRICE:
        return handle_price(user_id, text)
    elif intent == Intent.ORDER_TRACKING:
        return handle_order_tracking(user_id, text)
    elif intent == Intent.DELIVERY:
        return handle_delivery(user_id, text)
    elif intent == Intent.BUSINESS_HOURS:
        return handle_business_hours(text)
    elif intent == Intent.GREETING:
        return tone.greeting_reply()
    elif intent == Intent.CREDIT_CARD:
        return "抱歉我們沒有刷卡喔"
    elif intent == Intent.CONFIRMATION:
        # 購物車有東西 → 視為確認下單
        from storage import cart as cart_store
        if not cart_store.is_empty(user_id):
            return handle_checkout(user_id, line_api)
        # 只有真正的感謝語才回覆，純確認詞（好的/了解/收到）靜默不回
        _THANKS_KW = ["謝謝", "感謝", "感恩", "辛苦了", "謝啦", "謝了", "多謝"]
        if any(kw in text for kw in _THANKS_KW):
            return tone.confirmation_ack()
        return None  # 好的/了解/收到/OK → 靜默
    # ── 台型查詢：「巨無霸有什麼」「中巨有哪些」─────────────────────
    _machine_type = detect_machine_query(text)
    if _machine_type:
        return handle_machine_query(user_id, _machine_type, line_api)
    # ── 新場景 ─────────────────────────────────────
    elif intent == Intent.BARGAINING:
        return handle_bargaining(user_id, text)
    elif intent == Intent.SPEC:
        return handle_spec(user_id, text, line_api)
    elif intent == Intent.RETURN:
        return handle_return(user_id, text, line_api)
    elif intent == Intent.MULTI_PRODUCT:
        return handle_multi_product(user_id, text)
    elif intent == Intent.ADDRESS_CHANGE:
        return handle_address_change(user_id, text, line_api)
    elif intent == Intent.COMPLAINT:
        return handle_complaint(user_id, text, line_api)
    elif intent == Intent.URGENT_ORDER:
        return handle_urgent_order(user_id, text, line_api)
    elif intent == Intent.NOTIFY_REQUEST:
        return handle_notify_request(user_id, text, line_api)
    elif intent == Intent.MACHINE_SIZE:
        # 娃娃機尺寸詢問 → 靜默記錄，不回覆客戶
        issue_store.add(user_id, "machine_size", text)
        return None
    elif intent == Intent.VISIT_STORE:
        # 優先用 real_name（客戶提供的真實姓名），否則用 LINE 顯示名稱
        _cust = customer_store.get_by_line_id(user_id)
        _display = (_cust.get("real_name") if _cust else None) or ""
        if not _display:
            _prof = _get_profile_cached(line_api, user_id)
            _display = _prof.display_name if _prof else ""
        return handle_visit(user_id, text, _display)
    elif intent == Intent.CHECKOUT:
        from storage import cart as cart_store
        if cart_store.is_empty(user_id):
            # 購物車空的 → 當一般確認語處理
            return tone.confirmation_ack()
        return handle_checkout(user_id, line_api)
    else:
        # 完全不知道 → 通知真人客服
        return handle_unknown(user_id, text, line_api)


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
