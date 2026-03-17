"""
每小時待處理清單推送

整理所有未完成項目：
  1. 待確認轉帳（payment_store）
  2. 調貨等待 HQ 回覆 / 等待客戶確認（restock_store）
  3. 待確認配送詢問（delivery_store）
  4. 退換貨 / 投訴 / 地址更改（issue_store）

推送到內部人員群組，並附說明如何標記完成：
  ✅ P1  → 標記第 1 筆轉帳為已確認
  ✅ R2  → 標記第 2 筆調貨為已處理
  ✅ D3  → 標記第 3 筆配送詢問為已處理
  ✅ I4  → 標記第 4 筆問題（退換貨/投訴/地址更改）為已處理
"""

from datetime import datetime

import pytz
from linebot.v3.messaging import (
    ApiClient, Configuration, MessagingApi,
    PushMessageRequest, TextMessage,
)

from config import settings
from storage.customers import customer_store
from storage.payments import payment_store
from storage.restock import restock_store
from storage.delivery import delivery_store
from storage.issues import issue_store
from storage.pending import pending_store
from storage import queue as queue_store


def build_pending_text() -> str | None:
    """
    組合待處理清單文字並回傳（無未解決項目則回傳 None）。
    不推送，只回傳字串，供內部群指令回覆使用。
    """
    pending_payments = payment_store.get_pending()
    unresolved_restock = restock_store.get_unresolved()
    pending_deliveries = delivery_store.get_pending()
    pending_issues = issue_store.get_pending()
    queued_msgs = queue_store.get_unprocessed()
    pending_queries = pending_store.get_pending()

    if not pending_payments and not unresolved_restock and not pending_deliveries \
            and not pending_issues and not pending_queries and not queued_msgs:
        return None

    tz = pytz.timezone(settings.BUSINESS_TZ)
    now = datetime.now(tz)
    time_str = now.strftime("%m/%d %H:%M")

    lines = [f"📋 待處理清單（{time_str}）\n"]

    def _display_name(user_id: str) -> str:
        cust = customer_store.get_by_line_id(user_id)
        if not cust:
            return "未知客戶"
        if cust.get("chat_label"):
            return cust["chat_label"]
        dn = cust.get("display_name") or ""
        rn = cust.get("real_name") or ""
        if dn and rn and dn != rn:
            return f"{dn}（{rn}）"
        return dn or rn or "未知客戶"

    # ── 轉帳待確認 ──────────────────────────────────
    if pending_payments:
        lines.append("💰 待確認轉帳：")
        for p in pending_payments:
            raw = p["text_raw"][:30] + ("…" if len(p["text_raw"]) > 30 else "")
            lines.append(f"  • #P{p['id']}  {_display_name(p['user_id'])}　「{raw}」")

    # ── 調貨等待 HQ 回覆 ────────────────────────────
    pending_hq = [r for r in unresolved_restock if r["status"] == "pending"]
    if pending_hq:
        if pending_payments:
            lines.append("")
        lines.append("🔄 等待總公司回覆：")
        for r in pending_hq:
            lines.append(f"  • #R{r['id']}  {_display_name(r['user_id'])}　{r['prod_name']} × {r['qty']} 個")

    # ── 等待客戶確認是否願意等 ───────────────────────
    pending_wait = [r for r in unresolved_restock if r["status"] == "ordering"]
    if pending_wait:
        if pending_payments or pending_hq:
            lines.append("")
        lines.append("⏳ 等待客戶確認：")
        for r in pending_wait:
            lines.append(f"  • #R{r['id']}  {_display_name(r['user_id'])}　{r['prod_name']} × {r['qty']} 個（{r['wait_time']}）")

    # ── 配送待確認 ───────────────────────────────────
    if pending_deliveries:
        if pending_payments or unresolved_restock:
            lines.append("")
        lines.append("🚚 待確認配送詢問：")
        for d in pending_deliveries:
            raw = d["text_raw"][:30] + ("…" if len(d["text_raw"]) > 30 else "")
            lines.append(f"  • #D{d['id']}  {_display_name(d['user_id'])}　「{raw}」")

    # ── 退換貨 / 投訴 / 地址更改 ────────────────────────
    _ISSUE_LABEL = {
        "return":         "🔄 退換貨",
        "complaint":      "⚠️ 投訴",
        "address_change": "📍 地址更改",
        "urgent_order":   "🚛 催貨",
        "order_failed":   "❌ 訂單失敗",
        "unknown":        "❓ 無法識別",
        "spec_query":     "📋 規格詢問",
        "image_query":    "🖼️ 圖片詢問",
        "order_query":    "📦 訂單查詢",
        "address_pending": "📮 待確認地址",
        "machine_size":    "🎰 娃娃機尺寸",
    }
    _MACHINE_SIZE_NAMES = ["巨無霸", "中巨", "標準", "小K", "K霸", "小k", "k霸", "迷你機"]

    def _issue_label(issue: dict) -> str:
        if issue["type"] == "machine_size":
            txt = issue.get("text_raw", "")
            size = next((s for s in _MACHINE_SIZE_NAMES if s in txt), "娃娃機尺寸")
            return f"🎰 {size}"
        return _ISSUE_LABEL.get(issue["type"], issue["type"])

    if pending_issues:
        if pending_payments or unresolved_restock or pending_deliveries:
            lines.append("")
        lines.append("📌 待處理問題：")
        for i in pending_issues:
            raw = i["text_raw"][:30] + ("…" if len(i["text_raw"]) > 30 else "")
            label = _issue_label(i)
            lines.append(f"  • #I{i['id']}  {_display_name(i['user_id'])}　[{label}]「{raw}」")

    # ── 查無商品待確認 ───────────────────────────────────
    if pending_queries:
        if pending_payments or unresolved_restock or pending_deliveries or pending_issues:
            lines.append("")
        lines.append("🔍 待確認商品查詢：")
        for q in pending_queries:
            lines.append(f"  • #Q{q['id']}  {_display_name(q['user_id'])}　「{q['product']}」")

    # ── 離峰排隊訊息 ────────────────────────────────────
    if queued_msgs:
        if pending_payments or unresolved_restock or pending_deliveries or pending_issues:
            lines.append("")
        lines.append(f"💬 離峰累積訊息（共 {len(queued_msgs)} 則）：")
        for q in queued_msgs[:10]:
            if q["msg_type"] == "image":
                lines.append(f"  • {_display_name(q['user_id'])}　[圖片]")
            else:
                raw = (q["content"] or "")[:30] + ("…" if len(q["content"] or "") > 30 else "")
                lines.append(f"  • {_display_name(q['user_id'])}　「{raw}」")
        if len(queued_msgs) > 10:
            lines.append(f"  …另有 {len(queued_msgs) - 10} 則")

    return "\n".join(lines)


def send_pending_summary() -> None:
    """整理待處理清單並推送到內部群組（無未解決項目則不發送）"""
    if not settings.LINE_GROUP_ID:
        return
    msg = build_pending_text()
    if not msg:
        return
    config = Configuration(access_token=settings.LINE_CHANNEL_ACCESS_TOKEN)
    with ApiClient(config) as api_client:
        line_api = MessagingApi(api_client)
        try:
            line_api.push_message(
                PushMessageRequest(
                    to=settings.LINE_GROUP_ID,
                    messages=[TextMessage(text=msg)],
                )
            )
            print(f"[summary] 已推送待處理清單", flush=True)
        except Exception as e:
            print(f"[summary] 推送失敗: {e}", flush=True)


def build_full_report(days: int = 3) -> str:
    """
    產生「已處理 + 未處理」完整報表文字（不推送，回傳字串）。

    Args:
        days: 已處理部分顯示幾天內的記錄（預設 3 天）
    """
    tz = pytz.timezone(settings.BUSINESS_TZ)
    now = datetime.now(tz)
    time_str = now.strftime("%m/%d %H:%M")

    # ── 資料取得 ──────────────────────────────────────
    pending_payments   = payment_store.get_pending()
    resolved_payments  = payment_store.get_recent_resolved(days)

    unresolved_restock = restock_store.get_unresolved()
    completed_restock  = restock_store.get_recent_completed(days)

    pending_deliveries  = delivery_store.get_pending()
    resolved_deliveries = delivery_store.get_recent_resolved(days)

    pending_issues  = issue_store.get_pending()
    resolved_issues = issue_store.get_recent_resolved(days)

    pending_queries = pending_store.get_pending()

    queued_msgs = queue_store.get_unprocessed()

    def _cust_name(user_id: str) -> str:
        """LINE 暱稱優先；若有 real_name 且不同則附上（real_name）"""
        cust = customer_store.get_by_line_id(user_id)
        if not cust:
            return "未知客戶"
        if cust.get("chat_label"):
            return cust["chat_label"]
        dn = cust.get("display_name") or ""
        rn = cust.get("real_name") or ""
        if dn and rn and dn != rn:
            return f"{dn}（{rn}）"
        return dn or rn or "未知客戶"

    def _fmt_time(iso: str | None) -> str:
        if not iso:
            return ""
        try:
            return iso[11:16]   # HH:MM
        except Exception:
            return ""

    lines = [f"📊 完整記錄（{days} 天內已處理 + 未處理）　{time_str}", ""]

    _ISSUE_LABEL = {
        "return":         "退換貨",
        "complaint":      "投訴",
        "address_change": "地址更改",
        "urgent_order":   "催貨",
        "order_failed":   "訂單失敗",
        "unknown":        "無法識別",
        "spec_query":     "規格詢問",
        "image_query":    "圖片詢問",
        "order_query":    "訂單查詢",
        "machine_size":   "娃娃機尺寸",
    }

    # ━━━ 已處理 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    has_resolved = resolved_payments or completed_restock or resolved_deliveries or resolved_issues
    if has_resolved:
        lines.append("✅ 已處理")

        if resolved_payments:
            lines.append("  💰 轉帳確認：")
            for p in resolved_payments:
                raw = p["text_raw"][:25] + ("…" if len(p["text_raw"]) > 25 else "")
                t   = _fmt_time(p.get("resolved_at"))
                lines.append(f"    #P{p['id']}  {_cust_name(p['user_id'])}　「{raw}」  ✓{t}")

        if completed_restock:
            lines.append("  🔄 調貨：")
            for r in completed_restock:
                status_label = "已確認" if r["status"] == "confirmed" else "已取消"
                lines.append(
                    f"    #R{r['id']}  {_cust_name(r['user_id'])}　"
                    f"{r['prod_name']} × {r['qty']} 個　[{status_label}]"
                )

        if resolved_deliveries:
            lines.append("  🚚 配送詢問：")
            for d in resolved_deliveries:
                raw = d["text_raw"][:25] + ("…" if len(d["text_raw"]) > 25 else "")
                t   = _fmt_time(d.get("resolved_at"))
                lines.append(f"    #D{d['id']}  {_cust_name(d['user_id'])}　「{raw}」  ✓{t}")

        if resolved_issues:
            lines.append("  📌 退換貨/投訴/地址：")
            _MS = ["巨無霸", "中巨", "標準", "小K", "K霸", "小k", "k霸", "迷你機"]
            for i in resolved_issues:
                raw   = i["text_raw"][:25] + ("…" if len(i["text_raw"]) > 25 else "")
                t     = _fmt_time(i.get("resolved_at"))
                if i["type"] == "machine_size":
                    size = next((s for s in _MS if s in i.get("text_raw", "")), "娃娃機尺寸")
                    label = f"娃娃機-{size}"
                else:
                    label = _ISSUE_LABEL.get(i["type"], i["type"])
                lines.append(f"    #I{i['id']}  {_cust_name(i['user_id'])}　[{label}]「{raw}」  ✓{t}")

        lines.append("")

    # ━━━ 未處理 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    pending_hq   = [r for r in unresolved_restock if r["status"] == "pending"]
    pending_wait = [r for r in unresolved_restock if r["status"] == "ordering"]
    has_pending  = pending_payments or pending_hq or pending_wait or pending_deliveries or pending_issues or pending_queries or queued_msgs

    if has_pending:
        lines.append("⏳ 未處理")

        if pending_payments:
            lines.append("  💰 待確認轉帳：")
            for p in pending_payments:
                raw = p["text_raw"][:25] + ("…" if len(p["text_raw"]) > 25 else "")
                lines.append(f"    #P{p['id']}  {_cust_name(p['user_id'])}　「{raw}」")

        if pending_hq:
            lines.append("  🔄 等待總公司回覆：")
            for r in pending_hq:
                lines.append(
                    f"    #R{r['id']}  {_cust_name(r['user_id'])}　"
                    f"{r['prod_name']} × {r['qty']} 個"
                )

        if pending_wait:
            lines.append("  ⏳ 等待客戶確認：")
            for r in pending_wait:
                lines.append(
                    f"    #R{r['id']}  {_cust_name(r['user_id'])}　"
                    f"{r['prod_name']} × {r['qty']} 個（{r['wait_time']}）"
                )

        if pending_deliveries:
            lines.append("  🚚 待確認配送詢問：")
            for d in pending_deliveries:
                raw = d["text_raw"][:25] + ("…" if len(d["text_raw"]) > 25 else "")
                lines.append(f"    #D{d['id']}  {_cust_name(d['user_id'])}　「{raw}」")

        if pending_issues:
            lines.append("  📌 待處理問題：")
            _MS = ["巨無霸", "中巨", "標準", "小K", "K霸", "小k", "k霸", "迷你機"]
            for i in pending_issues:
                raw   = i["text_raw"][:25] + ("…" if len(i["text_raw"]) > 25 else "")
                if i["type"] == "machine_size":
                    size = next((s for s in _MS if s in i.get("text_raw", "")), "娃娃機尺寸")
                    label = f"娃娃機-{size}"
                else:
                    label = _ISSUE_LABEL.get(i["type"], i["type"])
                complaint_flag = " ⚠️" if i["type"] == "complaint" else ""
                lines.append(
                    f"    #I{i['id']}  {_cust_name(i['user_id'])}　"
                    f"[{label}]「{raw}」{complaint_flag}"
                )

        if pending_queries:
            lines.append("  🔍 待確認商品查詢：")
            for q in pending_queries:
                lines.append(f"    #Q{q['id']}  {_cust_name(q['user_id'])}　「{q['product']}」")

        if queued_msgs:
            lines.append(f"  💬 離峰累積訊息（共 {len(queued_msgs)} 則，上班自動補處理）：")
            for q in queued_msgs[:10]:
                name = _cust_name(q["user_id"])
                t    = (q.get("created_at") or "")[:16].replace("T", " ")
                if q["msg_type"] == "image":
                    lines.append(f"    {name}　[圖片]　{t}")
                else:
                    raw = (q["content"] or "")[:25] + ("…" if len(q["content"] or "") > 25 else "")
                    lines.append(f"    {name}　「{raw}」　{t}")
            if len(queued_msgs) > 10:
                lines.append(f"    …另有 {len(queued_msgs) - 10} 則")

        lines.append("")
        lines.append("———")
        lines.append("✅ P1 → 標記轉帳已確認")
        lines.append("✅ R2 → 標記調貨已處理")
        lines.append("✅ D3 → 標記配送已確認")
        lines.append("✅ I4 → 標記問題已處理")
        lines.append("✅ Q5 → 標記商品查詢已回覆")
    else:
        lines.append("✅ 目前無待處理項目")

    return "\n".join(lines)


def send_full_report(days: int = 3) -> None:
    """推送完整報表（已處理 + 未處理）到內部群組"""
    if not settings.LINE_GROUP_ID:
        return
    msg = build_full_report(days)
    config = Configuration(access_token=settings.LINE_CHANNEL_ACCESS_TOKEN)
    with ApiClient(config) as api_client:
        line_api = MessagingApi(api_client)
        try:
            line_api.push_message(
                PushMessageRequest(
                    to=settings.LINE_GROUP_ID,
                    messages=[TextMessage(text=msg)],
                )
            )
            print("[summary] 已推送完整報表")
        except Exception as e:
            print(f"[summary] 完整報表推送失敗: {e}")
