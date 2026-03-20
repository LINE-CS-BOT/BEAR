"""
回饋金計算引擎
- 從 data/rebate_sales.json 讀取銷貨明細（由 sync 腳本定期更新）
- 合併同名客戶（名稱-後綴 視為同一人，特殊合併組）
- 計算回饋金級距並分配到個人
"""
import json
from pathlib import Path
from datetime import datetime, timedelta

_BASE = Path(__file__).parent.parent
_SALES_PATH = _BASE / "data" / "rebate_sales.json"

# 特殊合併組：這些名稱視為同一組
_MERGE_GROUPS = {
    "WEI丞": ["WEI", "丞"],  # group_name: [member names]
}

# 回饋金級距
def _calc_rebate(total: float) -> float:
    """根據合計金額計算回饋金"""
    if total >= 100000:
        return round(total * 0.05)
    elif total >= 60000:
        return 2000
    elif total >= 30000:
        return 1000
    return 0


def _get_base_name(name: str) -> str:
    """取得合併用的基礎名稱：去掉 - 後面的部分"""
    if "-" in name:
        return name.split("-")[0].strip()
    if "\uff0d" in name:
        return name.split("\uff0d")[0].strip()
    return name.strip()


def _get_merge_group(name: str) -> str | None:
    """檢查是否屬於特殊合併組，回傳組名"""
    base = _get_base_name(name)
    for group_name, members in _MERGE_GROUPS.items():
        if base in members:
            return group_name
    return None


def load_sales() -> list[dict]:
    """載入銷貨資料。格式：[{"customer": "林子翔-基隆", "amount": 47387}, ...]"""
    if not _SALES_PATH.exists():
        return []
    try:
        return json.loads(_SALES_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []


def calculate_rebates(sales: list[dict] | None = None) -> dict:
    """
    計算回饋金。

    回傳：{
        "month": "2026-03",
        "groups": [
            {
                "group_name": "林子翔",
                "total": 61436,
                "rebate": 2000,
                "tier": "6萬",
                "members": [
                    {"name": "林子翔-基隆", "amount": 47387, "rebate": 2000},
                    {"name": "林子翔-樹林", "amount": 14049, "rebate": 0},
                ]
            },
            ...
        ],
        "summary": {"total_sales": ..., "total_rebate": ...}
    }
    """
    if sales is None:
        sales = load_sales()

    if not sales:
        return {"month": datetime.now().strftime("%Y-%m"), "groups": [], "summary": {"total_sales": 0, "total_rebate": 0}}

    # Step 1: Group by merged name
    groups: dict[str, list[dict]] = {}
    for item in sales:
        name = item.get("customer", "").strip()
        amount = float(item.get("amount", 0))
        if not name or amount <= 0:
            continue

        # Check special merge group first
        merge_group = _get_merge_group(name)
        if merge_group:
            group_key = merge_group
        else:
            group_key = _get_base_name(name)

        if group_key not in groups:
            groups[group_key] = []
        groups[group_key].append({"name": name, "amount": amount})

    # Step 2: Calculate rebate for each group and distribute
    result_groups = []
    total_sales = 0
    total_rebate = 0

    for group_name, members in sorted(groups.items(), key=lambda x: -sum(m["amount"] for m in x[1])):
        group_total = sum(m["amount"] for m in members)
        total_sales += group_total
        rebate = _calc_rebate(group_total)
        total_rebate += rebate

        # Determine tier label
        if group_total >= 100000:
            tier = "10萬"
        elif group_total >= 60000:
            tier = "6萬"
        elif group_total >= 30000:
            tier = "3萬"
        else:
            tier = "未達"

        # Distribute rebate to individuals
        distributed = []
        if rebate > 0 and len(members) > 1:
            if group_total >= 100000:
                # 10萬以上：按金額比例分配 5%
                for m in members:
                    m_rebate = round(rebate * m["amount"] / group_total)
                    distributed.append({**m, "rebate": m_rebate})
            else:
                # 3萬/6萬：誰達到門檻就給誰，都沒達到就平均
                threshold = 60000 if rebate == 2000 else 30000
                achievers = [m for m in members if m["amount"] >= threshold]
                if achievers:
                    per_rebate = round(rebate / len(achievers))
                    for m in members:
                        m_rebate = per_rebate if m["amount"] >= threshold else 0
                        distributed.append({**m, "rebate": m_rebate})
                else:
                    # 沒人個別達標 → 平均分
                    per_rebate = round(rebate / len(members))
                    for m in members:
                        distributed.append({**m, "rebate": per_rebate})
        elif rebate > 0:
            distributed = [{**members[0], "rebate": rebate}]
        else:
            distributed = [{**m, "rebate": 0} for m in members]

        result_groups.append({
            "group_name": group_name,
            "total": group_total,
            "rebate": rebate,
            "tier": tier,
            "members": sorted(distributed, key=lambda x: -x["amount"]),
        })

    return {
        "month": datetime.now().strftime("%Y-%m"),
        "groups": result_groups,
        "summary": {"total_sales": total_sales, "total_rebate": total_rebate},
    }


def get_approaching_customers(sales: list[dict] | None = None) -> list[dict]:
    """
    找出「快接近達成」的客戶（離下一級距差 20% 以內）。
    用於每月 20 號後每日提醒。
    """
    result = calculate_rebates(sales)
    approaching = []
    thresholds = [30000, 60000, 100000]

    for g in result["groups"]:
        total = g["total"]
        for t in thresholds:
            if total < t and total >= t * 0.8:  # 差 20% 以內
                approaching.append({
                    "group_name": g["group_name"],
                    "total": total,
                    "target": t,
                    "gap": t - total,
                    "current_tier": g["tier"],
                    "next_tier": f"{t // 10000}萬",
                })
                break

    return sorted(approaching, key=lambda x: x["gap"])
