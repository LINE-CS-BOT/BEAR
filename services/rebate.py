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
    "WEI丞": ["WEI", "丞"],
    "舒老闆": ["舒老闆", "寧寧", "冬冬", "夾鬥陣"],
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

    # Step 1: Group by merged name,記錄是否為特殊合併組
    groups: dict[str, list[dict]] = {}
    is_merge_group: dict[str, bool] = {}
    for item in sales:
        name = item.get("customer", "").strip()
        amount = float(item.get("amount", 0))
        if not name or amount == 0:
            continue

        # Check special merge group first
        merge_group = _get_merge_group(name)
        if merge_group:
            group_key = merge_group
            is_merge_group[group_key] = True
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

        # Distribute rebate
        distributed = []
        if is_merge_group.get(group_name) and len(members) > 1:
            # 特殊合併組（如 WEI+丞）：按人合計，保留各店明細
            person_totals: dict[str, float] = {}
            person_stores: dict[str, list[dict]] = {}
            for m in members:
                base = _get_base_name(m["name"])
                person_totals[base] = person_totals.get(base, 0) + m["amount"]
                person_stores.setdefault(base, []).append(
                    {"name": m["name"], "amount": m["amount"]}
                )

            if group_total >= 100000:
                # 合計 ≥ 10萬：合計的 5%，按各人比例分
                for person, p_total in person_totals.items():
                    p_rebate = round(rebate * p_total / group_total)
                    stores = sorted(person_stores[person], key=lambda x: -x["amount"])
                    distributed.append({"name": person, "amount": p_total, "rebate": p_rebate, "stores": stores})
            else:
                # 合計 < 10萬：各人獨立看是否達標
                total_rebate_actual = 0
                for person, p_total in person_totals.items():
                    p_rebate = _calc_rebate(p_total)
                    stores = sorted(person_stores[person], key=lambda x: -x["amount"])
                    distributed.append({"name": person, "amount": p_total, "rebate": p_rebate, "stores": stores})
                    total_rebate_actual += p_rebate
                # 修正 group 層級的回饋金為各人實際合計
                rebate = total_rebate_actual
                total_rebate += rebate - _calc_rebate(group_total)
        elif rebate > 0:
            # 同名各店（如 林子翔-基隆/樹林）：回饋金整筆歸本人，不拆分
            distributed = [{**m, "rebate": 0} for m in members]
        else:
            distributed = [{**m, "rebate": 0} for m in members]

        result_groups.append({
            "group_name": group_name,
            "total": group_total,
            "rebate": rebate,
            "tier": tier,
            "is_merge": is_merge_group.get(group_name, False),
            "members": sorted(distributed, key=lambda x: -x["amount"]),
        })

    return {
        "month": datetime.now().strftime("%Y-%m"),
        "groups": result_groups,
        "summary": {"total_sales": total_sales, "total_rebate": total_rebate},
    }


def get_approaching_customers(sales: list[dict] | None = None) -> list[dict]:
    """
    找出「快接近達成」的客戶。
    門檻：3萬→17000起算、6萬→45000起算、10萬→75000起算
    """
    result = calculate_rebates(sales)
    approaching = []
    # (目標, 起算門檻)
    thresholds = [(30000, 17000), (60000, 45000), (100000, 75000)]

    for g in result["groups"]:
        total = g["total"]
        for target, floor in thresholds:
            if total < target and total >= floor:
                approaching.append({
                    "group_name": g["group_name"],
                    "total": total,
                    "target": target,
                    "gap": target - total,
                    "current_tier": g["tier"],
                    "next_tier": f"{target // 10000}萬",
                })
                break

    return sorted(approaching, key=lambda x: -x["total"])


def get_last_month_achievers() -> dict:
    """
    取得上個月確定達標的客戶。
    每月 1~14 日顯示用。
    """
    # 讀取上月資料
    last_month_path = _BASE / "data" / "rebate_sales_lastmonth.json"
    if not last_month_path.exists():
        return {"month": "", "achievers": [], "total_rebate": 0}

    try:
        sales = json.loads(last_month_path.read_text(encoding="utf-8"))
    except Exception:
        return {"month": "", "achievers": [], "total_rebate": 0}

    result = calculate_rebates(sales)
    achievers = [g for g in result["groups"] if g["rebate"] > 0]

    # 上個月的月份
    now = datetime.now()
    if now.month == 1:
        month_str = f"{now.year - 1}-12"
    else:
        month_str = f"{now.year}-{now.month - 1:02d}"

    return {
        "month": month_str,
        "achievers": achievers,
        "total_rebate": sum(g["rebate"] for g in achievers),
    }
