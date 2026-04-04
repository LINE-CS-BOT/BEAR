"""
銷售分析模組

分析維度：
1. 銷售速度排行
2. 滯銷品偵測
3. 客戶分析
4. 補貨預測
5. 價位帶分析
6. 品類分析
7. 新品採購建議
"""

import sqlite3
import re
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

DB_PATH = Path(__file__).parent.parent / "data" / "sales_detail.db"


def _conn():
    return sqlite3.connect(str(DB_PATH))


# ── 排除規則 ──────────────────────────────────────────

def _is_excluded_product(prod_cd: str) -> bool:
    """排除不分析的品項：Z+英文開頭、HH開頭（娃娃機零件耗材）、X開頭（成人用品）、NN開頭"""
    cd = prod_cd.upper()
    if len(cd) >= 2 and cd[0] == 'Z' and cd[1].isalpha():
        return True
    if cd.startswith("HH"):
        return True
    if cd.startswith("X0"):
        return True
    if cd.startswith("NN"):
        return True
    return False


def _is_excluded_customer(customer: str) -> bool:
    """民享店不分析"""
    return "民享" in customer


# ── 品類分類 ──────────────────────────────────────────

_CATEGORY_RULES = [
    ("藍牙耳機", ["藍牙耳機", "藍芽耳機", "音樂耳機", "真無線", "耳機"]),
    ("音響", ["音響", "音箱", "喇叭"]),
    ("行動電源", ["行動電源", "充電寶", "移動電源"]),
    ("遙控車/飛機", ["遙控車", "遙控飛機", "遙控直升機", "攀爬車", "四軸", "無人機", "空拍機", "滑翔翼", "遙控坦克", "遙控工程", "遙控噴霧",
                   "攀岩王", "貨卡", "得利卡", "仿真比例車", "回力車", "賽車", "方程賽車", "環保尖兵車", "拖拉機", "直升飛機"]),
    ("涼風扇", ["涼風扇", "風扇", "散熱器"]),
    ("釣具", ["蝦竿", "釣竿", "釣蝦", "釣魚", "魚輪", "魚竿", "魚餌", "軟餌", "浮標", "魚缸", "螃蟹軟餌", "雷蛙", "海竿", "仿真餌", "銀龍桿"]),
    ("手錶", ["手錶", "手環", "電子錶"]),
    ("打火機/噴火槍", ["打火機", "噴火槍", "點火", "小噴槍", "火焰燈", "火焰噴槍"]),
    ("工具類", ["工具箱", "萬用表", "打氣筒", "充氣機", "膠槍", "射釘", "螺絲", "手電筒", "噴壺", "吸提器",
              "扳手", "洗車", "水管槍", "噴水槍", "探照燈", "封口機", "護貝機", "烘槍",
              "測電筆", "萬國充", "吸塵器", "壁燈", "抽水器", "鑽", "工具套裝"]),
    ("生活用品", ["蓮蓬頭", "插座", "充電頭", "數據線", "枕", "掛件", "小夜燈", "露營燈", "支架", "皮帶", "內褲", "刮鬍",
                "鞋墊", "拖鞋", "圍巾", "手套", "帽", "吹風機", "杯", "水壺", "太空壺", "便當", "收納", "濕紙巾", "拖把",
                "沙發床", "卡套", "集線器", "計算機", "筆記本", "削鉛筆", "標籤機", "雨鞋", "背包", "斜背",
                "止滑", "鞋", "澡盆", "飲水機", "保溫桶", "湯鍋", "刀具", "鍋", "卡式爐",
                "監視器", "攝像頭", "行車記錄", "滅火器", "耳塞", "噴水頭", "書桌燈"]),
    ("玩具", ["娃娃機", "遊戲機", "切切樂", "存錢筒", "存錢罐", "公仔", "掌上型", "彈弓", "投影", "戰鬥機",
             "恐龍", "積木", "滑步車", "溜溜車", "迴力車", "拼圖", "捏捏樂", "手辦", "奧特曼",
             "手槍", "三角龍", "烤鴨", "撞球", "練習球", "足球", "籃球", "兵工", "工具拉桿箱",
             "跳舞章魚", "昆蟲機甲", "戰甲", "章魚", "寶可夢", "家家酒"]),
    ("美容/保養", ["毛球機", "剃鬚", "內衣", "滅蚊"]),
    ("三麗鷗/IP", ["三麗鷗", "三麗歐", "迪士尼", "蠟筆小新", "野獸國", "PLAYBOY", "嘻小家",
                  "娃三歲", "庫洛米", "湯姆貓", "傑利鼠", "飛天小女警", "史迪奇", "海螺"]),
    ("零食飲料", ["樂事", "多力多滋", "泡芙", "桶麵", "品客", "咖啡", "嗨啾", "康貝特", "愛之味",
                 "滿漢", "繽紛樂", "麥香", "餅乾", "巧克力", "米餅", "樂連連", "樂天"]),
    ("耗材", ["洗衣球", "泡澡球", "衛生紙", "厚敏", "紅包袋", "紙盒"]),
    ("娃娃/絨毛", ["絨毛", "娃娃", "崽崽", "玩偶", "毛絨", "炸毛獅子", "小象菲菲", "水豚", "動物椅"]),
    ("合金模型", ["合金模型", "合金回力", "模型車", "摩托車", "合金車", "合金版", "合金自行車"]),
    ("雷射/燈光", ["雷射", "激光", "LED", "爆亮", "自行車燈"]),
    ("飾品配件", ["項鍊", "鑰匙圈", "鑰匙扣", "吊飾", "手鏈", "手環", "戒指", "磁鐵", "冰箱貼"]),
    ("節慶商品", ["財神", "醒獅", "春節", "新春", "賀歲", "擺飾", "擺件", "鞭炮", "炮竹", "宮燈", "媽祖", "招財貓", "開運",
                "祥獅", "年貨", "舞獅", "過年", "相框"]),
    ("電腦周邊", ["滑鼠", "鍵盤", "USB", "隨身碟"]),
    ("盲盒", ["盲盒"]),
    ("車載用品", ["車載", "車用", "車充"]),
    ("暖風/保暖", ["暖手", "暖風", "暖爐", "暖寶"]),
    ("香薰", ["香薰", "香氛", "擴香"]),
]


def _classify(prod_name: str) -> str:
    """根據品名分類"""
    for cat, keywords in _CATEGORY_RULES:
        for kw in keywords:
            if kw in prod_name:
                return cat
    return "其他"


# ── 1. 銷售速度排行 ──────────────────────────────────

def top_sellers(days: int = 30, limit: int = 20) -> list[dict]:
    """最近 N 天銷量最多的品項（用銷貨明細，不含調撥退貨）"""
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) as total_out,
                   COUNT(DISTINCT date) as active_days
            FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
            ORDER BY total_out DESC
            LIMIT ?
        """, (cutoff, limit * 2)).fetchall()
    results = [
        {"code": r[0], "name": r[1], "total_out": int(r[2]),
         "active_days": r[3], "daily_avg": round(r[2] / max(r[3], 1), 1),
         "category": _classify(r[1])}
        for r in rows if not _is_excluded_product(r[0])
    ]
    return results[:limit]


# ── 2. 滯銷品偵測 ──────────────────────────────────

def slow_movers(no_sale_days: int = 60, min_stock: int = 10) -> list[dict]:
    """庫存 > min_stock 但最近 N 天沒出庫的品項"""
    import json
    avail_path = Path(__file__).parent.parent / "data" / "available.json"
    if not avail_path.exists():
        return []
    avail = json.loads(avail_path.read_text(encoding="utf-8"))

    cutoff = (datetime.now() - timedelta(days=no_sale_days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        # 最近有銷售的品項（用銷貨明細）
        active = set(r[0] for r in conn.execute(
            "SELECT DISTINCT prod_cd FROM sales_detail WHERE date >= ? AND customer NOT LIKE '%民享%'",
            (cutoff,)
        ).fetchall())

        # 有銷售紀錄但最近沒賣的
        all_products = conn.execute("""
            SELECT prod_cd, prod_name, MAX(date) as last_sale
            FROM sales_detail
            WHERE customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """).fetchall()

    # 最近 30 天有入庫的品項（剛到貨，不算滯銷）
    recent_restock_cutoff = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    with _conn() as conn:
        recently_restocked = set(r[0] for r in conn.execute(
            "SELECT DISTINCT prod_cd FROM inventory_changes WHERE date >= ? AND qty_in > 0",
            (recent_restock_cutoff,)
        ).fetchall())

    # 計算每個產品的出庫量（用於周轉率判斷）
    with _conn() as conn:
        _sales_90d = {}
        for _sc, _sq in conn.execute(
            "SELECT prod_cd, SUM(qty) FROM sales_detail WHERE date >= ? AND customer NOT LIKE '%民享%' GROUP BY prod_cd",
            ((datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d"),)
        ).fetchall():
            _sales_90d[_sc] = _sq or 0

    results = []
    for code, name, last_out in all_products:
        if _is_excluded_product(code):
            continue
        if code in recently_restocked:
            continue  # 最近才進貨，不是滯銷

        stock = 0
        if code in avail:
            d = avail[code]
            stock = d.get("available", 0) if isinstance(d, dict) else d
        if stock < min_stock:
            continue

        sold_90d = _sales_90d.get(code, 0)
        turnover = round(sold_90d / stock, 2) if stock > 0 else 0

        # 滯銷判斷：完全沒賣 OR 周轉率 < 0.5 且庫存 > 30
        is_no_sale = code not in active
        is_low_turnover = turnover < 0.5 and stock > 30

        if is_no_sale or is_low_turnover:
            results.append({
                "code": code, "name": name, "stock": stock,
                "last_sale": last_out or "N/A",
                "sold_90d": sold_90d,
                "turnover": turnover,
                "category": _classify(name),
            })
    results.sort(key=lambda x: -x["stock"])
    return results


# ── 3. 客戶分析 ──────────────────────────────────

def customer_analysis(days: int = 90, limit: int = 20) -> list[dict]:
    """客戶購買分析：金額、頻率、最愛品類"""
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT customer, SUM(amount) as total, COUNT(*) as items,
                   COUNT(DISTINCT date) as order_days,
                   COUNT(DISTINCT prod_cd) as unique_products
            FROM sales_detail
            WHERE date >= ? AND customer != '' AND customer NOT LIKE '%民享%'
            GROUP BY customer
            ORDER BY total DESC
            LIMIT ?
        """, (cutoff, limit)).fetchall()

        results = []
        for cust, total, items, order_days, unique_prods in rows:
            if _is_excluded_customer(cust):
                continue
            # 最常買的品類
            cat_rows = conn.execute("""
                SELECT prod_name, SUM(qty) as total_qty
                FROM sales_detail
                WHERE date >= ? AND customer = ?
                GROUP BY prod_cd
                ORDER BY total_qty DESC
                LIMIT 3
            """, (cutoff, cust)).fetchall()
            top_items = [{"name": r[0], "qty": int(r[1])} for r in cat_rows]
            cats = defaultdict(int)
            for r in cat_rows:
                cats[_classify(r[0])] += int(r[1])
            fav_cat = max(cats, key=cats.get) if cats else "N/A"

            # 回購間隔
            dates = conn.execute("""
                SELECT DISTINCT date FROM sales_detail
                WHERE date >= ? AND customer = ? ORDER BY date
            """, (cutoff, cust)).fetchall()
            intervals = []
            for i in range(1, len(dates)):
                d1 = datetime.strptime(dates[i-1][0], "%Y-%m-%d")
                d2 = datetime.strptime(dates[i][0], "%Y-%m-%d")
                intervals.append((d2 - d1).days)
            avg_interval = round(sum(intervals) / len(intervals)) if intervals else 0

            # 取 base name
            base_name = cust.split("-")[0].strip() if "-" in cust else cust

            results.append({
                "name": base_name,
                "full_name": cust,
                "total_amount": int(total),
                "order_count": order_days,
                "unique_products": unique_prods,
                "avg_interval_days": avg_interval,
                "fav_category": fav_cat,
                "top_items": top_items,
            })
    return results


# ── 4. 補貨預測 ──────────────────────────────────

def restock_forecast(days_history: int = 30) -> list[dict]:
    """根據最近銷售速度預測幾天後斷貨（用銷貨明細）"""
    import json
    avail_path = Path(__file__).parent.parent / "data" / "available.json"
    if not avail_path.exists():
        return []
    avail = json.loads(avail_path.read_text(encoding="utf-8"))

    cutoff = (datetime.now() - timedelta(days=days_history)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) as total_out
            FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
            HAVING total_out > 0
        """, (cutoff,)).fetchall()

    results = []
    for code, name, total_out in rows:
        if _is_excluded_product(code):
            continue
        stock = 0
        if code in avail:
            d = avail[code]
            stock = d.get("available", 0) if isinstance(d, dict) else d
        if stock <= 0:
            continue
        daily_avg = total_out / days_history
        if daily_avg <= 0:
            continue
        days_left = round(stock / daily_avg)
        results.append({
            "code": code, "name": name,
            "stock": stock,
            "daily_avg_out": round(daily_avg, 1),
            "days_left": days_left,
            "category": _classify(name),
        })
    results.sort(key=lambda x: x["days_left"])
    return results


# ── 5. 價位帶分析 ──────────────────────────────────

def price_band_analysis(days: int = 90) -> list[dict]:
    """各價位帶的銷售表現"""
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT prod_cd, prod_name, unit_price, SUM(qty) as total_qty, SUM(amount) as total_amount
            FROM sales_detail
            WHERE date >= ? AND unit_price > 0 AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (cutoff,)).fetchall()

    bands = defaultdict(lambda: {"count": 0, "total_qty": 0, "total_amount": 0, "products": []})
    for code, name, price, qty, amount in rows:
        if _is_excluded_product(code):
            continue
        if price <= 50:
            band = "50以下"
        elif price <= 100:
            band = "51-100"
        elif price <= 150:
            band = "101-150"
        elif price <= 200:
            band = "151-200"
        elif price <= 300:
            band = "201-300"
        else:
            band = "300以上"

        bands[band]["count"] += 1
        bands[band]["total_qty"] += int(qty)
        bands[band]["total_amount"] += int(amount)
        bands[band]["products"].append({"code": code, "name": name, "price": int(price), "qty": int(qty)})

    results = []
    order = ["50以下", "51-100", "101-150", "151-200", "201-300", "300以上"]
    for band in order:
        if band in bands:
            b = bands[band]
            b["products"].sort(key=lambda x: -x["qty"])
            results.append({
                "band": band,
                "product_count": b["count"],
                "total_qty": b["total_qty"],
                "total_amount": b["total_amount"],
                "avg_qty_per_product": round(b["total_qty"] / max(b["count"], 1)),
                "top3": b["products"][:3],
            })
    return results


# ── 6. 品類分析 ──────────────────────────────────

def category_analysis(days: int = 90) -> list[dict]:
    """各品類銷售佔比"""
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) as total_qty, SUM(amount) as total_amount
            FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (cutoff,)).fetchall()

    cats = defaultdict(lambda: {"qty": 0, "amount": 0, "products": 0, "items": []})
    for code, name, qty, amount in rows:
        if _is_excluded_product(code):
            continue
        cat = _classify(name)
        cats[cat]["qty"] += int(qty)
        cats[cat]["amount"] += int(amount)
        cats[cat]["products"] += 1
        cats[cat]["items"].append({"name": name, "qty": int(qty)})

    total_amount = sum(c["amount"] for c in cats.values())
    results = []
    for cat, data in sorted(cats.items(), key=lambda x: -x[1]["amount"]):
        data["items"].sort(key=lambda x: -x["qty"])
        pct = round(data["amount"] / total_amount * 100, 1) if total_amount else 0
        results.append({
            "category": cat,
            "product_count": data["products"],
            "total_qty": data["qty"],
            "total_amount": data["amount"],
            "pct": pct,
            "top3": data["items"][:3],
        })
    return results


# ── 7. 月趨勢分析 ──────────────────────────────────

def monthly_trend() -> list[dict]:
    """每月銷售趨勢（金額、筆數、客戶數）"""
    with _conn() as conn:
        rows = conn.execute("""
            SELECT substr(date,1,7) as m,
                   CAST(SUM(amount) AS INTEGER),
                   COUNT(*),
                   COUNT(DISTINCT customer)
            FROM sales_detail
            WHERE customer NOT LIKE '%民享%'
            GROUP BY m ORDER BY m
        """).fetchall()
    return [{"month": r[0], "amount": r[1], "orders": r[2], "customers": r[3]} for r in rows]


# ── 8. 產品成長/衰退 ──────────────────────────────────

def product_trend(days: int = 90) -> dict:
    """比較前半段 vs 後半段銷量，找出成長和衰退的品項"""
    mid = (datetime.now() - timedelta(days=days//2)).strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")

    with _conn() as conn:
        first_half = {}
        for r in conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) FROM sales_detail
            WHERE date >= ? AND date < ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (start, mid)).fetchall():
            if not _is_excluded_product(r[0]):
                first_half[r[0]] = {"name": r[1], "qty": int(r[2])}

        second_half = {}
        for r in conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (mid,)).fetchall():
            if not _is_excluded_product(r[0]):
                second_half[r[0]] = {"name": r[1], "qty": int(r[2])}

    growing = []
    declining = []
    for code in set(list(first_half.keys()) + list(second_half.keys())):
        q1 = first_half.get(code, {}).get("qty", 0)
        q2 = second_half.get(code, {}).get("qty", 0)
        name = second_half.get(code, first_half.get(code, {})).get("name", code)
        if q1 > 10 and q2 > q1 * 1.5:
            growing.append({"code": code, "name": name, "before": q1, "after": q2,
                          "growth": round((q2-q1)/q1*100)})
        elif q1 > 20 and q2 < q1 * 0.5:
            declining.append({"code": code, "name": name, "before": q1, "after": q2,
                            "decline": round((q1-q2)/q1*100)})

    growing.sort(key=lambda x: -x["growth"])
    declining.sort(key=lambda x: -x["decline"])
    return {"growing": growing[:10], "declining": declining[:10]}


# ── 9. 庫存周轉率 ──────────────────────────────────

def stock_turnover(days: int = 90) -> list[dict]:
    """庫存周轉率：銷量 ÷ 目前庫存，越高代表賣得越快（用銷貨明細）"""
    import json
    avail_path = Path(__file__).parent.parent / "data" / "available.json"
    if not avail_path.exists():
        return []
    avail = json.loads(avail_path.read_text(encoding="utf-8"))

    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with _conn() as conn:
        rows = conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) as total_out
            FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (cutoff,)).fetchall()

    results = []
    for code, name, total_out in rows:
        if _is_excluded_product(code):
            continue
        stock = 0
        if code in avail:
            d = avail[code]
            stock = d.get("available", 0) if isinstance(d, dict) else d
        if stock <= 0:
            continue
        turnover = round(total_out / stock, 1)
        results.append({
            "code": code, "name": name,
            "total_out": int(total_out), "stock": stock,
            "turnover": turnover,
            "category": _classify(name),
        })
    results.sort(key=lambda x: -x["turnover"])
    return results


# ── 10. 客戶流失偵測 ──────────────────────────────────

def customer_churn(days_inactive: int = 60) -> list[dict]:
    """最近 N 天沒下單但之前有下單的客戶"""
    cutoff = (datetime.now() - timedelta(days=days_inactive)).strftime("%Y-%m-%d")
    with _conn() as conn:
        # 之前有下單但最近沒下的
        rows = conn.execute("""
            SELECT customer, MAX(date) as last_order,
                   COUNT(DISTINCT date) as total_days, CAST(SUM(amount) AS INTEGER) as total_amt
            FROM sales_detail
            WHERE customer NOT LIKE '%民享%' AND customer != ''
            GROUP BY customer
            HAVING last_order < ?
            ORDER BY total_amt DESC
        """, (cutoff,)).fetchall()

    from services.rebate import _get_base_name
    seen = set()
    results = []
    for cust, last, days_count, amt in rows:
        base = _get_base_name(cust)
        if base in seen:
            continue
        seen.add(base)
        inactive_days = (datetime.now() - datetime.strptime(last, "%Y-%m-%d")).days
        results.append({
            "name": base, "last_order": last,
            "inactive_days": inactive_days,
            "total_orders": days_count, "total_amount": amt,
        })
    return results[:20]


# ── 11. 不建議叫貨清單 ──────────────────────────────────

def do_not_restock() -> list[dict]:
    """綜合分析：哪些品項不建議再叫貨"""
    import json
    avail_path = Path(__file__).parent.parent / "data" / "available.json"
    if not avail_path.exists():
        return []
    avail = json.loads(avail_path.read_text(encoding="utf-8"))

    with _conn() as conn:
        # 近 90 天 vs 前 90 天的銷量
        now = datetime.now()
        mid = (now - timedelta(days=90)).strftime("%Y-%m-%d")
        old_start = (now - timedelta(days=180)).strftime("%Y-%m-%d")

        recent = {}
        for r in conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) FROM sales_detail
            WHERE date >= ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (mid,)).fetchall():
            recent[r[0]] = {"name": r[1], "qty": int(r[2])}

        old = {}
        for r in conn.execute("""
            SELECT prod_cd, SUM(qty) FROM sales_detail
            WHERE date >= ? AND date < ? AND customer NOT LIKE '%民享%'
            GROUP BY prod_cd
        """, (old_start, mid)).fetchall():
            old[r[0]] = int(r[1])

        # 最後出庫日
        last_out = {}
        for r in conn.execute("""
            SELECT prod_cd, MAX(date) FROM inventory_changes
            WHERE qty_out > 0 GROUP BY prod_cd
        """).fetchall():
            last_out[r[0]] = r[1]

    results = []
    for code, d in avail.items():
        if _is_excluded_product(code):
            continue
        stock = d.get("available", 0) if isinstance(d, dict) else d
        if stock <= 0:
            continue

        name = recent.get(code, {}).get("name", "")
        if not name:
            from services.ecount import ecount_client
            item = ecount_client.get_product_cache_item(code)
            name = (item.get("name") if item else None) or code

        recent_qty = recent.get(code, {}).get("qty", 0)
        old_qty = old.get(code, 0)
        last_sale = last_out.get(code, "N/A")
        days_since = 0
        if last_sale and last_sale != "N/A":
            try:
                days_since = (datetime.now() - datetime.strptime(last_sale, "%Y-%m-%d")).days
            except Exception:
                pass

        # 檢查是否「之前斷貨 → 最近才補貨」的情況
        # 方式1：庫存變更有近期入庫
        with _conn() as conn2:
            _recent_in = conn2.execute(
                "SELECT MIN(date) FROM inventory_changes WHERE prod_cd=? AND qty_in > 0 AND date >= ?",
                (code, mid)
            ).fetchone()
        if _recent_in and _recent_in[0]:
            try:
                _restock_date = datetime.strptime(_recent_in[0], "%Y-%m-%d")
                _days_in_stock = (datetime.now() - _restock_date).days
                if _days_in_stock < 30:
                    continue
            except Exception:
                pass

        # 方式2：以前有賣但中間斷貨（old_qty > 0 但 recent_qty 很低），且現在庫存突然很多
        # → 代表最近才補貨，不是賣不動
        if old_qty > 20 and recent_qty < 5 and stock > old_qty * 0.5:
            # 以前賣得不錯，最近沒賣，但庫存很多 → 可能剛補貨
            # 再查是否中間有斷貨（庫存變更中 balance 曾 <= 0）
            with _conn() as conn3:
                _was_zero = conn3.execute(
                    "SELECT 1 FROM inventory_changes WHERE prod_cd=? AND balance <= 0 AND date >= ?",
                    (code, old_start)
                ).fetchone()
            if _was_zero:
                continue  # 確認中間斷過貨，跳過

        # 評分：越高越不建議叫
        score = 0
        reasons = []

        # 銷量大幅衰退（前提：以前有賣過且現在有庫存）
        if old_qty > 20 and recent_qty < old_qty * 0.3:
            score += 3
            reasons.append(f"銷量衰退{round((1-recent_qty/old_qty)*100)}%")
        elif old_qty > 10 and recent_qty < old_qty * 0.5:
            score += 2
            reasons.append(f"銷量下滑{round((1-recent_qty/max(old_qty,1))*100)}%")

        # 庫存高但出貨慢
        if stock > 50 and recent_qty < 10:
            score += 3
            reasons.append(f"庫存{stock}但近90天只出{recent_qty}個")
        elif stock > 30 and recent_qty < 5:
            score += 2
            reasons.append(f"庫存{stock}近90天出{recent_qty}個")

        # 很久沒出貨（但要有庫存超過 30 天才算）
        if days_since > 90:
            score += 2
            reasons.append(f"已{days_since}天沒出貨")
        elif days_since > 60:
            score += 1
            reasons.append(f"已{days_since}天沒出貨")

        if score >= 3:
            results.append({
                "code": code, "name": name, "stock": stock,
                "recent_qty": recent_qty, "old_qty": old_qty,
                "last_sale": last_sale, "score": score,
                "reasons": "、".join(reasons),
                "category": _classify(name),
            })

    results.sort(key=lambda x: -x["score"])
    return results


# ── 12. 新品採購建議 ──────────────────────────────────

def new_product_suggestion(category: str, price: int) -> dict:
    """根據品類+價位，建議新品採購數量"""
    # 同品類歷史銷量
    with _conn() as conn:
        all_products = conn.execute("""
            SELECT prod_cd, prod_name, SUM(qty) as total, SUM(amount) as total_amt,
                   unit_price
            FROM sales_detail
            WHERE unit_price > 0
            GROUP BY prod_cd
        """).fetchall()

    same_cat = []
    same_band = []
    for code, name, qty, amt, up in all_products:
        if _is_excluded_product(code):
            continue
        cat = _classify(name)
        if cat == category:
            same_cat.append({"code": code, "name": name, "qty": int(qty), "price": int(up)})
        if abs(up - price) <= 50:
            same_band.append({"code": code, "name": name, "qty": int(qty), "price": int(up)})

    # 同品類平均月銷量
    months = 3  # 假設 3 個月資料
    cat_avg = round(sum(p["qty"] for p in same_cat) / max(len(same_cat), 1) / months) if same_cat else 0
    band_avg = round(sum(p["qty"] for p in same_band) / max(len(same_band), 1) / months) if same_band else 0

    # 建議數量
    suggested = max(cat_avg, band_avg)
    if suggested < 10:
        suggested = 10  # 最低建議
    # 取整到 10 的倍數
    suggested = round(suggested / 10) * 10

    return {
        "category": category,
        "price": price,
        "same_category_count": len(same_cat),
        "same_category_avg_monthly": cat_avg,
        "same_priceband_count": len(same_band),
        "same_priceband_avg_monthly": band_avg,
        "suggested_qty": suggested,
        "top_in_category": sorted(same_cat, key=lambda x: -x["qty"])[:5],
        "top_in_priceband": sorted(same_band, key=lambda x: -x["qty"])[:5],
    }


# ── 綜合報告 ──────────────────────────────────

def full_report() -> str:
    """生成綜合分析報告文字"""
    lines = ["📊 銷售分析報告\n"]

    # 1. 銷售排行
    top = top_sellers(30, 10)
    if top:
        lines.append("🔥 【近30天銷售排行 TOP 10】")
        for i, p in enumerate(top, 1):
            lines.append(f"  {i:2}. {p['code']} {p['name'][:15]} 出{p['total_out']}個 日均{p['daily_avg']}")
        lines.append("")

    # 2. 滯銷品
    slow = slow_movers(60, 20)
    if slow:
        lines.append(f"⚠️ 【滯銷品（60天無出庫，庫存≥20）】共 {len(slow)} 項")
        for p in slow[:5]:
            lines.append(f"  {p['code']} {p['name'][:15]} 庫存{p['stock']} 上次出庫{p['last_sale']}")
        lines.append("")

    # 3. 補貨預測
    forecast = restock_forecast(30)
    urgent = [f for f in forecast if f["days_left"] <= 7]
    if urgent:
        lines.append(f"🚨 【7天內可能斷貨】共 {len(urgent)} 項")
        for p in urgent[:5]:
            lines.append(f"  {p['code']} {p['name'][:15]} 庫存{p['stock']} 日均出{p['daily_avg_out']} 剩{p['days_left']}天")
        lines.append("")

    # 4. 價位帶
    pb = price_band_analysis(90)
    if pb:
        lines.append("💰 【價位帶表現（近90天）】")
        for b in pb:
            lines.append(f"  {b['band']:>8}  {b['product_count']}品項  出{b['total_qty']}個  ${b['total_amount']:,}")
        lines.append("")

    # 5. 品類
    cats = category_analysis(90)
    if cats:
        lines.append("📦 【品類銷售佔比（近90天）】")
        for c in cats[:8]:
            lines.append(f"  {c['category']:>10}  {c['product_count']}品項  {c['pct']}%  ${c['total_amount']:,}")
        lines.append("")

    # 6. 客戶 TOP 10
    custs = customer_analysis(90, 10)
    if custs:
        lines.append("👥 【客戶 TOP 10（近90天）】")
        for c in custs:
            interval = f"每{c['avg_interval_days']}天" if c['avg_interval_days'] > 0 else "單次"
            lines.append(f"  {c['name'][:8]}  ${c['total_amount']:,}  {c['order_count']}次  {interval}  愛買:{c['fav_category']}")

    return "\n".join(lines)
