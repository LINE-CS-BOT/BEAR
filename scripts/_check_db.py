"""Quick DB check - customers with ecount_cust_cd"""
import sqlite3
from pathlib import Path

db = Path(__file__).parent.parent / "data" / "customers.db"
conn = sqlite3.connect(str(db))
conn.row_factory = sqlite3.Row
cur = conn.cursor()

cur.execute("SELECT COUNT(*) FROM customers")
total = cur.fetchone()[0]

cur.execute("SELECT COUNT(*) FROM customers WHERE ecount_cust_cd IS NOT NULL AND ecount_cust_cd != ''")
with_code = cur.fetchone()[0]

print(f"Total customers: {total}")
print(f"With ecount_cust_cd: {with_code}")
print()
print("Samples with code:")
cur.execute("SELECT display_name, ecount_cust_cd FROM customers WHERE ecount_cust_cd IS NOT NULL AND ecount_cust_cd != '' LIMIT 8")
for row in cur.fetchall():
    print(f"  {row['display_name']!r:20s}  {row['ecount_cust_cd']}")

conn.close()
