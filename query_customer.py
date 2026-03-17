import sqlite3, os, sys
sys.stdout.reconfigure(encoding='utf-8')
os.chdir(r'C:\Users\bear\Desktop\code\line-cs-bot')
conn = sqlite3.connect('data/customers.db')
conn.row_factory = sqlite3.Row
rows = conn.execute("SELECT line_user_id, display_name, real_name FROM customers ORDER BY display_name").fetchall()
print(f"Total: {len(rows)}")
for r in rows:
    d = dict(r)
    dn = d['display_name'] or ''
    if '佛' in dn or '爺' in dn:
        print(f"  MATCH: display_name={repr(dn)} real_name={repr(d['real_name'])} line_user_id={d['line_user_id']}")
print("--- All display_names ---")
for r in rows:
    print(dict(r)['display_name'])
conn.close()
