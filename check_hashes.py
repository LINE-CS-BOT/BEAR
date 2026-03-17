import json, os
os.chdir(r'C:\Users\bear\Desktop\code\line-cs-bot')
path = 'data/image_hashes.json'
if not os.path.exists(path):
    print('image_hashes.json 不存在！需要執行 scripts/build_image_hashes.py')
else:
    db = json.loads(open(path, encoding='utf-8').read())
    print(f'共 {len(db)} 筆雜湊記錄')
    for entry in db[:5]:
        print(f"  code={entry.get('code')}  file={entry.get('file','')}")
