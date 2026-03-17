from pathlib import Path
import json

txt = Path(r'H:\其他電腦\我的電腦\小蠻牛\產品PO文.txt')
print('PO文.txt exists:', txt.exists())
if txt.exists():
    content = txt.read_text(encoding='utf-8', errors='ignore')
    print('file size:', len(content))
    # 用 \n\n 切塊
    blocks = [b.strip() for b in content.split('\n\n') if b.strip()]
    print('blocks (\\n\\n split):', len(blocks))
    found = False
    for b in blocks:
        if 'T1198' in b.upper():
            print('=== T1198 block found ===')
            print(repr(b[:200]))
            found = True
            break
    if not found:
        print('T1198 NOT found with \\n\\n split')
        # 試試別的分隔方式
        lines = content.splitlines()
        print(f'total lines: {len(lines)}')
        for i, line in enumerate(lines):
            if 'T1198' in line.upper():
                print(f'Found T1198 at line {i}: {repr(line)}')
                # 印前後5行
                start = max(0, i-2)
                end = min(len(lines), i+8)
                for l in lines[start:end]:
                    print(repr(l))
                break
        else:
            print('T1198 not found in file at all')
            print('First 3 lines:', lines[:3])

# 也檢查 available.json 的庫存值
avail = Path(r'C:\Users\bear\Desktop\code\line-cs-bot\data\available.json')
print('\navailable.json exists:', avail.exists())
if avail.exists():
    import time
    age = time.time() - avail.stat().st_mtime
    print(f'available.json age: {int(age)} seconds')
    data = json.loads(avail.read_text(encoding='utf-8'))
    print(f'T1198 in available.json: {data.get("T1198", "NOT FOUND")}')
    print(f'total entries: {len(data)}')
