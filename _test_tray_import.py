import sys
sys.path.insert(0, r'C:\Users\bear\Desktop\code\line-cs-bot')
try:
    import tray
    print("import OK")
except Exception as e:
    print(f"import FAILED: {e}")
    import traceback; traceback.print_exc()
