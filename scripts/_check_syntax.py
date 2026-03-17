import ast, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

files = [
    "main.py",
    "handlers/service.py",
    "services/vision.py",
    "storage/specs.py",
    "handlers/tone.py",
]

all_ok = True
for f in files:
    try:
        src = open(f, encoding="utf-8").read()
        ast.parse(src)
        print(f"  OK  {f}")
    except SyntaxError as e:
        print(f"  !! {f}: {e}")
        all_ok = False

sys.exit(0 if all_ok else 1)
