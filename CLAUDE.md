# CLAUDE.md — 小蠻牛 (BEAR) LINE Customer Service Bot

> AI assistant guide for working with this codebase.

---

## Project Overview

A **LINE Official Account bot** for a retail business (小蠻牛), handling customer service, order processing, inventory queries, and staff operations. Integrated with **Ecount ERP** for real-time inventory sync.

- **Language**: Python 3.10+
- **Framework**: FastAPI + uvicorn (async)
- **Database**: SQLite (multiple `.db` files in `data/`)
- **External services**: LINE Messaging API, Ecount ERP OAPI v2, Google Calendar
- **Deployment**: Windows machine with system tray daemon (`tray.py`) + Caddy reverse proxy
- **Domain**: `xmnline.duckdns.org` (DuckDNS + Caddy HTTPS)

---

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and configure environment
cp .env.example .env
# Edit .env with your LINE/Ecount/Google credentials

# Run development server
uvicorn main:app --reload --host 0.0.0.0 --port 8000

# Production (Windows)
python tray.py   # System tray + watchdog + Caddy
```

---

## Directory Structure

```
├── main.py                 # FastAPI app, webhooks, admin API, lifespan scheduler (~3000 lines)
├── config.py               # Pydantic Settings — all env vars
├── tray.py                 # System tray daemon + watchdog (auto-restart uvicorn + Caddy)
│
├── handlers/               # Business logic (intent-based routing)
│   ├── intent.py           # Intent detection (22 intents, keyword-based)
│   ├── internal.py         # Internal group + HQ commands (~2600 lines)
│   ├── inventory.py        # Customer inventory queries
│   ├── ordering.py         # Order flow + box/unit conversion
│   ├── restock.py          # HQ restocking flow
│   ├── summary.py          # Pending items summary ("清單" command)
│   ├── hours.py            # Business hours queries
│   ├── delivery.py         # Delivery timeline
│   ├── escalate.py         # Human handoff
│   ├── tone.py             # Human speech style simulation
│   ├── service.py          # Bargaining, specs, returns, complaints
│   ├── visit.py            # Store visit reservations
│   ├── followup.py         # Scheduled follow-ups (24h/48h)
│   ├── price.py            # Price queries
│   ├── payment.py          # Payment handling
│   ├── orders.py           # Order tracking
│   └── ad_maker.py         # Ad content generation
│
├── services/               # External integrations
│   ├── ecount.py           # Ecount ERP API client (Big5/GBK encoding)
│   ├── vision.py           # Image recognition (pHash + OCR)
│   ├── refresh.py          # Scheduled data refresh
│   ├── google_cal.py       # Google Calendar integration
│   └── inventory_csv.py    # CSV import
│
├── storage/                # Data access layer (raw SQLite, no ORM)
│   ├── customers.py        # Customer DB
│   ├── state.py            # In-memory conversation state
│   ├── persistent_state.py # Persistent state (survives restart)
│   ├── restock.py          # Restock requests
│   ├── new_products.py     # New product approvals
│   ├── pending.py          # Pending queries
│   ├── issues.py           # Returns/complaints/address changes
│   ├── visits.py           # Store visit records
│   ├── specs.py            # Product specs (JSON-backed)
│   ├── reserved.py         # Reserved inventory (in-memory)
│   ├── payments.py         # Payment records
│   ├── queue.py            # Off-peak message queue
│   ├── cart.py             # Shopping cart (in-memory)
│   ├── notify.py           # Notification log
│   ├── tags_config.py      # Tag configuration
│   └── delivery.py         # Delivery info
│
├── scripts/                # Utility scripts (sync, import, generation)
├── static/admin.html       # Admin dashboard SPA
├── data/                   # Runtime data (SQLite DBs, JSON caches) — gitignored
├── .env                    # Credentials — gitignored, never commit
├── ARCHITECTURE.md         # System architecture (Chinese)
└── BUGFIX_LOG.md           # Bug fix and feature log (Chinese)
```

---

## Architecture & Key Patterns

### Message Flow
```
LINE → POST /webhook → Signature verification → Intent detection → Handler dispatch → LINE reply
```

### Intent-Based Routing
`handlers/intent.py` defines 22 intents (`Intent` enum). Detection is keyword-based with regex. Each intent maps to a dedicated handler module.

### Conversation State
- **Transient**: `storage/state.py` — in-memory dict, lost on restart
- **Persistent**: `storage/persistent_state.py` — SQLite, survives restart
- State transitions drive multi-turn conversations (e.g., `awaiting_product_code` → `awaiting_quantity`)

### Data Storage
- **No ORM** — all SQLite access uses raw SQL with context managers
- Storage modules follow the pattern: class with `_init_db()`, `add()`, `query()`, `update()` methods
- Module-level `DB_PATH` constant pointing to `data/*.db`
- JSON files for caches (`data/available.json`, `data/specs.json`, etc.)

### Ecount ERP Integration
- `services/ecount.py` — async HTTP client with session management
- Encoding detection chain: UTF-8 → Big5 → GBK → GB18030
- `_safe_json()` handles mixed encoding responses
- Product cache: 6-hour refresh, excludes Z+letter product codes
- Thread lock for sync operations (`_sync_lock`)

### Background Tasks (Lifespan)
Defined in `main.py` lifespan handler:
- Inventory sync (every 2 min via `services/refresh.py`)
- Hourly summary reports
- Message queue processor
- Restock notification checker
- Follow-up reminders (24h/48h)

### Message Buffering
- Image messages held up to 6 seconds for follow-up text
- Text messages coalesced (5 sec timeout)
- Prevents fragmented order parsing

### Role-Based Access
| Context | Capabilities |
|---------|-------------|
| Customer (1:1 chat) | Inventory, order, hours, delivery, bargaining |
| Internal group (`LINE_GROUP_ID`) | All admin commands, proxy orders |
| HQ group (`LINE_GROUP_ID_HQ`) | Restock coordination |
| Admin panel (`/admin/*`) | HTTP Basic Auth dashboard |

---

## Internal Group Order Formats

Staff can place orders via the internal LINE group using these formats:

| Format | Example |
|--------|---------|
| A (single-line) | `楊庭瑋 訂 Z2095 30個` |
| B (multi-line with 訂) | `鄭鉅耀 訂\nZ3340 10個\n備註 送松山` |
| B2 (multi-line, name only) | `鄭鉅耀\nZ3340 10個\n備註 送松山` |
| C (direct) | `方力緯 Z3562 5個` |
| D (product name, single) | `曹竣智 要 洗衣球 5` |
| E (product name, multi) | `楊庭瑋 衛生紙30箱 泡澡球10件` |

Notes keyword: `備註` / `備誌` / `備记` + space or colon.

---

## Configuration

All settings are in `config.py` via **pydantic-settings** (`BaseSettings`), loaded from `.env`.

Key environment variables:
- `LINE_CHANNEL_ACCESS_TOKEN`, `LINE_CHANNEL_SECRET` — LINE API auth
- `LINE_GROUP_ID`, `LINE_GROUP_ID_HQ`, `LINE_GROUP_ID_SHOWCASE` — group routing
- `ECOUNT_COMPANY_NO`, `ECOUNT_USER_ID`, `ECOUNT_API_CERT_KEY` — ERP auth
- `ECOUNT_ZONE` (default: `IB`), `ECOUNT_BASE_URL`
- `BUSINESS_HOURS_START/END`, `BUSINESS_DAYS`, `BUSINESS_TZ`
- `ADMIN_USER`, `ADMIN_PASS` — admin dashboard auth
- `PRODUCT_MEDIA_PATH` — local path to product images

---

## Code Conventions

### Style
- **Type hints** throughout (Python 3.10+ `str | None` syntax)
- **Async/await** for all FastAPI endpoints and background tasks
- **Module-level constants** for regex patterns (verbose with comments)
- **Chinese** in user-facing strings, comments, and documentation

### Database Access Pattern
```python
DB_PATH = os.path.join("data", "example.db")

class ExampleStore:
    def __init__(self):
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("CREATE TABLE IF NOT EXISTS ...")

    def add(self, ...):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("INSERT INTO ...")

    def query(self, ...):
        with sqlite3.connect(DB_PATH) as conn:
            return conn.execute("SELECT ...").fetchall()
```

### Error Handling
- Try/except with logging to stdout / `server.log`
- Graceful fallbacks (Ecount API failure → use JSON cache)
- User-facing errors sent via LINE reply

### Naming
- Handler files named by feature: `inventory.py`, `ordering.py`, `visit.py`
- Storage files named by entity: `customers.py`, `payments.py`, `visits.py`
- Private functions prefixed with `_` (e.g., `_do_order()`, `_safe_json()`)

---

## Testing

No automated test framework (pytest). Testing is manual via scripts:

```bash
# Test scripts in scripts/ directory
python scripts/test_flow.py              # Conversation flow
python scripts/test_new_product_parse.py # Product parsing
python scripts/test_save_order.py        # Order saving

# Root-level test utilities
python test_ecount_api.py                # Ecount API
python simulate.py                       # Message simulation
```

---

## Common Tasks for AI Assistants

### Adding a New Intent
1. Add enum value to `Intent` in `handlers/intent.py`
2. Add keyword list `_NEW_KEYWORDS` in `handlers/intent.py`
3. Add detection logic in `detect_intent()` function
4. Create handler module `handlers/new_handler.py`
5. Wire handler in `main.py` webhook processing

### Adding a New Storage Module
1. Create `storage/new_entity.py` following the pattern above
2. Set `DB_PATH = os.path.join("data", "new_entity.db")`
3. Implement `_init_db()`, `add()`, `query()`, `update()` methods
4. Import and instantiate in the handler that needs it

### Adding an Admin API Endpoint
1. Add route in `main.py` under the admin section
2. Protect with HTTP Basic Auth (existing middleware)
3. Update `static/admin.html` to display data

### Modifying Order Format Parsing
- All format parsing lives in `handlers/internal.py`
- Regex patterns are module-level constants
- Test with `scripts/test_flow.py` or `simulate.py`

---

## Important Warnings

- **Never commit `.env`** — contains LINE tokens, Ecount API keys, admin credentials
- **`data/` directory is gitignored** — contains SQLite DBs and caches; will be recreated at runtime
- **`main.py` is ~3000 lines** — the main entry point; be careful with large edits
- **`handlers/internal.py` is ~2600 lines** — handles all internal group logic; complex regex parsing
- **Chinese text handling** — regex boundaries (`\b`) don't work at Chinese/English boundaries; use `(?<![A-Za-z])` lookahead/lookbehind instead
- **Ecount API encoding** — responses may be Big5/GBK; always use `_safe_json()` pattern
- **Message buffering** — image + text coalescing means handlers may receive combined input after a delay

---

## Known Issues (from BUGFIX_LOG.md)

| ID | Issue | Priority |
|----|-------|----------|
| P-01 | `_followup_loop` — `line_api` not in scope (follow-ups may fail) | High |
| P-02 | Startup script points to `start.bat` instead of `start_tray.bat` | Medium |
| P-03 | `customer_group_address` table empty (multi-address feature incomplete) | Low |
| P-04 | Google Calendar integration incomplete | Low |
| P-05 | Ambiguous product resolution adds extra confirmation step | Low |

---

## Dependencies

```
fastapi>=0.110.0
uvicorn[standard]>=0.27.0
line-bot-sdk>=3.5.0
pydantic-settings>=2.0.0
httpx>=0.25.0
pytz>=2024.1
google-auth-oauthlib>=1.0.0
google-api-python-client>=2.100.0
Pillow>=10.0.0
imagehash>=4.3.1
```

Additional tools (not in requirements.txt):
- **Playwright** — browser automation for Ecount data scraping
- **Selenium** — alternative automation
- **pystray** — system tray icon (Windows)
