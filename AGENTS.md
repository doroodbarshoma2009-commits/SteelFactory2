# SteelFactory2

سامانه مدیریت تختال (اسلب) فولاد سفید دشت — a Persian/RTL, Jalali-calendar steel-slab
management desktop application.

## Cursor Cloud specific instructions

### What is in this repo (important)
This checkout is a **partial snapshot**, not the full product. It contains only:
- `main (1).py` — the entire Tkinter desktop GUI (~19k lines, single file).
- `schema.sql` — SQL Server Express (T-SQL) schema for the real backend.
- `README.md` — title only.

The GUI **hard-imports** `db_backend` (line ~203: `from db_backend import ...`) and
lazily imports `client.*`, `server.*`, `shared.*`. **None of these modules are in this
repo** — they live in the external `SteelFactory2-v2` framework (SQL Server / HTTP API
backend on port 8080). Without that framework the app cannot connect to its real storage.

### Running the GUI (dev)
There is no build step, no test suite, and no lint config. It is a plain single-file app.
- **Lint / syntax check:** `python3 -m py_compile "main (1).py"`
- **Run:** `python3 "main (1).py"`

Non-obvious requirements to actually launch it on this headless Linux VM:
- **Display:** needs an X server. Use `DISPLAY=:1` (a noVNC display is already running).
- **`STF_ADMIN=1`:** run in admin mode. Otherwise the app forces kiosk fullscreen
  lockdown (`STF_KIOSK=1`) which hides the taskbar and grabs the keyboard — avoid on Linux.
- **`STF_DISABLE_SYNC=1`:** disables the optional legacy LAN P2P sync (UDP 57321 / TCP 57322).
- **Persian font:** module-level font detection `_get_main_font()` has **no fallback** if no
  "B Nazanin"-like family is installed — it raises `NameError` at import. A `B Nazanin` font
  (derived from farsiweb Nazli) is installed at `~/.local/share/fonts/BNazanin.ttf`; do not
  remove it, or the app will fail to import on a headless box.
- **`db_backend` stub:** because the real `db_backend` is missing, a minimal **local
  JSON-backed dev stub** exists at `~/devstub/db_backend.py` (loaded via
  `PYTHONPATH=~/devstub`). It is only for running the GUI locally — it is NOT the product
  and is intentionally not committed. It seeds one admin user (`admin` / password `admin`)
  and stores state in `~/devstub/dev_state.json`. Real deployments use SQL Server / the API.

Full launch command used for verification:
```
cd /workspace
DISPLAY=:1 PYTHONPATH=~/devstub STF_ADMIN=1 STF_DISABLE_SYNC=1 python3 "main (1).py"
```

### Core smoke test (hello-world)
Log in as `admin` / `admin`, open **ثبت ذوب جدید → ثبت دستی**, enter an 11-digit slab
number (validated by `validate_slab_id`, e.g. `40412250001`), click **✔ ثبت اسلب**. A
green success message appears and the record is appended to `melts` and persisted.

### To run the real backend
Add the `SteelFactory2-v2` framework (`db_backend`, `client/`, `server/`, `shared/`) to the
machine and point `STF_ROOT` at it, plus a SQL Server Express instance initialized with
`schema.sql` and the API server on `http://127.0.0.1:8080`. None of that server-side code is
present in this repository.
