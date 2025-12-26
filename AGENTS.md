# Repository Guidelines

## Project Structure & Module Organization

- `immich_holiday_album_collector.py` — Tkinter GUI that searches an Immich instance for assets taken around selected holidays and adds them to per-holiday albums via the Immich REST API.
- `immich-openapi-specs.json` — vendored OpenAPI 3 schema for the Immich API (reference file; avoid hand-editing).
- Runtime output: `immich_holiday_album_collector.log` is created/updated when the app runs.

## Build, Test, and Development Commands

- Create a venv: `python3 -m venv .venv && source .venv/bin/activate`
- Install dependencies: `pip install -r requirements.txt`
- Run locally: `python3 immich_holiday_album_collector.py`
- Quick sanity check: `python3 -m py_compile immich_holiday_album_collector.py`
- Build Windows `.exe` (CI): push a `v*` tag to trigger `.github/workflows/build-windows-exe.yml`

## Coding Style & Naming Conventions

- Python: 4-space indentation; follow PEP 8 naming (`snake_case` functions, `UPPER_SNAKE_CASE` constants).
- Keep UI work inside `create_gui()` and avoid blocking the Tkinter mainloop; long-running work should run in the background thread and report progress via `progress_queue`.
- Prefer small, reusable helpers for date logic and API calls.

## Testing Guidelines

This repo does not currently include an automated test suite. If you add one, place tests under `tests/` and start by covering:

- Holiday/date helpers (deterministic year-based outputs)
- API request payload construction (no network calls)

## Commit & Pull Request Guidelines

This directory does not currently include Git metadata, so there is no established commit-message convention. Recommended:

- Use Conventional Commits (`feat: …`, `fix: …`, `chore: …`) and keep changes focused.
- PRs should describe behavior changes and include screenshots for GUI updates.

## Security & Configuration Tips

- Create `app_config.json` (copy `app_config.example.json`) and set `api_base_url` to your Immich `/api` endpoint.
- Never commit API keys or tokens; the app is designed to store the API key in the OS keyring (`keyring`).
