# Immich Holiday Album Collector

![Immich Holiday Album Collector GUI](docs/gui.png)

A small Tkinter app that searches your Immich library for assets taken around holidays (and/or a custom date like a birthday) and automatically adds them to albums.

## Features

- **Holiday albums across years**: pick one or more holidays and a year range; the app searches each holiday for every year in the range and adds matches to a per-holiday album.
- **Delta windowing**: expand each holiday/date into a +/- day window (e.g. `7` searches two weeks centered on the date; `0` searches only that day).
- **Specific date**: run once or repeat the same month/day across all years (useful for birthdays/anniversaries).
- **Advanced filters**:
  - People filter with **OR (any)** / **AND (all)** matching
  - Raw `/search/metadata` JSON filters (e.g. `{"isFavorite": true}`)
- **People picker**: load all people, scroll, type-to-filter, multi-select, and add to the People filter.
- **Presets**: save/load UI settings to `config.json` (no API key stored there).

## Requirements

- Python 3 with Tk support (Tkinter).
- Immich server reachable from your machine.
- Immich API key with permissions for:
  - `asset.read`, `album.read`, `album.create`, `album.write` (for searching and adding to albums)
  - `person.read` (only if you use People filters/picker)

## Download (Windows)

1. Download the latest `immich-holiday-album-collector-windows.zip` from GitHub Releases (contains the `.exe` and `app_config.example.json`).
2. In the same folder as the `.exe`, copy `app_config.example.json` to `app_config.json` and set `api_base_url` (e.g. `https://immich.example.com/api`).
3. Double-click the `.exe`.

## Run From Source

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp app_config.example.json app_config.json
python3 immich_holiday_album_collector.py
```

If you hit Tk-related crashes on macOS, use a Python distribution that ships Tk 8.6 (e.g. Anaconda or python.org).

### Windows (Python)

Install Python 3 from python.org (includes Tk), then:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
Copy-Item app_config.example.json app_config.json
py immich_holiday_album_collector.py
```

Set `api_base_url` to your Immich API endpoint (e.g. `https://immich.example.com/api`). If you provide only the host, the app will assume `/api`.

## Advanced Search Notes

- **People picker search**
  - OR: `jack, jill` or `jack or jill`
  - AND: `jack jill` or `jack and jill`
- **Additional metadata filters (JSON)** are merged into the `/search/metadata` request. Avoid `takenAfter`, `takenBefore`, `page`, `size` (the app controls those). If you include `personIds`, it will be combined with the People filter.

## Privacy & Safety

- API keys are stored in your OS keyring (optional). Never commit keys/tokens.
- Logs can include request payloads and asset IDs; avoid sharing them. This repo’s `.gitignore` excludes common sensitive/local files (`app_config.json`, logs, presets, `.env`, certs, etc.).
