import requests
from datetime import datetime, timedelta
import json
import calendar
import tkinter as tk
from tkinter import ttk, messagebox
import re
from urllib.parse import urlsplit, urlunsplit
import os
import queue
import threading
from tkcalendar import DateEntry
import keyring  # For secure API key storage

# -------------------------------
# Configuration Defaults
# -------------------------------
APP_NAME = "Immich Holiday Album Collector"
APP_SLUG = "immich_holiday_album_collector"
APP_CONFIG_FILE = "app_config.json"
API_BASE_URL = ""  # Loaded from APP_CONFIG_FILE at runtime
SERVICE_NAME = "ImmichHolidayAlbumCollector"  # Keyring service name
LEGACY_SERVICE_NAME = "HolidayAssetCollector"  # Backwards-compatible keyring service name
KEY_NAME = "api_key"  # Key name for storing in keyring
REQUEST_TIMEOUT_SECONDS = 30

DEFAULT_HOLIDAYS = [
    "New Year's Day",
    "Martin Luther King Jr. Day",
    "Presidents' Day",
    "Easter",
    "Memorial Day",
    "Juneteenth",
    "Independence Day",
    "Labor Day",
    "Columbus Day",
    "Halloween",
    "Veterans Day",
    "Thanksgiving",
    "Christmas"
]

LOG_FILE = f"{APP_SLUG}.log"

# Global variables for inter-thread communication
stop_event = threading.Event()
progress_queue = queue.Queue()  # Global queue for inter-thread communication

def get_stored_api_key():
    for service_name in (SERVICE_NAME, LEGACY_SERVICE_NAME):
        try:
            value = keyring.get_password(service_name, KEY_NAME)
        except Exception:
            continue
        if value:
            return value
    return ""

def store_api_key_in_keyring(api_key):
    keyring.set_password(SERVICE_NAME, KEY_NAME, api_key)

def delete_api_key_from_keyring():
    deleted = False
    last_error = None
    for service_name in (SERVICE_NAME, LEGACY_SERVICE_NAME):
        try:
            keyring.delete_password(service_name, KEY_NAME)
            deleted = True
        except keyring.errors.PasswordDeleteError as e:
            last_error = e
    if not deleted and last_error:
        raise last_error

# -------------------------------
# App Config Helpers
# -------------------------------
def _normalize_api_base_url(value):
    if not isinstance(value, str):
        return ""
    url = value.strip()
    if not url:
        return ""

    url = url.rstrip("/")

    # Convenience: if the user provides only the host (e.g. https://immich.example.com),
    # default to the Immich API base path at /api.
    try:
        parts = urlsplit(url)
    except ValueError:
        return url

    if parts.scheme in {"http", "https"} and parts.netloc:
        path = (parts.path or "").rstrip("/")
        if path in {"", "/"}:
            return urlunsplit((parts.scheme, parts.netloc, "/api", "", "")).rstrip("/")

    return url

def load_app_config(filename=APP_CONFIG_FILE):
    """Load app configuration from disk (non-secret settings only)."""
    try:
        with open(filename, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data, None
        return {}, f"{filename} must contain a JSON object."
    except FileNotFoundError:
        return {}, None
    except json.JSONDecodeError as e:
        return {}, f"Invalid JSON in {filename}: {e}"
    except OSError as e:
        return {}, f"Failed to read {filename}: {e}"

def set_api_base_url_from_config():
    """Set API_BASE_URL from app_config.json; returns (url, error_message)."""
    global API_BASE_URL
    config, error = load_app_config()
    API_BASE_URL = _normalize_api_base_url(config.get("api_base_url", ""))
    return API_BASE_URL, error

# -------------------------------
# Logging and Status Functions
# -------------------------------
def log_message(message):
    """Log a message to both file and GUI queue."""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(message + "\n")
    progress_queue.put({"type": "log", "text": message})

def set_status(text):
    """Update the status label via the queue."""
    progress_queue.put({"type": "status", "text": text})

def set_progress(value, max_value):
    """Update the progress bar via the queue."""
    progress_queue.put({"type": "progress", "value": value, "max": max_value})

# -------------------------------
# Date Calculation Helpers
# -------------------------------
def get_fixed_date(year, month, day):
    return datetime(year, month, day)

def get_nth_weekday_of_month(year, month, weekday, nth):
    d = datetime(year, month, 1)
    while d.weekday() != weekday:
        d += timedelta(days=1)
    d += timedelta(weeks=(nth-1))
    return d

def get_last_weekday_of_month(year, month, weekday):
    days_in_month = calendar.monthrange(year, month)[1]
    d = datetime(year, month, days_in_month)
    while d.weekday() != weekday:
        d -= timedelta(days=1)
    return d

def get_easter_date(year):
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19*a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2*e + 2*i - h - k) % 7
    m = (a + 11*h + 22*l) // 451
    month = (h + l - 7*m + 114) // 31
    day = ((h + l - 7*m + 114) % 31) + 1
    return datetime(year, month, day)

def get_holiday_date(year, holiday_name):
    if holiday_name == "New Year's Day":
        return get_fixed_date(year, 1, 1)
    elif holiday_name == "Martin Luther King Jr. Day":
        return get_nth_weekday_of_month(year, 1, 0, 3)
    elif holiday_name == "Presidents' Day":
        return get_nth_weekday_of_month(year, 2, 0, 3)
    elif holiday_name == "Easter":
        return get_easter_date(year)
    elif holiday_name == "Memorial Day":
        return get_last_weekday_of_month(year, 5, 0)
    elif holiday_name == "Juneteenth":
        return get_fixed_date(year, 6, 19)
    elif holiday_name == "Independence Day":
        return get_fixed_date(year, 7, 4)
    elif holiday_name == "Labor Day":
        return get_nth_weekday_of_month(year, 9, 0, 1)
    elif holiday_name == "Columbus Day":
        return get_nth_weekday_of_month(year, 10, 0, 2)
    elif holiday_name == "Halloween":
        return get_fixed_date(year, 10, 31)
    elif holiday_name == "Veterans Day":
        return get_fixed_date(year, 11, 11)
    elif holiday_name == "Thanksgiving":
        return get_nth_weekday_of_month(year, 11, 3, 4)
    elif holiday_name == "Christmas":
        return get_fixed_date(year, 12, 25)
    else:
        raise ValueError(f"Unknown holiday: {holiday_name}")

# -------------------------------
# API Interaction Functions
# -------------------------------
def find_or_create_album(headers, album_name):
    url = f"{API_BASE_URL}/albums"
    log_message(f"GET {url}")
    try:
        r = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
        r.raise_for_status()
    except requests.RequestException as e:
        error_msg = f"Failed to fetch albums: {str(e)}"
        log_message(error_msg)
        set_status(error_msg)
        raise
    log_message(f"Response: {r.status_code}")
    albums = r.json()

    for album in albums:
        if album.get("albumName") == album_name:
            log_message(f"Found existing album for {album_name}: {album['id']}")
            return album["id"]

    payload = {
        "albumName": album_name,
        "assetIds": []
    }
    log_message(f"POST {url} with payload: {json.dumps(payload, indent=2)}")
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
        r.raise_for_status()
    except requests.RequestException as e:
        error_msg = f"Failed to create album: {str(e)}"
        log_message(error_msg)
        set_status(error_msg)
        raise
    new_album = r.json()
    log_message(f"Created new album for {album_name}: {new_album['id']}")
    return new_album["id"]

def search_people_by_name(headers, name, with_hidden=False):
    url = f"{API_BASE_URL}/search/person"
    params = {"name": name, "withHidden": with_hidden}
    log_message(f"GET {url} (person search)")
    try:
        r = requests.get(url, headers=headers, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
        r.raise_for_status()
    except requests.RequestException as e:
        error_msg = f"Failed to search people: {str(e)}"
        log_message(error_msg)
        set_status(error_msg)
        raise
    return r.json()

def get_all_people(headers, with_hidden=False):
    url = f"{API_BASE_URL}/people"
    page_size = 1000
    page = 1
    people = []

    while True:
        if stop_event.is_set():
            return []

        params = {"page": page, "size": page_size}
        if with_hidden:
            params["withHidden"] = True

        log_message(f"GET {url} (people list page {page})")
        try:
            r = requests.get(url, headers=headers, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
            r.raise_for_status()
        except requests.RequestException as e:
            error_msg = f"Failed to fetch people: {str(e)}"
            log_message(error_msg)
            set_status(error_msg)
            raise

        data = r.json() or {}
        page_people = data.get("people", []) or []
        if not page_people:
            break

        people.extend(page_people)

        has_next_page = data.get("hasNextPage")
        total = data.get("total")

        if has_next_page is False:
            break
        if isinstance(total, int) and len(people) >= total:
            break
        if has_next_page is None and len(page_people) < page_size:
            break

        page += 1

    return people

def search_assets_by_date_range(headers, start_date, end_date, additional_filters=None):
    page_size = 100
    current_page = 1
    all_asset_ids = []

    while True:
        if stop_event.is_set():
            log_message("Search interrupted by user.")
            return []
        payload = {"withDeleted": False}
        if additional_filters:
            payload.update(additional_filters)
        payload.update({
            "takenAfter": start_date.isoformat(),
            "takenBefore": end_date.isoformat(),
            "size": page_size,
            "page": current_page,
        })

        url = f"{API_BASE_URL}/search/metadata"
        log_message(f"POST {url} with payload: {json.dumps(payload, indent=2)}")
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
            r.raise_for_status()
        except requests.RequestException as e:
            error_msg = f"Failed to search assets: {str(e)}"
            log_message(error_msg)
            set_status(error_msg)
            raise
        log_message(f"Response: {r.status_code}")

        search_results = r.json()
        assets = search_results.get("assets", {}).get("items", [])
        if not assets:
            break

        all_asset_ids.extend([asset["id"] for asset in assets])

        if len(assets) < page_size:
            break

        current_page += 1

    log_message(f"Found {len(all_asset_ids)} total assets in the date range {start_date} - {end_date}")
    return all_asset_ids

def add_assets_to_album(headers, album_id, asset_ids):
    if not asset_ids:
        log_message("No assets to add to album.")
        return 0
    payload = {"ids": asset_ids}
    url = f"{API_BASE_URL}/albums/{album_id}/assets"
    log_message(f"PUT {url} with payload: {json.dumps(payload, indent=2)}")
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
        r.raise_for_status()
    except requests.RequestException as e:
        error_msg = f"Failed to add assets to album: {str(e)}"
        log_message(error_msg)
        set_status(error_msg)
        raise
    log_message(f"Added {len(asset_ids)} assets to album {album_id}.")
    return len(asset_ids)

# -------------------------------
# Advanced Search Helpers
# -------------------------------
UUID_PATTERN = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")

def parse_additional_filters_json(filters_text):
    if not filters_text or not filters_text.strip():
        return {}, None
    try:
        data = json.loads(filters_text)
    except json.JSONDecodeError as e:
        return {}, f"Invalid JSON for additional filters: {e}"
    if not isinstance(data, dict):
        return {}, "Additional filters must be a JSON object (e.g. {\"isFavorite\": true})."
    return data, None

def parse_people_input(people_text):
    if not people_text:
        return []
    normalized = people_text.replace("\n", ",")
    parts = [p.strip() for p in normalized.split(",")]

    tokens = []
    for part in parts:
        if not part:
            continue

        # Allow convenience formats like:
        # - "<uuid>  # Name"
        # - "<uuid> - Name"
        # by extracting the UUID prefix if present.
        candidate = part.split("#", 1)[0].strip()
        if not candidate:
            continue

        uuid_prefix = re.match(r"^([0-9a-fA-F-]{36})", candidate)
        if uuid_prefix:
            possible_uuid = uuid_prefix.group(1)
            if UUID_PATTERN.match(possible_uuid):
                tokens.append(possible_uuid)
                continue

        tokens.append(candidate)

    return tokens

def resolve_person_ids(headers, people_text, with_hidden=False):
    tokens = parse_people_input(people_text)
    if not tokens:
        return []

    resolved_ids = []
    seen = set()

    for token in tokens:
        if stop_event.is_set():
            return []

        if UUID_PATTERN.match(token):
            person_id = token
            if person_id not in seen:
                seen.add(person_id)
                resolved_ids.append(person_id)
            continue

        matches = search_people_by_name(headers, token, with_hidden=with_hidden)
        if not matches:
            raise ValueError(f"No person found matching '{token}'.")

        token_lower = token.lower()
        exact_matches = [
            p for p in matches
            if str(p.get("name", "")).strip().lower() == token_lower
        ]

        if len(exact_matches) == 1:
            person = exact_matches[0]
        elif len(matches) == 1:
            person = matches[0]
        else:
            match_names = ", ".join(sorted({p.get("name", "") for p in matches if p.get("name")}))
            raise ValueError(
                f"Person name '{token}' is ambiguous. Matches: {match_names}. "
                f"Use a more specific name or paste the person UUID."
            )

        person_id = person.get("id")
        if not person_id:
            raise ValueError(f"Person search result missing id for '{token}'.")

        if person_id not in seen:
            seen.add(person_id)
            resolved_ids.append(person_id)

    return resolved_ids

def search_assets_for_date_range(headers, start_date, end_date, additional_filters=None, person_ids=None, people_match_mode="any"):
    if stop_event.is_set():
        return []

    person_ids = person_ids or []

    if not person_ids:
        return search_assets_by_date_range(headers, start_date, end_date, additional_filters=additional_filters)

    if len(person_ids) == 1:
        filters = dict(additional_filters or {})
        filters["personIds"] = [person_ids[0]]
        return search_assets_by_date_range(headers, start_date, end_date, additional_filters=filters)

    if people_match_mode == "all":
        intersection = None
        for person_id in person_ids:
            if stop_event.is_set():
                return []
            filters = dict(additional_filters or {})
            filters["personIds"] = [person_id]
            ids = set(search_assets_by_date_range(headers, start_date, end_date, additional_filters=filters))
            intersection = ids if intersection is None else (intersection & ids)
        return list(intersection or set())

    # Default: any (OR)
    union = set()
    for person_id in person_ids:
        if stop_event.is_set():
            return []
        filters = dict(additional_filters or {})
        filters["personIds"] = [person_id]
        union.update(search_assets_by_date_range(headers, start_date, end_date, additional_filters=filters))
    return list(union)

# -------------------------------
# Main Logic Function
# -------------------------------
def run_search(
    api_key,
    delta_days,
    start_year,
    end_year,
    selected_items,
    specific_date_str,
    specific_date_album_name,
    specific_date_all_years,
    people_text="",
    people_match_mode="any",
    additional_filters_text="",
    people_with_hidden=False,
):
    stop_event.clear()  # Reset stop event

    if not API_BASE_URL:
        error_msg = f"API base URL is not configured. Set `api_base_url` in {APP_CONFIG_FILE}."
        set_status(error_msg)
        log_message(error_msg)
        return

    headers = {
        "x-api-key": api_key,
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    set_status("Starting asset collection...")
    log_message("Starting asset collection...")
    total_assets_added = 0

    additional_filters, filters_error = parse_additional_filters_json(additional_filters_text)
    if filters_error:
        set_status(f"Error: {filters_error}")
        log_message(filters_error)
        return

    people_match_mode = (people_match_mode or "any").strip().lower()
    if people_match_mode not in {"any", "all"}:
        people_match_mode = "any"

    combined_person_ids = []
    seen_person_ids = set()

    person_ids_from_filters = additional_filters.pop("personIds", None)
    if person_ids_from_filters is not None:
        if not isinstance(person_ids_from_filters, list) or not all(isinstance(p, str) for p in person_ids_from_filters):
            set_status("Error: additional_filters.personIds must be a list of UUID strings.")
            log_message("Invalid additional_filters: personIds must be a list of UUID strings.")
            return
        for person_id in person_ids_from_filters:
            if person_id not in seen_person_ids:
                seen_person_ids.add(person_id)
                combined_person_ids.append(person_id)

    if people_text and people_text.strip():
        try:
            set_status("Resolving people filter...")
            resolved_ids = resolve_person_ids(headers, people_text, with_hidden=people_with_hidden)
        except ValueError as e:
            set_status(f"Error: {str(e)}")
            log_message(str(e))
            return
        for person_id in resolved_ids:
            if person_id not in seen_person_ids:
                seen_person_ids.add(person_id)
                combined_person_ids.append(person_id)

    # Calculate total tasks for progress bar
    total_tasks = 0
    if "Specific Date" in selected_items:
        if specific_date_all_years:
            total_tasks += (end_year - start_year + 1)
        else:
            total_tasks += 1
    for holiday_name in selected_items:
        if holiday_name == "Specific Date":
            continue
        total_tasks += (end_year - start_year + 1)

    set_progress(0, total_tasks)
    current_progress = 0

    # Helper function for date range based on delta_days
    def get_date_range(base_date, delta):
        if delta == 0:
            day_start = datetime(base_date.year, base_date.month, base_date.day)
            return day_start, day_start + timedelta(days=1)
        else:
            return (base_date - timedelta(days=delta), base_date + timedelta(days=delta))

    try:
        # Specific Date
        if "Specific Date" in selected_items:
            specific_date = datetime.strptime(specific_date_str, "%Y-%m-%d")
            album_name = specific_date_album_name
            album_id = find_or_create_album(headers, album_name)

            if specific_date_all_years:
                for year in range(start_year, end_year + 1):
                    if stop_event.is_set():
                        set_status("Operation Cancelled")
                        return
                    day_replaced = specific_date.replace(year=year)
                    start_search_date, end_search_date = get_date_range(day_replaced, delta_days)
                    set_status(f"Searching {album_name} for {year}...")
                    asset_ids = search_assets_for_date_range(
                        headers,
                        start_search_date,
                        end_search_date,
                        additional_filters=additional_filters,
                        person_ids=combined_person_ids,
                        people_match_mode=people_match_mode,
                    )
                    total_assets_added += add_assets_to_album(headers, album_id, asset_ids)
                    current_progress += 1
                    set_progress(current_progress, total_tasks)
            else:
                if stop_event.is_set():
                    set_status("Operation Cancelled")
                    return
                start_search_date, end_search_date = get_date_range(specific_date, delta_days)
                set_status(f"Searching {album_name}...")
                asset_ids = search_assets_for_date_range(
                    headers,
                    start_search_date,
                    end_search_date,
                    additional_filters=additional_filters,
                    person_ids=combined_person_ids,
                    people_match_mode=people_match_mode,
                )
                total_assets_added += add_assets_to_album(headers, album_id, asset_ids)
                current_progress += 1
                set_progress(current_progress, total_tasks)

        # Holidays
        for holiday_name, album_name in selected_items.items():
            if holiday_name == "Specific Date":
                continue
            album_id = find_or_create_album(headers, album_name)
            for year in range(start_year, end_year + 1):
                if stop_event.is_set():
                    set_status("Operation Cancelled")
                    return
                holiday_date = get_holiday_date(year, holiday_name)
                start_search_date, end_search_date = get_date_range(holiday_date, delta_days)
                set_status(f"Searching {holiday_name} for {year}...")
                asset_ids = search_assets_for_date_range(
                    headers,
                    start_search_date,
                    end_search_date,
                    additional_filters=additional_filters,
                    person_ids=combined_person_ids,
                    people_match_mode=people_match_mode,
                )
                total_assets_added += add_assets_to_album(headers, album_id, asset_ids)
                current_progress += 1
                set_progress(current_progress, total_tasks)

        set_status(f"Completed: Added {total_assets_added} assets to {len(selected_items)} albums")
        log_message(f"Completed: Added {total_assets_added} assets to {len(selected_items)} albums")
    except Exception as e:
        set_status(f"Error: {str(e)}")
        log_message(f"Error occurred: {str(e)}")
    finally:
        set_progress(0, 0)  # Reset progress bar

# -------------------------------
# Preset Config Helpers
# -------------------------------
def save_config(config, filename="config.json"):
    """Save current configuration to a JSON file (no secrets)."""
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)

def load_config(filename="config.json"):
    """Load configuration from a JSON file."""
    try:
        with open(filename, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return None, f"{filename} must contain a JSON object."
        return data, None
    except FileNotFoundError:
        return None, f"{filename} not found."
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON in {filename}: {e}"
    except OSError as e:
        return None, f"Failed to read {filename}: {e}"

def show_help():
    """Display a help dialog."""
    help_text = (
        f"{APP_NAME}\n\n"
        "Summary\n"
        "- For each selected holiday, searches every year from Start year → End year and adds matching assets to a per-holiday album.\n"
        "- Delta days turns each holiday/date into a +/- window (e.g. delta=7 searches a 2-week range centered on the date).\n"
        "- Specific date (useful for birthdays/anniversaries) can run once, or repeat that month/day across all years when “All years” is enabled.\n"
        "- Creates albums if needed and adds any found assets.\n"
        "- Advanced filters: People (OR/AND) and extra `/search/metadata` JSON filters; presets can be saved/loaded via `config.json`.\n\n"
        "Connection\n"
        "- API endpoint: loaded from `app_config.json` (`api_base_url`). Use a full `/api` URL (e.g. "
        "`https://immich.example.com/api`). If you provide only the host, the app assumes `/api`.\n"
        "- API key: required for all actions. To use People filters, the key must have `person.read`.\n\n"
        "Search Options\n"
        "- Delta days: expands each target date into a range. Example: delta=7 searches from 7 days before "
        "to 7 days after the date. Delta=0 searches only that day.\n"
        "- Start year / End year: repeats the search for each selected holiday across every year in the range.\n\n"
        "Holidays Tab\n"
        "- Each selected holiday creates/finds an album (name is editable) and adds matching assets.\n"
        "- Specific date (e.g. a birthday): searches the chosen date once, or (if “All years” is checked) repeats that month/day "
        "across the year range.\n\n"
        "Advanced Tab\n"
        "People filter\n"
        "- Paste names/UUIDs, or click “Browse people…” to pick from a searchable list.\n"
        "- Match any (OR): assets containing any selected person.\n"
        "- Match all (AND): assets containing all selected people (slower; multiple searches).\n"
        "People picker search\n"
        "- OR: `jack, jill` or `jack or jill`\n"
        "- AND: `jack jill` or `jack and jill`\n\n"
        "Additional metadata filters (JSON)\n"
        "- JSON object merged into the `/search/metadata` request.\n"
        "- Example: {\"isFavorite\": true, \"city\": \"Boston\"}\n"
        "- Avoid `takenAfter`, `takenBefore`, `page`, `size` (the app controls these). "
        "`personIds` is combined with the People filter.\n\n"
        "Presets\n"
        "- Save preset / Load preset reads and writes `config.json` (no API key).\n\n"
        "Cancel\n"
        "- Cancels the current run; in-flight requests may take a moment to stop.\n"
    )

    if "root" in globals() and isinstance(globals().get("root"), tk.Tk) and root.winfo_exists():
        win = tk.Toplevel(root)
        win.title("Help")
        win.geometry("760x560")
        win.transient(root)

        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=1)

        frame = ttk.Frame(win, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        text = tk.Text(frame, wrap="word")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        text.insert("1.0", help_text)
        text.configure(state="disabled")

        buttons = ttk.Frame(frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))

        def copy_help():
            root.clipboard_clear()
            root.clipboard_append(help_text)

        ttk.Button(buttons, text="Copy", command=copy_help).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(buttons, text="Close", command=win.destroy).grid(row=0, column=1)
        return

    messagebox.showinfo("Help", help_text)

# -------------------------------
# GUI Creation
# -------------------------------
def create_gui():
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)

    global root
    root = tk.Tk()
    root.title(APP_NAME)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("Title.TLabel", font=("TkDefaultFont", 14, "bold"))
    style.configure("Error.TLabel", foreground="#b00020")

    api_base_url, config_error = set_api_base_url_from_config()

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    main = ttk.Frame(root, padding=12)
    main.grid(row=0, column=0, sticky="nsew")
    main.columnconfigure(0, weight=1)
    main.rowconfigure(3, weight=1)

    header = ttk.Frame(main)
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(0, weight=1)
    ttk.Label(header, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Button(header, text="Help", command=show_help).grid(row=0, column=1, sticky="e")

    # -------------------------------
    # Connection
    # -------------------------------
    connection_frame = ttk.Labelframe(main, text="Connection", padding=10)
    connection_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
    connection_frame.columnconfigure(1, weight=1)

    api_url_var = tk.StringVar(value=api_base_url or "")
    api_error_var = tk.StringVar(
        value=config_error or ("" if api_base_url else f"Set `api_base_url` in {APP_CONFIG_FILE}.")
    )

    def reload_app_config():
        api_url, err = set_api_base_url_from_config()
        api_url_var.set(api_url or "")
        api_error_var.set(err or ("" if api_url else f"Set `api_base_url` in {APP_CONFIG_FILE}."))

    ttk.Label(connection_frame, text="API endpoint:").grid(row=0, column=0, sticky="w")
    ttk.Label(connection_frame, textvariable=api_url_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
    ttk.Button(connection_frame, text="Reload config", command=reload_app_config).grid(row=0, column=2, sticky="e")
    ttk.Label(connection_frame, textvariable=api_error_var, style="Error.TLabel").grid(row=1, column=1, sticky="w", padx=(8, 0))

    api_key_var = tk.StringVar()
    stored_api_key = get_stored_api_key()
    if stored_api_key:
        api_key_var.set(stored_api_key)

    show_key_var = tk.BooleanVar(value=False)

    api_key_entry = ttk.Entry(connection_frame, textvariable=api_key_var, width=50, show="•")
    api_key_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
    ttk.Label(connection_frame, text="API key:").grid(row=2, column=0, sticky="w", pady=(8, 0))

    def toggle_key_visibility():
        api_key_entry.configure(show="" if show_key_var.get() else "•")

    def store_api_key():
        if not API_BASE_URL:
            messagebox.showerror(
                "Error",
                f"API base URL is not configured.\n\n"
                f"Create `{APP_CONFIG_FILE}` and set `api_base_url` to your Immich `/api` endpoint "
                f"(see `app_config.example.json`)."
            )
            return

        api_key = api_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "API Key is required to store.")
            return
        try:
            headers = {"x-api-key": api_key}
            r = requests.get(f"{API_BASE_URL}/albums", headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
            r.raise_for_status()
            store_api_key_in_keyring(api_key)
            messagebox.showinfo("Success", "API Key stored successfully.")
            set_status("Stored API key in keyring")
        except requests.RequestException as e:
            messagebox.showerror("Error", f"Failed to validate API Key: {str(e)}")

    def delete_stored_api_key():
        try:
            delete_api_key_from_keyring()
            api_key_var.set("")
            messagebox.showinfo("Success", "Stored API Key deleted.")
            set_status("Stored API key deleted")
        except keyring.errors.PasswordDeleteError:
            messagebox.showinfo("Info", "No stored API Key to delete.")

    key_actions = ttk.Frame(connection_frame)
    key_actions.grid(row=2, column=2, sticky="e", pady=(8, 0))
    ttk.Checkbutton(key_actions, text="Show", variable=show_key_var, command=toggle_key_visibility).grid(row=0, column=0, padx=(0, 10))
    ttk.Button(key_actions, text="Store", command=store_api_key).grid(row=0, column=1, padx=(0, 6))
    ttk.Button(key_actions, text="Delete", command=delete_stored_api_key).grid(row=0, column=2)

    # -------------------------------
    # Search Options
    # -------------------------------
    options_frame = ttk.Labelframe(main, text="Search Options", padding=10)
    options_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))

    delta_var = tk.StringVar(value="7")
    start_year_var = tk.StringVar(value="2007")
    end_year_var = tk.StringVar(value=str(datetime.now().year))

    ttk.Label(options_frame, text="Delta days:").grid(row=0, column=0, sticky="w")
    ttk.Spinbox(options_frame, from_=0, to=365, width=6, textvariable=delta_var).grid(row=0, column=1, sticky="w", padx=(6, 18))
    ttk.Label(options_frame, text="Start year:").grid(row=0, column=2, sticky="w")
    ttk.Spinbox(options_frame, from_=1900, to=2100, width=8, textvariable=start_year_var).grid(row=0, column=3, sticky="w", padx=(6, 18))
    ttk.Label(options_frame, text="End year:").grid(row=0, column=4, sticky="w")
    ttk.Spinbox(options_frame, from_=1900, to=2100, width=8, textvariable=end_year_var).grid(row=0, column=5, sticky="w")

    # -------------------------------
    # Tabs
    # -------------------------------
    notebook = ttk.Notebook(main)
    notebook.grid(row=3, column=0, sticky="nsew", pady=(10, 0))

    holidays_tab = ttk.Frame(notebook, padding=10)
    advanced_tab = ttk.Frame(notebook, padding=10)
    log_tab = ttk.Frame(notebook, padding=10)
    notebook.add(holidays_tab, text="Holidays")
    notebook.add(advanced_tab, text="Advanced")
    notebook.add(log_tab, text="Log")

    # Holidays tab
    holidays_tab.columnconfigure(1, weight=1)

    ttk.Label(holidays_tab, text="Select holidays and optionally rename the target album.").grid(row=0, column=0, columnspan=3, sticky="w")

    header_row = ttk.Frame(holidays_tab)
    header_row.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 4))
    header_row.columnconfigure(1, weight=1)
    ttk.Label(header_row, text="Use").grid(row=0, column=0, sticky="w", padx=(0, 10))
    ttk.Label(header_row, text="Holiday / Specific Date").grid(row=0, column=1, sticky="w")
    ttk.Label(header_row, text="Album name").grid(row=0, column=2, sticky="w", padx=(10, 0))

    specific_date_enabled = tk.BooleanVar(value=False)
    specific_date_all_years_var = tk.BooleanVar(value=False)
    specific_date_album_var = tk.StringVar(value="Specific Date Search")

    ttk.Checkbutton(holidays_tab, variable=specific_date_enabled).grid(row=2, column=0, sticky="w", pady=(0, 4))

    sd_frame = ttk.Frame(holidays_tab)
    sd_frame.grid(row=2, column=1, sticky="ew", pady=(0, 4))
    ttk.Label(sd_frame, text="Specific date:").grid(row=0, column=0, sticky="w")
    specific_date_entry = DateEntry(sd_frame, width=12, date_pattern="yyyy-mm-dd")
    specific_date_entry.grid(row=0, column=1, sticky="w", padx=(6, 10))
    ttk.Checkbutton(sd_frame, text="All years", variable=specific_date_all_years_var).grid(row=0, column=2, sticky="w")

    ttk.Entry(holidays_tab, textvariable=specific_date_album_var, width=30).grid(row=2, column=2, sticky="w", pady=(0, 4), padx=(10, 0))

    holiday_vars = {}
    holiday_album_vars = {}

    for i, holiday in enumerate(DEFAULT_HOLIDAYS, start=3):
        var = tk.BooleanVar(value=False)
        album_var = tk.StringVar(value=holiday)
        holiday_vars[holiday] = var
        holiday_album_vars[holiday] = album_var

        ttk.Checkbutton(holidays_tab, variable=var).grid(row=i, column=0, sticky="w", pady=2)
        ttk.Label(holidays_tab, text=holiday).grid(row=i, column=1, sticky="w", pady=2)
        ttk.Entry(holidays_tab, textvariable=album_var, width=30).grid(row=i, column=2, sticky="w", pady=2, padx=(10, 0))

    def set_all_holidays(enabled):
        for h in DEFAULT_HOLIDAYS:
            holiday_vars[h].set(enabled)

    holiday_actions = ttk.Frame(holidays_tab)
    holiday_actions.grid(row=len(DEFAULT_HOLIDAYS) + 3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    ttk.Button(holiday_actions, text="Select all", command=lambda: set_all_holidays(True)).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(holiday_actions, text="Clear all", command=lambda: set_all_holidays(False)).grid(row=0, column=1)

    # Advanced tab
    advanced_tab.columnconfigure(0, weight=1)

    people_frame = ttk.Labelframe(advanced_tab, text="People filter", padding=10)
    people_frame.grid(row=0, column=0, sticky="ew")
    people_frame.columnconfigure(0, weight=1)

    ttk.Label(people_frame, text="Enter names/UUIDs, or pick from a searchable list.").grid(row=0, column=0, sticky="w")
    people_text = tk.Text(people_frame, height=4, width=60)
    people_text.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    people_match_var = tk.StringVar(value="any")
    match_row = ttk.Frame(people_frame)
    ttk.Radiobutton(match_row, text="Match any (OR)", variable=people_match_var, value="any").grid(row=0, column=0, padx=(0, 10))
    ttk.Radiobutton(match_row, text="Match all (AND)", variable=people_match_var, value="all").grid(row=0, column=1)

    people_with_hidden_var = tk.BooleanVar(value=False)

    def open_people_picker():
        if not API_BASE_URL:
            messagebox.showerror(
                "Error",
                f"API base URL is not configured.\n\n"
                f"Create `{APP_CONFIG_FILE}` and set `api_base_url` to your Immich `/api` endpoint "
                f"(see `app_config.example.json`)."
            )
            return

        api_key = api_key_var.get().strip()
        if not api_key:
            api_key = get_stored_api_key() or ""
        if not api_key:
            messagebox.showerror("Error", "API Key is required.")
            return

        headers = {"x-api-key": api_key, "Accept": "application/json"}

        picker = tk.Toplevel(root)
        picker.title("Select People")
        picker.geometry("720x520")
        picker.transient(root)
        picker.grab_set()

        picker.columnconfigure(0, weight=1)
        picker.rowconfigure(1, weight=1)

        top = ttk.Frame(picker, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        filter_var = tk.StringVar()
        include_hidden_var = tk.BooleanVar(value=bool(people_with_hidden_var.get()))
        status_var = tk.StringVar(value="Loading people…")

        ttk.Label(top, text="Search (OR: comma/or, AND: space/and):").grid(row=0, column=0, sticky="w")
        filter_entry = ttk.Entry(top, textvariable=filter_var)
        filter_entry.grid(row=0, column=1, sticky="ew", padx=(6, 10))
        include_hidden_check = ttk.Checkbutton(top, text="Include hidden", variable=include_hidden_var)
        include_hidden_check.grid(row=0, column=2, sticky="e")
        ttk.Label(top, textvariable=status_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        body = ttk.Frame(picker, padding=(10, 0, 10, 10))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        listbox = tk.Listbox(body, selectmode=tk.EXTENDED)
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        footer = ttk.Frame(picker, padding=10)
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        all_people = []
        filtered_people = []

        def parse_people_search_expression(text):
            expr = (text or "").strip()
            if not expr:
                return []

            or_parts = [
                p.strip()
                for p in re.split(r"\s*(?:,|;|\|\|?|\bor\b)\s*", expr, flags=re.IGNORECASE)
                if p.strip()
            ]

            clauses = []
            for part in or_parts:
                part = re.sub(r"\band\b", " ", part, flags=re.IGNORECASE)
                part = part.replace("&&", " ").replace("&", " ")
                terms = [
                    t.strip().strip("\"'").casefold()
                    for t in re.split(r"\s+", part)
                    if t.strip()
                ]
                if terms:
                    clauses.append(terms)

            return clauses

        def apply_filter(*_):
            clauses = parse_people_search_expression(filter_var.get())
            listbox.delete(0, tk.END)
            filtered_people.clear()

            for person in all_people:
                person_id = str(person.get("id", "")).strip()
                if not person_id:
                    continue

                name = str(person.get("name", "")).strip()
                if not name:
                    continue

                name_cf = name.casefold()
                if clauses and not any(all(term in name_cf for term in clause) for clause in clauses):
                    continue

                is_hidden = bool(person.get("isHidden", False))
                label = f"{name} [hidden]" if is_hidden else name
                short_id = person_id.split("-", 1)[0]
                listbox.insert(tk.END, f"{label}  —  {short_id}")
                filtered_people.append(person)

            if all_people:
                status_var.set(f"Showing {len(filtered_people)} of {len(all_people)} people")

        def load_people():
            status_var.set("Loading people…")
            listbox.delete(0, tk.END)
            filtered_people.clear()

            def worker():
                try:
                    people = get_all_people(headers, with_hidden=bool(include_hidden_var.get()))
                except Exception as e:
                    error_text = str(e).strip() or repr(e)
                    root.after(0, lambda msg=error_text: status_var.set(f"Error loading people: {msg}"))
                    return

                people_sorted = sorted(
                    [p for p in people if str(p.get("name", "")).strip()],
                    key=lambda p: (str(p.get("name", "")).casefold(), str(p.get("id", ""))),
                )

                def done():
                    all_people.clear()
                    all_people.extend(people_sorted)
                    apply_filter()
                    if not all_people:
                        status_var.set("No people found.")

                root.after(0, done)

            threading.Thread(target=worker, daemon=True).start()

        def add_selected(close_after=False):
            indices = listbox.curselection()
            if not indices:
                return

            existing_tokens = parse_people_input(people_text.get("1.0", tk.END))
            existing_ids = {t for t in existing_tokens if UUID_PATTERN.match(t)}

            lines_to_add = []
            for idx in indices:
                person = filtered_people[int(idx)]
                person_id = str(person.get("id", "")).strip()
                if not person_id or person_id in existing_ids:
                    continue

                name = str(person.get("name", "")).strip() or "(Unnamed)"
                lines_to_add.append(f"{person_id}  # {name}")
                existing_ids.add(person_id)

            if lines_to_add:
                current = people_text.get("1.0", tk.END).rstrip()
                if current and not current.endswith("\n"):
                    people_text.insert(tk.END, "\n")
                people_text.insert(tk.END, "\n".join(lines_to_add) + "\n")

            if close_after:
                picker.destroy()

        ttk.Button(footer, text="Reload", command=load_people).grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Add selected", command=lambda: add_selected(False)).grid(row=0, column=1, sticky="e", padx=(6, 0))
        ttk.Button(footer, text="Add & close", command=lambda: add_selected(True)).grid(row=0, column=2, sticky="e", padx=(6, 0))
        ttk.Button(footer, text="Close", command=picker.destroy).grid(row=0, column=3, sticky="e", padx=(6, 0))

        filter_var.trace_add("write", apply_filter)
        include_hidden_check.configure(command=load_people)
        listbox.bind("<Double-Button-1>", lambda _e: add_selected(True))
        picker.after(0, load_people)
        filter_entry.focus_set()

    people_actions = ttk.Frame(people_frame)
    people_actions.grid(row=2, column=0, sticky="w", pady=(6, 8))
    ttk.Button(people_actions, text="Browse people…", command=open_people_picker).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(people_actions, text="Clear", command=lambda: people_text.delete("1.0", tk.END)).grid(row=0, column=1)

    match_row.grid(row=3, column=0, sticky="w")

    ttk.Checkbutton(
        people_frame,
        text="Include hidden people when resolving names",
        variable=people_with_hidden_var,
    ).grid(row=4, column=0, sticky="w", pady=(6, 0))

    filters_frame = ttk.Labelframe(advanced_tab, text="Additional metadata filters (JSON)", padding=10)
    filters_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
    filters_frame.columnconfigure(0, weight=1)

    ttk.Label(filters_frame, text='Example: {"isFavorite": true, "city": "Boston"}').grid(row=0, column=0, sticky="w")
    additional_filters_text = tk.Text(filters_frame, height=6, width=60)
    additional_filters_text.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

    def validate_filters():
        _, err = parse_additional_filters_json(additional_filters_text.get("1.0", tk.END))
        if err:
            messagebox.showerror("Invalid JSON", err)
        else:
            messagebox.showinfo("OK", "Additional filters JSON is valid.")

    ttk.Button(filters_frame, text="Validate JSON", command=validate_filters).grid(row=2, column=0, sticky="w", pady=(8, 0))

    # Log tab
    log_tab.columnconfigure(0, weight=1)
    log_tab.rowconfigure(0, weight=1)

    log_viewer = tk.Text(log_tab, height=12, width=80)
    log_scrollbar = ttk.Scrollbar(log_tab, command=log_viewer.yview)
    log_viewer.configure(yscrollcommand=log_scrollbar.set)
    log_viewer.grid(row=0, column=0, sticky="nsew")
    log_scrollbar.grid(row=0, column=1, sticky="ns")

    def clear_log():
        log_viewer.delete("1.0", tk.END)

    ttk.Button(log_tab, text="Clear log view", command=clear_log).grid(row=1, column=0, sticky="w", pady=(8, 0))

    # -------------------------------
    # Actions + Status
    # -------------------------------
    action_bar = ttk.Frame(main)
    action_bar.grid(row=4, column=0, sticky="ew", pady=(10, 0))
    action_bar.columnconfigure(1, weight=1)

    progress_bar = ttk.Progressbar(action_bar, orient="horizontal", mode="determinate")
    progress_bar.grid(row=0, column=1, sticky="ew", padx=10)

    current_thread = None

    def run_search_in_background():
        nonlocal current_thread

        if not API_BASE_URL:
            messagebox.showerror(
                "Error",
                f"API base URL is not configured.\n\n"
                f"Create `{APP_CONFIG_FILE}` and set `api_base_url` to your Immich `/api` endpoint "
                f"(see `app_config.example.json`)."
            )
            return

        api_key = api_key_var.get().strip()
        if not api_key:
            api_key = get_stored_api_key() or ""
        if not api_key:
            messagebox.showerror("Error", "API Key is required.")
            return

        try:
            delta_days = int(delta_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Delta days must be an integer.")
            return

        try:
            start_year = int(start_year_var.get().strip())
            end_year = int(end_year_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Start year and end year must be integers.")
            return

        if start_year > end_year:
            messagebox.showerror("Error", "Start year must be <= end year.")
            return

        selected_items = {}
        specific_date_str = ""

        if specific_date_enabled.get():
            specific_date_str = specific_date_entry.get_date().strftime("%Y-%m-%d")
            album_name = specific_date_album_var.get().strip() or "Specific Date Search"
            selected_items["Specific Date"] = album_name

        for holiday in DEFAULT_HOLIDAYS:
            if holiday_vars[holiday].get():
                album_name = holiday_album_vars[holiday].get().strip() or holiday
                selected_items[holiday] = album_name

        if not selected_items:
            messagebox.showerror("Error", "Select at least one holiday or a specific date.")
            return

        people_query = people_text.get("1.0", tk.END).strip()
        people_match_mode = people_match_var.get()
        filters_text = additional_filters_text.get("1.0", tk.END).strip()
        with_hidden_people = bool(people_with_hidden_var.get())
        _, filters_error = parse_additional_filters_json(filters_text)
        if filters_error:
            messagebox.showerror("Invalid additional filters", filters_error)
            return

        run_button.configure(state="disabled")
        cancel_button.configure(state="normal")

        current_thread = threading.Thread(
            target=run_search,
            args=(
                api_key,
                delta_days,
                start_year,
                end_year,
                selected_items,
                specific_date_str,
                selected_items.get("Specific Date"),
                bool(specific_date_all_years_var.get()),
                people_query,
                people_match_mode,
                filters_text,
                with_hidden_people,
            ),
            daemon=True,
        )
        current_thread.start()

        def check_thread():
            if current_thread and not current_thread.is_alive():
                run_button.configure(state="normal")
                cancel_button.configure(state="disabled")
            else:
                root.after(150, check_thread)

        root.after(150, check_thread)

    def cancel_run():
        stop_event.set()
        set_status("Cancelling...")
        cancel_button.configure(state="disabled")

    def save_preset():
        selected_items = {}
        if specific_date_enabled.get():
            selected_items["Specific Date"] = specific_date_album_var.get().strip() or "Specific Date Search"
        for holiday in DEFAULT_HOLIDAYS:
            if holiday_vars[holiday].get():
                selected_items[holiday] = holiday_album_vars[holiday].get().strip() or holiday

        preset = {
            "delta_days": delta_var.get().strip(),
            "start_year": start_year_var.get().strip(),
            "end_year": end_year_var.get().strip(),
            "selected_items": selected_items,
            "specific_date": specific_date_entry.get_date().strftime("%Y-%m-%d") if specific_date_enabled.get() else "",
            "specific_date_all_years": 1 if specific_date_all_years_var.get() else 0,
            "specific_date_album_name": specific_date_album_var.get().strip() or "Specific Date Search",
            "people": people_text.get("1.0", tk.END).strip(),
            "people_match_mode": people_match_var.get(),
            "people_with_hidden": 1 if people_with_hidden_var.get() else 0,
            "additional_filters": additional_filters_text.get("1.0", tk.END).strip(),
        }
        save_config(preset)
        set_status("Saved preset to config.json")

    def load_preset():
        preset, err = load_config()
        if err:
            messagebox.showerror("Load preset failed", err)
            return
        if not preset:
            return

        delta_var.set(str(preset.get("delta_days", "7")))
        start_year_var.set(str(preset.get("start_year", "2007")))
        end_year_var.set(str(preset.get("end_year", str(datetime.now().year))))

        selected = preset.get("selected_items", {}) or {}
        specific_date_enabled.set("Specific Date" in selected or bool(preset.get("specific_date")))
        if preset.get("specific_date"):
            try:
                specific_date_entry.set_date(datetime.strptime(preset.get("specific_date"), "%Y-%m-%d"))
            except ValueError:
                pass
        specific_date_all_years_var.set(bool(preset.get("specific_date_all_years", 0)))
        specific_date_album_var.set(str(preset.get("specific_date_album_name", selected.get("Specific Date", "Specific Date Search"))))

        for holiday in DEFAULT_HOLIDAYS:
            holiday_vars[holiday].set(holiday in selected)
            holiday_album_vars[holiday].set(str(selected.get(holiday, holiday)))

        people_text.delete("1.0", tk.END)
        people_text.insert(tk.END, str(preset.get("people", "")).strip())
        people_match_var.set(str(preset.get("people_match_mode", "any")))
        people_with_hidden_var.set(bool(preset.get("people_with_hidden", 0)))

        additional_filters_text.delete("1.0", tk.END)
        additional_filters_text.insert(tk.END, str(preset.get("additional_filters", "")).strip())

        set_status("Loaded preset from config.json")

    preset_buttons = ttk.Frame(action_bar)
    preset_buttons.grid(row=0, column=0, sticky="w")
    ttk.Button(preset_buttons, text="Save preset", command=save_preset).grid(row=0, column=0, padx=(0, 6))
    ttk.Button(preset_buttons, text="Load preset", command=load_preset).grid(row=0, column=1)

    run_controls = ttk.Frame(action_bar)
    run_controls.grid(row=0, column=2, sticky="e")
    run_button = ttk.Button(run_controls, text="Run", command=run_search_in_background)
    run_button.grid(row=0, column=0, padx=(0, 6))
    cancel_button = ttk.Button(run_controls, text="Cancel", command=cancel_run, state="disabled")
    cancel_button.grid(row=0, column=1)

    status_var = tk.StringVar(value="Ready" if api_base_url else f"Set API base URL in {APP_CONFIG_FILE}")
    status_label = ttk.Label(main, textvariable=status_var)
    status_label.grid(row=5, column=0, sticky="ew", pady=(10, 0))

    # Progress Queue and Update Function
    def update_progress():
        try:
            while True:
                msg = progress_queue.get_nowait()
                if msg["type"] == "status":
                    status_var.set(msg["text"])
                elif msg["type"] == "progress":
                    progress_bar["value"] = msg["value"]
                    progress_bar["maximum"] = msg["max"]
                elif msg["type"] == "log":
                    log_viewer.insert(tk.END, msg["text"] + "\n")
                    log_viewer.see(tk.END)
        except queue.Empty:
            pass
        root.after(100, update_progress)

    update_progress()

    return root

# -------------------------------
# Validation Functions
# -------------------------------
def validate_year(value):
    """Validate that the year is between 1900 and 2100."""
    if value == "":
        return True
    try:
        year = int(value)
        return 1900 <= year <= 2100
    except ValueError:
        return False

def validate_delta(value):
    """Validate that delta days is a non-negative integer."""
    if value == "":
        return True
    try:
        delta = int(value)
        return delta >= 0
    except ValueError:
        return False

# -------------------------------
# Tooltip Class
# -------------------------------
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        x, y, _, _ = self.widget.bbox("insert") if self.widget.winfo_exists() else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tw, text=self.text, background="yellow", relief="solid", borderwidth=1)
        label.pack()

    def hide(self, event=None):
        if hasattr(self, 'tw') and self.tw.winfo_exists():
            self.tw.destroy()

if __name__ == "__main__":
    app = create_gui()
    app.mainloop()
