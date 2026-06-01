import os
import re
import requests
import csv
import datetime
from datetime import timedelta
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed

# Load environment variables
load_dotenv()

API_KEY = os.getenv("SAM_API_KEY")
if not API_KEY:
    raise ValueError("Missing SAM_API_KEY. Please set it in your .env file.")

PIPEDRIVE_TOKEN = os.getenv("PIPEDRIVE_API_TOKEN")

# URL for SAM.gov Opportunities Search 
BASE_URL = "https://api.sam.gov/prod/opportunities/v2/search"
PIPEDRIVE_BASE = "https://api.pipedrive.com/v1"

# SYTE Corp's Pipedrive pipeline: "RFP/Solicitation" (ID=12)
# Starting stage: "Potential Opp" (ID=69)
PIPEDRIVE_PIPELINE_ID = 12
PIPEDRIVE_STAGE_ID = 69

NAICS_CODES = ["237110", "237120", "237310", "237990",
               "213112", "237130", "238210", "561210"]

# Target states for place of performance (2-letter codes).
# Opportunities with no location info ("Not Specified") are kept by default.
TARGET_STATES = {
    "AL", "AR", "CA", "FL", "GA", "IL", "IN", "KS", "KY", "LA",
    "MI", "MN", "MO", "MS", "NC", "ND", "NY", "OH", "PA", "SC",
    "SD", "TN", "TX", "WI",
}

# Core keywords related to SYTE Corp's fields
SYTE_KEYWORDS = [
    "infrastructure", "construction", "utilities", "methane", 
    "water", "sewer", "pipeline", "highway", "bridge", "civil", "gas"
]

def extract_rom_from_description(notice_id):
    """Fetch the full description text for one opportunity and extract dollar amounts."""
    if not notice_id:
        return ""
    try:
        url = f"https://api.sam.gov/prod/opportunities/v1/noticedesc?noticeid={notice_id}"
        r = requests.get(url, params={"api_key": API_KEY}, timeout=10)
        if r.status_code != 200:
            return ""
        text = r.text  # JSON-wrapped HTML description

        # Pattern 1: Standard dollar amounts like $500,000 or $1.2 million / $45M
        p1 = r'\$[\d,]+(?:[.][\d]+)?(?:\s*(?:million|billion|M|B|K|thousand))?'
        # Pattern 2: Range like "between $250,000 and $500,000" or "$100K - $500K"
        p2 = r'\$[\d,]+(?:[.][\d]+)?(?:\s*(?:million|billion|M|B|K|thousand))?\s*(?:to|-|and|–)\s*\$[\d,]+(?:[.][\d]+)?(?:\s*(?:million|billion|M|B|K|thousand))?'
        # Pattern 3: "estimated value/cost is $X"
        p3 = r'[Ee]stimated\s+(?:value|cost|amount|price).{0,40}\$[\d,]+(?:[.][\d]+)?(?:\s*(?:million|billion|M|B|K))?'
        # Pattern 4: "approximately $X"
        p4 = r'[Aa]pproximately\s+\$[\d,]+(?:[.][\d]+)?(?:\s*(?:million|billion|M|B|K))?'

        found = []
        for pattern in [p2, p3, p4, p1]:  # Check ranges & context first, then simple
            matches = re.findall(pattern, text, re.IGNORECASE)
            found.extend([m.strip() for m in matches if m.strip()])

        if found:
            seen = set()
            unique = []
            for f in found:
                key = f.lower()
                if key not in seen:
                    seen.add(key)
                    unique.append(f)
            return " | ".join(unique[:3])
    except Exception:
        pass
    return ""
def parse_rom_midpoint(rom_text):
    """Parse a ROM text string into a single numeric midpoint.
    e.g. '$500,000 and $1,000,000' → 750000
         '$45 million'             → 45000000
         '$400M | $200M'           → 300000 (avg of all found)
    Returns a float or None.
    """
    if not rom_text:
        return None

    suffix_map = {
        'thousand': 1_000, 'k': 1_000,
        'million':  1_000_000, 'm': 1_000_000,
        'billion':  1_000_000_000, 'b': 1_000_000_000,
    }

    # Extract all individual dollar amounts
    token_pattern = r'\$([\d,]+(?:[.][\d]+)?)\s*(million|billion|thousand|M|B|K)?'
    amounts = []
    for raw_num, suffix in re.findall(token_pattern, rom_text, re.IGNORECASE):
        try:
            value = float(raw_num.replace(',', ''))
            if suffix:
                value *= suffix_map.get(suffix.lower(), 1)
            amounts.append(value)
        except ValueError:
            pass

    if not amounts:
        return None
    # Return midpoint (average) of all extracted amounts
    return sum(amounts) / len(amounts)


def calculate_relevance(title, department):
    score = 0
    text = f"{title} {department}".lower()
    for kw in SYTE_KEYWORDS:
        if kw in text:
            score += 1
    return score

LAST_RUN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_run.txt")

def read_last_run_date():
    """Return the last run date as MM/DD/YYYY, or 90 days ago if file doesn't exist."""
    if os.path.exists(LAST_RUN_FILE):
        with open(LAST_RUN_FILE) as f:
            date_str = f.read().strip()
        try:
            dt = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            print(f"  [Incremental] Last run: {date_str}. Fetching only new postings since then.")
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            pass
    fallback = (datetime.datetime.now() - timedelta(days=7)).strftime("%m/%d/%Y")
    print(f"  [Incremental] No last_run.txt found. Fetching last 7 days (from {fallback}).")
    return fallback

def write_last_run_date():
    """Write today's date to last_run.txt."""
    with open(LAST_RUN_FILE, "w") as f:
        f.write(datetime.datetime.now().strftime("%Y-%m-%d"))

def get_opportunities(limit=50, posted_from=None, posted_to=None):
    if posted_to is None:
        posted_to = datetime.datetime.now().strftime("%m/%d/%Y")
    if posted_from is None:
        posted_from = read_last_run_date()
        
    all_opportunities = []
    
    headers = {
        "x-api-key": API_KEY,
        "accept": "application/json"
    }
    
    for ncode in NAICS_CODES:
        print(f"Fetching opportunities for NAICS Code: {ncode} (From {posted_from} to {posted_to})...")
        ncode_total = 0
        offset = 0
        try:
            while True:
                params = {
                    "api_key": API_KEY,
                    "ncode": ncode,
                    "postedFrom": posted_from,
                    "postedTo": posted_to,
                    "ptype": "o,p,s",
                    "limit": limit,
                    "offset": offset,
                }
                response = requests.get(BASE_URL, params=params, headers=headers, timeout=30)
                response.raise_for_status()
                data = response.json()

                opportunities = data.get("opportunitiesData", [])
                for opp in opportunities:
                    opp["Queried_NAICS"] = ncode
                    all_opportunities.append(opp)
                ncode_total += len(opportunities)

                # Stop when fewer than `limit` returned (last page)
                if len(opportunities) < limit:
                    break
                offset += limit
            print(f"  -> Found {ncode_total} entries for {ncode}.")
        except Exception as e:
            print(f"  -> Error for NAICS {ncode}: {e}")
            
    return all_opportunities

def filter_and_sort_opportunities(opportunities):
    print("\nFiltering for SDVOSB & Small Business, and ranking by relevance...")
    filtered = []
    
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    max_deadline = today + timedelta(days=45)
    for opp in opportunities:
        set_aside_desc = str(opp.get("typeOfSetAsideDescription", "")).lower()
        set_aside_code = str(opp.get("typeOfSetAside", "")).lower()

        # Keep only Small Business or SDVOSB
        is_small_biz = "small business" in set_aside_desc or "sba" in set_aside_code
        is_sdvosb = "service-disabled veteran" in set_aside_desc or "sdvosb" in set_aside_desc or "sdvosb" in set_aside_code or "veteran" in set_aside_desc

        if is_small_biz or is_sdvosb:
            # 1. Deadline Check: for Solicitations, keep only those due within next 45 days
            deadline_str = opp.get("responseDeadLine") or opp.get("responseDate")
            parsed_date = None
            if deadline_str:
                try:
                    parsed_date = datetime.datetime.strptime(deadline_str[:10], "%Y-%m-%d")
                except:
                    pass

            notice_type = str(opp.get("type", "")).lower()
            if "solicitation" in notice_type and "presolicitation" not in notice_type:
                if parsed_date is None:
                    continue  # No deadline info, skip
                if parsed_date < today or parsed_date > max_deadline:
                    continue  # Already expired or beyond 45-day window
            
            # 2. Extract Location
            place = opp.get("placeOfPerformance") or {}
            city = place.get("city", {}).get("name", "") if isinstance(place.get("city"), dict) else ""
            state_name = place.get("state", {}).get("name", "") if isinstance(place.get("state"), dict) else ""
            state_code = place.get("state", {}).get("code", "") if isinstance(place.get("state"), dict) else ""

            # Filter by target states — keep if state matches OR location is unspecified
            if state_code and state_code.upper() not in TARGET_STATES:
                continue  # Outside SYTE's target geography

            if city and state_name:
                opp["location"] = f"{city}, {state_name}"
            elif city or state_name:
                opp["location"] = city or state_name
            else:
                opp["location"] = "Not Specified"
                
            # 3. site_visit + cleanup (ROM filled in concurrently below)
            opp["site_visit_date"] = opp.get("siteVisitDate", "")
            opp.pop("description", None) # Remove useless link

            # Calculate Relevance Score
            opp["Relevance_Score"] = calculate_relevance(opp.get("title", ""), opp.get("department", ""))
            filtered.append(opp)

    # Remove duplicates based on noticeId in case multiple NAICS returned the same opp
    unique_opps = {opp["noticeId"]: opp for opp in filtered}.values()
    filtered = list(unique_opps)

    # 3b. Concurrent ROM extraction (10 workers — 5–10x speedup vs. serial)
    if filtered:
        print(f"  Extracting ROM from {len(filtered)} notice descriptions (concurrent)...")
        with ThreadPoolExecutor(max_workers=10) as ex:
            future_map = {
                ex.submit(extract_rom_from_description, opp.get("noticeId", "")): opp
                for opp in filtered
            }
            for future in as_completed(future_map):
                opp = future_map[future]
                try:
                    opp["ROM"] = future.result()
                except Exception:
                    opp["ROM"] = ""

    # Sort
    # 1. Notice Type (group Solicitations, Presolicitations, etc.)
    # 2. Relevance Score (descending: highest first)
    # 3. Due Date (ascending: closest deadline first)
    def sort_key(opp):
        notice_type = str(opp.get("type", "")).lower()
        score = opp.get("Relevance_Score", 0)
        
        if "presolicitation" in notice_type:
            type_rank = 2
        elif "solicitation" in notice_type:
            type_rank = 0
        else:
            type_rank = 1
        
        deadline_str = opp.get("responseDeadLine") or opp.get("responseDate")
        parsed_date = datetime.datetime.max # Default to max so NULLs go to the end
        if deadline_str:
            try:
                parsed_date = datetime.datetime.strptime(deadline_str[:10], "%Y-%m-%d")
            except:
                pass
                
        return (type_rank, -score, parsed_date)

    filtered.sort(key=sort_key)
    return filtered

def save_to_csv(opportunities, filename="syte_opportunities.csv"):
    if not opportunities:
        print("\nNo opportunities found matching your criteria to save.")
        return set(), set()

    base_dir = os.path.dirname(os.path.abspath(__file__))

    primary_cols = [
        "type",
        "title",
        "location",
        "site_visit_date",
        "ROM",
        "solicitationNumber",
        "department",
        "typeOfSetAsideDescription",
        "responseDeadLine",
        "postedDate",
        "naicsCode"
    ]
    exclude_cols = {"archiveType", "classificationCode", "noticeId", "Relevance_Score"}

    # Build fieldnames from new data
    keys = set()
    for opp in opportunities:
        keys.update(opp.keys())
    fieldnames = primary_cols + [
        k for k in keys
        if k not in primary_cols and k not in exclude_cols and not isinstance(opportunities[0].get(k), (list, dict))
    ]

    # Load existing syte_opportunities.csv to merge (incremental)
    existing_path = os.path.join(base_dir, filename)
    existing = {}
    if os.path.exists(existing_path):
        with open(existing_path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                key = row.get("solicitationNumber") or row.get("title", "")
                if key:
                    existing[key] = row
        print(f"  Loaded {len(existing)} existing records from {filename} for merging.")

    # Merge: new data takes priority over existing
    new_by_key = {}
    for opp in opportunities:
        key = opp.get("solicitationNumber") or opp.get("title", "")
        if key:
            new_by_key[key] = {k: opp.get(k, "") for k in fieldnames}

    new_keys     = set(new_by_key.keys()) - set(existing.keys())
    updated_keys = set(new_by_key.keys()) & set(existing.keys())

    merged = {**existing, **new_by_key}  # new overwrites old on same key
    merged_rows = list(merged.values())

    # Extend fieldnames to cover any extra columns from existing data
    existing_cols = set()
    for row in merged_rows:
        existing_cols.update(row.keys())
    for col in existing_cols:
        if col not in fieldnames:
            fieldnames.append(col)

    print(f"\nProcessing {len(merged_rows)} total records ({len(new_by_key)} new/updated, "
          f"{len(existing) - len(set(existing) & set(new_by_key))} unchanged from history)...")

    # Write date-stamped file
    today_str = datetime.datetime.now().strftime("%Y%m%d")
    dated_filename = f"syte_opportunities_{today_str}.csv"
    dated_path = os.path.join(base_dir, dated_filename)
    with open(dated_path, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(merged_rows)
    print(f"Saved dated snapshot: {dated_path}")

    # Also update the canonical syte_opportunities.csv
    with open(existing_path, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(merged_rows)
    print(f"Updated canonical file: {existing_path}")
    return new_keys, updated_keys

def push_to_pipedrive(opportunities):
    if not PIPEDRIVE_TOKEN:
        print("\n[Pipedrive] Skipped: PIPEDRIVE_API_TOKEN not set in .env")
        return

    print("\n=== Syncing to Pipedrive (Pipeline: RFP/Solicitation) ===")

    # Fetch existing deal titles to avoid duplicates
    existing_titles = set()
    page = 0
    while True:
        r = requests.get(
            f"{PIPEDRIVE_BASE}/deals",
            params={"api_token": PIPEDRIVE_TOKEN, "limit": 500, "start": page * 500,
                    "pipeline_id": PIPEDRIVE_PIPELINE_ID, "status": "open"}
        )
        deals_data = r.json().get("data") or []
        for d in deals_data:
            existing_titles.add(d["title"].strip().lower())
        if not r.json().get("additional_data", {}).get("pagination", {}).get("more_items_in_collection"):
            break
        page += 1
    print(f"  -> Found {len(existing_titles)} existing deals in pipeline (for dedup check).")

    created, skipped = 0, 0
    for opp in opportunities:
        title = (opp.get("title") or "").strip()
        if not title:
            continue

        if title.lower() in existing_titles:
            skipped += 1
            continue

        # Parse deadline for expected_close_date
        deadline_str = opp.get("responseDeadLine") or ""
        close_date = ""
        if deadline_str:
            try:
                close_date = datetime.datetime.strptime(deadline_str[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
            except:
                pass

        payload = {
            "title": title,
            "pipeline_id": PIPEDRIVE_PIPELINE_ID,
            "stage_id": PIPEDRIVE_STAGE_ID,
        }
        if close_date:
            payload["expected_close_date"] = close_date

        # Set deal value as ROM midpoint
        rom_text = opp.get("ROM", "")
        midpoint = parse_rom_midpoint(rom_text)
        if midpoint:
            payload["value"] = int(midpoint)
            payload["currency"] = "USD"

        r = requests.post(
            f"{PIPEDRIVE_BASE}/deals",
            params={"api_token": PIPEDRIVE_TOKEN},
            json=payload
        )
        result = r.json()

        if result.get("success"):
            deal_id = result["data"]["id"]
            existing_titles.add(title.lower())
            created += 1

            # Add key details as a note on the deal
            sol_num = opp.get("solicitationNumber", "N/A")
            notice_type = opp.get("type", "N/A")
            location = opp.get("location", "N/A")
            set_aside = opp.get("typeOfSetAsideDescription", "N/A")
            naics = opp.get("naicsCode", "N/A")
            dept = opp.get("fullParentPathName", "N/A")
            sam_link = opp.get("uiLink", "")
            posted = opp.get("postedDate", "N/A")

            note_content = (
                f"<b>📋 SAM.gov Federal Contract Opportunity</b><br>"
                f"<b>Type:</b> {notice_type}<br>"
                f"<b>Solicitation #:</b> {sol_num}<br>"
                f"<b>Location:</b> {location}<br>"
                f"<b>Set-Aside:</b> {set_aside}<br>"
                f"<b>NAICS Code:</b> {naics}<br>"
                f"<b>Department:</b> {dept}<br>"
                f"<b>Posted Date:</b> {posted}<br>"
                f"<b>Response Deadline:</b> {deadline_str[:10] if deadline_str else 'N/A'}<br>"
                f"<b>SAM.gov Link:</b> <a href='{sam_link}'>{sam_link}</a>"
            )
            requests.post(
                f"{PIPEDRIVE_BASE}/notes",
                params={"api_token": PIPEDRIVE_TOKEN},
                json={"content": note_content, "deal_id": deal_id, "pinned_to_deal_flag": True}
            )
        else:
            print(f"  [!] Failed to create deal for: {title[:60]} — {result.get('error', 'unknown error')}")

    print(f"  -> Done! Created {created} new deals, skipped {skipped} duplicates.")


def safe_request(method, url, **kwargs):
    """HTTP request with up to 3 retries on connection errors."""
    import time
    for attempt in range(3):
        try:
            return requests.request(method, url, timeout=15, **kwargs)
        except requests.exceptions.ConnectionError:
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)


def backfill_pipedrive_values():
    """For existing Pipedrive deals that have no value set,
       scan their notes for dollar amounts and update the deal value with the midpoint."""
    if not PIPEDRIVE_TOKEN:
        return

    print("\n=== Backfilling Deal Values from Notes ===")
    updated, skipped, errors = 0, 0, 0
    start = 0

    while True:
        r = safe_request("GET", f"{PIPEDRIVE_BASE}/deals",
            params={"api_token": PIPEDRIVE_TOKEN, "limit": 100, "start": start,
                    "pipeline_id": PIPEDRIVE_PIPELINE_ID, "status": "open"})
        deals = r.json().get("data") or []
        if not deals:
            break

        for deal in deals:
            deal_id = deal["id"]
            current_value = deal.get("value") or 0

            if current_value and float(current_value) > 0:
                skipped += 1
                continue  # Already has a value

            try:
                notes_r = safe_request("GET", f"{PIPEDRIVE_BASE}/notes",
                    params={"api_token": PIPEDRIVE_TOKEN, "deal_id": deal_id, "limit": 5})
                notes = notes_r.json().get("data") or []
            except Exception:
                errors += 1
                continue
            combined_notes = " ".join(n.get("content", "") for n in notes)

            midpoint = parse_rom_midpoint(combined_notes)
            if midpoint and midpoint > 0:
                requests.put(f"{PIPEDRIVE_BASE}/deals/{deal_id}",
                    params={"api_token": PIPEDRIVE_TOKEN},
                    json={"value": int(midpoint), "currency": "USD"})
                updated += 1

        if not r.json().get("additional_data", {}).get("pagination", {}).get("more_items_in_collection"):
            break
        start += 100

    print(f"  -> Updated {updated} deals with ROM midpoint values, skipped {skipped} (already had value).")


if __name__ == "__main__":
    print("===========================================")
    print("SYTE Corp - SAM.gov Federal Contract Search")
    print("===========================================\n")

    raw_results = get_opportunities(limit=100)
    processed_results = filter_and_sort_opportunities(raw_results)
    save_to_csv(processed_results)
    write_last_run_date()
    print(f"\nLast run date updated to {datetime.datetime.now().strftime('%Y-%m-%d')}.")
    # push_to_pipedrive(processed_results)
    # backfill_pipedrive_values()
