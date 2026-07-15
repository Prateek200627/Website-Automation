import re
#
import requests
import pandas as pd
from pathlib import Path
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# ============================================
# CONFIG
# ============================================

OUTPUT_FILE = "NSDL_Nondelivery_PAN_Part_26.xlsx"
URL = "https://nsdl.com/resources/regulatory-action/list-of-pan-demat-accounts-frozen"

API_URL = "https://nsdl.com/web/api/v1/pan-demat-listing"

HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Referer": "https://nsdl.com/resources/regulatory-action/list-of-pan-demat-accounts-frozen",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
    "sec-ch-ua": '"Google Chrome";v="149", "Chromium";v="149", "Not?A_Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}

session = requests.Session()

# ============================================
# EXISTING VARIABLES
# ============================================
MARK_DUPLICATES = True
existing_records = set()
existing_data = []
new_records = []
# =====================================================
# HELPERS
# =====================================================

def clean_text(text):
    if text is None:
        return ""

    text = str(text)
    text = ILLEGAL_CHARACTERS_RE.sub("", text)

    return text.strip()

def normalize_company_words(text):
    text = clean_text(text)

    # Normalize separators
    text = re.sub(r"[./]+", " ", text)

    # Normalize multiple spaces
    text = re.sub(r"\s+", " ", text).strip()

    # Replace abbreviations
    replacements = {
        "pvt": "Private",
        "ltd": "Limited",
        "p": "Private",          # Handles P/Ltd
    }

    words = []

    for word in text.split():
        key = word.lower()

        if key in replacements:
            words.append(replacements[key])
        else:
            words.append(word)

    return " ".join(words)

male_titles = {
    "mr", "mr.", "mr/",
    "shri", "shree", "sri"
}

female_titles = {
    "ms", "ms.",
    "mrs", "mrs.",
    "miss",
    "smt", "smt.",
    "kumari","M s"
}

persons = {
    "dr", "dr.",
    "prof", "prof.",
    "adv", "adv."
}
entity_keywords = {
    "m/s","M/s.","M/S",
    "limited",
    "private",
    "llp",
    "corporation",
    "corp",
    "trader",
    "traders",
    "merchant",
    "capital",
    "company",
    "co",
    "enterprises",
    "industries",
    "developers",
    "holdings",
    "finance",
    "financial",
    "securities",
    "HUF","LLP","Trust","Pvt.", "Ltd.", "Ltd", "Pvt Ltd.", "Pvt. Ltd.", "Pvt Ltd",
    "Co.", "Corp.", "Inc.",
    "PRIVATE LIMITED", "PRIVATE LTD", "LIMITED", "PVT", "LTD",
    "LLP", "OPC", "COMPANY",
    "Trader", "Traders", "Trading", "Tradecom", "Tradevine",
    "Commodeal", "Commodities", "Securities", "Finance",
    "Investments", "Marketing", "Realty", "Infrastructure",
    "Developers", "Builders", "Enterprises", "Industries",
    "Projects", "Holdings", "Ventures", "Broking",
    "Consultants", "Agencies", "Distributors", "Merchants",
    "Securities",
    "HUF", "LLP", "Trust", "Family Trust",
    "Beneficiary Trust", "Foundation", "Society", "Association"
}
def get_category(full_name):
    name = clean_text(full_name).lower()
    if any(keyword.lower() in name for keyword in entity_keywords):
        return "E"

    return "P"
def get_gender(full_name):

    name = clean_text(full_name).lower()
    if any(keyword.lower() in name for keyword in entity_keywords):
        return ""

    words = name.split()
    if not words:
        return ""

    first_word = (
    words[0]
    .lower()
    .replace("/", "")
    .replace(".", "")
    )

    if first_word in {"mr", "shri", "shree", "sri"}:
        return "Male"

    if first_word in {"ms", "mrs", "smt", "miss", "kumari"}:
        return "Female"

    return ""
def load_existing_data():
    print("Checking if file exists...")

    print("OUTPUT_FILE =", OUTPUT_FILE)
    print("Exists =", Path(OUTPUT_FILE).exists())
    if not Path(OUTPUT_FILE).exists():

        print("No existing excel found. Fresh run.")
        # write_log("No existing excel found. Fresh run.")

        return

    df = pd.read_excel(
    OUTPUT_FILE,
    dtype=str,
    keep_default_na=False
    )

    existing_data.extend(df.to_dict("records"))

    for _, row in df.iterrows():

        key = (
            clean_text(row.get("FULL_NAME", "")).lower(),
            clean_text(row.get("IDENTIFY NUMBER", "")).upper(),
            clean_text(row.get("DETAILS", "")).lower(),
            clean_text(row.get("REM1", "")).lower()
        )

        existing_records.add(key)
    # write_log(f"Loaded {len(existing_records)} existing records")

def save_excel():

    if not new_records:
        # write_log("No new records found")
        return

    all_records = existing_data + new_records

    df = pd.DataFrame(all_records)

    if MARK_DUPLICATES:

        df["REMARKS"] = ""

        df["_dup_key"] = (
            df["FULL_NAME"].astype(str).str.lower().str.strip() + "|" +
            df["IDENTIFY NUMBER"].astype(str).str.upper().str.strip() + "|" +
            df["DETAILS"].astype(str).str.lower().str.strip() + "|" +
            df["REM1"].astype(str).str.lower().str.strip()
        )

        dup_mask = df.duplicated(
        subset=["_dup_key"],
        keep=False
        )
        df.loc[dup_mask, "REMARKS"] = "Duplicate"

        df.loc[dup_mask, "REMARKS"] = "Duplicate"

        df.drop(columns=["_dup_key"], inplace=True)

    df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    # write_log(f"Total records in file : {len(df)}")
    print(len(df))
    # write_log(f"New records added     : {len(new_records)}")

# ============================================
# API FETCH
# ============================================

def fetch_data():

    # ---------------- First API ----------------
    first_response = session.get(
        API_URL,
        headers=HEADERS,
        params={
            "page": 0,
            "limit": 1
        }
    )

    first_response.raise_for_status()

    print("=" * 80)
    print("FIRST API")
    print("Status :", first_response.status_code)
    print("URL    :", first_response.url)

    first_json = first_response.json()

    total_records = int(first_json["total_count"])

    print("Total Records :", total_records)

    # ---------------- Second API ----------------
    second_response = session.get(
        API_URL,
        headers=HEADERS,
        params={
            "page": 0,
            "limit": total_records
        }
    )

    second_response.raise_for_status()

    print("=" * 80)
    print("SECOND API")
    print("Status :", second_response.status_code)
    print("URL    :", second_response.url)

    second_json = second_response.json()

    rows = second_json["data"]

    print(f"Records received from API : {len(rows)}")

    # ---------------- Process Records ----------------
    for idx, row in enumerate(rows, start=1):

        try:

            full_name = normalize_company_words(
                clean_text(row.get("client_name", ""))
            )

            pan_no = clean_text(
                row.get("client_pan", "")
            )

            detail = clean_text(
                row.get("particulars_sebi_direction", "")
            )

            ref_sebi = clean_text(
                row.get("reference_sebi_direction", "")
            )

            remark1 = clean_text(
                row.get("remarks", "")
            )

            gender = get_gender(full_name)
            category = get_category(full_name)

            record_key = (
                full_name.lower().strip(),
                pan_no.upper().strip(),
                detail.lower().strip(),
                remark1.lower().strip()
            )

            if record_key in existing_records:
                continue

            new_records.append({
                "FULL_NAME": full_name,
                "CATEGORY": category,
                "F_NAME": "",
                "M_NAME": "",
                "L_NAME": "",
                "GENDER": gender,
                "DOB": "",
                "ADD_COUNTRY": "INDIA",
                "STATE": "",
                "NATIONALITIES": "",
                "ADDRESS": "",
                "IDENTIFY NUMBER": pan_no,
                "IDENTIFY TYPE": "PAN",
                "REF_DATE": "",
                "DETAILS": detail,
                "WEB_LINK": URL,
                "VIOLATION_ID": "",
                "SOURCE": "Nondelivery PAN NSDL",
                "ALIAS": "",
                "ASSOCIATES": "",
                "MAINACTIVITY": (
                    "Non-delivery of SCN / Orders as per "
                    f"SEBI Circular ref. no - {ref_sebi}"
                ),
                "CITIZENSHIP INFORMATION": "",
                "STATUS": "",
                "REM1": remark1,
                "REM2": "",
                "REM3": "",
                "REMARKS": ""
            })

        except Exception as e:
            print(f"Failed row {idx}: {e}")

    print(f"New records collected : {len(new_records)}")

def main():

    try:

        print("Step 1: Loading existing data")
        load_existing_data()    
        print("Step 2: Fetching API data")
        fetch_data()

        print(f"New records collected: {len(new_records)}")

        print("Step 3: Saving Excel")
        save_excel()

        print("Completed Successfully")

    except Exception as e:
        print(f"Fatal Error: {e}")
        # write_log(f"Fatal Error: {e}")

if __name__ == "__main__":
    main()