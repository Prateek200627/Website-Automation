import os
import sys
from urllib.parse import urljoin
from flask import logging
import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import re
import logging

# Get executable/script name
if getattr(sys, "frozen", False):
    app_name = os.path.splitext(os.path.basename(sys.executable))[0]
    app_dir = os.path.dirname(sys.executable)
else:
    app_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    app_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

ENABLE_LOGS = "withlogs" in app_name.lower()

if ENABLE_LOGS:
    logs_folder = os.path.join(app_dir, "logs")
    os.makedirs(logs_folder, exist_ok=True)

    logging.basicConfig(
        filename=os.path.join(logs_folder, "CBI_Old_Records.log"),
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
else:
    logging.disable(logging.CRITICAL)

URL = "https://cbi.gov.in/old-records"
EXCEL_FILE = "CBI_Old_Records.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

import re

def clean_name(name):
    pattern = r'^(Mr\.?|Mrs\.?|Ms\.?|Miss|Shri|Sh\.?|Smt\.?|Sri|Dr\.?|Shri\.?)\s+'
    return re.sub(pattern, '', name, flags=re.IGNORECASE).strip()

def clean_father_name(father_name):
    if not father_name:
        return ""

    father_name = father_name.strip()

    # Remove invalid values
    if father_name.upper() in [
        "N/A", "NA", "NIL", "NONE", "NAME N/A", "NAM", "-"
    ]:
        return ""

    # Remove common prefixes repeatedly
    pattern = (
        r'^(?:'
        r'late\s+'
        r'|mr\.?\s+'
        r'|mrs\.?\s+'
        r'|ms\.?\s+'
        r'|miss\s+'
        r'|shri\.?\s+'
        r'|sh\.?\s+'
        r'|smt\.?\s+'
        r'|sri\s+'
        r'|dr\.?\s+'
        r')+'
    )

    father_name = re.sub(pattern, "", father_name, flags=re.IGNORECASE)

    # Remove extra spaces
    father_name = re.sub(r"\s+", " ", father_name).strip()

    return father_name

def get_soup():
    """
    Fetch the webpage and return BeautifulSoup object.
    """
    response = requests.get(URL, headers=HEADERS)
    logging.info("Website opened successfully")
    response.raise_for_status()

    return BeautifulSoup(response.text, "html.parser")


def extract_records(soup):

    logging.info("=" * 80)
    logging.info("CBI OLD RECORDS SCRAPER STARTED")
    """
    Extract all wanted person records from the webpage.
    """

    records = []

    cards = soup.find_all("div", class_="oldRecordContent")

    print(f"Records Found : {len(cards)}")
    logging.info(f"Records Found : {len(cards)}")

    for card in cards:

        # Temporary variables
        name = ""
        dob = ""
        gender = ""
        charges = ""
        branch = ""
        address = ""
        pdf_link = ""

        father_name = ""

        father_remarks = ""

        # Process father name: handle aliases (using '@') and multiple names
        if father_name:
            father_name = father_name.strip()

            # Case 1: Father alias using '@' (e.g. 'John Doe@Alias1@Alias2')
            if "@" in father_name:
                parts = [p.strip() for p in father_name.split("@")]

                # First part is the actual father name
                father_name = parts[0]

                # Remaining parts are aliases
                father_remarks = "Father Alias: " + "; ".join(parts[1:])

            # Case 2: Multiple father names separated by /, &, , or ;
            elif any(sep in father_name for sep in ["/", "&", ",", ";"]):
                father_remarks = f"Father Name: {father_name}"
                father_name = ""
            
        boxes = card.find_all("div", class_="box")

        for box in boxes:

            paragraphs = box.find_all("p")

            if len(paragraphs) >= 2:

                key = paragraphs[0].get_text(strip=True).replace(":", "")
                value = paragraphs[1].get_text(" ", strip=True)

                if key == "Name":
                    name = value

                elif key == "Father Name":
                    father_name = clean_father_name(value)

                elif key == "Gender":
                    gender = value

                elif key == "Charges":
                    charges = " ".join(value.split())
                    # charges = value

                elif key == "Branch Name":
                    branch = value

                elif key == "Date Of Birth":
                    dob = value.strip()

                    if dob.upper() == "N/A":
                        dob = ""
                    else:
                        try:
                            dob = datetime.strptime(dob, "%Y-%m-%d").strftime("%d/%m/%Y")
                        except ValueError:
                            pass

            # Address
            heading = box.find("h3")

            if heading and "Address" in heading.get_text():

                p = box.find("p")

                if p:
                    address = p.get_text("\n", strip=True)
                    address = re.sub(r'\n+', '; ', address)
                    address = re.sub(r'\s*;\s*', '; ', address)
                    address = re.sub(r'\s+', ' ', address).strip()

            # PDF Link
            a = box.find("a")

            if a:
                pdf_link = urljoin(URL, a["href"])

        # Alias
        alias = ""

        if "@" in name:
            parts = [part.strip() for part in name.split("@")]

            # First part is the actual name
            name = parts[0]

            # Remaining parts are aliases separated by ;
            alias = "; ".join(parts[1:])

        # Remove prefixes like Mr., Shri, etc.
        name = clean_name(name)
        # Excel Format
        data = {
            "FULL_NAME": name,
            "CATEGORY": "P",
            "F_NAME": "",
            "M_NAME": "",
            "L_NAME": "",
            "GENDER": gender,
            "DOB": f"{dob}" if dob else "",
            "ADD_CITY": "",
            "ADD_COUNTRY": "",
            "State": "",
            "Nationalities": "",
            "ADDRESS": address if address else "",
            "Identity Number": "",
            "Identity Type": "",
            "REF_DATE": "",
            "DETAILS": f"Case ID date of Registration and Section of Law - {charges}",
            "WEB_LINK": "https://cbi.gov.in/old-records",
            "VIOLATION_ID": "",
            "SOURCE": "CBI- Rewards (Old)",
            "Alias": alias,
            "Associates": "",
            "MainActivity": "",
            "Citizenship information": "",
            "STATUS": "",
            "Rem1": f"S/o {father_name}".replace("@", "; ") if father_name else "",
            "Rem2": "",
            "Rem3": f"Branch Name - {branch}" if branch else "",
            "Remarks": father_remarks
        }

        records.append(data)

        logging.info(f"Added : {name}")

    return records

logging.info("=" * 80)
logging.info("SCRAPER COMPLETED")
logging.info("=" * 80)

def save_to_excel(records):

    new_df = pd.DataFrame(records)

    # If Excel doesn't exist, create it
    if not os.path.exists(EXCEL_FILE):
        new_df.to_excel(EXCEL_FILE, index=False)
        print(f"Scraped Records : {len(new_df)}")
        print("New Records     :", len(new_df))
        print("Duplicates      : 0")
        print("Total in Excel  :", len(new_df))
        return

    # Read existing Excel
    old_df = pd.read_excel(EXCEL_FILE).fillna("")
    new_df = new_df.fillna("")

    # Create unique keys from existing records
    existing_keys = set(
        zip(
            old_df["FULL_NAME"].str.strip().str.lower(),
            old_df["DOB"].astype(str).str.strip(),
            old_df["SOURCE"].str.strip().str.lower()
        )
    )

    new_rows = []
    added = 0
    skipped = 0

    for _, row in new_df.iterrows():

        key = (
            row["FULL_NAME"].strip().lower(),
            str(row["DOB"]).strip(),
            row["SOURCE"].strip().lower()
        )

        if key in existing_keys:
            skipped += 1
        else:
            new_rows.append(row)
            existing_keys.add(key)
            added += 1

    if new_rows:
        final_df = pd.concat(
            [old_df, pd.DataFrame(new_rows)],
            ignore_index=True
        )
    else:
        final_df = old_df

    final_df.to_excel(EXCEL_FILE, index=False)

    print("\n========== SUMMARY ==========")
    print(f"Scraped Records : {len(new_df)}")
    print(f"New Records     : {added}")
    print(f"Duplicates      : {skipped}")
    print(f"Total in Excel  : {len(final_df)}")
    print("=============================")
    logging.info("=" * 80)
    logging.info("SUMMARY")
    logging.info(f"Scraped Records : {len(new_df)}")
    logging.info(f"New Records     : {added}")
    logging.info(f"Duplicate Records: {skipped}")
    logging.info(f"Total in Excel  : {len(final_df)}")
    logging.info("=" * 80)

    
def main():

    soup = get_soup()

    records = extract_records(soup)
    save_to_excel(records)

    
if __name__ == "__main__":
    main()