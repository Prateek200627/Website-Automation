#cbi_announced_new_reward
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime
from datetime import date
import re
import os

URL = "https://cbi.gov.in/announce-rewards"

def format_dob(dob):
    if not dob:
        return ""

    dob = dob.strip()

    # Keep DOB blank if it is NA/N.A./N/A
    if dob.upper() in ["NA", "N/A", "N.A.", "NOT AVAILABLE", "-"]:
        return ""

    formats = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d %b %Y",
        "%d %B %Y",
        "%Y-%m-%d"
    ]

    for fmt in formats:
        try:
            return datetime.strptime(dob, fmt).strftime("%d/%m/%Y")
        except:
            continue

    return dob

def calculate_age(dob):

    if not dob:
        return ""

    try:
        dob = datetime.strptime(dob, "%d/%m/%Y").date()

        today = date.today()

        age = today.year - dob.year - (
            (today.month, today.day) < (dob.month, dob.day)
        )

        return age

    except:
        return ""
    
INDIAN_STATES = [
    "Andhra Pradesh",
    "Arunachal Pradesh",
    "Assam",
    "Bihar",
    "Chhattisgarh",
    "Goa",
    "Gujarat",
    "Haryana",
    "Himachal Pradesh",
    "Jharkhand",
    "Karnataka",
    "Kerala",
    "Madhya Pradesh",
    "Maharashtra",
    "Manipur",
    "Meghalaya",
    "Mizoram",
    "Nagaland",
    "Odisha",
    "Punjab",
    "Rajasthan",
    "Sikkim",
    "Tamil Nadu",
    "Telangana",
    "Tripura",
    "Uttar Pradesh",
    "Uttarakhand",
    "West Bengal",
    "Delhi",
    "Jammu and Kashmir",
    "Ladakh",
    "Chandigarh",
    "Puducherry",
    "Andaman and Nicobar Islands",
    "Dadra and Nagar Haveli and Daman and Diu",
    "Lakshadweep"
]

STATE_MAPPING = {

    # West Bengal
    "kolkata": "West Bengal",
    "calcutta": "West Bengal",
    "howrah": "West Bengal",
    "purulia": "West Bengal",
    "asansol": "West Bengal",
    "siliguri": "West Bengal",
    "durgapur": "West Bengal",

    # Bihar
    "patna": "Bihar",
    "begusarai": "Bihar",
    "gaya": "Bihar",
    "bhagalpur": "Bihar",
    "munger": "Bihar",
    "muzaffarpur": "Bihar",

    # Uttar Pradesh
    "agra": "Uttar Pradesh",
    "lucknow": "Uttar Pradesh",
    "kanpur": "Uttar Pradesh",
    "meerut": "Uttar Pradesh",
    "varanasi": "Uttar Pradesh",
    "ghaziabad": "Uttar Pradesh",
    "prayagraj": "Uttar Pradesh",

    # Delhi
    "delhi": "Delhi",
    "new delhi": "Delhi",

    # Punjab
    "ludhiana": "Punjab",
    "amritsar": "Punjab",
    "jalandhar": "Punjab",
    "moga": "Punjab",

    # Haryana
    "panipat": "Haryana",
    "gurgaon": "Haryana",
    "gurugram": "Haryana",
    "faridabad": "Haryana",

    # Maharashtra
    "mumbai": "Maharashtra",
    "pune": "Maharashtra",
    "nagpur": "Maharashtra",
    "thane": "Maharashtra",

    # Tamil Nadu
    "chennai": "Tamil Nadu",
    "coimbatore": "Tamil Nadu",
    "madurai": "Tamil Nadu",
    "thanjavur": "Tamil Nadu",

    # Karnataka
    "bangalore": "Karnataka",
    "bengaluru": "Karnataka",
    "mysore": "Karnataka",

    # Telangana
    "hyderabad": "Telangana",

    # Andhra Pradesh
    "visakhapatnam": "Andhra Pradesh",
    "vijayawada": "Andhra Pradesh",
    "tirupati": "Andhra Pradesh",

    # Odisha
    "bhubaneswar": "Odisha",
    "cuttack": "Odisha",
    "puri": "Odisha",

    # Kerala
    "kochi": "Kerala",
    "ernakulam": "Kerala",
    "allepy": "Kerala",
    "alappuzha": "Kerala",

    # Gujarat
    "ahmedabad": "Gujarat",
    "surat": "Gujarat",
    "rajkot": "Gujarat",

    # Rajasthan
    "jaipur": "Rajasthan",
    "jodhpur": "Rajasthan",

    # Madhya Pradesh
    "bhopal": "Madhya Pradesh",
    "indore": "Madhya Pradesh",

    # Chhattisgarh
    "raipur": "Chhattisgarh",

    # Jharkhand
    "ranchi": "Jharkhand",

    # Assam
    "guwahati": "Assam",

    # Tripura
    "agartala": "Tripura",
    "north tripura": "Tripura"
}

def extract_state(address):

    if not address:
        return ""

    address = address.lower()

    # 1. First check city/district names
    for location, state in STATE_MAPPING.items():
        for location, state in STATE_MAPPING.items():
            if re.search(rf"\b{re.escape(location)}\b", address):
                return state

    # 2. Then check explicit state names
    for state in INDIAN_STATES:
        if state.lower() in address:
            return state

    return ""

def split_name_alias(name):
    if not name:
        return "", ""

    name = re.sub(r"\s+", " ", name).strip()

    if "@" not in name:
        return name, ""

    parts = [p.strip() for p in name.split("@") if p.strip()]

    full_name = parts[0]
    aliases = "; ".join(parts[1:])

    return full_name, aliases



def cbi_new_rewards():

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)

    driver.get(URL)

    main_window = driver.current_window_handle

    total = len(wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'mostWnatedBox')]"))))

    print(f"Records Found : {total}")

    cards = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'mostWnatedBox')]//a")))

    href_list = []

    for card in cards:
        href = card.get_attribute("href")
        if href:
            href_list.append(href)

    print("Total URLs:", len(href_list))

    output_file = "CBI_Announced_Rewards.xlsx"

    if os.path.exists(output_file):

        existing_df = pd.read_excel(output_file, dtype=str).fillna("")

        existing_links = set(existing_df["WEB_LINK"].str.strip())

        all_records = existing_df.to_dict("records")

        print(f"Existing Records : {len(existing_links)}")

    else:

        existing_links = set()

        all_records = []


    main_window = driver.current_window_handle

    for i, href in enumerate(href_list, start=1):
        if href in existing_links:
            print(f"Already Exists : {href}")
            continue

        print(f"Processing Record {i} of {len(href_list)}")
        print(href)

        # Open the URL in a new tab
        driver.execute_script("window.open(arguments[0], '_blank');", href)

        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[-1])

        # Wait for page to load
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # -------------------------
        # Scrape your data here
        # -------------------------

        boxes = driver.find_elements(By.XPATH, "(//div[contains(@class,'oldRecordContent')]//div[contains(@class,'box')])[position() < last()]")

        record = {}

        for box in boxes:
            print("=" * 80)
            print(box.get_attribute("outerHTML"))

            try:
                label = box.find_element(By.XPATH, "./p[1]").text.strip().replace(":", "")
                value = box.find_element(By.XPATH, "./p[2]").text.strip()

                print("LABEL :", label)
                print("VALUE :", value)

                record[label] = value

            except Exception as e:
                print("ERROR :", e)

        print(record.keys())

        name, alias = split_name_alias(record.get("Name", ""))
        father_name = record.get("Father Name", "")

        if father_name.strip().upper() in [
            "NA",
            "N/A",
            "N.A.",
            "NOT AVAILABLE",
            "UNKNOWN",
            "-",
            "--"
        ]:
            father_name = ""
        else:
            father_name = re.sub(r"\s*@\s*", "; ", father_name)
            father_name = re.sub(r"^(late\s+|late\.?\s+|mr\.?\s+|mrs\.?\s+|ms\.?\s+|shri\s+|sh\.?\s+|smt\.?\s+)+","",father_name,flags=re.IGNORECASE).strip()

        dob = format_dob(record.get("Date Of Birth", ""))
        age = calculate_age(dob)
        gender = record.get("Gender", "")
        address = record.get("Address", "")
        if address.strip().upper() in ["NA", "N/A", "N.A.", "NOT AVAILABLE", "-"]:
            address = ""
        state = extract_state(address)
        reward = record.get("Reward", "")
        details = record.get("Details", "")

        if details.strip().upper() in ["NA", "N/A", "N.A.", "NOT AVAILABLE", "-", "--"]:
            details = ""

        address_lower = address.lower()

        # If the address contains Singapore, keep nationality blank
        if "singapore" in address_lower:
            nationality = ""
            country = ""
        else:
            nationality = "Indian"
            country = "India"

        if age != "":
            details = f"Age - {age} Years ; {details}"

        else:
            details = f"Age - {age} Years"

        all_records.append({
            "FULL_NAME": name,
            "CATEGORY": "P",
            "F_NAME": "",
            "M_NAME": "",
            "L_NAME": "",
            "GENDER": gender,
            "DOB": dob,
            "ADD_CITY": "",
            "ADD_COUNTRY": country,
            "State": state,
            "Nationalities": nationality,
            "ADDRESS": address,
            "Identity Number": "",
            "Identity Type": "",
            "REF_DATE": "",
            "DETAILS": details,
            "WEB_LINK": href,
            "VIOLATION_ID": "",
            "SOURCE": "CBI ANNOUNCED REWARDS",
            "Alias": alias,
            "Associates": "",
            "MainActivity": "",
            "Citizenship information": "",
            "STATUS": "",
            "Rem1": f"Father Name - {father_name}" if father_name else "",
            "Rem2": "",
            "Rem3": "",
            "Remarks": ""
        })
        
        existing_links.add(href)

        print(driver.title)

        # Example:
        # name = driver.find_element(By.XPATH, "//h5").text
        # print(name)

        # Close the tab
        driver.close()

        # Go back to the main page
        driver.switch_to.window(main_window)

        # Close the browser
    driver.quit()

    # Column order
    columns = [
        "FULL_NAME",
        "CATEGORY",
        "F_NAME",
        "M_NAME",
        "L_NAME",
        "GENDER",
        "DOB",
        "ADD_CITY",
        "ADD_COUNTRY",
        "State",
        "Nationalities",
        "ADDRESS",
        "Identity Number",
        "Identity Type",
        "REF_DATE",
        "DETAILS",
        "WEB_LINK",
        "VIOLATION_ID",
        "SOURCE",
        "Alias",
        "Associates",
        "MainActivity",
        "Citizenship information",
        "STATUS",
        "Rem1",
        "Rem2",
        "Rem3",
        "Remarks"
    ]

    # Create DataFrame
    df = pd.DataFrame(all_records)

    # Ensure all columns exist and are in the desired order
    df = df.reindex(columns=columns)

    # Replace NaN values with empty strings
    df.fillna("", inplace=True)

    # Save to Excel
    output_file = "CBI_Announced_Rewards.xlsx"
    df.to_excel(output_file, index=False)

    print(f"\nExcel file saved successfully: {output_file}")
    print(f"Total Records Saved: {len(df)}")



if __name__ == "__main__":
    cbi_new_rewards()