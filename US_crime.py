from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from datetime import datetime
import time,os,re,logging
import pandas as pd
log_folder = "logs"
os.makedirs(log_folder, exist_ok=True)
#

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(log_folder, f"US_Secret_Service_{timestamp}.log")

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
    force=True
)

def us_secret_service(existing_urls):
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    logging.info("US Secret Service scraper started")
    driver.get("https://www.secretservice.gov/investigations/mostwanted")
    logging.info("Website opened successfully")
    time.sleep(1)
    wait = WebDriverWait(driver,2)
    criminals = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'wanted-card')]")))
    total_criminals = len(criminals)
    items = []
    for i in range(total_criminals):
        criminals = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'wanted-card')]")))
        criminal = criminals[i]
        logging.info("clicking on View more")
        view_more = criminal.find_element(By.XPATH, ".//a[contains(.,'View More')]")
        view_more.click()

        url = driver.current_url
        logging.info(f"URL: {url}")

        criminal_name = wait.until(EC.presence_of_element_located((By.XPATH, "//h1[contains(@class,'uswds-page-title')]/span"))).text.strip()
        logging.info(f"Processing: {criminal_name}")

        if url in existing_urls:
            logging.info(f"Skipped existing record: {url}")

            driver.back()
            wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'wanted-card')]")))
            continue

        time.sleep(1)
        p = driver.find_element(By.XPATH, "//h2[normalize-space()='CASE SUMMARY']/preceding-sibling::p[1]")

        data = {}
        for line in p.text.splitlines():
            line = line.strip()
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            data[key.strip()] = value.strip()

        aliases_raw = data.pop("ALIASES", "")
        dob_raw = data.get("DOB", "")
        nationality = data.get("NATIONALITY", "")
        citizenship = data.get("CITIZENSHIP", "")
        english_aliases = []
        english_aliases = re.findall(r'[A-Za-z0-9][A-Za-z0-9 .\'"()/-]*',aliases_raw)
        aliases = "; ".join(alias.strip(' "“”') for alias in english_aliases)
        
        dob = ""
        if dob_raw and dob_raw != "N/A":
            try:
                dob = datetime.strptime(dob_raw, "%B %d, %Y").strftime("%d-%m-%Y")
            except Exception:
                dob = dob_raw
        
        
        details = p.text.strip()
        details = re.sub(r'ALIASES:.*?(?=DOB:)', '', details, flags=re.DOTALL)
        details = re.sub(r'DOB:.*?(?=NATIONALITY:)', '', details, flags=re.DOTALL)
        details = re.sub(r'NATIONALITY:.*?(?=CITIZENSHIP:)', '', details, flags=re.DOTALL)
        details = re.sub(r'CITIZENSHIP:.*?(?=HEIGHT:)', '', details, flags=re.DOTALL)
        details = re.sub(r'\r?\n+', '; ', details).strip()
        if nationality == citizenship:
            citizenship = ""
        else:
            nationality= ""
            
        siblings = driver.find_elements(By.XPATH,"//h2[normalize-space()='CASE SUMMARY']/following-sibling::*")
        # check2 = driver.find_elements(By.XPATH,"//h2[normalize-space()='CASE SUMMARY']/following-sibling::p[following-sibling::h2]")
        # check3 = driver.find_elements(By.XPATH,"//h2[normalize-space()='CASE SUMMARY']/following-sibling::p[following-sibling::h3]")
        case_summary = []

        for element in siblings:
            tag = element.tag_name.lower()

            if tag in ["h2", "h3"]:
                break

            if tag == "p":
                # case_summary.append(element.text)
                if element.find_elements(By.TAG_NAME, "a"):
                    continue

                case_summary.append(element.text.strip())
        summary = " ".join(case_summary)
        driver.back()
        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'wanted-card')]")))

        items.append({
            "FULL_NAME": criminal_name,
            "CATEGORY": "P",
            "F_NAME": "",
            "M_NAME": "",
            "L_NAME": "",
            "GENDER": "",
            "DOB": dob,
            "ADD_CITY": "",
            "ADD_COUNTRY": "",
            "State": "",
            "Nationalities": nationality,
            "ADDRESS": "",
            "Identity Number": "",
            "Identity Type": "",
            "REF_DATE": "",
            "DETAILS": details,
            "WEB_LINK": url,
            "VIOLATION_ID": "",
            "SOURCE": "US SS wntd Fugitives",
            "Alias": aliases,
            "Associates": "",
            "MainActivity": "",
            "Citizenship information": citizenship,
            "STATUS": f"Case Summary-{summary}",
            "Rem1": "",
            "Rem2": "",
            "Rem3": "",
            "Remarks": ""
        })
        logging.info(f"Successfully scraped: {criminal_name}")

    return items


if __name__ == "__main__":
    output_file = "US_Secret_Service_Most_Wanted.xlsx"

if os.path.exists(output_file):
    existing_df = pd.read_excel(output_file, dtype=str).fillna("")
    existing_urls = set(existing_df["WEB_LINK"])
else:
    existing_df = pd.DataFrame()
    existing_urls = set()

items =  us_secret_service(existing_urls)
new_df = pd.DataFrame(items)

if not existing_df.empty:
        final_df = pd.concat([existing_df, new_df], ignore_index=True)
else:
        final_df = new_df

final_df.drop_duplicates(subset=["WEB_LINK"], keep="first", inplace=True)
final_df.to_excel(output_file, index=False)

logging.info(f"Excel saved successfully: {output_file}")
