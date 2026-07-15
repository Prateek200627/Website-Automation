# CBI Announced Rewards Scraper

## Overview

This Python automation script scrapes the **CBI (Central Bureau of Investigation) Announced Rewards** page and extracts detailed information about each wanted individual. The extracted data is cleaned, standardized, and stored in an Excel file while preventing duplicate records during future executions.

---

# Workflow

## 1. Launch the Website

- Opens the CBI Announced Rewards webpage using Selenium Chrome WebDriver.
- Waits until all reward cards are loaded.
- Collects the URL of every reward announcement.

---

## 2. Duplicate Record Detection

Before scraping begins, the script checks whether an Excel file (`CBI_Announced_Rewards.xlsx`) already exists.

If the file exists:

- Reads all previously scraped records.
- Stores existing `WEB_LINK` values.
- Skips records that have already been scraped.

This ensures that only newly added records are processed.

---

## 3. Visit Every Record

For every new reward announcement:

- Opens the record in a new browser tab.
- Waits until the page is fully loaded.
- Extracts all information displayed on the page.

The script dynamically reads every information box instead of using fixed XPaths, making it more robust against layout changes.

---

## 4. Data Cleaning

Several helper functions are used to clean and standardize the extracted information.

### Date Formatting

The script converts different date formats into a standard format:

```
DD/MM/YYYY
```

Invalid or unavailable dates such as:

- NA
- N/A
- Not Available
- -

are stored as blank values.

---

### Age Calculation

After formatting the Date of Birth, the script calculates the person's current age using today's date.

Example:

```
DOB : 17/05/1951

Age : 74 Years
```

The calculated age is appended to the Details column.

---

### Name & Alias Extraction

Some names contain aliases separated using the `@` symbol.

Example:

```
Rahul Sharma @ Sonu @ Rocky
```

is converted into

**Full Name**

```
Rahul Sharma
```

**Alias**

```
Sonu; Rocky
```

---

### Father Name Cleaning

Unavailable values such as

- NA
- N/A
- Unknown

are converted into blank values.

If multiple names are separated using `@`, they are converted into semicolon-separated values.

Example:

```
Ram @ Shyam
```

becomes

```
Ram; Shyam
```

---

### Address Cleaning

Invalid addresses are removed.

The script also extracts the Indian State from the address.

Example

```
Village XYZ,
Begusarai,
Bihar
```

returns

```
State = Bihar
```

If only a city is present, a predefined city-to-state mapping is used.

---

### Nationality Detection

If the address belongs to Singapore,

```
Nationality = Blank
Country = Blank
```

Otherwise,

```
Nationality = Indian
Country = India
```

---

## 5. Record Preparation

Each record is converted into a structured format containing fields such as

- Full Name
- Gender
- Date of Birth
- Address
- State
- Nationality
- Reward Details
- Alias
- Source
- Web Link
- Remarks

Additional information like Father's Name is stored inside the Remarks field.

---

## 6. Save the Data

After all records have been processed,

- Creates a Pandas DataFrame.
- Maintains a fixed column order.
- Replaces missing values with blank strings.
- Saves the data into

```
CBI_Announced_Rewards.xlsx
```

---

# Features

- Selenium-based automation
- Dynamic data extraction
- Automatic duplicate detection
- Date standardization
- Automatic age calculation
- Alias separation
- Father name cleaning
- State extraction from address
- Nationality detection
- Structured Excel export
- Future-proof incremental scraping

---

# Technologies Used

- Python
- Selenium
- Pandas
- Regular Expressions (Regex)
- OpenPyXL
- Datetime
- Chrome WebDriver

---

# Output

The generated Excel file contains structured information including:

- FULL_NAME
- GENDER
- DOB
- ADDRESS
- State
- Nationalities
- DETAILS
- WEB_LINK
- Alias
- SOURCE
- Remarks

This file is automatically updated during future executions without duplicating previously scraped records.

---

# Advantages

- No duplicate records
- Automatically detects newly published reward notices
- Standardized output format
- Easy to maintain
- Suitable for scheduled automation
- Handles missing and inconsistent data gracefully
