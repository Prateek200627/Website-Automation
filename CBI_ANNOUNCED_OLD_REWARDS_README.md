# CBI Old Records Scraper

## Overview

This Python automation script scrapes the **CBI (Central Bureau of Investigation) Old Records** webpage and extracts information about wanted individuals. The extracted data is cleaned, standardized, and exported to an Excel file. The script also prevents duplicate entries by comparing newly scraped records with the existing Excel data.

---

# Workflow

## 1. Fetch the Website

- Sends an HTTP GET request to the CBI Old Records webpage using the `requests` library.
- Uses a custom User-Agent to mimic a browser request.
- Parses the webpage using **BeautifulSoup**.

---

## 2. Locate All Records

- Searches for every record available on the webpage.
- Each wanted person's information is contained inside an `oldRecordContent` section.
- Counts the total number of available records before processing.

---

## 3. Extract Record Details

For every wanted person's record, the script extracts:

- Name
- Father's Name
- Date of Birth
- Gender
- Charges
- Branch Name
- Address
- PDF Link (if available)

The data is dynamically identified using field labels, making the scraper more reliable if the order of fields changes.

---

## 4. Data Cleaning

Several helper functions are used to clean and standardize the extracted information.

### Name Cleaning

The script removes common prefixes from names.

Example:

```
Mr. Rahul Sharma
```

becomes

```
Rahul Sharma
```

Supported prefixes include:

- Mr.
- Mrs.
- Ms.
- Miss
- Shri
- Sh.
- Smt.
- Sri
- Dr.

---

### Alias Extraction

If the name contains aliases separated using the `@` symbol, the script separates them.

Example:

```
Rahul Sharma @ Sonu @ Rocky
```

becomes

**Full Name**

```
Rahul Sharma
```

**Alias**

```
Sonu; Rocky
```

---

### Father Name Processing

The script intelligently handles different Father Name formats.

#### Alias Handling

Example

```
Ram Kumar @ Ramu @ RK
```

becomes

Father Name

```
Ram Kumar
```

Remarks

```
Father Alias: Ramu; RK
```

#### Multiple Father Names

If multiple names are separated using

- /
- &
- ,
- ;

the Father Name field is left blank, and the complete value is stored in the Remarks column.

---

### Date Formatting

The Date of Birth is converted from

```
YYYY-MM-DD
```

to

```
DD/MM/YYYY
```

Unavailable values such as

- N/A

are stored as blank values.

---

## 5. Prepare Structured Data

The extracted information is converted into a standardized format with fields such as:

- Full Name
- Gender
- Date of Birth
- Details
- Alias
- Source
- Address
- Remarks

Additional information is stored as follows:

- **Rem1** → Father's Name
- **Rem2** → Address
- **Rem3** → Branch Name
- **Remarks** → Father aliases or multiple father names

---

## 6. Duplicate Detection

Before writing data to Excel, the script checks whether the Excel file already exists.

If it exists, previously stored records are loaded.

A unique key is created using:

- FULL_NAME
- DOB
- SOURCE

Every newly scraped record is compared against these keys.

If a matching key already exists:

- The record is skipped.

Otherwise:

- The new record is added to the dataset.

This ensures that running the scraper multiple times never creates duplicate entries.

---

## 7. Save to Excel

The cleaned records are converted into a Pandas DataFrame.

If the Excel file does not exist:

- A new Excel file is created.

If the Excel file already exists:

- Only new records are appended.

At the end of execution, the script displays a summary including:

- Total Records Scraped
- Newly Added Records
- Duplicate Records Skipped
- Total Records Stored in Excel

---

# Features

- BeautifulSoup-based web scraping
- Dynamic field extraction
- Name prefix removal
- Alias extraction
- Father name normalization
- Date standardization
- Duplicate detection
- Incremental Excel updates
- Structured data export
- Handles missing or inconsistent values gracefully

---

# Technologies Used

- Python
- Requests
- BeautifulSoup (bs4)
- Pandas
- Regular Expressions (Regex)
- Datetime
- OpenPyXL

---

# Output

The generated Excel file contains structured information including:

- FULL_NAME
- GENDER
- DOB
- DETAILS
- WEB_LINK
- Alias
- SOURCE
- Rem1 (Father Name)
- Rem2 (Address)
- Rem3 (Branch Name)
- Remarks

The Excel file is automatically updated during future executions without duplicating existing records.

---

# Advantages

- Fast HTTP-based scraping without browser automation
- Prevents duplicate records
- Automatically updates existing datasets
- Produces clean and standardized output
- Easy to maintain and extend
- Robust handling of inconsistent or incomplete data
- Suitable for scheduled or recurring data collection
