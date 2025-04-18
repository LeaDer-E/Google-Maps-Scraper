# Google Maps Scraper

This repository contains two Python scripts for scraping place details from Google Maps:
1. **`G-Maps_Scrapper.py`**: Scrapes data for a specific place in a single area.
2. **`G-Maps_Multiply_Scrapper.py`**: Scrapes data for a specific place across multiple areas listed in `Areas.txt`.

# Important Note:
* When Code Told You `Found 244 places` and Extract in Excel Sheet only `122 Place` That Mean It's Found Only `122 Place` From G-Maps but It's Read `244 XPATH`.
* When You Found `Error extracting details for element 16` That Mean Ther's No XPATH Clickable Here and Ther's No Place Found.
* You Can Change `chrome_options.headless` To `False` in The Code to be `chrome_options.headless = False` and You Will Know That The Script Extract All The Places Found From G-Maps.
* So Any `Error extracting details for element` Don't Care About It.
* Any Duplicated Value Founds In Terminal It Will Not Extract At Excel Sheet. **(Fixed)**

## Features
* Scrapes business details such as name, description, address, phone number, and website.
* Supports both single-area and multiple-area searches.
* Saves extracted data into well-formatted Excel files.
* Includes automatic handling of duplicates.
* Customizable search parameters.
* Error logging to help debug issues.

## Prerequisites
- Python 3.x
- Google Chrome installed
- undetected-chromedriver
- Chrome WebDriver compatible with your Chrome version
- selenium
- pandas
- openpyxl
- colorama


## Installation
* Clone this repository:
  ```bash
  git clone https://github.com/your-repo/google-maps-scraper.git
  cd google-maps-scraper
  ```
* Install dependencies:
  ```bash
  pip install selenium pandas openpyxl colorama undetected-chromedriver
  ```


## Usage
### 1. `G-Maps_Scrapper.py`
**Purpose**: Scrapes place details for a specific place in a single area.

**How to run**:
```bash
python G-Maps_Scrapper.py
```

**Input**:
- Enter the place to search for (e.g., "Store").
- Enter the area to search in (e.g., "Saudi Arabia").

**Output**: An Excel file named `<place> in <area> at <timestamp>.xlsx` with formatted data.

### 2. `G-Maps_Multiply_Scrapper.py`
**Purpose**: Scrapes place details for a specific place across multiple areas listed in `Areas.txt`.

**How to run**:
```bash
python G-Maps_Multiply_Scrapper.py
```

**Input**:
- Enter the place to search for (e.g., "Store").
- The areas are read from `Areas.txt` (one area per line, e.g., "Saudi Arabia", "Egypt", "Libya").

**Output**: A single Excel file named `<place> in Multiple Areas at <timestamp>.xlsx` with data from all areas.

## Excel Formatting
- **Background**: Black for all cells.
- **Text Color**:
  - White for most cells.
  - Orange for the first row (headers).
  - Dark yellow for column D (from row 2 onwards).
- **Borders**: Dark gray borders around all cells.

## Common Issues and Solutions
### Why Errors?
You might see errors like:
```text
Error extracting details for element 1
Failed to load details panel for element 2
Failed to load details panel for element 3
```
**Explanation**: These errors occur because Google Maps' structure can vary by country or region. The script attempts to handle different XPaths to accommodate these differences.

**Impact**: These errors are non-critical and are included to ensure no place is skipped. They do not affect the overall functionality.

### Why Duplicated Items? **(It Should Be Fixed)**
You might notice duplicated entries in the console output:
```text
────────────────────────────────────────────────────────────
Extracted: Place Name
Description: #####
Address: ##############################
Phone Number: +20 ### ### #####
Website: https://www.#########.com
────────────────────────────────────────────────────────────
Extracted: Place Name
Description: #####
Address: ##############################
Phone Number: +20 ### ### #####
Website: https://www.#########.com
```

**Explanation**: Duplicates can occur due to how Google Maps lists places, especially across different regions. This is intentional to ensure no data is missed.

**Solution**: Duplicates are automatically removed in the final Excel file using `pandas.drop_duplicates()`, ensuring the output is clean and unique.


## Logging & Debugging
* The script logs errors like missing data or XPath issues.
  Common errors include:
  ```
  Error extracting details for element 1
  Failed to load details panel for element 2
  ```

  ```
  Error extracting details for element 231
  Extracted: #############
  Description: مقهى 
  Address: ################################3
  Phone Number: 011 ############
  Website: https://www.facebook.com/####################
  ────────────────────────────────────────────────────────────
  Error extracting details for element 233
  Extracted: #############
  Description: مقهى 
  Address: ################################3
  Phone Number: 011 ############
  Website: https://www.facebook.com//####################
  ────────────────────────────────────────────────────────────
  Error extracting details for element 235
  Extracted: #############
  Description: مقهى 
  Address: ################################3
  Phone Number: 011 ############
  Website: https://www.facebook.com//####################
  ────────────────────────────────────────────────────────────
  Error extracting details for element 237
  Extracted: #############
  Description: مقهى 
  Address: ################################3
  Phone Number: 011 ############
  Website: https://www.facebook.com//####################
  ────────────────────────────────────────────────────────────
  Error extracting details for element 239
  Extracted: #############
  Description: مقهى 
  Address: ################################3
  Phone Number: 011 ############
  Website: https://www.facebook.com//####################
  ────────────────────────────────────────────────────────────
  Error extracting details for element 241
  Error extracting details for element 242
  ```
* These errors are non-critical and occur due to Google Maps' structure changes.
* The script logs duplicates but removes them in the final output.
* you have to wait the code untill it's finished to get the stored data.


## Notes
- **First Script (`G-Maps_Scrapper.py`)**: Prompts the user for both the place and area, then extracts data for that specific area.
- **Second Script (`G-Maps_Multiply_Scrapper.py`)**: Prompts only for the place, reads areas from `Areas.txt` (e.g., "Saudi Arabia", "Egypt", "Libya"), searches each area sequentially, and combines results into one Excel file.
- **Excel Output**: Both scripts format the Excel file with a black background, white/orange/dark yellow text, and dark gray borders. The second script adds an "Area" column.
- **README.md**: Fully in English, it explains usage, formatting, and addresses your specific questions about errors and duplicates.

## License
* This project is licensed under the MIT License - see the LICENSE file for details.


# **Merge and Format Excel Sheets**  

This script processes multiple Excel files from a specified folder, merges their data, removes duplicates, and applies professional formatting for better readability.  

## **Features**  

* Reads all Excel files in the `ready/` folder.  
* Extracts only columns **A to F** from each file.  
* Cleans data by removing extra spaces and hidden characters.  
* Removes rows where both **Phone Number** and **Website** are empty.  
* Deduplicates records based on **Phone Number** and **Website** (keeping only the first occurrence).  
* Sorts data for better organization.  
* Formats the final Excel sheet with:  
   - **Black background** for all cells.  
   - **White, orange, and dark yellow fonts** based on content.  
   - **Column width adjustments** and **row height styling**.  
   - **Centered text alignment** for better readability.  
   - **Frozen panes** at cell B2 for easier navigation.  

## **Usage**  

1. Place all Excel files inside the `ready/` folder.  
2. Run the script to generate a cleaned, merged, and formatted Excel file.  
3. The final file will be saved as `merged.xlsx` in the same directory as the script.  

## **Code Reference**  

```python
folder_path = "ready/"  # Change this if needed
output_file = "merged.xlsx"
merge_and_filter(folder_path, output_file)
```

Run the script and let it handle the rest! 🚀


