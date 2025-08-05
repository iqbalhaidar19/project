# Sales Data Analysis and Cleaning Project

## Brief Description

This project demonstrates the ability to clean, process, and analyze data using Python with the Pandas library. The script takes messy fictitious sales data in CSV format, systematically cleans it, enriches it with category data, performs some analysis, and generates a clean and professional multi-page Excel report, complete with summaries and graphs.

---

## Main Features

- **Comprehensive Data Cleaning:** Handles a variety of raw data issues such as duplicate data, inconsistent formats (date, invoice number), typos, capitalization, and missing values.
- **Data Enrichment:** Combines transaction data with product master data to add category information, demonstrating data `merge` capability.
- **Business Analysis:** Apply business logic to calculate new columns such as “Total Price” and “Commission”.
- **Report Automation:** Automatically generates `.xlsx` report files containing:
  - Sheet 1: Clean and complete data.
  - Sheet 2: Sales summary per category with visual graphs.
- **Professional Presentation:** Apply Rupiah format (Rp#,##0) to the currency column in the output Excel file for easy reading.

---

## Project Folder Structure

.
├── .gitignore                          # File to ignore unnecessary files/folders
├── README.md                           # File you are currently reading this
├── data_processing_script.py           # Python main script
├── product_categories.csv              # Product category master data
├── raw_sales_data.csv                  # raw sales data
└── requirements.txt                    # List of required Python libraries

---

## How to use

Here's how to run this project on your local computer.

### Prerequisites

- Python 3.8 or later
- Git

### Installation

1.  **Clone this repository:**
 ```bash
 git clone https://github.com/iqbalhaidar19/project.git
 cd project
 ```

2.  **Create and activate a virtual environment:**
 ```bash
    # Create venv
 python -m venv venv

    # Activate on Windows
.\venv\Scripts\activate

    # Activate on macOS/Linux
 source venv/bin/activate
 ```

3.  **Install all required libraries:**
```bash
 pip install -r requirements.txt
 ```

### Running the Script

After the installation is complete, run the main script from the terminal:
```bash
python data_processing_script.py


After the script finishes running, a new file called Sales_Report.xlsx will be created in the project folder. This file contains the final results of the data processing.
