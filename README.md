# AIP Subscription Processor

## Overview

The **AIP Subscription Processor** is a Python script designed to process subscription data from Excel files, categorize products into Backfiles, Subscriptions, and Packages, and generate a consolidated subscription period range for each entry. It reads data from two Excel files (`AIP_Master_List.xlsx` and `AIP_Client_Pivot.xlsx`), processes product information, and outputs the results to a new Excel file (`ordered_table.xlsx`).

This tool is particularly useful for managing and analyzing subscription-based product data, ensuring that subscription periods are normalized and presented in a readable format (e.g., "2005~2010").

## Features

-   **Product Categorization**: Classifies products into Backfiles, Subscriptions, and Packages based on predefined rules and names.
-   **Text Normalization**: Standardizes product names and titles by converting full-width characters to half-width, removing extra spaces, and applying title case with exceptions.
-   **Year Range Processing**: Merges and filters subscription years into continuous ranges (e.g., "2004~2006, 2010~2013") within specified start and end dates.
-   **Data Cleaning**: Removes rows where the subscription period is empty and reorders columns for better readability.
-   **Output Generation**: Saves the processed data to an Excel file with a user-friendly structure.

## Prerequisites

-   **Python 3.6+**
-   **Required Libraries**:
    -   `pandas`
    -   `numpy`
    -   `re` (built-in)
    -   `unicodedata` (built-in)

Install the required libraries using pip:

```bash
pip install pandas numpy
```

## Installation

Clone or download this repository to your local machine.

Ensure the input Excel files (AIP_Master_List.xlsx and AIP_Client_Pivot.xlsx) are placed in the same directory as the script (`betterSolution.py`).

## Usage

### Prepare Input Files:

* `AIP_Master_List.xlsx`: Contains the master list of journal subscriptions with columns like "Title", "Subject Collection", "Online date, start", "Online date, End", etc.

* `AIP_Client_Pivot.xlsx`: Contains client subscription data with columns like "Product Name" and "Sub Year".

### Run the Script:

Open a terminal in the project directory and execute:

```
python betterSolution.py
```

### Output:

The script generates `ordered_table.xlsx` in the same directory, containing the processed data with subscription periods in the fifth column.

## Script Workflow

### Data Loading:

* Reads `AIP_Master_List.xlsx` into `master` DataFrame and `AIP_Client_Pivot.xlsx` into `client` DataFrame.

* Drops rows with any missing values in `client`.

### Text Processing:

* Normalizes text (e.g., "and" to "&", removes "eJournal") and applies title case with exceptions (e.g., "to", "and" remain lowercase unless at the start).

### Product Classification:

* Identifies Backfiles (names containing "backfile" or "backfiles"), Packages (predefined list `p`), and Subscriptions (remaining products).

* Creates dictionaries of product objects (Backfile, Subscription, Package) with associated years.

### Data Enrichment:

* Updates `master` with Backfiles, Single Subscriptions, and Package periods based on product dictionaries.

* Merges all subscription periods into continuous ranges, constrained by "Online date, start" and "Online date, End".

### Final Processing:

* Filters out rows where "Subscribe Period" is empty.

* Moves "Subscribe Period" to the fifth column.

* Saves the result to `ordered_table.xlsx`.

## File Structure

```
AIP_Subscription_Processor/
│
├── betterSolution.py         # Main script
├── AIP_Master_List.xlsx      # Input: Master product list
├── AIP_Client_Pivot.xlsx     # Input: Client subscription data
├── ordered_table.xlsx        # Output: Processed subscription data
└── README.md                 # This file

```

## Example

### Input (`AIP_Master_List.xlsx` excerpt):

| Title | Subject Collection | Online date, start | Online date, End | product a |
| ----- | ------------------ | ------------------ | ---------------- | --------- |
| Journal of Mgmt | Management | 2005 | 2020 | \* |
| Tech Review | Engineering | 2018 | 2022 | \[2018, 2019] |

### Input (`AIP_Client_Pivot.xlsx` excerpt):

| Product Name | Sub Year |
| ------------ | -------- |
| Management Backfiles | 2006 |
| Management Core | 2010 |

### Output (`ordered_table.xlsx` excerpt):

| Title | Subject Collection | Single Subscribe | Backfiles | Subscribe Period | product a | ... |
| ----- | ------------------ | ---------------- | --------- | ---------------- | --------- | --- |
| Journal of Mgmt | Management | \[\] | @Management Backfiles | 2005\~2010 | \[2010\] | ... |
| Tech Review | Engineering | \[\] | | 2018\~2019 | \[2018, 2019\] | ... |

### Column Descriptions:

* **Title**: Name of the journal.

* **Subject Collection**: The academic field or category of the journal.

* **Single Subscribe**: Subscription years for individual subscriptions.

* **Backfiles**: Backfile product package with a prefixed "@" symbol (e.g., "@Management Backfiles").

* **Subscribe Period**: Consolidated year ranges (e.g., "2005\~2010").

* **product a**: One of the package products with its subscription years.

## Notes

* **Backfiles**: Automatically assigned fixed ranges (0-2006 or 1994-2013 for "Additions").

* **Error Handling**: Ensure input files exist and contain required columns ("Title", "Subject Collection", "Online date, start", "Online date, End", etc.) to avoid runtime errors.

* **Customization**: Modify the `p` list in the script to adjust Package product names.

## Contributing

Feel free to submit issues or pull requests if you have suggestions
