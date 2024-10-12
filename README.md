# Seveso Estimation

## What is Seveso / ICPE?
Seveso and ICPE are two European laws that regulate the use of chemicals in chemical industries or research laboratories.  
The full text of the law can be found here:  
[https://www.legifrance.gouv.fr/jorf/id/JORFTEXT000026306231](https://www.legifrance.gouv.fr/jorf/id/JORFTEXT000026306231)

## Implications
Every five years, regulatory control ensures that the law is respected. To guarantee compliance with the law, several key points must be checked:

- For each compound stored on-site, a document detailing each specific danger related to the chemical must be prepared.
    - **Chemical-specific dangers** are specified using the HXXX phrases given by the **Safety Data Sheets (SDS)**.
    - Chemicals must be classified into three danger types:
      1. Physical dangers
      2. Environmental dangers
      3. Toxic dangers  
        These classifications depend on the nature of the hazard.

- For each compound, the total amount of chemicals must be compared to specific thresholds provided by the Seveso and ICPE regulations.
    - The thresholds may vary depending on the physical state of the compounds.

- Our research center possesses around **40,000 different compounds**, representing approximately 9 tons of chemicals.

Due to this, the Python script automates the classification of each compound by cleaning and parsing our entire database.

## Implementation
The script processes three Excel sheets. (Direct data fetching from databases or APIs could be implemented by adjusting the data loading process.)

- **VC+ Extraction**: Columns used:  
    `["PRODUCT_NAME", "STR_ID", "STOCK_AMOUNT", "STOCK_AMOUNT_UNIT", "CAS_NO", "STORAGE_CLASS", "SAFETY_PHRASES", "PHYSICAL_STATE"]`
    
- **Named Substances Table**: Contains all named substances, their classifications (a, b, c), and corresponding CAS numbers.
    - This table was prepared by extracting data from European legislation.

- **Threshold Table**: Contains all Seveso and ICPE thresholds applicable to the compounds detected in the VC+ extraction, along with their classification.
    - This table was also extracted and prepared from European legislation.
    - In the future, automation of extracting inner HTML text from documents using RegEx expressions and/or an HTML parser (e.g., BeautifulSoup) could be implemented to improve Seveso estimation.

The Jupyter Notebook performs several operations, including data cleaning, parsing, and homogenization, followed by various joining methods to match each type of reagent with its correct classification according to the legislation. All algorithms based on the decision trees outlined in the legislation have been implemented.

Finally, the script generates a new Excel file where critical chemicals are listed in a separate sheet, all fractions are calculated, and each compound is detailed line by line.
