# Seveso estimation 

https://www.legifrance.gouv.fr/jorf/id/JORFTEXT000026306231

This script help the redaction of the Seveso / ICPE report to estimate the danger of all products in the reasearch center.

It works by processing 3 excel sheets. (Direct data fetching from databases or API could be implemented by adjusting the loading data process.)
- VC+ extraction: column used : ["PRODUCT_NAME", "STR_ID", "STOCK_AMOUNT","STOCK_AMOUNT_UNIT", "CAS_NO", "STORAGE_CLASS", "SAFETY_PHRASES", "PHYSICAL_STATE"]
- A table containing all named_substances, their classification (a, b, c), and all cas numbers
    - This table was prepared  by extracting datas from the european legislation
- A table containing all thresholds (seveso and ICPE) detected in the VC+ extraction, and on wich classification they apply
    - This table was also extracted and prepared  from  the european  legisltation.
    - Automation of extracting Inner Html text for the documents by using  RegEx expression and / or HTML parser (beautiful soup?) may be implemented  for future Seveso estimation.

 This Jupiter Notebook contains multiple operations of cleaning, parsing, homogenise data, followed by all the joining method to make each type of reagent correspond to the correct classification according to the legislation. All algorigrams from the decisionnal trees exemplified in the legislation were implemented. 

 Finally, The script generate a new excel file, were the critical chemicals are written in a separate sheet, all fractions calculated, and all compounds detailled line by line.
