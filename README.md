#This script help the redaction of the Seveso / ICPE report for all products in stock + laboratories in lyon.

#It works by processing 3 excel sheets. (Direct data fetching from databases or API could be implemented by adjusting the loading data process.)
#   -VC+ extraction: column used : ["PRODUCT_NAME", "STR_ID", "STOCK_AMOUNT","STOCK_AMOUNT_UNIT", "CAS_NO", "STORAGE_CLASS", "SAFETY_PHRASES", "PHYSICAL_STATE"]

#   -A table containing all named_substances, their classification (a, b, c), and all cas numbers
#   -A table containing all thresholds (seveso and ICPE) detected in the VC+ extraction, and on wich classification they apply
