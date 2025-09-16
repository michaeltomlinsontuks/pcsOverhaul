# Directory and File Changes for Macro Compatibility

This list details the required file and directory renaming/moving actions to ensure the old system's hardcoded VBA macros function without errors. Only changes necessary for macro compatibility are included.

---

## Main Directory

_Interface.xls -> Interface.xls
Operation.xls -> Operations.xls  (if macros expect only one name, keep both if unsure)

---

## Archive Directory
No changes required if macros reference by number (e.g., 284446.xls).

---

## Contracts Directory
- Remove spaces and unify naming for contract files if macros expect specific patterns.
- Example: 'ACT 11 4 SPRING WASHERS.xls' -> 'ACT_11_4_SPRING_WASHERS.xls'
- 'Copy of JA 1069432-0194  J33-1008.xls' -> 'JA_1069432-0194_J33-1008.xls'

---

## Customers Directory
- Ensure all customer files match the expected naming convention in macros (e.g., no spaces, consistent case).
- Example: 'ACRON ENGINEERING.xls' -> 'ACRON_ENGINEERING.xls'

---

## Enquiries Directory
- Ensure all enquiry files are named as numbers only if referenced by number (e.g., 384775.xls).
- If macros expect a prefix, rename accordingly (e.g., 'enquiry_384775.xls').

---

## Job Templates Directory
- Standardize template file names if macros expect specific names (e.g., 'A Gen Job Card.xls' -> 'A_Gen_Job_Card.xls').

---

## Templates Directory
- Ensure all template files match macro expectations (e.g., 'Price List.xls' -> 'Price_List.xls').

---

## VBA Directory
- Ensure every .xls.vba file matches its parent .xls file name exactly (except for the .vba extension).
- Example: 'Job_Number.xls.vba' should match 'Job_Number.xls'.
- Remove duplicate or backup files (e.g., 'Search - Copy.xls.vba').

---

## WIP Directory
- Ensure all WIP files are named as numbers only if referenced by number (e.g., 29784.xls).

---

## Quotes Directory
- Ensure all quote files are named as numbers only if referenced by number (e.g., 284139.xls).

---

## General
- Remove spaces and special characters from all file and directory names if macros are case/space sensitive.
- Use underscores or camel case as needed to match macro references.
- Keep both singular and plural forms of files if unsure which is referenced (e.g., Operation.xls and Operations.xls).

---

**Note:** Only rename/move files if you encounter macro errors referencing missing files. Always back up before making changes.

