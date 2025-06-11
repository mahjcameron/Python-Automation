# Project Python (Excel) Automation

## Business Need
As part of a recurring monthly HR access management process, there was a need to efficiently process and transform Excel documents containing staff access records. 
Each Excel file included updates for employees who experienced changes to their access status — including increases, reductions, removals, or other record updates.
Pre-processing this excel document take considerable amount of time and with the amount of steps is prone to human error, in which this project would minimize both issues.

## Technical Need
1. The raw data required manual cleaning and reformatting to make it readable and suitable for further processing. This included:
 - Removing unnecessary columns
 - Adding new fields (e.g., “Notes/Action”, “Access”)
 - Reordering and renaming columns for clarity and consistency
2. Additionally, specific columns needed to be extracted and formatted to generate dynamic SQL queries used to retrieve current access permissions from a database.
3. The output of these SQL queries was then manually copied into the Excel workbook to run VLOOKUP formulas, allowing for comparison between the updated position access requirements and existing access levels.
