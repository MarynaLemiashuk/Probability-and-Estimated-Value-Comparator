# Probability-and-Estimated-Value-Comparator

What the Script Does
Loads two Excel files:

0_PPO_Report 7-16-2025 4-21-23 PM Python.xlsx (CRM data)

Power BI Opp Python.xlsx (Power BI data)

Filters Power BI data to include only rows where ITIS_FLAG == 1.

Renames columns in the CRM data for consistent matching.

Merges the datasets on the opportunity name (name).

Compares the sey_probability and estimatedvalue_eur fields between the two datasets.

Marks differences by showing CRM values when they differ from Power BI; leaves blank if values match.

Checks if opportunity names in Power BI exist in CRM, flags missing ones.

Saves the comparison result to result_comparison.xlsx.

Requirements
Install required packages via pip:
python compare_probabilities.py

How to Use
Place both Excel files in the same directory as the script.

Run the script with Python:
python compare_probabilities.py

Find the results in result_comparison.xlsx.

Output Columns
| Column               | Description                                                        |
| -------------------- | ------------------------------------------------------------------ |
| `name`               | Opportunity name                                                   |
| `Probability`        | CRM probability value if it differs from Power BI; otherwise blank |
| `estimatedvalue_eur` | CRM estimated value if it differs from Power BI; otherwise blank   |
| `Right name`         | Shows opportunity name if missing in CRM dataset; otherwise blank  |
