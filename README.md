# DOI-Resolver
This PowerShell script checks DOIs exported in CSV format from the
Talis Aspire Reading List system.  It was made by Jai Parker, Information
Access Librarian at the Queensland University of Technology with help from
Microsoft Copilot.  As per the [license](./LICENSE), caveat emptor.
Please note that the output of this script has approximately a 50% success rate and is guaranteed to to not resolve target DOIs for Taylor & Francis.

## Prerequisites
PowerShell requires an execution policy set to run this script. To set the Execution Policy:
1. Click the Windows start button and search for "Windows PowerShell" then select **Run as administrator**
2. Type `Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser` and press Enter
3. When asked "Do you want to change the execution policy?" press Y

## Running the script
To run this script:

1. Download a csv of items for checking from Talis. The CSV file name format
   must start with all_list_items_ for the script to pick them up.
2. Open the CSV file and do a cleanup of the DOIs in Column O. The script
   automatically prepends https://doi.org/ to anything starting with 10. in
   this field so all the duplicated DOIs, mostly delimited by a ; need to be
   removed first.

### Input

URLs are fed into the script via a CSV file. The files must be named to match the
pattern `all_list_items_*.csv` (for example `all_list_items_2025.csv`).  Only one file with this naming pattern should be in the folder with the script.

### Output
Once completed a file named doi_resolutions_[yyyyMMdd_HHmm].csv will appear in the same folder and the script and all_list_items*.csv file.
The output contains the following rows from the all_list_items*.csv file plus a row named 'ResolvedURL'
- Title
- Chapter/Article Title
- Item Link
- List Link
- DOI
