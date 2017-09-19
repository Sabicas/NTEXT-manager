# NTEXT-manager
This script was built to handle an issue with NTEXT fields exceeding the character limit allowed in an .xlsx cell. The script reads a directory of CSV files, creates a spreadsheet with each tab representing a file, and creates separate directories/files for NTEXT files that exceeded the aforementioned limit. The file is referenced and highlighted in the resulting .xlsx

This was built to be command-line driven from an SSIS package that creates a CSV repository. To run the script and example directory/files from cmd:

node ntextManager.js genericclient

I used the Complete Works of Shakespeare (5 million chars+) in some of the csv fields for the demonstration.  
