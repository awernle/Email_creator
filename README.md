# Email_creator
7/7/2021
#Purpose:
The purpose of this script is to automatically generate emails (as word docs) for SSSP by pulling superintendent and school information from preexisting excel files. There is an initial version (V1) for the SSSP introduction letter, and the final version (V2) for the SSSP project update letter. V2 is slightly more optimized, it features a table generator and prints all necessary contact information at the bottom of the document.

#Setup:
In order to run this script, you must have it within the same directory as your preexisting excel files. Preexisting excel files should be up to date and organized so that the script can pull from them without errors.

#Output:
A word document named after the school district the email will be sent to, for each school district. Email text will be printed within the word documents. 

#Potential Errors:
Bad data entry in spreadsheets leads to faulty output from script. Calling incorrect names of spreadsheets or dataframes. Variable may not be correct type, I sometimes change from Pandas dataframe to list and back. Incorrect formatting of text section. See for more details on using docx module: https://python-docx.readthedocs.io/en/latest/ 

#Potential Improvements:
The script may not need to pull from 2-3 separate spreadsheets if the necessary information is optimized within one spreadsheet.
With the correct permissions, emails may be sent automatically with no user interface, however this may not be possible at the DNR due to IT policies.


