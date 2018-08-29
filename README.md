# Price Update Script
Updates items' prices in excel with the [python script](updatePriceScript.py).  
2 logs are produced in addition to the excel output.  
(1) The [update log](update_log.txt) which shows the items that are updated   
(2) The [error log](error_log.txt) which shows the unexpected errors such as items not found, or unexpected files.  

The script uses xlrd and xlutils to parse and write to excel spreadsheets.
The source excel spreadsheets are not included.

#### Weak ID matching:  
TRUE --> matches IDs that begin in the same sequence  
eg. 3051 can match to 3051o  
Turn off weak ID matching if exact matching is desired (i.e. 3051o must match with 3051o)
