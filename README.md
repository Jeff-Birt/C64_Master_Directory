# C64_Master_Directory
Master Directory3 - Converts CSV file output from Directory Master into an Excel file. The excel file included macros to check for duplciate directories and duplciate files.

First use DirMaster ***Link to export a CSV of your C64 disk images. 

1) Go to Disk->Options and check the 'Output MD5 has in CSV'.

2) Go to Disk->Batch Processing, select parent folder, check 'Process Sub Folder', select 'Export Directories' from drop down, select 'Comma Seperated' from 'Save Directory as' drop down box, check 'Include Sub Directory' and 'Save as Single File'.

This will generate a CSV file that contains a dump of all your image paths, dirctory contents and a MD5 hash of each file. 

3) Run MasterDirectory3.py. Click the 'Source File' button to select the CSV file to process. Then click the 'Convert' button to convert it to an Excel file.

The generated Excel file contains macros so you will likely see the 'Eable Content' button pop up. Click this to enable macros. The first shee is a 'Master_Index'. This sheet will show you the name and path of each image, a link to the indivisual sheet for that image, and a MD5 hash of the directory hash (hash of directory text only). This hash serves to help locate duplicate directories, i.e. having two images of same disk. Duplicate hashes as color coded in yellow. If you click the 'Duplicates' button only the rows with matches will be shown.

Each disk image is given its own sheet. This sheet contains the directory text, a MD5 hash of the contents of each file and a Column called 'File#'. The 'File#' column helps to identify which directory entry row a fiel is on. When you click the 'Duplciates' button all the MD5 hash values on the current sheet will be compared to all other sheets. The 'F' colum will be populated with the results. The 'G' column will be populated with a button for each row. Click the button on a row to pop up a text box showing the matches, if any. The format is 'Image#.File#', where 'Image#' is the sheet name and 'File#' is the number in the 'E' column.

The 'Master_Index' page 'C' colum contains a hyper link to each disk image sheet and the top-left cell in each image sheet contains a link back to teh 'Master_Index'. 

The macros and links may or may not work without modification on Open Office, Libre Office as their VBA macro support is not 100%.