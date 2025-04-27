# Checkbook Analysis 
This repository is home to the python script that will compare two check registers (i.e a personal register to an export from the bank).  

## Description 
This script will compare two checkbook registers: 1) A personal register and 2) one exported from the bank.  The script will work to identify where differences lie between the two registers.  <br>

**For either register, we cannot have headers. Also, delete extra columns or rows of data.  For example, if there's a note at the bottom of the personal register, these additional rows at the bottom will break the script**

The location (column positions) is hardcoded.  It is assumed that, for the personal register, the format is <br>
**ITEM NUMBER | TYPE | MONTH | DAY | YEAR | DESCRIPTION | CLEARED | AMOUNT | BALANCE** <br>
Therefore, for the personal register, the amount column will be column number eight (the opnxyl module starts at one, not zero)

while the information exported from the bank is assumed to be in the following format: <br>
**DATE | TYPE | DESCRIPTION | STATUS | AMOUNT** <br>
Therefore, for the bank's export, the amount column will be column number five  (the opnxyl module starts at one, not zero)

**For both, the headers shall be deleted**


## Running
Be sure that each register excel workbook contains only one sheet

The format of the information exported from the bank will be in *.csv format -- this is to be expected, and the script will handle this.

For the personal register, delete rows containing information from previous transactions. For example, let's say we want to start at Sep 2024 which is on Excel row 1400.  This means that Excel rows 1 through 1399 shall be deleted.  


## Revisions
v0.0.1 -- Initial version. 