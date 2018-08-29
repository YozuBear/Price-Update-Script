# This script updates the prices in order_xxxxx.xlsm with the
# price lists in folder "lists"

# Usage:
# cd xlrd-1.1.0
# python setup.py install
# python updatePriceScript.py

import xlrd
from xlutils.copy import copy
import os
import sys
import datetime
import re

# User Configurations
# TODO: column 5 or 6 update price?
#
# Path to price_update directory
PATH = "/Users/yozubear/Desktop/price_update"

# Weak ID matching: TRUE --> matches IDs that begin in the same sequence
# eg. 3051 can match to 3051o
# Turn off weak ID matching if exact matching is desired
#   (i.e. 3051o must match with 3051o)
WEAK_ID_MATCHING = False

# Global variables
#
# Path to the new price lists (poultry, chinese etc.)
priceListsFolder = PATH + "/lists"

# Regular expression that item ID must match
idPattern = re.compile("^([0-9]{4,6}[A-Z]{0,3})$")
priceFormat = '{0:.2f}'
pricePattern = re.compile("^([0-9]{0,2}\.[0-9]{2})$")

# the name of order spreadsheet
orderFileName = ""

# Create / overwrite the error log file
errorLog = open(PATH + "/error_log.txt", "w")
errorLogNum = 1
print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n\n", file=errorLog)

# Create / overwrite the updated log file
updateLog = open(PATH + "/update_log.txt", "w")
updateLogNum = 1
print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), file=updateLog)
print("WEAK ID MATCHING: " + str(WEAK_ID_MATCHING) + "\n\n", file=updateLog)

# Error function
def error(errorMessage):
    global errorLogNum
    errorMessage = str(errorLogNum) + ". Error: " + errorMessage
    print(errorMessage)
    print(errorMessage, file=errorLog)
    errorLogNum += 1

# Fatal error function, program terminates after fatal error
def fatal_error(errorMessage):
    print("Fatal Error: " + errorMessage)
    print("Fatal Error: " + errorMessage, file=errorLog)
    sys.exit()

# itemID: ID on the item sheet
# oldPrice: old price on the item sheet
# codeID: ID on the price list
# newPrice: new price on the price list
def updatedItemLog(itemID, oldPrice, codeID, newPrice):
    global updateLogNum
    updateMessage = "%d. %s: $%s replaced with %s: $%s" % (updateLogNum, itemID, oldPrice, codeID, newPrice)
    print(updateMessage)
    print(updateMessage, file=updateLog)
    updateLogNum += 1

# Find the items needing updates in the price list (i.e. Poultry list)
# filePath: the file path of the price list to parse
# startRow: the row of the price list to start parsing
# priceCol: the column number of the price list that contains the new price
# orderPriceCol: the column number of the order list that contains the old price
def parsePriceList(filePath, fileName, startRow, priceCol, orderPriceCol):
    sh = xlrd.open_workbook(filePath).sheet_by_index(0)
    for row in range(startRow, sh.nrows):
        # Convert codeID to string format without decimal
        codeID = sh.cell_value(rowx=row, colx=0)
        if isinstance(codeID, float):
            codeID = int(codeID)
        codeID = str(codeID).strip()

        if idPattern.match(codeID):
            newPrice = sh.cell_value(rowx=row, colx=priceCol)

            # try to convert the price to 2 decimal places if it's an int or float
            try:
                newPrice = priceFormat.format(float(newPrice))
            except ValueError:
                pass
            success = updateItemPrice(codeID, newPrice, orderPriceCol, False)  # first round doesn't check weak ID
            if success:
                continue
            else:
                if WEAK_ID_MATCHING:
                    # check again
                    success = updateItemPrice(codeID, newPrice, orderPriceCol, True)
                    if success:
                        continue
            error("Cannot map ID: %s to order's Item from %s, row %d" % (str(codeID), fileName, row+1))

# update the order list's item sheet
# codeID: the ID of item to be updated
# newPrice: the new price of the item
# priceColumn: the column number of the order list that contains the price
# return true if a match is found
def updateItemPrice(codeID, newPrice, priceColumn, weakIDChecking):
    global r_itemSheet
    global w_itemSheet
    global priceFormat
    for row in range(0, r_itemSheet.nrows):

        # Convert itemID to string form without decimal for proper matching
        itemID = r_itemSheet.cell(rowx=row, colx=0).value
        if isinstance(itemID, float):
            itemID = int(itemID)
        itemID = str(itemID).strip()

        # Eliminate the rows that don't have proper item formats
        if not idPattern.match(itemID):
            continue

        # convert old price to string, to 2 decimal places
        oldPrice = r_itemSheet.cell(rowx=row, colx=priceColumn).value
        try:
            oldPrice = priceFormat.format(float(oldPrice))
        except ValueError:
            pass

        # Weakly match the codeID and itemID if weakIDChecking = True
        check = codeID == itemID
        if (not check) and weakIDChecking:
            check = itemID.startswith(codeID) or codeID.startswith(itemID)
        if check:
            if oldPrice != newPrice:
                if pricePattern.match(newPrice):
                    newPrice = float(newPrice)
                w_itemSheet.write(row, priceColumn, newPrice)
                updatedItemLog(itemID, oldPrice, codeID, newPrice)
            return True
    return False


# Find the file name of the order list
for file in os.listdir(PATH):
    if file.startswith("order", 0, 5):
        orderFileName = os.path.join(PATH, file)

if not orderFileName:
    fatal_error("Cannot find the order spreadsheet")


try:
    # Open the order in read and write modes
    rb = xlrd.open_workbook(orderFileName)
    wb = copy(rb)
    sheets = rb.sheets()
    itemIndex = [s.name for s in sheets].index("Item")
    r_itemSheet = rb.sheet_by_index(itemIndex)
    w_itemSheet = wb.get_sheet(itemIndex)

    # Loop through each price list
    for file in os.listdir(priceListsFolder):
        priceListPath = os.path.join(priceListsFolder, file)
        if file.startswith("POULTRY"):
            parsePriceList(priceListPath, "POULTRY", startRow=4, priceCol=8, orderPriceCol=5)
        elif file.startswith("SEAFOOD"):
            # update case rather than lb, in price list and order list
            parsePriceList(priceListPath, "SEAFOOD", startRow=9, priceCol=7, orderPriceCol=6)
        elif file.startswith("CHINESE"):
            parsePriceList(priceListPath, "CHINESE", startRow=5, priceCol=6, orderPriceCol=5)
        elif file.startswith("BUTCHER"):
            parsePriceList(priceListPath, "BUTCHER", startRow=7, priceCol=3, orderPriceCol=5)
        else:
            error("Unrecognized file " + priceListPath)

    # Save the processed file onto a different file
    wb.save(PATH + "/output.xls")

except Exception as e:
    fatal_error(e)

# closing the log files
errorLog.close()

print("\n\nsuccessfully updated", file=updateLog)
updateLog.close()
