# ProcessData
# Written by Andrew Bouchard
# This program provides access to product database information and is written for Medline Industries, Inc.
# Date Started: 7/10/2008
# Last Updated: 8/12/2008

import os
import sys
import string
import cPickle
import datetime
from win32com.client import Dispatch

# Define the dictionary data type
class Item(object):
    
    def __init__(self):
        self.Sales = 0
        self.Orders = 0
        self.B02Status = ""
        self.DirStatus = ""
        self.StockLevel = 0
        self.Description = ""
        self.B02Vendor = ""
        self.DirVendor = ""
        self.Competitors = {}

def main(script):
    
    print "Opening Excel for data import..."
    
    # Get the current path in order to identify the location of files
    pathname = os.path.dirname(sys.argv[0])
    pathprefix = os.path.abspath(pathname)
    
    # Initialize Excel
    xlApp = Dispatch("Excel.Application")
    #xlApp.Visible = 1
    
    print "Excel open."
    
    # Initiatize the data structure
    data = {}
    
    # Store the creation date of this database in the file
    data['Date'] = datetime.datetime.now().strftime("%m/%d/%Y")
    
    print "Opening stock data..."
    
    # Open the item data
    try:
        fullpath = pathprefix + "\stock.xls"
        stock = xlApp.Workbooks.Open(fullpath)
    except:
        fullpath = pathprefix + "\stock.dbf"
        stock = xlApp.Workbooks.Open(fullpath)
    
    print "Stock data open. Importing..."
    
    # Parse through the item file and store the information in the database
    for row in range(2, 65536):
        
        if row%1000==0:
            print "Processing row " + str(row) + "..."
        
        # Get the part number from the first column
        pn = str(stock.Sheets(1).Cells(row, 1))
        
        # If there is no part number, we have reached the end of the list
        if (pn == "None"):
            break
        
        #Don't include part numbers ending in 'ET' or 'FR'
        if pn.endswith('ET') or pn.endswith('FR'):
            continue
        
        #Store the pieces of information for the part
        entry = Item()
        entry.StockLevel = int(stock.Sheets(1).Cells(row, 5).Value)
        entry.Description = str(stock.Sheets(1).Cells(row, 6).Value)
        entry.B02Vendor = str(stock.Sheets(1).Cells(row, 8).Value)
        entry.DirVendor = str(stock.Sheets(1).Cells(row, 9).Value)
        
        # Process the stock status
        entry.DirStatus = str(stock.Sheets(1).Cells(row, 4).Value)
        entry.B02Status = str(stock.Sheets(1).Cells(row, 3).Value)
        
        data[pn] = entry
        
    print "Stock data import complete. Closing file..."
    
    # Close the item data workbook
    stock.Close(SaveChanges=0)
    
    print "Opening cross-reference data..."
    
    # Open the cross reference data
    try:
        fullpath = pathprefix + "\cross.xls"
        cross = xlApp.Workbooks.Open(fullpath)
    except:
        fullpath = pathprefix + "\cross.dbf"
        cross = xlApp.Workbooks.Open(fullpath)
    
    print "Cross-reference data open. Importing..."
    
    # Parse through the cross reference file and store the information in the database
    for row in range(2, 65536):
        
        if row%1000==0:
            print "Processing row " + str(row) + "..."
        
        # Get the Medline part number from the first column
        pn = str(cross.Sheets(1).Cells(row, 1))
        
        # If there is no part number, we have reached the end of the list
        if (pn == "None"):
            break
        
        # Pull the part number's entry from the dictionary and add the vendor name as the value to the 
        #   cross-ref number key in the Competitors dictionary
        vendor = str(cross.Sheets(1).Cells(row, 2))
        cfpn = str(cross.Sheets(1).Cells(row, 3))
        if vendor in data[pn].Competitors.keys():
            data[pn].Competitors[vendor].append(cfpn)
        else:
            data[pn].Competitors[vendor] = [cfpn]
    
    print "Cross-reference data import complete. Closing file..."
    
    # Close the cross reference data workbook and Excel
    cross.Close(SaveChanges=0)
    
    print "Opening sales data..."
    
    # Open the sales data
    fullpath = pathprefix + "\sales.xls"
    sales = xlApp.Workbooks.Open(fullpath)
    
    print "Sales data open. Importing data..."
    
    # Parse through the sales file and store the information in the database
    for row in range(2, 65536):
        
        if row%1000==0:
            print "Processing row " + str(row) + "..."
        
        # Get the Medline part number from the eighteenth column
        pn = str(sales.Sheets(1).Cells(row, 18))
        
        # If there is no part number, we have reached the end of the list
        if (pn == "None"):
            break
        
        # Pull the part number's entry from the dictionary, add the sale amount to the sales and one to the order quantity
        data[pn].Orders = data[pn].Orders+1
        data[pn].Sales = data[pn].Sales+int(sales.Sheets(1).Cells(row, 25).Value)
        
    print "Sales data import complete. Closing file..."
    
    sales.Close(SaveChanges=0)
    
    print "Data importing complete. Closing Excel..."
    
    xlApp.Quit()
    
    print "Writing to data file..."
    
    # Open a file for writing the data to
    output = open('Div15_Manager.dat', 'wb')
    
    # Pickle the output to the file and close the data stream
    cPickle.dump(data, output, 2)
    output.close()
    
    print "Program complete."

if __name__=='__main__':
    main(*sys.argv)