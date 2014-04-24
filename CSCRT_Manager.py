# Central Sterile Cross Reference Tool
# Written by Andrew Bouchard
# This program is written for a device company and provides an interface to access data from the main database
#   The program assumes available data that has already been processed by its companion program.
# Date Started: 7/10/2008
# Last Updated: 8/12/2008

# Import needed libraries
import wx
import cPickle
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

# This is the GUI class for the Guide box
class Guide(wx.Frame):
    def __init__(self, parent, id, title):
        
        #Initialize the frame
        wx.Frame.__init__(self, parent, id, title, size=(550, 300))
        
        # Define the panel and the sizer
        panel = wx.Panel(self, -1)
        sizer = wx.GridBagSizer(0, 0)
        
        # Declare fonts
        title = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        subtitle = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        
        # Insert an icon
        iconFile = "Logo.ico"
        CompanyIcon = wx.Icon(iconFile, wx.BITMAP_TYPE_ICO)
        self.SetIcon(CompanyIcon)
        
        # Insert a title text
        lbltitle = wx.StaticText(panel, -1, 'Central Sterile Cross Reference Tool', style=wx.ALIGN_CENTER)
        sizer.Add(lbltitle, (0, 0), (1,2), wx.ALL | wx.EXPAND, 5)
        lbltitle.SetFont(title)
        
        #Insert a subtitle text
        lblsubtitle = wx.StaticText(panel, -1, 'Guide for Use', style=wx.ALIGN_CENTER)
        sizer.Add(lblsubtitle, (1, 0), (1,2), wx.ALL | wx.EXPAND, 5)
        lblsubtitle.SetFont(subtitle)
        
        # Separation line
        sepline = wx.StaticLine(panel, -1 )
        sizer.Add(sepline, (2, 0), (1, 2), wx.ALL | wx.EXPAND, 5)
        
        # Help tasks combo box
        tasks = ['Look up a competitor part number', 'Look up a Company part number', 'Add information to Excel spreadsheet', 'Browse product family', 'Update data', 'Report errors']
        lblHelpTask = wx.StaticText(panel, -1, 'Select a task:')
        sizer.Add(lblHelpTask, (3, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.cmbHelpTask = wx.ComboBox(panel, -1, 'Look up a competitor part number', choices=tasks, style=wx.CB_READONLY)
        sizer.Add(self.cmbHelpTask, (3, 1), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.Bind(wx.EVT_COMBOBOX,  self.GetHelp, id=self.cmbHelpTask.GetId())
        
        # Help display box
        helpfont=wx.Font(9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,  faceName="Courier New")
        self.txtHelp = wx.TextCtrl (panel, -1, style = wx.TE_MULTILINE )
        self.txtHelp.SetFont(helpfont)
        self.txtHelp.SetEditable(False)
        sizer.Add(self.txtHelp, (4, 0), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        sizer.AddGrowableCol(1)
        sizer.AddGrowableRow(4)
        panel.SetSizerAndFit(sizer)
        self.Centre()
        self.Show(True)
        
        self.GetHelp(1)
    
    def GetHelp(self, event):
        
        # Get the option selected from the drop-down
        choice = self.cmbHelpTask.GetValue()
        
        # Instructions for looking up a Company part number
        if choice == 'Look up a Company part number':
            text = '- Type the Company number in the \'Part Number\' text box\n'
            text = text+'- Press the \'Enter\' key\n'
            text = text+'- The item description will be displayed in the \'Description\' text box\n'
            text = text+'- The stock status and stock level will be displayed in the \'Stock Status\' section\n'
            text = text+'- All competitors with the item will appear in the \'Competitor Information\' with their part numbers\n'
            text = text+'- The product family will be displayed in the \'Product Family\' section\n'
        
        # Instructions for looking up a competitor part number
        elif choice == 'Look up a competitor part number':
            text = '- Type the competitor number in the \'Part Number\' text box using the formatting guide below\n'
            text = text+'  -----Vendor-----------Item Number-------------------Tip--------------\n'
            #text = text+'  Aesculap        BT041R                 Numbers are always preceded by\n'
	    text = text+'  Uibroli         BT041R                 Numbers are always preceded by\n'
            text = text+'                  (Cross Ex: BT041)      two letters and end with an R;\n'
            text = text+'                                         do not use the R at the end of\n'
            text = text+'                                         the item number\n'
            #text = text+'  Codman (J&J)    50-4071                Always use the dash\n'
	    text = text+'  Frazo           50-4071                Always use the dash\n'
            #text = text+'  Jarit           200-105                Always use the dash\n'
            text = text+'  Zpemebi         200-105                Always use the dash\n'
            #text = text+'  KMedic          KM51-372               Numbers are always preceded by\n'
            text = text+'  Pbrajol         KM51-372               Numbers are always preceded by\n'
            text = text+'                                         KM; always use the dash\n'
            #text = text+'  Miltex          11-122                 Must add MX prefix in front of\n'
            text = text+'  Miltex          11-122                 Must add MX prefix in front of\n'
            text = text+'                  (Cross Ex: MX11-122)   number; always use the dash\n'
            #text = text+'  Pilling Weck    Pilling: 16-4715       Pilling: Do not use dash in \n'
            text = text+'  Gomev Iemda     Gomev: 16-4715         Gomev: Do not use dash in \n'
            text = text+'                  (Cross Ex: 164715)              number\n'
            #text = text+'                  Weck: 480190           Weck: There is no dash in \n'
            text = text+'                  Iemda: 480190          Iemda: There is no dash in \n'
            text = text+'                                               number; usually five to \n'
            text = text+'                                               six digits\n'
            #text = text+'  Sklar           60-1085                Always use the dash\n'
            text = text+'  Lagiso          60-1085                Always use the dash\n'
            #text = text+'  Storz           N4847                  Numbers typically start with \n'
            text = text+'  Oineza          N4847                  Numbers typically start with \n'
            text = text+'                                         N or E\n'
            #text = text+'  V. Mueller      SU3660                 Numbers are always preceded by\n'
            #text = text+'    (Allegiance)                         two letters; include dash if \n'
            text = text+'  Pdrabiyu        SU3660                 Numbers are always preceded by\n'
            text = text+'    (Qiaqe)                              two letters; include dash if \n'
            text = text+'                                         the number has one\n'
            text = text+'- Press the \'Enter\' key\n'
            text = text+'- The associated Company number will be displayed in the \'Part Number\' text box\n'
            text = text+'- The item description will be displayed in the \'Description\' text box\n'
            text = text+'- The stock status and stock level will be displayed in the \'Stock Status\' section\n'
            text = text+'- All competitors with the item will appear in the \'Competitor Information\' with their part numbers\n'
            text = text+'- The product family will be displayed in the \'Product Family\' section\n'
        
        # Instructions for adding information to an Excel spreadsheet
        elif choice == 'Add information to Excel spreadsheet':
            text = '- Make sure that all desired information is in the appropriate text box\n'
            text = text+'- Choose the \'Add Item to List\' option from the File menu, press the button, or press Ctrl+E\n'
            text = text+'- The part number, description, stock status, and stock level are inserted into a new Excel spreadsheet with appropriate headings\n'
            text = text+'- Continue to add entries until you are finished\n'
            text = text+'- Note that items that are not found can still be added to the Excel sheet for later reference\n'
            text = text+'- Save the spreadsheet when finished and close Excel\n'
            text = text+'- Choosing \'Add Item to List\' after closing Excel will create a new spreadsheet\n'
        
        # Instructions for browsing a product family
        elif choice == 'Browse product family':
            text = '- Find an entry for a Company or competitor part number\n'
            text = text+'- Double-click on any of the items in the \'Product Family\' text box\n'
            text = text+'- The selected Company number will be displayed in the \'Part Number\' text box\n'
            text = text+'- The item description will be displayed in the \'Description\' text box\n'
            text = text+'- The stock status and stock level will be displayed in the \'Stock Status\' section\n'
            text = text+'- All competitors with the item will appear in the \'Competitor Information\' with their part numbers\n'
            text = text+'- The product family will be displayed in the \'Product Family\' section\n'
        
        # Instructions for updating product data
        elif choice == 'Update data':
            text = '- The data for the program is stored in the Div15_Sales.dat file in the program folder\n'
            text = text+'- To update the data in the program, this file must be replaced with a more recent version\n'
            text = text+'- Contact Div. 15 to download and install the new file\n'
        
        # Instructions for reporting errors
        elif choice == 'Report errors':
            text = '- Any and all errors should be reported to Div. 15 directly\n'
            text = text+'- Company Information Technology does not support this software\n'
        
        # Otherwise, something is wrong, display nothing
        else:
            text = ''
        
        # Display the help text in the text box
        self.txtHelp.SetValue(text)

# This is the GUI class for the About box
class About(wx.Frame):
    def __init__(self, parent, id, title):
        
        # Initialize the frame
        wx.Frame.__init__(self, parent, id, title, size=(320, 200))
        
        # Define the panel and the sizer
        panel = wx.Panel(self, -1)
        sizer = wx.GridBagSizer(0, 0)
        
        # Declare fonts
        title = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        subtitle = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        
        # Insert an icon
        iconFile = "Logo.ico"
        CompanyIcon = wx.Icon(iconFile, wx.BITMAP_TYPE_ICO)
        self.SetIcon(CompanyIcon)
        
        # Insert a title text
        lbltitle = wx.StaticText(panel, -1, 'Central Sterile Cross Reference Tool', style=wx.ALIGN_CENTER)
        sizer.Add(lbltitle, (0, 0), (1,1), wx.ALL | wx.EXPAND, 5)
        lbltitle.SetFont(title)
        
        #Insert a subtitle text
        lblsubtitle = wx.StaticText(panel, -1, 'Written for Company, Inc.\nAuthored by Andrew T. Bouchard\n(c)2008', style=wx.ALIGN_CENTER)
        sizer.Add(lblsubtitle, (1, 0), (1,1), wx.ALL | wx.EXPAND, 5)
        lblsubtitle.SetFont(subtitle)
        
        # Separation line
        sepline = wx.StaticLine(panel, -1 )
        sizer.Add(sepline, (2, 0), (1, 1), wx.ALL | wx.EXPAND, 5)
        
        #Insert content
        lblcontent = wx.StaticText(panel, -1, 'The program is designed to provide access to Company database\ninformation, provided in a companion file generated by a\nseparate program.', style=wx.ALIGN_CENTER)
        sizer.Add(lblcontent, (3, 0), (1,1), wx.ALL | wx.EXPAND, 5)
        lblcontent.Wrap(-1)
        
        panel.SetSizerAndFit(sizer)
        self.Centre()
        self.Show(True)

# This is the main GUI class
class CSCRT(wx.Frame):
    def __init__(self, parent, id, title):
        
        # Initialize the frame
        wx.Frame.__init__(self, parent, id, title, size=(620, 700))
        
        # Declare global variables
        global data
        global date
        
        try:
            # Import the data
            datafile = open("Div15_Manager.dat", 'rb')
            data = cPickle.load(datafile)
            datafile.close()

            # Define the date and remove the entry from the dictionary - it causes problems
            date = data['Date']
            del data['Date']
        
        except:
            # Data importing error
            errstring = 'Data file Div15_Manager.dat could not be found.\nMake sure this file is in the installation directory and\nrefresh data before using program.\n\nCall Div. 15 for assistance.'
            wx.MessageBox(errstring, 'Info')
        
        # Declare fonts
        title = wx.Font(14, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        subtitle = wx.Font(11, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        
        # Define the panel and the sizer
        panel = wx.Panel(self, -1)
        sizer = wx.GridBagSizer(0, 0)
        
        # Insert an icon
        iconFile = "Logo.ico"
        CompanyIcon = wx.Icon(iconFile, wx.BITMAP_TYPE_ICO)
        self.SetIcon(CompanyIcon)

        # Create the menu
        menubar = wx.MenuBar()
        file = wx.Menu()
        self.mAddToList = wx.MenuItem(file, -1, 'Add It&em to List\tCtrl+E')
        file.AppendItem(self.mAddToList)
        self.Bind(wx.EVT_MENU, self.AddToList, id=self.mAddToList.GetId())
        self.mRefresh = wx.MenuItem(file, -1, '&Refresh data\tCtrl+R')
        file.AppendItem(self.mRefresh)
        self.Bind(wx.EVT_MENU, self.OnRefresh, id=self.mRefresh.GetId())
        self.mQuit = wx.MenuItem(file, -1, '&Quit\tCtrl+Q')
        file.AppendItem(self.mQuit)
        self.Bind(wx.EVT_MENU, self.OnQuit, id=self.mQuit.GetId())
        menubar.Append(file, '&File')
        help = wx.Menu()
        self.mGuide = wx.MenuItem(help, -1, '&Guide\tCtrl+G')
        help.AppendItem(self.mGuide)
        self.Bind(wx.EVT_MENU, self.OnGuide, id=self.mGuide.GetId())
        self.mAbout = wx.MenuItem(help, -1, 'About')
        help.AppendItem(self.mAbout)
        self.Bind(wx.EVT_MENU, self.OnAbout, id=self.mAbout.GetId())
        menubar.Append(help, '&Help')
        self.SetMenuBar(menubar)

        # Declare the row for the gui widgets
        guirow = 0
        
        # Main Title
        lblmaintitle = wx.StaticText(panel, -1, 'Central Sterile Cross Reference Tool', style=wx.ALIGN_CENTER)
        sizer.Add(lblmaintitle, (guirow, 0), (1,5), wx.ALL | wx.EXPAND, 5)
        lblmaintitle.SetFont(title)
        guirow = guirow+1
        lblmainsubtitle = wx.StaticText(panel, -1, 'Company, Inc.', style=wx.ALIGN_CENTER)
        sizer.Add(lblmainsubtitle, (guirow, 0), (1,5), wx.ALL | wx.EXPAND, 5)
        lblmainsubtitle.SetFont(subtitle)

        guirow = guirow+1
        
        # Separation line
        line1 = wx.StaticLine(panel, -1 )
        sizer.Add(line1, (guirow, 0), (1, 5), wx.ALL | wx.EXPAND, 5)

        guirow = guirow+1
        
        # Part number entry
        lblPartNo = wx.StaticText(panel, -1, 'Part Number:')
        sizer.Add(lblPartNo, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtPartNo = wx.TextCtrl(panel, -1, style=wx.PROCESS_ENTER)
        sizer.Add(self.txtPartNo, (guirow, 1), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.Bind(wx.EVT_TEXT_ENTER, self.PassOn, id = self.txtPartNo.GetId())

        # Description display
        lblDescription = wx.StaticText(panel, -1, 'Description:')
        sizer.Add(lblDescription, (guirow, 2), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtDescription = wx.TextCtrl(panel, -1)
        self.txtDescription.SetEditable(False)
        sizer.Add(self.txtDescription, (guirow, 3), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        guirow = guirow+1
        
        # Data validity label
        datavalid = 'Data Valid As Of ' + date
        lblDataValid = wx.StaticText(panel, -1, datavalid)
        sizer.Add(lblDataValid, (guirow, 0), (1, 2), wx.ALL | wx.ALIGN_CENTER, 5)

        # Add item to list button
        cmdAddToList = wx.Button(panel, -1, 'Add It&em To List (Ctrl+E)')
        sizer.Add(cmdAddToList, (guirow, 3), (1,2), flag=wx.ALL | wx.ALIGN_CENTER | wx.EXPAND, border=5)
        self.Bind(wx.EVT_BUTTON,  self.AddToList, id=cmdAddToList.GetId())
        
        guirow = guirow+1
        
        # Separation line
        line2 = wx.StaticLine(panel, -1)
        sizer.Add(line2, (guirow, 0), (1, 5), wx.ALL | wx.EXPAND, 5)

        guirow = guirow+1
        
        # Stock status title
        lblstocktitle = wx.StaticText(panel, -1, 'Stock Status', style=wx.ALIGN_CENTER)
        sizer.Add(lblstocktitle, (guirow, 0), (1,5), wx.ALL | wx.ALIGN_CENTER | wx.EXPAND, 5)
        lblstocktitle.SetFont(subtitle)
        
        guirow = guirow+1
        
        # B02 Fixed Vendor
        lblB02Vendor = wx.StaticText(panel, -1, 'B02 Fixed Vendor:')
        sizer.Add(lblB02Vendor, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtB02Vendor = wx.TextCtrl(panel, -1)
        self.txtB02Vendor.SetEditable(False)
        sizer.Add(self.txtB02Vendor, (guirow, 1), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        # DIR Fixed Vendor
        lblDirVendor = wx.StaticText(panel, -1, 'DIR Fixed Vendor:')
        sizer.Add(lblDirVendor, (guirow, 3), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtDirVendor = wx.TextCtrl(panel, -1)
        self.txtDirVendor.SetEditable(False)
        sizer.Add(self.txtDirVendor, (guirow, 4), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        guirow = guirow+1
        
        # Stock status display
        lblB02Status = wx.StaticText(panel, -1, 'B02 Status:')
        sizer.Add(lblB02Status, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtB02Status = wx.TextCtrl(panel, -1)
        self.txtB02Status.SetEditable(False)
        sizer.Add(self.txtB02Status, (guirow, 1), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        # Stock level display
        lblStockLevel = wx.StaticText(panel, -1, 'Total Company Stock:')
        sizer.Add(lblStockLevel, (guirow, 3), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtStockLevel = wx.TextCtrl(panel, -1)
        self.txtStockLevel.SetEditable(False)
        sizer.Add(self.txtStockLevel, (guirow, 4), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        guirow = guirow+1
        
        # DIR status display
        lblDirStatus = wx.StaticText(panel, -1, 'DIR Status:')
        sizer.Add(lblDirStatus, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtDirStatus = wx.TextCtrl(panel, -1)
        self.txtDirStatus.SetEditable(False)
        sizer.Add(self.txtDirStatus, (guirow, 1), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        # Spacer for rightmost column
        lblSpace = wx.StaticText(panel, -1, '', size = (200, 5))
        sizer.Add(lblSpace, (guirow, 4), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)

        guirow = guirow+1
        
        # Separation line
        line3 = wx.StaticLine(panel, -1)
        sizer.Add(line3, (guirow, 0), (1, 5), wx.ALL | wx.EXPAND, 5)

        guirow = guirow+1
        
        # Sales title
        lblsalestitle = wx.StaticText(panel, -1, 'Sales Data', style=wx.ALIGN_CENTER)
        sizer.Add(lblsalestitle, (guirow, 0), (1,5), wx.ALL | wx.ALIGN_CENTER | wx.EXPAND, 5)
        lblsalestitle.SetFont(subtitle)
        
        guirow = guirow+1
        
        # 12-month sales display
        lbl12MoSales = wx.StaticText(panel, -1, 'Sales Past 12 Mo.:')
        sizer.Add(lbl12MoSales, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txt12MoSales = wx.TextCtrl(panel, -1)
        self.txt12MoSales.SetEditable(False)
        sizer.Add(self.txt12MoSales, (guirow, 1), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        # Stock level display
        lblOrders = wx.StaticText(panel, -1, 'No. Orders:')
        sizer.Add(lblOrders, (guirow, 3), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtOrders = wx.TextCtrl(panel, -1)
        self.txtOrders.SetEditable(False)
        sizer.Add(self.txtOrders, (guirow, 4), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        guirow = guirow+1
        
        # Separation line
        line4 = wx.StaticLine(panel, -1)
        sizer.Add(line4, (guirow, 0), (1, 5), wx.ALL | wx.EXPAND, 5)

        guirow = guirow+1
        
        # Competitor information title
        lblcomptitle = wx.StaticText(panel, -1, 'Competitor Information', style=wx.ALIGN_CENTER)
        sizer.Add(lblcomptitle, (guirow, 0), (1,2), wx.ALL | wx.ALIGN_CENTER | wx.EXPAND, 5)
        lblcomptitle.SetFont(subtitle)
        
        guirow = guirow+1
        
        # Unique vendor display
        lblUniqueVendors = wx.StaticText(panel, -1, 'Unique Vendors:')
        sizer.Add(lblUniqueVendors, (guirow, 0), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.txtUniqueVendors = wx.TextCtrl(panel, -1)
        self.txtUniqueVendors.SetEditable(False)
        sizer.Add(self.txtUniqueVendors, (guirow, 1), (1, 1), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        guirow = guirow+1
        
        # Vendor cross-reference display
        self.lbCrossRef = wx.ListBox(panel, -1, style= wx.LB_SINGLE)
        sizer.Add(self.lbCrossRef, (guirow, 0), (1,2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER, 5)
        
        guirow = guirow-2
        
        # Product family title
        lblProductFamily = wx.StaticText(panel, -1, 'Product Family', style=wx.ALIGN_CENTER)
        sizer.Add(lblProductFamily, (guirow, 2), (1,3), wx.ALL | wx.ALIGN_CENTER | wx.EXPAND, 5)
        lblProductFamily.SetFont(subtitle)
        
        guirow = guirow+1
        
        # Family sorting combo box
        sortchoices = ['Part Number', 'Competitors', 'Sales']
        lblFamilySort = wx.StaticText(panel, -1, 'Sort By:')
        sizer.Add(lblFamilySort, (guirow, 2), flag=wx.ALL | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.cmbFamilySort = wx.ComboBox(panel, -1, 'Part Number', choices=sortchoices, style=wx.CB_READONLY)
        sizer.Add(self.cmbFamilySort, (guirow, 3), (1, 2), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.Bind(wx.EVT_COMBOBOX,  self.SortFamily, id=self.cmbFamilySort.GetId())
        
        guirow = guirow+1
        
        # Product family display
        self.lbFamily = wx.ListBox(panel, -1, style= wx.LB_SINGLE)
        sizer.Add(self.lbFamily, (guirow, 2), (1,3), wx.EXPAND | wx.ALL | wx.ALIGN_CENTER, 5)
        self.Bind(wx.EVT_LISTBOX_DCLICK, self.GoToEntry, id=self.lbFamily.GetId())
    
        sizer.AddGrowableCol(3)
        sizer.AddGrowableRow(guirow)
        panel.SetSizerAndFit(sizer)
        self.Centre()
        self.Show(True)
        
        # Bind an event for switching back to this window
        self.Bind(wx.EVT_ACTIVATE, self.OnActivate)
    
    def OnActivate(self, event):
        
        # Select the part number when the focus returns to this window
        if event.GetActive():
            self.txtPartNo.SetFocus()
            self.txtPartNo.SelectAll()
        event.Skip()

    def GoToEntry(self, event):
        
        # Get the information back from the list box
        item = self.lbFamily.GetSelection()
        fulltext = self.lbFamily.GetString(item)
        
        # Trim to the first space to get the item number
        parts = fulltext.rsplit(' ')
        pn = parts[0]
        self.txtPartNo.SetValue(pn)
        self.GetData()
    
    def PassOn(self, event):
        
        self.GetData()
        
    def AddToList(self, event):
        
        global myBook
        
        # Make sure that the data in the textboxes is representative of the part number in the entry
        self.GetData()
        
        # If an Excel sheet is not already open, open one
        xlApp = Dispatch("Excel.Application")
        
        try:
            
            myBook.Sheets(1).Cells(1,255).Value = ''
            
        except:          
            
            # Create the workbook
            try:
                xlApp.Visible = 1
                
            except:
                print("xlApp.Visible Exception handled")
            
            myBook = xlApp.Workbooks.Add()
            
            # Create headings
            myBook.Sheets(1).Cells(1,1).Value = 'Part No.'
            myBook.Sheets(1).Cells(1,2).Value = 'Description'
            myBook.Sheets(1).Cells(1,3).Value = 'B02 Fixed Vendor'
            myBook.Sheets(1).Cells(1,4).Value = 'DIR Fixed Vendor'
            myBook.Sheets(1).Cells(1,5).Value = 'B02 Status'
            myBook.Sheets(1).Cells(1,6).Value = 'Stock Level'
            myBook.Sheets(1).Cells(1,7).Value = 'DIR Status'
            myBook.Sheets(1).Cells(1,8).Value = '12 Mo. Sales'
            myBook.Sheets(1).Cells(1,9).Value = 'No. Orders'
        
        # Identify the column to align with
        for col in range(1, 255):
            header = str(myBook.Sheets(1).Cells(1, col))
            if header == 'Part No.':
                break
        
        # Find the next empty row
        for row in range(2, 65536):
            empty = str(myBook.Sheets(1).Cells(row, col))
            if (empty == 'None'):
                break
        
        # Insert the data at the open row
        myBook.Sheets(1).Cells(row,col).Value = self.txtPartNo.GetValue()
        myBook.Sheets(1).Cells(row,col+1).Value = self.txtDescription.GetValue()
        myBook.Sheets(1).Cells(row,col+2).Value = self.txtB02Vendor.GetValue()
        myBook.Sheets(1).Cells(row,col+3).Value = self.txtDirVendor.GetValue()
        myBook.Sheets(1).Cells(row,col+4).NumberFormat = "@"
        myBook.Sheets(1).Cells(row,col+4).Value = self.txtB02Status.GetValue()
        myBook.Sheets(1).Cells(row,col+5).Value = self.txtStockLevel.GetValue()
        myBook.Sheets(1).Cells(row,col+6).NumberFormat = "@"
        myBook.Sheets(1).Cells(row,col+6).Value = self.txtDirStatus.GetValue()
        myBook.Sheets(1).Cells(row,col+7).Value = self.txt12MoSales.GetValue()
        myBook.Sheets(1).Cells(row,col+8).Value = self.txtOrders.GetValue()
        
        # Fit the column widths
        for i in range(col, col+9):
            myBook.Sheets(1).Cells(1, i).EntireColumn.AutoFit()
        
    def SortFamily(self, event):
        
        # Find the entry for the entered part number in the data dictionary
        pn = self.txtPartNo.GetValue()
        
        # Make sure that the current part number entry is valid
        if pn not in data.keys():
            
            #Warn the user the part was not found
            errstring = 'Part number ' + pn + ' not found.'
            wx.MessageBox(errstring, 'Info')
            
            # Put dummy data out
            info = Item()
        
        else:
            
            # Get the data associated with the part number
            info = data[pn]
        
        # Display the family data
        sortme = {}
        choice = self.cmbFamilySort.GetValue()
        
        
        # Build the family dictionary
        for part in data.keys():
            if pn[0:7] == part[0:7]:
                if choice == 'Part Number':
                    sortme[part] = data[part].Description
                elif choice == 'Competitors':
                    sortme[len(data[part].Competitors.keys())] = part
                elif choice == 'Sales':
                    try:
                        sortme[data[part].Sales].append(part)
                    except:
                        sortme[data[part].Sales] = []
                        sortme[data[part].Sales].append(part)
                else:
                    sortme[part] = part
        
        # Sort by the dictionary keys
        order = sortme.keys()
        order.sort(reverse = True)
        
        # Clear the family listbox
        self.lbFamily.Clear()
        
        place = 0
        for index in order:
            
            # Format and display the family information, then increment the insertion point
            if choice == 'Sales':
                for entry in sortme[index]:
                    famstring = entry + ' - ' + str(index)
                    self.lbFamily.Insert(famstring, place)
                    place = place+1
            elif choice == 'Part Number':
                famstring = str(index) + ' - ' + sortme[index]
                self.lbFamily.Insert(famstring, place)
                place = place+1
            else:
                famstring = sortme[index] + ' - ' + str(index)
                self.lbFamily.Insert(famstring, place)
                place = place+1
    
    def GetData(self):
        
        # Find the entry for the entered part number in the data dictionary
        pn = self.txtPartNo.GetValue()
        
        # Make sure that all letters are upper-case and that there are no spaces
        pn = pn.strip()
        pn = pn.upper()
        self.txtPartNo.SetValue(pn)
        
        # Check if the part number exists in the data
        while pn not in data.keys():
            
            # Check to see if it is a cross-reference number
            results = []
            for part in data.keys():
                for entry in data[part].Competitors.keys():
                    if (pn in data[part].Competitors[entry]) and (part not in results):
                        results.append(part)
            
            if len(results)>1:
                
                outstring = 'Multiple matching part numbers were found.\n' + str(results)
                wx.MessageBox(outstring, 'Info')
            
            # If parts were found
            if len(results)==1:
                
                pn = results[0]
                self.txtPartNo.SetValue(pn)
            
            # Multiple parts were found - use logic to determine which to return
            elif len(results)>1:
                
                done = False
                
                # Look for a stocked MDS number first
                for mds in results:
                    
                    if mds.startswith('MDS'):
                        
                        if data[mds].StockStatus == 1:
                            
                            if not done:
                                pn = mds
                            done = True
                        
                        results.remove(mds)
                
                # Look for a stocked MDG number
                if not done:
                    
                    for mdg in results:
                        
                        if mdg.startswith('MDG'):
                            
                            if data[mdg].StockStatus == 1:
                                
                                if not done:
                                    pn = mdg
                                done = True
                            
                            results.remove(mdg)
                
                # No stocked options - call division
                if not done:
                    
                    if AddingList:
                        errstring = 'Although part number ' + pn + ' is not stocked, it will be added to the list. \nCall Div. 15 for assistance.'
                        wx.MessageBox(errstring, 'Info')
                    else:
                        errstring = 'Part number ' + pn + ' not stocked. \nCall Div. 15 for assistance.'
                        wx.MessageBox(errstring, 'Exclamation')
                
                self.txtPartNo.SetValue(pn)
            
            else:
                
                #Warn the user the part was not found
                if AddingList:
                    errstring = 'Although part number ' + pn + ' was not found, it will be added to the list. \nCall Div. 15 for assistance.'
                    wx.MessageBox(errstring, 'Info')
                else:
                    errstring = 'Part number ' + pn + ' not found. \nCall Div. 15 for assistance.'
                    wx.MessageBox(errstring, 'Exclamation')
                
                break
        
        # Get the data associated with the part number
        try:
            info = data[pn]
        except:
            # Put dummy data out on error
            info = Item()
        
        # Display the description
        self.txtDescription.SetValue(info.Description)
        
        # Display the sales data
        self.txt12MoSales.SetValue(str(info.Sales))
        self.txtOrders.SetValue(str(info.Orders))
        
        # Display the fixed vendor
        self.txtB02Vendor.SetValue(info.B02Vendor)
        self.txtDirVendor.SetValue(info.DirVendor)
        
        # Display the stock status
        self.txtB02Status.SetValue(info.B02Status)
        self.txtDirStatus.SetValue(info.DirStatus)
        
        # Display the available stock
        self.txtStockLevel.SetValue(str(info.StockLevel))
        
        # Display the number of unique competing vendors
        numVendors = str(len(info.Competitors.keys()))
        self.txtUniqueVendors.SetValue(numVendors)
        
        # Display the cross reference information
        # Define the insertion point in the listbox
        place = 0
        
        # Go through the vendor entries in the part information
        self.lbCrossRef.Clear()
        for vendor in info.Competitors.keys():
            
            # Insert the vendor name and increment the insertion counter
            self.lbCrossRef.Insert(vendor, place)
            place = place+1
            
            # Go through the part numbers for this vendor
            for cpn in info.Competitors[vendor]:
                
                # Format and display the vendor's crossref number, then increment the insertion point
                pnstring = '- ' + cpn
                self.lbCrossRef.Insert(pnstring, place)
                place = place+1
        
        # Display the family data
        sortme = {}
        choice = self.cmbFamilySort.GetValue()
        
        # Build the family dictionary
        for part in data.keys():
            if pn[0:7] == part[0:7]:
                if choice == 'Part Number':
                    sortme[part] = data[part].Description
                elif choice == 'Competitors':
                    sortme[len(data[part].Competitors.keys())] = part
                else:
                    sortme[part] = part
        
        # Sort by the dictionary keys
        order = sortme.keys()
        order.sort()
        
        # Clear the family listbox
        self.lbFamily.Clear()
        
        place = 0
        for index in order:
            
            # Format and display the family information, then increment the insertion point
            if choice == 'Part Number':
                famstring = str(index) + ' - ' + sortme[index]
            else:
                famstring = sortme[index] + ' - ' + str(index)
            self.lbFamily.Insert(famstring, place)
            place = place+1

        # If the part does not exist, delete the contents of the family list
        if info.Description=='':
            self.lbFamily.Clear()
    
    def OnQuit(self, event):
        self.Close()
    
    def OnRefresh(self, event):
        
        global data
        global date
        
        try:
            # Import the data
            datafile = open("Div15_Manager.dat", 'rb')
            data = cPickle.load(datafile)
            datafile.close()

            # Define the date and remove the entry from the dictionary - it causes problems
            date = data['Date']
            del data['Date']
        
        except:
            # Data importing error
            errstring = 'Data file Div15_Manager.dat could not be found.\nMake sure this file is in the installation directory and\nrefresh data before using program.\n\nCall Div. 15 for assistance.'
            wx.MessageBox(errstring, 'Info')
    
    def OnAbout(self, event):
        
        description = """The Product Cross Reference Tool is written for Company \nas a tool to locally access product database information."""

        info = wx.AboutDialogInfo()

        info.SetIcon(wx.Icon('Company.png', wx.BITMAP_TYPE_PNG))
        info.SetName('Product Cross Reference Tool')
        info.SetVersion('1.0')
        info.SetDescription(description)
        info.SetCopyright('(C) 2008 Andrew T. Bouchard')
        info.SetLicence(licence)
        info.AddDeveloper('Andrew T. Bouchard')
        info.AddDocWriter('Andrew T. Bouchard')

        wx.AboutBox(info)
    
    def OnGuide(self, event):
        
        Guide(None, -1, 'Guide')

def TranslateStatus(Status):
   
    # Assign useful statuses based on part type and codes
    if Status == 0:
        return ""
    elif Status == 1:
        return "Stocked"
    elif Status == 2:
        return "Call Div. 15 for sub."
    else:
        return "Error: " + str(Status)

# Define global variables
global data
global myBook
global date
global main

# Initialize the globals
data = {}
date = ""

# Start the GUI
app = wx.App()
main = CSCRT(None, -1, 'Company, Inc.')
app.MainLoop()