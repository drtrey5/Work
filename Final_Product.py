#!/usr/bin/env python
# coding: utf-8

#imports for the application

import numpy as np    
import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
from scipy import stats
import tkcalendar as cal
import pandastable as ps

import matplotlib.pyplot as plt
import datetime as dt
#import statsmodels as sm
    
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
    
xlsx = pd.ExcelFile("Data5_23_Correct.xlsx", engine = "openpyxl")
dataOEE_Page = pd.read_excel(xlsx,'Twin Screw OEE')
dataTS_Page = pd.read_excel(xlsx, "Twin Screw ")
pairedDate = pd.DataFrame()


def startUp():
    #This section of code takes filler sections of the excel spreadsheet from both tabs and removes them

    global dataOEE_Page
    global dataTS_Page

    dataOEE_Page.replace('x', np.nan, inplace= True)
    dataTS_Page.replace('x', np.nan, inplace = True)
    
    dataOEE_Page.drop(labels = 0, axis = 0, inplace = True)
    dataTS_Page.drop(labels = 0, axis = 0, inplace = True)

    #print(dataOEE_Page)

    #Names the indexes for the columns of the dataframes

    dataOEE_Page.columns = ['Date', 'Lbs Produced', '%OEE', 'Labor Hours', 'Lb/Man Hour', 'Pails Filled', 'Pails/Man Hour', 'Removal', 'Week of', 'Weekly %OEE']
    dataTS_Page.columns = ['Date', 'Num of Ops', 'Total Minutes', 'Maintenance', 'Making Powder', 'Making Prepolymer', 'Startup', 'Breaks', 'LBS Produced', 'Pails Filled', 'Minutes Worked', 'Efficiency', 'Unnamed', 'Unnamed', 'Unnamed']

    #dataOEE_Page

    #This section trims the fat off the data set and removes entries that are effectively empty from both the OEE and Efficiency
    #tab. 

    dataTS_Page = dataTS_Page.dropna(how = 'all', axis = 1)
    dataTS_Page = dataTS_Page.dropna(subset=['Num of Ops'])
    dataTS_Page = dataTS_Page.drop(labels='Pails Filled', axis = 1)

    dataOEE_Page = dataOEE_Page.dropna(subset=['Labor Hours'])
    dataOEE_Page = dataOEE_Page.dropna(how = 'all', axis = 1)
    dataOEE_Page = dataOEE_Page.drop(labels='Weekly %OEE', axis = 1)
    dataOEE_Page = dataOEE_Page.drop(labels='Week of', axis = 1)

    #Assigns type data to the values in the columns, certain columns are uninterpretable due to the extra data within them 
    
    dataOEE_Page = dataOEE_Page.astype({"Lbs Produced": np.float64, "%OEE": np.float64, "Labor Hours": np.float64, "Lb/Man Hour": np.float64, "Pails Filled": np.float64, "Pails/Man Hour": np.float64})
    dataTS_Page = dataTS_Page.astype({'Num of Ops': np.float64, 'Total Minutes': np.float64, 'Maintenance': np.float64, 'Efficiency': np.float64})
    
    #This section will identify the percent values that are based as a decimal in the updated dataframes
    #From here the new data is passed down onto future methods for stronger correlations

    dataOEE_Page.loc[dataOEE_Page['%OEE'] < 1, '%OEE'] = round(dataOEE_Page['%OEE']*100, 0)
    dataTS_Page.loc[dataTS_Page['Efficiency'] < 1, 'Efficiency'] = round(dataTS_Page['Efficiency']*100, 0)

    #dataOEE_Page.loc[str(dataOEE_Page['Pails Filled']).find('(') > -1, 'Pails Filled'] 
    


def pairDateData(OEEdata, TSdata):
    #This section of code takes the first 13 entries of the Twin Screw OEE tab of the excel sheet and removes them from the data pool
    #This effectively removes the logic error that occurred due to the lack of the # of operators on the - tab, and realigns the dates
    #Following this section, data that doesn't have a matching pair datewise will be removed from each respective dataframe

    global pairedDate

    OEEdata.drop(dataOEE_Page.index[range(0,13)], axis = 0, inplace =True)
    
    OEEdata = OEEdata.reset_index(drop = True)
    TSdata = TSdata.reset_index(drop = True)
    
    #Using the dates from the trimmed lists, then comparing them to make a condensed list of data points
    #This resultant list should contain all eligible data points
    #From this comparison dataframe we should be able to pull data that correlates directly
    #Should only be done in order to compare data cross sheet due to lack of variables

    dateCheck = [OEEdata['Date'], TSdata['Date']] #Makes a list of the columns from either dataFrame
    dateCheck = pd.concat(dateCheck, axis=1, ignore_index=False, verify_integrity=False) #Combines the dataCheck list into a dataFrame
    dateCheck['Removal_OEE'] = np.nan #Adds columns of nan values
    dateCheck['Removal_TS'] = np.nan 
    dateCheck.columns = ['OEE_Date', 'TS_Date','Removal_OEE','Removal_TS'] #Renames the columns of the dataframe
    
    #Checks whether or not the dates correspond to each other throughout either column, result is recorded in respective removal          #column
    #Recorded as True if there is matching pair, False if there isn't
    dateCheck['Removal_OEE'] = (dateCheck['OEE_Date'].isin(dateCheck['TS_Date']))
    dateCheck['Removal_TS'] = (dateCheck['TS_Date'].isin(dateCheck['OEE_Date']))

    OEEindex = dateCheck[dateCheck['Removal_OEE'] == False].index #Identifies the exact index of the date cell needing removed
    TSindex = dateCheck[dateCheck['Removal_TS'] == False].index

    #print(OEEindex)
    #print(TSindex)

    #Drops erroneous cells from both data sets
    TSdata = TSdata.drop(labels = ['Date'], axis = 1)
    OEEdata = OEEdata.drop(OEEindex, inplace=False)
    TSdata = TSdata.drop(TSindex,inplace=False)

    OEEdata = OEEdata.reset_index(drop = True)
    TSdata = TSdata.reset_index(drop = True)
    

    pairedDate = pd.concat([OEEdata, TSdata], axis = 1, ignore_index = False)

def narrowDates(date, userInput):

    #Will act as test parameters to prepare method for true use
    global pairedDate
    global tableData
    global beginningOfInterval
    global passedDate

    df = pairedDate.copy(deep = True)
    passedDate = date
    #testDay = True #Will be the program default
    #testWeek = False
    #testMonth = False
    testDate = pd.Timestamp(date + " 00:00:00")
    daysInMonth = testDate.daysinmonth
    beginningOfInterval = pd.Timestamp(2020)
    #print(daysInMonth)
    #userInput = 'Tier 3'

    
    df['Date'] = pd.to_datetime(df['Date'])

    #Basically decides how far back the date retrieval will go.

    if (userInput == "Tier 2"):
        #dateInterval = 0
        beginningOfInterval = testDate
    elif (userInput == 'Tier 3'):
        #dateInterval = 7
        beginningOfInterval = testDate - pd.Timedelta('7 day')
    elif (userInput == 'Tier 5'): #If not month's end it will retrieve all prior data up to that point
        beginningOfInterval = testDate.replace(day=1)
        #print (beginningOfInterval)
    else:
        print('System Error: Invalid selection')

    #if (testDate in df['Date'].values):
     #   print (\"It's here\")
    #else:
     #   print('Not in dataset')


    #Will set index of the dataframe to that of the dates to be easily moved through and sorted by
    df.set_index('Date', inplace=True)
    tableData = df.loc[beginningOfInterval:testDate]

#Strips down the tableData dataframe to the selected variables
def cleanseData(userX, userY, userY2):

    global tableData #used to change the tableData dataframe w/o a return

    #Doesn't allow interaction with an empty dataframe
    if tableData.empty:
        print('Action not executable with selected dataset')
    
    #Doesn't try to take the date indexes as a positional argument preventing an error
    elif userX == 'Date':
        if (userY == ""):
            tableData  = tableData[[userY2]]
        elif (userY2 == ''):
            tableData = tableData[[userY]]
        else:
            tableData = tableData[[userY, userY2]]
    #Avoids the date positional argument usage through no use of Date
    else:
        if (userY == ""):
            tableData  = tableData[[userX, userY2]]
        elif (userY2 == ''):
            tableData = tableData[[userX, userY]]
        else:
            tableData = tableData[[userX, userY, userY2]]

def findCorr(X, Y, Y2):

    global tableData
    global corrText
    ax = plt.gca()
    ax.set_ylim(bottom=0)
    ax.set_xlim(left = 0)
    #ax.show()

    corrText.config(state = 'normal')
    corrText.delete('1.0', tk.END)

    if (X == 'Date'):
        #df = pd.to_datetime(df['date']).map(dt.datetime.toordinal)
        return
    else:
        if Y2 == "":
            slope, intercept, r_value, p_value, std_err = stats.linregress(tableData[X], tableData[Y])
            x_vals = np.array(ax.get_xlim())
            print(x_vals)
            y_vals = intercept + slope * x_vals
            ax.plot(x_vals, y_vals, '--')
            corrText.insert(index=tk.INSERT, chars= "R^2 Value: \n" + str(r_value) + "\nSlope: \n" + str(slope) + "\nIntercept: \n" + str(intercept) + "\nStandard Error: \n" + str(std_err))
        elif Y == "":
            slope, intercept, r_value, p_value, std_err = stats.linregress(tableData[X], tableData[Y2])
            x_vals = np.array(ax.get_xlim())
            y_vals = intercept + slope * x_vals
            ax.plot(x_vals, y_vals, '--')
            corrText.insert(index=tk.INSERT, chars= "R^2 Value: \n" + str(r_value) + "\nSlope: \n" + str(slope) + "\nIntercept: \n" + str(intercept) + "\nStandard Error: \n" + str(std_err))
        else:
            slope, intercept, r_value, p_value, std_err = stats.linregress(tableData[X], tableData[Y])
            slope2, intercept2, r_value2, p_value2, std_err2 = stats.linregress(tableData[X], tableData[Y2])
            x_vals = np.array(ax.get_xlim())
            y_vals = intercept + slope * x_vals
            y_vals2 = intercept2 + slope2 * x_vals
            ax.plot(x_vals, y_vals, '--', label = str(slope)+ "x + " + str(intercept) + " (Set 1)", color = 'red')
            ax.plot(x_vals, y_vals2, '--', label = str(slope2)+ "x + " + str(intercept2) + " (Set 2)", color = 'blue')

    corrText.config(state='disabled')


    #dat = sm.tableData.get_rdataset(\"\", \"\").data

    # Fit regression model (using the natural log of one of the regressors)
    #results = smf.ols(data=dat).fit()

    # Inspect the results
    #print(results.summary())\n"
 
def applyFunction(date, userX, userY, userY2, tier, corrVal):
    
    global tableData
    global table
    global fig
    #global figIndex
    global canvas
    global lastX
    global lastY
    global lastY2
    global beginningOfInterval

    
    #for widget in gFrame.winfo_children():
        #widget.destroy()
    #figIndex = figIndex + 10
    plt.gca().remove()
    fig.add_subplot()
    ax = plt.gca()


    if userX == '' or (userY == '' and userY2 == '')  or (userY == 'Date' or userY2 == 'Date'):
        print('Invalid Input, please select both an X and Y variable')
    else:
        lastX = userX
        lastY = userY
        lastY2 = userY2

        narrowDates(date, tier)
        cleanseData(userX, userY, userY2)

        #N = len(tableData)
        #print(N)
        #ind = np.arange(N)
        #print(ind)
        #width = np.min(np.diff(ind))/3
        #width = 0.1

        if (userX == 'Date'):
            if (userY == ''):
                ax.scatter(tableData.index, tableData[userY2], label = userY2)
            elif (userY2 == ''):
                ax.scatter(tableData.index, tableData[userY], label = userY)
            else:
                ax.scatter(tableData.index, tableData[userY], width, color = 'red', label = userY, align = 'edge')
                ax.scatter(tableData.index, tableData[userY2], width, color = 'blue', label = userY2, align = 'edge')

        else:
            if (userY == ''):
                ax.scatter(tableData[userX], tableData[userY2], label = userY2)
            elif (userY2 == ''):
                ax.scatter(tableData[userX], tableData[userY], label = userY)
            else:
                ax.scatter(tableData[userX], tableData[userY], color = 'red', label = userY)
                ax.scatter(tableData[userX], tableData[userY2], color = 'blue', label = userY2)

        if (corrVal):
            #tableData.sort_values(userX, axis = 0, ascending = True, inplace = True, na_position ='last')
            #line = findCorr(userX)
            #ax.plot(line)
            findCorr(userX, userY, userY2)
            #print(\"Worked as plan\")


        ax.set_title(beginningOfInterval.strftime("%Y-%m-%d") + " to " + date)   
        ax.set_xlabel(userX)
        plt.setp(ax.get_xticklabels(), rotation=45)
        ax.legend()
        
        canvas.draw()
        
        table.model.df = tableData
        table.redraw()

#Doesn't function as intended... this is due to the list not retrieving the chart as a datatype
def exportReady(X, Y, Y2, ax):

    global exportList
    global exportCount
    global exportText
    global fig
    global passedDate
    global beginningOfInterval

    #added benefit of protecting the formating of the original graph

    exportList.append(ax)

    if(Y == ''):
        Y = "null"
    elif(Y2 == ''):
        Y2 = 'null'

    #Adds parameters to the export list textbox so removal if chosen can occur
    exportText.config(state='normal')
    exportText.insert(index=tk.INSERT, chars=(str(exportCount) + ". " + X + ", " + Y + ", " + Y2 + "   " +  beginningOfInterval.strftime("%Y-%m-%d")+ ' to  ' + passedDate + "\\n"))
    exportText.config(state='disabled')
    
    exportCount = exportCount + 1 #count that controls the referrence to list elements\n"
 
#Saves the chart to a defined place, look for code that pulls up a dropdown menu on where exactly to store it.
    def saveCharts(exports):
    
        global exportList
        global exportCount
        global exportText
    
        files = [('PDF File','.pdf')]
    
        file_path = fd.asksaveasfile(filetypes = files, defaultextension = files)
        if file_path is None:
            return
    
        with PdfPages(file_path.name) as pdf:
            exportIndex = 0
            figList = []
            for chart in exports:
                # get a reference to the old figure context so we can release it
                #old_fig = chart.figure #Likely prevents error in the following line should the fig be non-local
                figList.append(plt.figure(figsize=(8,5.5)))
    
                # remove the Axes from it's original Figure context
                #chart.remove() #Places the ax as an independent variable
                #exportList[exportCount].add_subplot(chart)
    
                # set the pointer from the Axes to the new figure
                chart.figure = figList[exportIndex] #Sets new home for the ax
    
                # add the Axes to the registry of axes for the figure
                figList[exportIndex].axes.append(chart) #adds it to the chart
                figList[exportIndex].add_axes(chart) #likely queues it into place
    
                pdf.savefig(figList[exportIndex]) #prevents overlap that occurred when attempting to output a single figure
    
                exportIndex = exportIndex + 1
                
            #outputFile = fd.open(file_path.name, 'wb')
            #file_path.write(pdf)
            exportText.config(state='normal')
            exportText.delete('1.0', tk.END)
            exportText.config(state='disabled')
    
            exportList.clear()
    
            exportCount = 0
            
    startUp()
    pairDateData(dataOEE_Page, dataTS_Page)
    todaysDate = dt.datetime.today()
    beginningOfInterval = dt.datetime.today()
    todaysDate = todaysDate.strftime("%Y-%m-%d")
    passedDate = todaysDate
    narrowDates(passedDate, 'Tier 3')
    exportList = []
    #axList = []
    exportCount = 0
    #figIndex = 111
    #lastX = ''
    #lastY = ''
    #lastY2 = ''
    
    columnList = pairedDate.columns.values.tolist()
    columnList.insert(0, '')

    root = tk.Tk()
    root.title('Data Analysis for Twin Screw')
    #width, height = root.winfo_screenwidth(), root.winfo_screenheight()
    root.state('zoomed')
    
    #Sets up the menu bar for tabs to be mounted on
    selectMenu = tk.Menu(root)
    root.config(menu = selectMenu)
        
    #Makes dropdown menu for the File tab
    fileMenu = tk.Menu(selectMenu)
    fileMenu.add_command(label="Help")
    fileMenu.add_command(label="Save as", command = lambda: saveCharts(exportList))#will call a saving function
    fileMenu.add_command(label="Print") #call both a saving and printing function
    fileMenu.add_command(label="Exit", command = root.destroy)
    selectMenu.add_cascade(label="File", menu=fileMenu)
    
    x = tk.StringVar(root)
    y1 = tk.StringVar(root)
    y2 = tk.StringVar(root)
    select = tk.StringVar(root)
    corrCheck = tk.BooleanVar(root)

#Where all the labels for dropdown menus and the calender objects will go
    labelDate = tk.Label(root, text = "Date")
    labelX = tk.Label(root, text = "X-Variable")
    labelY = tk.Label(root, text = "1. Y-Variable")
    labelY2 = tk.Label(root, text = "2. Y-Variable")
    
    labelDate.grid(row = 0, column = 0, sticky = 'W')
    labelX.grid(row = 1, column = 0, sticky = 'W')
    labelY.grid(row = 2, column = 0, sticky = 'W')
    labelY2.grid(row = 3, column = 0, sticky = 'W')
    
    #Makes calender widget, and dropdown menus for variables
    dateEntry = cal.DateEntry(root)
    selectX = tk.OptionMenu(root, x, *columnList)
    selectY1 = tk.OptionMenu(root, y1, *columnList)
    selectY2 = tk.OptionMenu(root, y2, *columnList)

    dateEntry.grid(row = 0, column = 1, sticky = 'W', padx = 10)
    selectX.grid(row = 1, column= 1, sticky='W')
    selectY1.grid(row=2, column=1, sticky='W')
    selectY2.grid(row=3, column=1, sticky='W')
    
    #Builds frame for image
    graphFrame = tk.Frame(root)
    graphFrame.grid(row = 0, column = 2, rowspan = 4, columnspan = 5, padx=20) #Spans column 2 to 6 and row 0 to 3
    
    #Generates graph for the frame and is capable of being redrawn
    #tableData = tableData.sort_values(by='Efficiency') #Sorts values to be more easily understandable by graph view
    fig = plt.figure(figsize=(8, 5.5))
    fig.add_subplot()
    
    canvas = FigureCanvasTkAgg(fig, master=graphFrame)  # A tk.DrawingArea.
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    
    toolbar = NavigationToolbar2Tk(canvas, graphFrame)
    toolbar.update()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    
    #Checks for correlation between numbers through sorting by x-axis (if needed)
    correlationCheck = tk.Checkbutton(root, variable = corrCheck, text="Find Correlation")
    correlationCheck.grid(row = 4, column = 5, sticky = 'E')
    
    #Builds in the buttons to make the application function (Will likely lose some functionality in the final form)
    tierList = ['Tier 2', 'Tier 3', 'Tier 5']
    exportButton = tk.Button(root, text = "Export to PDF", command = lambda: exportReady(lastX, lastY, lastY2, plt.gca()))
    selectTier = tk.OptionMenu(root, select, *tierList)
    applyButton = tk.Button(root, text = "Apply", command = lambda: applyFunction(dateEntry.get_date().strftime("%Y-%m-%d"), x.get(), y1.get(), y2.get(), select.get(), corrCheck.get()))
    #refreshButton = tk.Button(root, text = \"Refresh\")
    #removeButton = tk.Button(root, text = \"Remove Row\")
    select.set('Tier 3')
    
    applyButton.grid(row = 4, column = 0) #Will recieve commands in final version
    exportButton.grid(row = 4, column = 1, sticky = 'W')
    selectTier.grid(row=4, column=2, sticky='W')
    #refreshButton.grid(row = 4, column = 3, sticky = 'E')
    #removeButton.grid(row = 4, column = 4, sticky = 'E')
    
    #Draws frame for datatable
    datatableFrame = tk.Frame(root)
    datatableFrame.grid(row = 6, column = 0, columnspan=7, sticky='WE')
    
    #Draws in Datatable
    table = ps.Table(datatableFrame, dataframe = tableData, showtoolbar= False)
    table.showIndex()
    table.show()
    
    #Draws label for export scrollable area
    corrLabel = tk.Label(root, text='Correlations:')
    corrLabel.grid(row=6, column=7)
    
    #Draws Correlation Frame
    corrFrame = tk.Frame(root)
    corrFrame.grid(row = 6, column = 8, sticky='WE')
    
    corrScroll = tk.Scrollbar(corrFrame)
    corrScroll.pack(side=tk.RIGHT)
    
    corrText = tk.Text(corrFrame, height=10, width=75)
    corrText.pack(side=tk.LEFT)
    corrText.config(state='disabled')

    #Draws label for export scrollable area
    exportLabel = tk.Label(root, text='Exporting:')
    exportLabel.grid(row=0, column=7)

    #Scrollable frame for export charts
    exportFrame = tk.Frame(root)
    exportFrame.grid(row = 0, column = 8, sticky = 'WE')

    exportScroll = tk.Scrollbar(exportFrame)
    exportScroll.pack(side=tk.RIGHT)

    exportText = tk.Text(exportFrame, height=4, width=75)
    exportText.pack(side=tk.LEFT)
    #exportText.config(state='readonly')
    exportText.config(state='disabled')
    
    #Activates Application\n"
    root.mainloop()




