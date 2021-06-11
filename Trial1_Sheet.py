# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
#imports for behind the scenes computations and organization, and import of the excel sheet into the code

import pandas as pd
import numpy as np
import tkinter as tk
import datetime as dt

xlsx = pd.ExcelFile("Data5_23_Correct.xlsx", engine = "openpyxl")
dataOEE_Page = pd.read_excel(xlsx,'Twin Screw OEE')
dataTS_Page = pd.read_excel(xlsx, "Twin Screw ")


# %%
#This section of code takes filler sections of the excel spreadsheet from both tabs and removes them

dataOEE_Page.replace('x', np.nan, inplace= True)
dataTS_Page.replace('x', np.nan, inplace = True)

dataOEE_Page.drop(labels = 0, axis = 0, inplace = True)
dataTS_Page.drop(labels = 0, axis = 0, inplace = True)

#print(dataOEE_Page)


# %%
#Names the indexes for the columns of the dataframes

dataOEE_Page.columns = ['OEE_Date', 'Lbs Produced', '%OEE', 'Labor Hours', 'Lb/Man Hour', 'Pails Filled', 'Pails/Man Hour', 'Removal', 'Week of', 'Weekly %OEE']
dataTS_Page.columns = ['TS_Date', 'Num of Ops', 'Total Minutes', 'Maintenance', 'Making Powder', 'Making Prepolymer', 'Startup', 'Breaks', 'LBS Produced', 'Pails Filled', 'Unnamed', 'Efficiency', 'Unnamed', 'Unnamed', 'Unnamed']

#dataOEE_Page


# %%
#This section trims the fat off the data set and removes entries that are effectively empty from both the OEE and Efficiency
#tab. 

dataTS_Page = dataTS_Page.dropna(how = 'all', axis = 1)
dataTS_Page = dataTS_Page.dropna(subset=['Num of Ops'])

dataOEE_Page = dataOEE_Page.dropna(subset=['Labor Hours'])
dataOEE_Page = dataOEE_Page.dropna(how = 'all', axis = 1)

#dataOEE_Page


# %%
#This section of code takes the first 13 entries of the Twin Screw OEE tab of the excel sheet and removes them from the data pool
#This effectively removes the logic error that occurred due to the lack of the # of operators on the - tab, and realigns the dates
#Following this section, data that doesn't have a matching pair datewise will be removed from each respective dataframe

dataOEE_Page.drop(dataOEE_Page.index[range(0,13)], axis = 0, inplace =True)

dataOEE_Page = dataOEE_Page.reset_index(drop = True)
dataTS_Page = dataTS_Page.reset_index(drop = True)


# %%
#Using the dates from the trimmed lists, then comparing them to make a condensed list of data points
#This resultant list should contain all eligible data points
#From this comparison dataframe we should be able to pull data that correlates directly
#Should only be done in order to compare data cross sheet due to lack of variables

dateCheck = [dataOEE_Page['OEE_Date'], dataTS_Page['TS_Date']] #Makes a list of the columns from either dataFrame
dateCheck = pd.concat(dateCheck, axis=1, ignore_index=False, verify_integrity=False) #Combines the dataCheck list into a dataFrame
dateCheck['Removal_OEE'] = np.nan #Adds columns of nan values
dateCheck['Removal_TS'] = np.nan 
#dateCheck.columns = ['OEE_Date', 'TS_Date','Removal_OEE','Removal_TS'] #Renames the columns of the dataframe

#Checks whether or not the dates correspond to each other throughout either column, result is recorded in respective removal column
#Recorded as True if there is matching pair, False if there isn't
dateCheck['Removal_OEE'] = (dateCheck['OEE_Date'].isin(dateCheck['TS_Date']))
dateCheck['Removal_TS'] = (dateCheck['TS_Date'].isin(dateCheck['OEE_Date']))

OEEindex = dateCheck[dateCheck['Removal_OEE'] == False].index #Identifies the exact index of the date cell needing removed
TSindex = dateCheck[dateCheck['Removal_TS'] == False].index

#print(OEEindex)
#print(TSindex)

#Drops erroneous cells from both data sets
dataOEE_Page = dataOEE_Page.drop(OEEindex, inplace=False)
dataTS_Page = dataTS_Page.drop(TSindex,inplace=False)

dataOEE_Page = dataOEE_Page.reset_index(drop = True)
dataTS_Page = dataTS_Page.reset_index(drop = True)


# %%
#This section will identify the percent values that are based as a decimal in the updated dataframes
#From here the new data is passed down onto future methods for stronger correlations

dataOEE_Page.loc[dataOEE_Page['%OEE'] < 1, '%OEE'] = dataOEE_Page['%OEE']*100
dataTS_Page.loc[dataTS_Page['Efficiency'] < 1, 'Efficiency'] = dataTS_Page['Efficiency']*100


# %%
#This section will include data manipulations that are in accordance with the value of period
#This will be to also organize data for the weekly and monthly values, won't display values from outside the month
#Possibly make this a setting: Include last month, exclude last month

testPeriod = 1
testInput = [dataOEE_Page['%OEE'], dataTS_Page['Efficiency'], dataOEE_Page['OEE_Date'], dataTS_Page['TS_Date']]
testInput = pd.concat(testInput, axis = 1, ignore_index = False)
#testInput

if testPeriod == 1:
    print('')
elif testPeriod == 7:
    print('')
elif testPeriod == 30:
    print('')
else:
    print('Error: Resetting. Please attempt operation again.')
    testPeriod = 1 


# %%
#This section of code effectively pairs the the %OEE numbers and Efficiency numbers with other vars to collect the data of both sets
#Will be working to make the user have input on stuff like this to make it easier to use

OEEdata = dataOEE_Page['%OEE']
TS_PgData = dataTS_Page[['Num of Ops','Efficiency', 'Total Minutes']]
#OEEdata
#TS_PgData

effVsOEE_data = [OEEdata, TS_PgData]
effVsOEE_data = pd.concat(effVsOEE_data, axis=1, ignore_index=True, verify_integrity=False)

effVsOEE_data.columns = ['%OEE', 'Num of Ops', 'Efficiency', "Total Minutes"]
#effVsOEE_data = effVsOEE_data.sort_values(by = ['Num of Ops'])
#effVsOEE_data


# %%
#Assigns type data to the values in the columns, certain columns are uninterruptable due to the 

dataOEE_Page = dataOEE_Page.astype({"Lbs Produced": np.float64, "%OEE": np.float64, "Labor Hours": np.float64, "Lb/Man Hour": np.float64, "Pails Filled": np.float64, "Pails/Man Hour": np.float64, 'Weekly %OEE': np.float64})
dataTS_Page = dataTS_Page.astype({'Num of Ops': np.float64, 'Total Minutes': np.float64, 'Maintenance': np.float64, 'Efficiency': np.float64})
effVsOEE_data = effVsOEE_data.astype({"Num of Ops": np.float64, "%OEE": np.float64, "Efficiency": np.float64, 'Total Minutes': np.float64})


# %%



# %%
#This section attains attributes of the dataset: Average, correlation values, etc.
#Will be changed to a method that can change to weekly, monthly, and yearly

#Is a list of the values read from the files, just the averages and no index names are attached
listNames = effVsOEE_data.columns.tolist()
meanAllTimeCalc = [(effVsOEE_data['%OEE'].sum())/(effVsOEE_data['%OEE'].count()+1), (effVsOEE_data['Num of Ops'].sum())/(effVsOEE_data['Num of Ops'].count()+1), (effVsOEE_data['Efficiency'].sum())/(effVsOEE_data['Efficiency'].count()+1), (effVsOEE_data['Total Minutes'].sum())/(effVsOEE_data['Total Minutes'].count()+1)]

#print(meanAllTimeCalc)

#Identifies the period of time that variables are being viewed over, and then identifies and relays how data point are being used
#if period == 1:
    #print("The daily values are being displayed with 1 dataset being used for display.")
#elif period == 7:
    #print("The weekly values are being displayed with -[variable for true # in dataset]- dataset(s) being used for display.")
#elif period == 30:
    #print("The monthly values are being displayed with -[variable for true # in dataset]- dataset(s) being used for display.")

print(listNames)

#Finds averages of all the values of the dataset
for x in meanAllTimeCalc:
    if x == meanAllTimeCalc[0]:
        print("The average value of " + listNames[0] + " is: " + round(x, 3).astype('str') + '%')
    elif x == meanAllTimeCalc[1]:
        print("The average value of " + listNames[1] + " is: " + round(x, 2).astype('str'))
    elif x == meanAllTimeCalc[2]:
        print("The average value of " + listNames[2] + " is: " + round(x,3).astype('str') + '%')
    else:
        print("The average value of " + listNames[3] + " is: " + round(x,2).astype('str'))

#Finds correlations between respective number of the sheet

print(effVsOEE_data.corr()[["Efficiency"]])
print(effVsOEE_data.corr()[["%OEE"]])


# %%
#Import libraries for graphing the dataframes properly as well a method for plotting the data. 
#This section of code could definitely use more work so as to improve the visibility of the x-axis.
#Possibly look to group numbers of operators and then begin to average them for the all time view.

#import seaborn as sns
import matplotlib.pyplot as plt

#fig, axs = plt.subplots(ncols=1, figsize=(30,5))
#figure = plt.figure(figsize = (6,5))
#ax = figure.add_subplot(111)

#ax = plt.gca()
#effVsOEE_data.set_index('Num of Ops', inplace=True)
#effVsOEE_data.groupby('Num of Ops')['Efficiency'].plot(legend = True)
#plt.x_axis(label = 'date')
#plt.axis([0,10,0,100])
#plt.show()
#plt.close()

#effVsOEE_data.groupby('Num of Ops')['%OEE'].plot(legend = True)
#plt.show()
#plt.close()
#sns.pointplot(x="Num of Ops", y="Efficiency",data=effVsOEE_data, ax=axs[0])


# %%
#Will become the GUI in time. Will work off the code above this.

import pandas as pd
import numpy as np
import tkinter as tk
import tkcalendar as cal
import datetime
from matplotlib.backends.backend_pdf import PdfPages

class Application(tk.Frame):
    global userInputX
    global userInputY
    global dateSelect
    global period
    #global varsMenu

    userInputX = ''
    userInputY = ''
    dateSelect = ''
    #varsMenu = 0
    period = 1


    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        master.geometry('500x500')
        master.title('Graphing for Twin Screw')
        self.create_widgets()
        #self.dataCleaning()
        

    def create_widgets(self): #Commented out buttons were removed due to being elements deemed unnecessary

        #Sets up the menu bar for tabs to be mounted on
        self.selectMenu = tk.Menu(self.master)
        self.master.config(menu = self.selectMenu)
        
        #Makes dropdown menu for the File tab
        self.fileMenu = tk.Menu(self.selectMenu)
        self.fileMenu.add_command(label="Save As", command=self.writeToPDFtest)
        self.fileMenu.add_command(label="Print", command=self.printTest)
        self.fileMenu.add_command(label="Exit", command=self.master.destroy)
        self.selectMenu.add_cascade(label="File", menu=self.fileMenu)
        
        #Makes dropdown menu for the Data tab
        self.editMenu = tk.Menu(self.selectMenu)
        self.editMenu.add_command(label="Change Vars", command=self.changeVars)
        self.editMenu.add_command(label="Pull up Data")
        self.editMenu.add_command(label="etc.")
        self.selectMenu.add_cascade(label="Data", menu=self.editMenu)

        #Make dropdown menu for generation of Tier 2 and Tier 3 
        self.reportMenu = tk.Menu(self.selectMenu)
        self.reportMenu.add_command(label="Tier 2", command=self.changeReport2)
        self.reportMenu.add_command(label="Tier 3", command=self.changeReport3)
        #self.reportMenu.add_command(label="etc.")
        self.selectMenu.add_cascade(label="Generate", menu=self.reportMenu)

        self.calendar = cal.DateEntry(self.master)
        self.calendar.pack()

        self.lookUpButton = tk.Button(self.master, text = 'Look Up', command = self.findDate)
        self.lookUpButton.pack()

    def writeToPDFtest(self):
        print("This function would take the displayed chart and put it on a pdf")

    def printTest(self):
        print("This function would take the saved pdf file and print it")

    def chartFunction(self):
        print("Displaying chart")

    def findDate(self):
        dateSelect = self.calendar.get()
        print(dateSelect)

    #Will collect the weekly data and put them on a graph
    def changeReport2(self):
        period = 7
    
    #Will collect the monthly data and put them on a scatterplot due to the number of data points (Up for debate)
    def changeReport3(self):
        period = 30

    #Retrieves the chosen values from the variable selector, and then checks to make sure variables have valid input
    def retrieveVars(self):
        print("Retrieving Values")

    def changeVars(self):
        varsMenu = tk.Tk()
        varsMenu.title("Variables Select Menu")
        varsMenu.geometry('500x500')

        userInput1 = tk.StringVar()
        userInput2 = tk.StringVar()

        #Sets the variables to make sure there's a single var selected
        #userInput1 = userInput1.set('Date')
        #userInput2 = userInput2.set('Lbs Produced')

        def applyAndClose():
            userX = False
            userY = False
            tempX = userInput1.get()
            tempY = userInput2.get()

            for x in radioButtListNames:

                if x == tempX:
                    userX = True
                if x == tempY:
                    userY = True

            if userX and userY:
                varsMenu.destroy()

        #canvas = tk.Canvas(varsMenu, borderwidth=0, background="#ffffff")
        frameX = tk.Frame(varsMenu, background="#ffffff")
        frameY = tk.Frame(varsMenu, background="#ffffff")
        frameX.pack(side='left')
        frameY.pack(side='right')
        #vsb = tk.Scrollbar(varsMenu, orient="vertical", command=canvas.yview)
        #canvas.configure(yscrollcommand=vsb.set)

        #vsb.pack(side = 'right', fill="y")
        #canvas.pack(fill="both", expand=True)
        #canvas.create_window((2,4), window=frameX, anchor="nw")
        #canvas.create_window((2,4), window=frameY, anchor="ne")

        #frameX.bind("<Configure>", lambda event, canvas=canvas: canvas.configure(scrollregion=canvas.bbox("left")))
        #frameY.bind("<Configure>", lambda event, canvas=canvas: canvas.configure(scrollregion=canvas.bbox("right")))

        radioButtListNames = ['Date', 'Lbs Produced', '%OEE', 'Labor Hours', 'Lb/Man Hour', 'Pails Filled', 'Pails/Man Hour', 'Num of Ops','Total Minutes', 'Maintenance', 'Making Powder', 'Making Prepolymer', 'Startup', 'Breaks', 'LBS Produced', 'Efficiency']
        radioButtList_x = []

        label_x = tk.Label(frameX, text = 'X-Variable')
        label_x.pack()
        buttonCount = 0

        for x in radioButtListNames:
            temp = tk.Radiobutton(frameX, text = x, variable = userInput1, value = x)
            radioButtList_x.append(temp)
            radioButtList_x[buttonCount].pack()
            buttonCount = buttonCount+1

        

        label_y = tk.Label(frameY, text = 'Y-Variable')
        label_y.pack()
        radioButtList_y = []

        buttonCount = 0

        for y in radioButtListNames:
            #print("y: " + y)
            temp = tk.Radiobutton(frameY, text = y, variable = userInput2, value = y)
            radioButtList_y.append(temp)
            radioButtList_y[buttonCount].pack()
            buttonCount = buttonCount+1

        #The button that closes the window and signals for the values to be collected
        closeButton = tk.Button(frameY, text = 'Apply and Close', command = applyAndClose)
        closeButton.pack(side = 'bottom')




root = tk.Tk()
app = Application(master=root)
app.mainloop()


# %%



# %%
def writeToPDF(self):
    with PdfPages('multipage_pdf.pdf') as pdf:
        plt.figure(figsize=(3, 3))
        plt.plot(range(7), [3, 1, 4, 1, 5, 9, 2], 'r-o')
        plt.title('Page One')
        pdf.savefig()  # saves the current figure into a pdf page
        plt.close()

        # if LaTeX is not installed or error caught, change to `False`
        plt.rcParams['text.usetex'] = False
        plt.figure(figsize=(8, 6))
        x = np.arange(0, 5, 0.1)
        plt.plot(x, np.sin(x), 'b-')
        plt.title('Page Two')
        pdf.attach_note("plot of sin(x)")  # attach metadata (as pdf note) to page
        pdf.savefig()
        plt.close()

        plt.rcParams['text.usetex'] = False
        fig = plt.figure(figsize=(4, 5))
        plt.plot(x, x ** 2, 'ko')
        plt.title('Page Three')
        pdf.savefig(fig)  # or you can pass a Figure object to pdf.savefig
        plt.close()

        # We can also set the file's metadata via the PdfPages object:
        d = pdf.infodict()
        d['Title'] = 'Multipage PDF Example'
        d['Author'] = 'Jouni K. Sepp\xe4nen'
        d['Subject'] = 'How to create a multipage pdf file and set its metadata'
        d['Keywords'] = 'PdfPages multipage keywords author title subject'
        d['CreationDate'] = datetime.datetime.today()
        


