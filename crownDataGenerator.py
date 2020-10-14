import os, re, openpyxl
#setting intial condition for excel sheet
column_y=1
row_x = 2 
#variables
password=''

#user input
pathForConfig = input('Enter absolute path for config Files: ')
pathForOutputExcel = input('Enter absolute path for excel sheet: ')
BFXX= input('Enter your FC example-BFI3 ')

#setting directories and getting config files
os.chdir(pathForConfig)
filesList = os.listdir()

# writing header for the excel sheet
workBook = openpyxl.Workbook()
sheet = workBook['Sheet']
sheet['A1']="Display Name"
sheet['B1']="Description"
sheet['C1']="Password"

macAddress = re.compile("([0-9A-F]{2}-){5}[0-9A-F]{2}") #regex object for mac-address

for fileName in filesList:
    column_y =1 #reset column
    match = macAddress.search(fileName) #regex grabbing mac-address
    if match is not None:
        #display name field
        sheet.cell(row=row_x, column=column_y).value = 'ASIN456_'+BFXX.upper()+'_'+match.group(0)[12:14]+"_"+match.group(0)[15:18]
        column_y=column_y+1

        #description field
        sheet.cell(row=row_x, column=column_y).value = 'Config generated for '+match.group(0)+" at "+BFXX.upper()
        column_y = column_y+1

        #get the password out of config
        configFile = open(fileName)
        words = configFile.readlines()
        passwordRaw = words[12]
        password= passwordRaw[9:len(passwordRaw)-2]

        #password field
        sheet.cell(row=row_x, column=column_y).value = password
        workBook.save(pathForOutputExcel+".\\data.xlsx")
    row_x=row_x+1 #increase row



