#Code in Progress
#Last Update: 8/27/2014 8:36
#Combination of all components needed to automate PAOP reports

import os, shutil,os.path, datetime, xlrd, xlsxwriter, datetime, calendar,string

def FileCopy(OrigDir, NewDir):
    #This function will copy files from one folder(OrigDir) to a new folder(NewDir).
    #This is need to take the automated SQL reports from an FTP folder to a local folder.
    dirOrigDir = os.listdir(OrigDir)

    for file_name in dirOrigDir:
        full_file_name = os.path.join(OrigDir, file_name)
        if (os.path.isfile(full_file_name)):
            shutil.copy(full_file_name, NewDir)
            
    print 'File Copied'


def FileRenameModDate(directory, filetype):
    #This function will rename a specified file(filetype) to include the month it was last modified.
    #If the file is not available in the specified folder(directory) it will return that the file is not available.
    #This is needed so that we will know which months location report the file is. Also needed so the new files
    #copied over will not overwrite the previous month.
    dirLocPAOP = directory
    os.chdir(dirLocPAOP)
    Tempallfiles = os.listdir(dirLocPAOP)
    Lallfiles= []
    #Modifies the original list of items in directory so folders are excluded.
    for x in Tempallfiles:
        y = x[-4]
        if y=='.':
            Lallfiles.append(x)

    for filename in Lallfiles:
        if not filetype in filename:
            print 'File not Available'
            continue
        Ttime = os.path.getmtime(filename)
        Tdate =datetime.datetime.fromtimestamp(Ttime)
        moddate = Tdate.strftime('%Y%m')
        temp_name = filetype + "_" + str(moddate) + '.xlsx'
        os.rename(filename,temp_name)
        print 'File Renamed'
        
def ReportsAvail(directory):
    #Returns a string of just the numeric month date in file names in the directory.
    dirLocPAOP = directory
    os.chdir(dirLocPAOP)
    Lallfiles = os.listdir(dirLocPAOP)

    Lmonths= []
    for x in Lallfiles:
        y = x[-7]+x[-6]
        Lmonths.append(y)

    print Lmonths





# Open the spreadsheet.

#Add a way to create variable for filename below





#Variable Identification
#FileCopy function
FileCopy("C:\Users\sjolly\Desktop\Test Folder 1","C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14")
#FileRenameModDate function
FileRenameModDate("C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14","DufferinPeel_Teaching_Location_PAOP")
#FileCopy function
FileCopy("C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14","C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14\Individual Months")
#ReportsAvail function
ReportsAvail("C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14\Individual Months")
