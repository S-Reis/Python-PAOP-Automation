#Complete 8/27/3014 8:35
#Returns a string of just the numeric month date in file names in the directory.
import os.path, string
def ReportsAvail(directory):
    dirLocPAOP = directory
    os.chdir(dirLocPAOP)
    Lallfiles = os.listdir(dirLocPAOP)

    Lmonths= []
    for x in Lallfiles:
        y = x[-7]+x[-6]
        Lmonths.append(y)

    print Lmonths



ReportsAvail("C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14\Individual Months")
