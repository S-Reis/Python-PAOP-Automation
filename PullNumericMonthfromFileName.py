#Testing to return just the numeric month date in file name

import os.path, string

directory = "C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14\Individual Months"

dirLocPAOP = directory
os.chdir(dirLocPAOP)
Lallfiles = os.listdir(dirLocPAOP)

for x in Lallfiles:
    y = x.rstrip('5')
    print y
    print y[-3:]
