import xlrd
import xlsxwriter
import datetime
import calendar

# Open the spreadsheet.
oWorkbook = xlrd.open_workbook('Test Location PAOP.xls')

# Get the sheet we want.
oSheet = oWorkbook.sheet_by_name('Sheet1')

# Create a dictionary to store the results.
dLocationPAOP = {}
vLocationDays= oSheet.cell(1,0).value

for iRow in xrange(5, oSheet.nrows):
    # Pull out the values we are interested in mapping.
    stLocationCode = oSheet.cell(iRow, 0).value
    stLocationName = oSheet.cell(iRow, 1).value
    stSubPool = oSheet.cell(iRow, 2).value
    stEmployees = oSheet.cell(iRow, 3).value
    stTotalJobs = oSheet.cell(iRow, 4).value
    stAbsences = oSheet.cell(iRow, 5).value
    stFilled = oSheet.cell(iRow, 6).value
    stNSR = oSheet.cell(iRow, 7).value
    stUnfilled = oSheet.cell(iRow, 8).value
    stSchoolType = oSheet.cell(iRow, 9).value
    stRegion = oSheet.cell(iRow, 10).value
    dLocationPAOP[stLocationCode]= stLocationName,stSubPool,stEmployees,stTotalJobs,stAbsences,stFilled,stNSR,stUnfilled,stSchoolType,stRegion,vLocationDays

print len(dLocationPAOP)


# Create a workbook
Cworkbook = xlsxwriter.Workbook('TestLocationPAOP1.xlsx')

#Create Sheet Month Name
calendar.setfirstweekday(calendar.SUNDAY)
month = datetime.datetime.now().strftime("%B")

#Create workbook sheet
Cworksheet = Cworkbook.add_worksheet(month)


#Header formats
HdrForm = Cworkbook.add_format()
HdrForm.set_bg_color('#66B2FF')
HdrForm.set_bold()
HdrForm.set_border(style=1)
#Border formats
ShBorder = Cworkbook.add_format()
ShBorder.set_border(style=1)
#Percentage formats
PercentForm = Cworkbook.add_format({'num_format': '0.00%', 'border': 1})

#Set Loc Groups
Region1= "Y - Area East"
Region2= "Y - Area North"
Region3= "Y - Area West"
Type1="Y - Type High School"
Type2= "Y - Type Elementary/Middle"


#Create Header
for iColumn, stHeader in enumerate(['Code','Location','SubPool', 'Emps', 'TotalJobs','Absences', 'Filled',
                                    'NSR', 'Unfilled', 'SchType', 'Region', 'Days', 'Fill Rate', 'AbsenceRate',
                                    Region1,Region2,Region3,Type1,Type1]):
    Cworksheet.write(0, iColumn, stHeader, HdrForm)


#Create Main Table
for iRow, stLocationCode in enumerate(dLocationPAOP):
    #Unpack Data
    stLocationName,stSubPool,stEmployees,stTotalJobs,stAbsences,stFilled,stNSR,stUnfilled,stSchoolType,stRegion,vLocationDays = dLocationPAOP[stLocationCode]
    Cworksheet.write(iRow + 1,0,stLocationCode,ShBorder)
    Cworksheet.write(iRow + 1,1,stLocationName,ShBorder)
    Cworksheet.write(iRow + 1,2,stSubPool,ShBorder)
    Cworksheet.write(iRow + 1,3,stEmployees,ShBorder)
    Cworksheet.write(iRow + 1,4,stTotalJobs,ShBorder)
    Cworksheet.write(iRow + 1,5,stAbsences,ShBorder)
    Cworksheet.write(iRow + 1,6,stFilled,ShBorder)
    Cworksheet.write(iRow + 1,7,stNSR,ShBorder)
    Cworksheet.write(iRow + 1,8,stUnfilled,ShBorder)
    Cworksheet.write(iRow + 1,9,stSchoolType,ShBorder)
    Cworksheet.write(iRow + 1,10,stRegion,ShBorder)
    Cworksheet.write(iRow + 1,11,vLocationDays,ShBorder)

    #Create Fill Rate and Absence Rate Columns
    Fillrate = '=IF(E{}>0,(G{}+H{})/E{},0)'.format(*([iRow + 2] * 12))
    Cworksheet.write_formula(iRow + 1, 12, Fillrate, PercentForm)
    Absrate = '=IF(D{}>0,(F{}/L{})/D{},0)'.format(*([iRow + 2] * 13))
    Cworksheet.write_formula(iRow + 1, 13, Absrate, PercentForm)

    #Create Location Group Columns
    FRegion1= '=IF(AND( E{}>0,J{} ="'+Region1+'"),M{})'
    ColRegion1 = FRegion1.format(*([iRow + 2] * 14))
    Cworksheet.write_formula(iRow + 1, 14, ColRegion1,PercentForm)

    FRegion2= '=IF(AND( E{}>0,J{} ="'+Region2+'"),M{})'
    ColRegion2 = FRegion2.format(*([iRow + 2] * 15))
    Cworksheet.write_formula(iRow + 1, 15, ColRegion2,PercentForm)

    FRegion3= '=IF(AND( E{}>0,J{} ="'+Region3+'"),M{})'
    ColRegion3 = FRegion3.format(*([iRow + 2] * 16))
    Cworksheet.write_formula(iRow + 1, 16, ColRegion3,PercentForm)

    FType1= '=IF(AND( E{}>0,K{} ="'+Type1+'"),M{})'
    ColType1 = FType1.format(*([iRow + 2] * 17))
    Cworksheet.write_formula(iRow + 1, 17, ColType1,PercentForm)

    FType2= '=IF(AND( E{}>0,K{} ="'+Type2+'"),M{})'
    ColType1 = FType1.format(*([iRow + 2] * 18))
    Cworksheet.write_formula(iRow + 1, 18, ColType1,PercentForm)  


#Create Total Row
SnRows=str(len(dLocationPAOP)+1)
SnRowsTot=str(len(dLocationPAOP)+2)
InRows=len(dLocationPAOP)
Cworksheet.write(InRows+1, 0, '',ShBorder)
Cworksheet.write(InRows+1, 1,'',ShBorder)
Cworksheet.write(InRows+1, 2, '=SUM(C2:C'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 3, '=SUM(D2:D'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 4, '=SUM(E2:E'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 5, '=SUM(F2:F'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 6, '=SUM(G2:G'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 7, '=SUM(H2:H'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 8, '=SUM(I2:I'+SnRows+')', ShBorder)
Cworksheet.write(InRows+1, 9, '',ShBorder)
Cworksheet.write(InRows+1, 10,'',ShBorder)
Cworksheet.write(InRows+1, 11,vLocationDays,ShBorder)
Cworksheet.write(InRows+1, 12, '=IF(E'+SnRowsTot+'>0,(G'+SnRowsTot+'+H'+SnRowsTot+')/E'+SnRowsTot+',0)', PercentForm)
Cworksheet.write(InRows+1, 13, '=IF(D'+SnRowsTot+'>0,(F'+SnRowsTot+'/L'+SnRowsTot+')/D'+SnRowsTot+',0)', PercentForm)
Cworksheet.write(InRows+1, 14, '=AVERAGE(O2:O'+SnRows+')', PercentForm)
Cworksheet.write(InRows+1, 15, '=AVERAGE(P2:P'+SnRows+')', PercentForm)
Cworksheet.write(InRows+1, 16, '=AVERAGE(Q2:Q'+SnRows+')', PercentForm)
Cworksheet.write(InRows+1, 17, '=AVERAGE(R2:R'+SnRows+')', PercentForm)
Cworksheet.write(InRows+1, 18, '=AVERAGE(S2:S'+SnRows+')', PercentForm)

    
Cworkbook.close()

print 'Excel Created'


















