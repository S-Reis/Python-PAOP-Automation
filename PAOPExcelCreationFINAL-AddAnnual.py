#This is the final completed version that can be succesfully run. It includes additional of the Annual results
#Last modified: 8/27/2014 16:57


import os, shutil,os.path, datetime, xlrd, xlsxwriter, datetime, calendar,string


def FileReadandCreate(dirLocPAOP,FiletoWrite,AnnualFile,SheettoRead,Region1,Region2,Region3,Type1,Type2):
    #This function will open up files in a speicfied directory, read them into a dictioanry and unpack them
    #into a newly created excel file. Each original file will have its own sheet in the excel file named by month.

    #This series will change the working directory, and create a list of the files in that directory then sort them by modfied date/time.          
    os.chdir(dirLocPAOP)
    Tempallfiles = os.listdir(dirLocPAOP)
    Lallfiles= []
    for x in Tempallfiles:
        y = x[-6]
        if y!='l':
            Lallfiles.append(x)
    Lallfiles.sort(key=lambda x: os.stat(os.path.join(dirLocPAOP, x)).st_mtime)

    # Create a workbook
    Cworkbook = xlsxwriter.Workbook(FiletoWrite)
    #Create months for sheet names
    LMonths =['Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun']

    oWorkbook = xlrd.open_workbook(AnnualFile)
    oSheet = oWorkbook.sheet_by_name(SheettoRead)

    #Create empty dictonary and specify which cell contains the number of days
    dAnnualLocationPAOP = {}
    vAnnualLocationDays= oSheet.cell(1,0).value

    for iRow in xrange(5, oSheet.nrows):
    #This loop will run through the main table rows and load them into the empty dictionary
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
        dAnnualLocationPAOP[stLocationCode]= stLocationName,stSubPool,stEmployees,stTotalJobs,stAbsences,stFilled,stNSR,stUnfilled,stSchoolType,stRegion,vAnnualLocationDays

    print len(dAnnualLocationPAOP)

    Cworksheet = Cworkbook.add_worksheet('Annual')

    #Create workbook sheet with month name

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

    #Create Header
    for iColumn, stHeader in enumerate(['Code','Location','SubPool', 'Emps', 'TotalJobs','Absences', 'Filled',
                                        'NSR', 'Unfilled', 'SchType', 'Region', 'Days', 'Fill Rate', 'AbsenceRate',
                                        Region1,Region2,Region3,Type1,Type1]):
        Cworksheet.write(0, iColumn, stHeader, HdrForm)

    #Create Main Table
    for iRow, stLocationCode in enumerate(dAnnualLocationPAOP):
        #Unpack Data
        stLocationName,stSubPool,stEmployees,stTotalJobs,stAbsences,stFilled,stNSR,stUnfilled,stSchoolType,stRegion,vAnnualLocationDays = dAnnualLocationPAOP[stLocationCode]
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
        Cworksheet.write(iRow + 1,11,vAnnualLocationDays,ShBorder)

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
    SnRows=str(len(dAnnualLocationPAOP)+1)
    SnRowsTot=str(len(dAnnualLocationPAOP)+2)
    InRows=len(dAnnualLocationPAOP)
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
    Cworksheet.write(InRows+1, 11,vAnnualLocationDays,ShBorder)
    Cworksheet.write(InRows+1, 12, '=IF(E'+SnRowsTot+'>0,(G'+SnRowsTot+'+H'+SnRowsTot+')/E'+SnRowsTot+',0)', PercentForm)
    Cworksheet.write(InRows+1, 13, '=IF(D'+SnRowsTot+'>0,(F'+SnRowsTot+'/L'+SnRowsTot+')/D'+SnRowsTot+',0)', PercentForm)
    Cworksheet.write(InRows+1, 14, '=AVERAGE(O2:O'+SnRows+')', PercentForm)
    Cworksheet.write(InRows+1, 15, '=AVERAGE(P2:P'+SnRows+')', PercentForm)
    Cworksheet.write(InRows+1, 16, '=AVERAGE(Q2:Q'+SnRows+')', PercentForm)
    Cworksheet.write(InRows+1, 17, '=AVERAGE(R2:R'+SnRows+')', PercentForm)
    Cworksheet.write(InRows+1, 18, '=AVERAGE(S2:S'+SnRows+')', PercentForm)


  
    for report, month in zip(Lallfiles,LMonths):
    #This loop will walk through each file in the directory and each month in the list above as long as their are still files
        #Open file and sheet to read
        oWorkbook = xlrd.open_workbook(report)
        oSheet = oWorkbook.sheet_by_name(SheettoRead)

        #Create empty dictonary and specify which cell contains the number of days
        dLocationPAOP = {}
        vLocationDays= oSheet.cell(1,0).value

        for iRow in xrange(5, oSheet.nrows):
        #This loop will run through the main table rows and load them into the empty dictionary
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



        #Create workbook sheet with month name
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



FileReadandCreate("C:\Users\sjolly\Desktop\TestFolder 2\Dufferin-Peel\Year13-14\Individual Months",'TestLocationPAOP1.xlsx', "DufferinPeel_Teaching_Location_PAOP_Annual.xlsx",'Sheet1', "Y - Area East", "Y - Area North", "Y - Area West", "Y - Type High School","Y - Type Elementary/Middle")









