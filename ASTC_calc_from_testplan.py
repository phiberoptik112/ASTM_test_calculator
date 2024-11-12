import pandas as pd
import numpy as np
import matplotlib.pyplot as plot
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from matplotlib.dates import (AutoDateLocator, YearLocator, MonthLocator,
                              DayLocator, WeekdayLocator, HourLocator,
                              MinuteLocator, SecondLocator, MicrosecondLocator,
                              RRuleLocator, rrulewrapper, MONTHLY,
                              MO, TU, WE, TH, FR, SA, SU, DateFormatter,
                              AutoDateFormatter, ConciseDateFormatter)
import datetime
from os import listdir, walk
from os.path import isfile, join
import xlsxwriter
from bokeh.plotting import figure, show

def retrieve_meter_paths(foundtest):
    currsrs = foundtest['Source']
    if currsrs[0] == 'D':
            datafile_num = currsrs[1:]
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            srs_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(srs_slm_found)
            srs_slm_file.append(rawDtestpath+srs_slm_found[0])
    elif currsrs[0] == 'E':
            datafile_num = currsrs[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            srs_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            srs_slm_file.append(rawEtestpath+srs_slm_found[0])
    # file to read is srs
    currrec = foundtest['Recieve '] 
    if currrec[0] == 'D':
            datafile_num = currrec[1:]
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            rec_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(rec_slm_found)
            receive_slm_file.append(rawDtestpath+rec_slm_found[0])
    elif currrec[0] == 'E':
            datafile_num = currrec[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            rec_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            receive_slm_file.append(rawEtestpath+rec_slm_found[0])

    bkgnd_list = foundtest['BNL']  
    if bkgnd_list[0] == 'D':
            datafile_num = bkgnd_list[1:]
            ##D meter, pull D_datafiles
            ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            bkgrnd_slm_found = [x for x in D_datafiles if datafile_num in x]
            bkgrnd_slm_file.append(rawDtestpath+bkgrnd_slm_found[0])
    elif bkgnd_list[0] == 'E':
            datafile_num = bkgnd_list[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            bkgrnd_slm_found = [x for x in E_datafiles if datafile_num in x]
            bkgrnd_slm_file.append(rawEtestpath+bkgrnd_slm_found[0])
    rt_list = foundtest['RT']
    if rt_list[0] == 'D':
            datafile_num = rt_list[1:]
            # print(datafile_num)
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            rt_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(rec_slm_found)
            rt_slm_file.append(rawDtestpath+rt_slm_found[0])
    elif rt_list[0] == 'E':
            datafile_num = rt_list[1:]
            # print(datafile_num)
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            rt_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            rt_slm_file.append(rawEtestpath+rt_slm_found[0])

# testplan file loading. Enter testplan with test numbers in 'Test Label' Column 
testplan_path ='//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Documents/Python_devstuff/INCOMING - PAFWC STC Testing Plan List_SRBR_Aprilrev1.xlsx'
testplanfile = pd.read_excel(testplan_path)
# copy all entries from  testplan and put them in a list 
# testlist = testplanfile.active

# change these for each project files to add 
rawDtestpath = '//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Test Data/Raw Data/SLM D - 3784/'
rawEtestpath = '//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Test Data/Raw Data/SLM E - 4328/'
D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]

print('Using hardcoded SLM Paths for D and E meters')

# Test files to calculate ASTC for:
tests_to_run = [
    '2.1',
    '2.2',
    '2.3',
    '2.4',
    '2.5',
]
# apparently, in the excel doc these strings need to be 2.1.0 to show up... weird. 
# need to establish room volumes per test - in testplan file

# looping through all tests 
for find_test in tests_to_run:
  
    mask = testplanfile.applymap(lambda x: find_test in x if isinstance(x,str) else False).to_numpy()
    indices = np.argwhere(mask) 
    # print(indices)
    index = indices[0,0]
    # print(index)
    print(testplanfile.iloc[index])
    foundtest = testplanfile.iloc[index] # structure with test info
    srs_slm_file = list()
    receive_slm_file = list()
    bkgrnd_slm_file = list()
    rt_slm_file = list()
    retrieve_meter_paths(foundtest)
    # found all the files associated with that test
    #room volumes in cubic meters
    recieve_roomvol = foundtest['source room vol']
    source_roomvol = foundtest['receive room vol']
    parition_area = foundtest['partition area']
    ASTC_vollimit = 883
    if recieve_roomvol > ASTC_vollimit:
        print('Using NIC calc, room volume too large')
    OBAdatasheet = 'OBA'
    RTsummarysheet = 'Summary'
    srs_OBAdata = pd.read_excel(srs_slm_file[0],OBAdatasheet)
    recive_OBAdata = pd.read_excel(receive_slm_file[0],OBAdatasheet)
    bkgrd_OBAdata = pd.read_excel(bkgrnd_slm_file[0],OBAdatasheet)
    rt = pd.read_excel(rt_slm_file[0],RTsummarysheet)
    #pulling raw OBA data from files 
    srs_overalloct = srs_OBAdata.iloc[7]
    srs_overalloct = srs_overalloct[14:30]
    rec_overalloct = recive_OBAdata.iloc[7]
    rec_overalloct = rec_overalloct[14:30]
    bkgrnd_overalloct = bkgrd_OBAdata.iloc[7]
    bkgrnd_overalloct = bkgrnd_overalloct[14:30]
    rt_thirty = rt['Unnamed: 10']
    rt_thirty = rt_thirty[25:41]
    rt_thirty = rt_thirty/1000 #convert to seconds
    constant = np.int32(20.047*np.sqrt(273.15+20))
    #constant contour 
    # STC_contour = [28,31,34,37,40,43,44,45,46,47,48,48,48,48,48,48]
    # rt_thirty = np.array(rt_thirty, dtype=int)
    # print(rt_thirty)
    intermed = 30/rt_thirty
    thisval = np.int32(recieve_roomvol*intermed)
    sabines =thisval/constant
    sabines = np.round(sabines*(0.921))
    # reciever correction 
    # print('sabines: ', sabines)
    recieve_corr = list()
    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
    print('rec vs background:',recieve_vsBkgrnd)
    for i in range(len(recieve_vsBkgrnd)):
        val = recieve_vsBkgrnd[i]
        print('val:', recieve_vsBkgrnd[i])
        print('count: ',i)
        if val < 5:
            recieve_corr.append(rec_overalloct.iloc[i]-2)
        elif val < 10:
            recieve_corr.append(np.log10(10**(rec_overalloct.iloc[i]/10)-10**(bkgrnd_overalloct.iloc[i]/10)))
        else:
            recieve_corr.append(rec_overalloct.iloc[i])
        # print('-=-=-=-=-')
        # print('recieve_corr: ',recieve_corr)
    recieve_corr = np.round(recieve_corr,1)
    #ATL calc
    ATL_val = srs_overalloct - recieve_corr+10*(np.log10(parition_area/sabines))
    # print(ATL_val)
    # ASTC curve fit 
    pos_diffs = list()
    diff_negative=0
    diff_positive=0 
    ASTC_start = 20
    New_curve =list()
    new_sum = 0
    STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]

    while (diff_negative < 8 and new_sum < 32):
        # print('starting loop')
        print('ASTC fit test value: ', ASTC_start)
        for vals in STCCurve:
            New_curve.append(vals+ASTC_start)
        ASTC_curve = New_curve - ATL_val
        # print('ASTC curve: ',ASTC_curve)

        diff_negative =  np.max(ASTC_curve - ATL_val)
        print('Max, single diff: ', diff_negative)
        for val in ASTC_curve:
            if val > 0:
                pos_diffs.append(np.round(val))
            else:
                pos_diffs.append(0)
        # print(pos_diffs)
        new_sum = np.sum(pos_diffs)
        print('Sum Positive diffs: ', new_sum)
        
        if new_sum > 32 or diff_negative > 8:
            print('Curve too high! ASTC fit: ', ASTC_start-1)
            print('Result for test: ', find_test) 
            print('-=-=-=-=-=-=-=-=-')
            break 
        pos_diffs = []
        New_curve = []
        ASTC_start = ASTC_start + 1
        if ASTC_start >80: break
        

import pandas as pd
from pandas import ExcelWriter
import numpy as np
import matplotlib.pyplot as plot
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
from matplotlib.dates import (AutoDateLocator, YearLocator, MonthLocator,
                              DayLocator, WeekdayLocator, HourLocator,
                              MinuteLocator, SecondLocator, MicrosecondLocator,
                              RRuleLocator, rrulewrapper, MONTHLY,
                              MO, TU, WE, TH, FR, SA, SU, DateFormatter,
                              AutoDateFormatter, ConciseDateFormatter)
import datetime
from os import listdir, walk
from os.path import isfile, join
import xlsxwriter
from bokeh.plotting import figure, show
from dataclasses import dataclass


def write_testdata(find_datafile, reportfile, newsheetname):
    if find_datafile[0] =='D':
        datafile_num = find_datafile[1:]
        datafile_num = '-831_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in D_datafiles if datafile_num in x]
        slm_found[0] = rawDtestpath+slm_found[0]
        # print(srs_slm_found)
    elif find_datafile[0] == 'E':
        datafile_num = find_datafile[1:]
        datafile_num = '-831_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in E_datafiles if datafile_num in x]
        slm_found[0] = rawEtestpath+slm_found[0]

    print(slm_found[0])

    srs_data = pd.read_excel(slm_found[0],sheet_name='OBA')
    with ExcelWriter(
    rawReportpath+reportfile,
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
    ) as writer:
        srs_data.to_excel(writer, sheet_name=newsheetname) 

def write_RTtestdata(find_datafile, reportfile,newsheetname):
    if find_datafile[0] =='D':
        datafile_num = find_datafile[1:]
        datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in D_datafiles if datafile_num in x]
        slm_found[0] = rawDtestpath+slm_found[0]
        # print(srs_slm_found)
    elif find_datafile[0] == 'E':
        datafile_num = find_datafile[1:]
        datafile_num = '-RT_Data.'+datafile_num+'.xlsx'
        slm_found = [x for x in E_datafiles if datafile_num in x]
        slm_found[0] = rawEtestpath+slm_found[0]

    print(slm_found[0])

    srs_data = pd.read_excel(slm_found[0],sheet_name='Summary')
    with ExcelWriter(
    rawReportpath+reportfile,
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
    ) as writer:
        srs_data.to_excel(writer, sheet_name=newsheetname) 

def retrieve_meter_paths(foundtest):
    currsrs = foundtest['Source']
    if currsrs[0] == 'D':
            datafile_num = currsrs[1:]
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            srs_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(srs_slm_found)
            srs_slm_file.append(rawDtestpath+srs_slm_found[0])
    elif currsrs[0] == 'E':
            datafile_num = currsrs[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            srs_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            srs_slm_file.append(rawEtestpath+srs_slm_found[0])
    # file to read is srs
    currrec = foundtest['Recieve '] 
    if currrec[0] == 'D':
            datafile_num = currrec[1:]
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            rec_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(rec_slm_found)
            receive_slm_file.append(rawDtestpath+rec_slm_found[0])
    elif currrec[0] == 'E':
            datafile_num = currrec[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            rec_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            receive_slm_file.append(rawEtestpath+rec_slm_found[0])

    bkgnd_list = foundtest['BNL']  
    if bkgnd_list[0] == 'D':
            datafile_num = bkgnd_list[1:]
            ##D meter, pull D_datafiles
            ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            bkgrnd_slm_found = [x for x in D_datafiles if datafile_num in x]
            bkgrnd_slm_file.append(rawDtestpath+bkgrnd_slm_found[0])
    elif bkgnd_list[0] == 'E':
            datafile_num = bkgnd_list[1:]
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            bkgrnd_slm_found = [x for x in E_datafiles if datafile_num in x]
            bkgrnd_slm_file.append(rawEtestpath+bkgrnd_slm_found[0])
    rt_list = foundtest['RT']
    if rt_list[0] == 'D':
            datafile_num = rt_list[1:]
            # print(datafile_num)
            ##D meter, pull D_datafiles
                ## (D_datafiles) for filename
            datafile_num = datafile_num+'.xlsx'
            rt_slm_found = [x for x in D_datafiles if datafile_num in x]
            # print(rec_slm_found)
            rt_slm_file.append(rawDtestpath+rt_slm_found[0])
    elif rt_list[0] == 'E':
            datafile_num = rt_list[1:]
            # print(datafile_num)
            ##E meter, pull E_datafiles
            datafile_num = datafile_num+'.xlsx'
            rt_slm_found = [x for x in E_datafiles if datafile_num in x]
            # print(rec_slm_found)
            rt_slm_file.append(rawEtestpath+rt_slm_found[0])

def ASTC_curve_fitplotter(ASTC_curve,New_curve):
      #plot the two curves to debug the fit
    plot.plot(ASTC_curve)
# testplan file loading. Enter testplan with test numbers in 'Test Label' Column 
testplan_path ='//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Documents/Python_devstuff/INCOMING - PAFWC STC Testing Plan List_SRBR_Aprilrev1.xlsx'
testplanfile = pd.read_excel(testplan_path)
# copy all entries from  testplan and put them in a list 
# testlist = testplanfile.active

# change these for each project files to add 
rawDtestpath = '//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Test Data/Raw Data/SLM D - 3784/'
rawEtestpath = '//DLA-04/Shared/KAILUA PROJECTS/2023/23-001 Hickam AFB Acoustical Testing/Test Data/Raw Data/SLM E - 4328/'
D_datafiles = [f for f in listdir(rawDtestpath) if isfile(join(rawDtestpath,f))]
E_datafiles = [f for f in listdir(rawEtestpath) if isfile(join(rawEtestpath,f))]

print('Using hardcoded SLM Paths for D and E meters')

# Test files to calculate ASTC for:
tests_to_run = [
    '2.1'
]
# apparently, in the excel doc these strings need to be 2.1.0 to show up... weird. 
# need to establish room volumes per test - in testplan file

# looping through all tests 
for find_test in tests_to_run:
  
    mask = testplanfile.applymap(lambda x: find_test in x if isinstance(x,str) else False).to_numpy()
    indices = np.argwhere(mask) 
    # print(indices)
    index = indices[0,0]
    # print(index)
    print(testplanfile.iloc[index])
    foundtest = testplanfile.iloc[index] # structure with test info, found from excel testplan

    srs_slm_file = list()
    receive_slm_file = list()
    bkgrnd_slm_file = list()
    rt_slm_file = list()
    retrieve_meter_paths(foundtest) # pull data from files into lists 
    # found all the files associated with that test
    #room volumes in cubic meters
    recieve_roomvol = foundtest['source room vol']
    source_roomvol = foundtest['receive room vol']
    parition_area = foundtest['partition area']
    ASTC_vollimit = 883 # cubic feet?
    if recieve_roomvol > ASTC_vollimit:
        print('Using NIC calc, room volume too large')
    OBAdatasheet = 'OBA'
    RTsummarysheet = 'Summary'
    srs_OBAdata = pd.read_excel(srs_slm_file[0],OBAdatasheet)
    recive_OBAdata = pd.read_excel(receive_slm_file[0],OBAdatasheet)
    bkgrd_OBAdata = pd.read_excel(bkgrnd_slm_file[0],OBAdatasheet)
    rt = pd.read_excel(rt_slm_file[0],RTsummarysheet)
    #pulling raw OBA data from files 
    srs_overalloct = srs_OBAdata.iloc[7]
    srs_overalloct = srs_overalloct[14:30]
    rec_overalloct = recive_OBAdata.iloc[7]
    rec_overalloct = rec_overalloct[14:30]
    bkgrnd_overalloct = bkgrd_OBAdata.iloc[7]
    bkgrnd_overalloct = bkgrnd_overalloct[14:30]
    rt_thirty = rt['Unnamed: 10']
    rt_thirty = rt_thirty[25:41]
    rt_thirty = rt_thirty/1000 #convert to seconds
    constant = np.int32(20.047*np.sqrt(273.15+20))
  
 
    # rt_thirty = np.array(rt_thirty, dtype=int)
    # print(rt_thirty)
    intermed = 30/rt_thirty
    thisval = np.int32(recieve_roomvol*intermed)
    sabines =thisval/constant
    sabines = np.round(sabines*(0.921))
    # reciever correction 
    # print('sabines: ', sabines)
    recieve_corr = list()
    recieve_vsBkgrnd = rec_overalloct - bkgrnd_overalloct
    # print('rec vs background:',recieve_vsBkgrnd)
    for i in range(len(recieve_vsBkgrnd)):
        val = recieve_vsBkgrnd[i]
        # print('val:', recieve_vsBkgrnd[i])
        # print('count: ',i)
        if val < 5:
            recieve_corr.append(rec_overalloct.iloc[i]-2)
        elif val < 10:
            recieve_corr.append(np.log10(10**(rec_overalloct.iloc[i]/10)-10**(bkgrnd_overalloct.iloc[i]/10)))
        else:
            recieve_corr.append(rec_overalloct.iloc[i])
        # print('-=-=-=-=-')
        # print('recieve_corr: ',recieve_corr)
    recieve_corr = np.round(recieve_corr,1)
    #Apparent Transmission Loss calc
    ATL_val = srs_overalloct - recieve_corr+10*(np.log10(parition_area/sabines))
    # print(ATL_val)
    # ASTC curve fit 
    pos_diffs = list()
    diff_negative=0
    diff_positive=0 
    ASTC_start = 20
    New_curve =list()
    new_sum = 0
    STCCurve = [-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4]

    while (diff_negative < 8 and new_sum < 32):
        # print('starting loop')
        print('ASTC fit test value: ', ASTC_start)
        for vals in STCCurve:
            New_curve.append(vals+ASTC_start)
        ASTC_curve = New_curve - ATL_val
        # print('ASTC curve: ',ASTC_curve)

        diff_negative =  np.max(ASTC_curve - ATL_val)
        print('Max, single diff: ', diff_negative)
        for val in ASTC_curve:
            if val > 0:
                pos_diffs.append(np.round(val))
            else:
                pos_diffs.append(0)
        # print(pos_diffs)
        new_sum = np.sum(pos_diffs)
        print('Sum Positive diffs: ', new_sum)
        ASTC_curve_fitplotter(ASTC_curve,New_curve)
        if new_sum > 32 or diff_negative > 8:
            print('Curve too high! ASTC fit: ', ASTC_start-1)
            print('Result for test: ', find_test) 
            print('-=-=-=-=-=-=-=-=-')
            break 
        pos_diffs = []
        New_curve = []
        ASTC_start = ASTC_start + 1
        if ASTC_start >80: break
        
## Once this is all properly working, need to format and plot the ASTC curve with the ATL