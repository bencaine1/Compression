# -*- coding: utf-8 -*-
"""
Created on Wed May 28 08:46:16 2014

@author: bcaine
"""

import os
import csv
import sys
import re
import datetime
import pyodbc
from win32com.client import Dispatch
import matplotlib.pyplot as plt

def timeStamped(fname, fmt='%Y-%m-%d-%H-%M-%S_{fname}'):
    return datetime.datetime.now().strftime(fmt).format(fname=fname)

def fillTemplate(sh, lst, overwrite=False):
    for i in xrange(len(lst)):
        if ((sh.Range('A' + str(i+2)).Value == None) or overwrite):
            print lst[i].specimen
            sh.Range('A' + str(i+2)).Value = lst[i].date
            sh.Range('B' + str(i+2)).Value = lst[i].time
            sh.Range('C' + str(i+2)).Value = lst[i].specimen
            sh.Range('D' + str(i+2)).Value = lst[i].mean
            sh.Range('E' + str(i+2)).Value = lst[i].stddev
            sh.Range('F' + str(i+2)).Value = lst[i].active_ingredient
            sh.Range('G' + str(i+2)).Value = lst[i].percent_solids
            sh.Range('H' + str(i+2)).Value = lst[i].batchno
            sh.Range('I' + str(i+2)).Value = lst[i].conductive_additive
            sh.Range('J' + str(i+2)).Value = lst[i].electrolyte

# Changes to DataPt: Changed ketjen to conductive_additive, added percent_solids after active_ingredient
class DataPt:
    def __init__(self, specimen, mean, stddev, date, time, active_ingredient, percent_solids, batchno, conductive_additive, electrolyte):
        self.specimen = specimen
        self.mean = mean
        self.stddev = stddev
        self.date = date
        self.time = time
        self.active_ingredient = active_ingredient
        self.percent_solids = percent_solids
        self.batchno = batchno
        self.conductive_additive = conductive_additive
        self.electrolyte = electrolyte

data = []
specimen, mean, stddev, date, time, active_ingredient, percent_solids, batchno, conductive_additive, electrolyte, comments = ('',)*11

rootdir = 'C:/Users/bcaine/Documents/Compression/csv'

for folder in os.listdir(os.getcwd()):
    if (os.path.isfile(folder + '/' + folder + '.csv')):
        with open(folder + '/' + folder + '.csv', 'rb') as f:
            i = 0
            reader = csv.reader(f, delimiter=',')
            reader = list(reader)
            for row in reader:
                for field in row:
                    if field == ' Date':
                        date = reader[i+3][1][1:]
                    elif field == ' Time':
                        time = reader[i+3][2][1:]
                    elif field == ' Specimen':
                        specimen = reader[i+3][3][1:]
                    elif field == ' Mean:':
                        mean = row[6]
                    elif field == ' Std. Dev.:':
                        stddev = row[6]
                    elif field == 'Comments:':
                        comments = reader[i+1][1]
                    elif re.match('[A-Z]{5}[0-9]{2}[A-Z][0-9]{4}', field):
                        batchno = field
                i+=1
        if (specimen != '' and mean != '' and mean != ' N/A' and stddev != '' and stddev != ' N/A'):
            try:
                mean = float(mean)
            except ValueError:
                mean = 0
            try:
                stddev = float(stddev)
            except ValueError:
                stddev = 0
            #specimen = specimen.replace('_','') # get rid of underscores
            active_ingredient = ''
            percent_solids = ''
            conductive_additive = ''
            m = None
            if ('LFP' in comments):
                active_ingredient = 'LFP'
            elif ('NMC' in comments):
                active_ingredient = 'NMC'
            elif ('MGPA' in comments or 'MGP-A' in comments or 'GPA' in comments or 'MPGA' in comments):
                active_ingredient = 'MGPA'
                comments_no_spaces = comments.replace(' ', '')
                percent_solids = comments_no_spaces[comments_no_spaces.find('A') + 1:comments_no_spaces.find('A') + 3]
            else:
                active_ingredient = comments
            if (percent_solids == ''):
                if ('45' in comments):
                    percent_solids = '45'
                elif ('50' in comments):
                    percent_solids = '50'
            m = re.search('Ket([ ]?)0([.]?)[0-9][0-9]', comments)
            if m:
                conductive_additive = m.group()
            else:
                m = re.search('C45[0 ](?P<additive>[0-9]\.?[0-9]?)', comments)
                if m:
                    conductive_additive = 'C45 ' + m.group('additive') + '%'
            m = re.search('E[1-9]+', comments)
            if m:
                electrolyte = m.group()
            t = DataPt(specimen, mean, stddev, date, time, active_ingredient, percent_solids, batchno, conductive_additive, electrolyte)
            
            #print 'Specimen: ' + specimen
            #print 'Mean: ' + str(mean)
            #print 'Std. Dev: ' + str(stddev)

            data.append(t)

# sort data by date
data.sort(key=lambda t: datetime.datetime.strptime(t.date, '%m-%d-%y'))

# Separate data into cathodes and anodes
cathodes, anodes = [], []
for t in data:
    if ('C' in t.specimen):
        cathodes.append(t)
    elif ('A' in t.specimen):
        anodes.append(t)    

################### FUNCTIONS ####################

# 'new': generate new compression charts from the template.
if (sys.argv[1] == 'new'):
    xl = Dispatch('Excel.Application')
    book = xl.Workbooks.Open(rootdir + '/compression_data_template.xls')
    c_sh = book.Sheets(1)
    a_sh = book.Sheets(2)
    lfp50_sh = book.Sheets(5)
    lfp45_sh = book.Sheets(7)
    lfp50_recent_sh = book.Sheets(9)
    mgpa50_sh = book.Sheets(11)
    mgpa45_sh = book.Sheets(13)
    mgpa50_recent_sh = book.Sheets(15)
    
    fillTemplate(c_sh, cathodes)
    fillTemplate(a_sh, anodes)
    
    lfp50, lfp45, lfp50_recent = [], [], []
    for t in cathodes:
        if (t.active_ingredient == 'LFP' and t.percent_solids == '50'):
            lfp50.append(t)
            t_date = datetime.datetime.strptime(t.date, '%m-%d-%y')
            from_date = datetime.date.today()-datetime.timedelta(days=14)
            if (t_date.date() >= from_date):
                lfp50_recent.append(t)
        elif (t.active_ingredient == 'LFP' and t.percent_solids == '45'):
            lfp45.append(t)
    fillTemplate(lfp50_sh, lfp50)
    fillTemplate(lfp45_sh, lfp45)
    fillTemplate(lfp50_recent_sh, lfp50_recent, True)

    mgpa50, mgpa45, mgpa50_recent = [], [], []
    for t in anodes:
        if (t.active_ingredient == 'MGPA' and t.percent_solids == '50'):
            mgpa50.append(t)
            t_date = datetime.datetime.strptime(t.date, '%m-%d-%y')
            from_date = datetime.date.today()-datetime.timedelta(days=14)
            if (t_date.date() >= from_date):
                mgpa50_recent.append(t)
        elif (t.active_ingredient == 'MGPA' and t.percent_solids == '45'):
            mgpa45.append(t)
    fillTemplate(mgpa50_sh, mgpa50)
    fillTemplate(mgpa45_sh, mgpa45)
    fillTemplate(mgpa50_recent_sh, mgpa50_recent, True)
    
    book.Save() # Template file should always be the same as the most recent file.
    book.SaveAs(timeStamped('compression_data.xls'))
    
# 'add_to_cell_test_data': Put the data into the 'cell test data' spreadsheets
elif (sys.argv[1] == 'add_to_cell_test_data'):

    xl = Dispatch("Excel.Application")
    
    for datapt in data:
        print datapt.specimen
        filestr = 'C:/Users/bcaine/Documents/Compression/excel/' + datapt.specimen[:6] + '/' + datapt.specimen[:6] + '.xlsx'
        if (os.path.isfile(filestr)):
            try:
                wbk = xl.Workbooks.Open(filestr)
                for sh in wbk.Sheets:
                    if sh.Name == "Slurry Data":
                        type_range = sh.Range("B9:B49")
                        for type in type_range:
                            if (type.Value == cathodes[type.Row].specimen[6:7]):
                                row, column = type.Row, type.Column
                                sh.Cells(74, type.Row).Value = cathodes[type.Row].mean
                                sh.Cells(75, type.Row).Value = cathodes[type.Row].stddev
                                print "Success"
                        break
                wbk.Save()
                wbk.Close()
            except:
                print "Probably password protected"
    
    xl.Application.Quit()

# 'add_to_dbb': Add the data to the db
elif (sys.argv[1] == 'add_to_db'):
    # connect to db
    cnxn_str =    """
    Driver={SQL Server Native Client 11.0};
    Server=172.16.111.235\SQLEXPRESS;
    Database=CellBuild;
    UID=sa;
    PWD=Welcome!;
    """
    cnxn = pyodbc.connect(cnxn_str)
    cnxn.autocommit = True
    cursor = cnxn.cursor()
    
    # MSSQL has no "on duplicate key" option, so you need to make a new temp
    # table, merge the tables, and the delete the temp table.
    cursor.execute("""
    create table StiffnessDataTmp (
    Specimen varchar(50) primary key,
    Mean float,
    StdDev float,
    Date varchar(50),
    Time varchar(50),
    Active_Ingredient varchar(50),
    Percent_Solids varchar(50),
    BatchNo varchar(50),
    Conductive_Additive varchar(50),
    Electrolyte varchar(50)
    );
    """)
#    cursor.execute("""
#    create table anode_tmp (
#    Specimen varchar(50) primary key,
#    Mean float,
#    StdDev float,
#    Date varchar(50),
#    Time varchar(50),
#    Active_Ingredient varchar(50),
#    Percent_Solids varchar(50),
#    BatchNo varchar(50),
#    Conductive_Additive varchar(50),
#    Electrolyte varchar(50)
#    );
#    """)
    # insert records into temp tables
    for cathode in cathodes:
        cursor.execute("""insert into dbo.cathode_tmp(Specimen, Mean, StdDev, Date, Time, Active_Ingredient, Percent_Solids, BatchNo, Conductive_Additive, Electrolyte)
        values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, cathode.specimen, cathode.mean, cathode.stddev, cathode.date, cathode.time, cathode.active_ingredient, cathode.percent_solids, cathode.batchno, cathode.conductive_additive, cathode.electrolyte)
    for anode in anodes:
        cursor.execute("""insert into dbo.anode_tmp(Specimen, Mean, StdDev, Date, Time, Active_Ingredient, Percent_Solids, BatchNo, Conductive_Additive, Electrolyte)
        values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, anode.specimen, anode.mean, anode.stddev, anode.date, anode.time, anode.active_ingredient, anode.percent_solids, anode.batchno, anode.conductive_additive, anode.electrolyte)
    # merge temp table with actual table
    cursor.execute("""
    merge cathode_compression_data as T
    using cathode_tmp as S
    on S.Specimen = T.Specimen
    when not matched then insert(Specimen, Mean, StdDev, Date, Time, Active_Ingredient, Percent_Solids, BatchNo, Conductive_Additive, Electrolyte)
    values (S.Specimen, S.Mean, S.StdDev, S.Date, S.Time, S.Active_Ingredient, S.Percent_Solids, S.BatchNo, S.Conductive_Additive, S.Electrolyte); 
    """)
#    cursor.execute("""
#    merge anode_compression_data as T
#    using anode_tmp as S
#    on S.Specimen = T.Specimen
#    when not matched then insert(Specimen, Mean, StdDev, Date, Time, Active_Ingredient, Percent_Solids, BatchNo, Conductive_Additive, Electrolyte)
#    values (S.Specimen, S.Mean, S.StdDev, S.Date, S.Time, S.Active_Ingredient, S.Percent_Solids, S.BatchNo, S.Conductive_Additive, S.Electrolyte); 
#    """)
    # delete temp tables
    cursor.execute("drop table dbo.cathode_tmp, dbo.anode_tmp")
    cnxn.commit()
    
    #close up shop
    cursor.close()
    del cursor
    cnxn.close()

elif (sys.argv[1] == 'visualize'):
    plt.plot([1,2,3,4])
    plt.ylabel('some numbers')
    plt.show()