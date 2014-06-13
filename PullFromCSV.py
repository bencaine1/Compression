# -*- coding: utf-8 -*-
"""
Created on Wed May 28 08:46:16 2014

@author: bcaine
"""

import os
from os.path import getmtime
import csv
import sys
import re
import datetime
import pyodbc
from win32com.client import Dispatch
import matplotlib.pyplot as plt

def timeStamped(fname, fmt='%Y-%m-%d-%H-%M-%S_{fname}'):
    return datetime.datetime.now().strftime(fmt).format(fname=fname)
    
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

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
    def __str__(self):
        s = 'Specimen: ' + str(self.specimen) + '\n'
        s += 'Mean: ' + str(self.mean) + '\n'
        s += 'Std Dev: ' + str(self.stddev) + '\n'
        s += 'Date: ' + str(self.date) + '\n'
        s += 'Time: ' + str(self.time) + '\n'
        s += 'Active Ingredient: ' + str(self.active_ingredient) + '\n'
        s += 'Percent Solids: ' + str(self.percent_solids) + '\n'
        s += 'Batch No: ' + str(self.batchno) + '\n'
        s += 'Conductive Additive: ' + str(self.conductive_additive) + '\n'
        s += 'Electrolyte: ' + str(self.electrolyte) + '\n'
        return s

data = []

rootdir = 'R:\\Characterization\\Compression Test'

folders = [f for f in os.listdir(rootdir) if re.search('[0-9]{6}[AC]', f)]

for folder in folders:
    # check last update, skip if already in FileUpdate db
    filedate = datetime.datetime.fromtimestamp(getmtime(os.path.join(rootdir, f))).strftime("%Y-%m-%d %H:%M:%S")
    row = cursor.execute("""
    select * from FileUpdate
    where Filename = ? and LastUpdate = ?;
    """, f, filedate).fetchone()
    if row:
        sys.stdout.write('^')
        continue

    specimen, mean, stddev, date, time, active_ingredient, percent_solids, batchno, conductive_additive, electrolyte, comments = ('',)*11
    paths = []
    paths.append(rootdir + '\\' + folder + '\\' + folder + '.csv')
    paths.append(rootdir + '\\' + folder + '\\' + folder + ' Results.csv')
    for path in paths:
        if (os.path.isfile(path)):
            with open(path, 'rb') as f:
                specimen = folder
                i = 0
                reader = csv.reader(f, delimiter=',')
                reader = list(reader)
                for row in reader:
                    for field in row:
                        if field == ' Date':
                            date = reader[i+3][1][1:]
                        elif field == ' Time':
                            time = reader[i+3][2][1:]
                        elif field == ' Mean:':
                            mean = row[6]
                        elif field == ' Std. Dev.:':
                            stddev = row[6]
                        elif field == 'Comments:':
                            comments = reader[i+1][1]
                        elif re.match('[A-Z]{5}[0-9]{2}[A-Z][0-9]{4}', field):
                            batchno = field
                    i+=1
            if (specimen != '' and '_' not in specimen and is_number(mean) and is_number(stddev)):
                try:
                    mean = float(mean)
                except ValueError:
                    mean = 0
                try:
                    stddev = float(stddev)
                except ValueError:
                    stddev = 0
                #specimen = specimen.replace('_','') # get rid of underscores
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
                #print t
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
    book = xl.Workbooks.Open(rootdir + '\\compression_data_template.xls')
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
    
# 'add_to_cell_test_data': Put the mean and stddev into the 'cell test data' spreadsheets
elif (sys.argv[1] == 'add_to_cell_test_data'):
    exceldir = 'C:\\Users\\bcaine\\Documents\\Compression\\excel'
    xl = Dispatch("Excel.Application")

    errorfiles = []

    for datapt in data:
        testreq = datapt.specimen[:6]
        #print testreq
        spec_no = datapt.specimen[7] if len(str(datapt.specimen)) > 7 else '1'
        print "Specimen: ", datapt.specimen, "Spec no: ", spec_no
        paths = []
        paths.append(exceldir + '\\' + testreq + '\\' + testreq + '.xlsx')
        paths.append(exceldir + '\\' + testreq + '\\' + testreq + '_shared.xlsx')
        paths.append(exceldir + '\\' + testreq + '\\' + testreq + '.xlsm')
        paths.append(exceldir + '\\' + testreq + '\\' + testreq + '_shared.xlsm')
        for path in paths:
            if (os.path.isfile(path)):
                try:
                    wbk = xl.Workbooks.Open(path)
                    sh = wbk.Worksheets('Slurry Data')
                    for i in xrange(9,49):
                        # E.g., if B9 == 'C' and I9 == '1'. zfill takes care of '1' vs '0001'
                        number = str(sh.Range('I' + str(i)).Value)
                        number = re.sub('\.0', '', number)
                        if (sh.Range('B' + str(i)).Value == datapt.specimen[6] and number.zfill(4) == spec_no.zfill(4)):
                            sh.Range('BV' + str(i)).Value = datapt.mean
                            sh.Range('BW' + str(i)).Value = datapt.stddev
                            print "Success"
                            break
                    wbk.Save()
                    wbk.Close()
                except:
                    print "Probably password protected"
                    errorfiles.append(testreq)
                    pass
    
    xl.Application.Quit()
    print 'the following files did not process:', errorfiles

# 'add_to_db': Add the data to the db
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
    
    # Delete everything and reset "auto-increment" counters
    # Slower, but ultimately more robust
    cursor.execute("""
    DELETE FROM StiffnessData;
    DELETE FROM SlurryBatch;
    DBCC CHECKIDENT (SlurryBatch, RESEED, 0);
    DBCC CHECKIDENT (StiffnessData, RESEED, 0);
    """)
    
    # Populate SlurryBatch table
    for t in data:
        c_or_a = 'O'
        if 'C' in t.specimen:
            c_or_a = 'C'
        elif 'A' in t.specimen:
            c_or_a = 'A'
        
        # Repeated tests on the same slurry do not get put in the SlurryBatch table
        test_req_number = t.specimen[:6]

#        cursor.execute("""
#        merge SlurryBatch as T
#        using (select ?, ?, ?, ?, ?, ?, ?, ?, ?) as S (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
#        on S.Batch = T.Batch and S.Active = T.Active and S.TestRequest = T.TestRequest
#        when not matched then insert (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
#        values (S.TestRequest, S.Batch, S.Active, S.ActiveVolumePercent, S.CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent);
#        """, test_req_number, t.batchno, t.active_ingredient, t.percent_solids, c_or_a, t.conductive_additive, t.percent_solids, None, None)
        
        cursor.execute("""
        insert into SlurryBatch (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
        values (?,?,?,?,?,?,?,?,?);
        """, test_req_number, t.batchno, t.active_ingredient, t.percent_solids, c_or_a, t.conductive_additive, t.percent_solids, None, None)
        
    # Populate StiffnessData table
    for t in data:
        d = datetime.datetime.strptime(date + ' ' + time, '%m-%d-%y %H:%M:%S')
        timestamp = d.strftime('%Y-%m-%d %H:%M:%S')
        
        # Determine SlurryBatchUID
        slurry_batch_uid = None
        row = cursor.execute("""
        select UID from SlurryBatch
        where Batch = ?
        """, t.batchno).fetchone()
        if row:
            slurry_batch_uid = row[0]
        
        cursor.execute("""
        merge StiffnessData as T
        using (select ?, ?, ?, ?, ?) as S (SlurryBatchUID, Specimen, Mean, StdDev, Timestamp)
        on S.Specimen = T.Specimen
        when not matched then insert(SlurryBatchUID, Specimen, Mean, StdDev, Timestamp)
        values (S.SlurryBatchUID, S.Specimen, S.Mean, S.StdDev, S.Timestamp);
        """, slurry_batch_uid, t.specimen, t.mean, t.stddev, timestamp)
        
    cnxn.commit()
    
    #close up shop
    cursor.close()
    del cursor
    cnxn.close()

elif (sys.argv[1] == 'visualize'):
    plt.plot([1,2,3,4])
    plt.ylabel('some numbers')
    plt.show()