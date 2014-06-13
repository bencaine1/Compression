# -*- coding: utf-8 -*-
"""
Created on Fri June 6 03:31:18 2014

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

# Changes to DataPt: Changed ketjen to conductive_additive, added percent_solids after active, added conductive_additive_percent after conductive_additive,
# added A_or_C at the end
class DataPt:
    def __init__(self, testreq, mean, stddev, date, active, percent_solids, batch, conductive_additive, conductive_additive_percent, electrolyte, A_or_C):
        self.testreq = testreq
        self.mean = mean
        self.stddev = stddev
        self.date = date
        self.active = active
        self.percent_solids = percent_solids
        self.batch = batch
        self.conductive_additive = conductive_additive
        self.conductive_additive_percent = conductive_additive_percent
        self.electrolyte = electrolyte
        self.A_or_C = A_or_C
        
def timeStamped(fname, fmt='%Y-%m-%d-%H-%M-%S_{fname}'):
    return datetime.datetime.now().strftime(fmt).format(fname=fname)

def is_number(s):
    if s:
        try:
            float(s)
            return True
        except ValueError:
            return False
    else:
        return False
def roundto5(x, base=5):
    return int(base * round(float(x)/base))

def fillTemplate(sh, lst, overwrite=True):
    print 'fillTemplate'
    used = sh.UsedRange
    nrows = used.Row + used.Rows.Count - 1
    if overwrite: nrows = 1
    for i in xrange(len(lst)):
#        if ((sh.Range('A' + str(i+2)).Value == None) or overwrite):
        print lst[i].testreq
        sh.Range('A' + str(i+nrows+1)).Value = lst[i].date
        sh.Range('C' + str(i+nrows+1)).Value = lst[i].testreq
        sh.Range('D' + str(i+nrows+1)).Value = lst[i].mean
        sh.Range('E' + str(i+nrows+1)).Value = lst[i].stddev
        sh.Range('F' + str(i+nrows+1)).Value = lst[i].active
        sh.Range('G' + str(i+nrows+1)).Value = lst[i].percent_solids
        sh.Range('H' + str(i+nrows+1)).Value = lst[i].batch
        sh.Range('I' + str(i+nrows+1)).Value = lst[i].conductive_additive
        sh.Range('J' + str(i+nrows+1)).Value = lst[i].conductive_additive_percent
        sh.Range('K' + str(i+nrows+1)).Value = lst[i].electrolyte

############## Scrape Excel files ################

data = []
rootdir = 'C:\\Users\\bcaine\\Documents\\Compression\\excel'

# connect to db
cnxn_str = """
Driver={SQL Server Native Client 11.0};
Server=172.16.111.235\SQLEXPRESS;
Database=CellBuild;
UID=sa;
PWD=Welcome!;
"""
cnxn = pyodbc.connect(cnxn_str)
cnxn.autocommit = True
cursor = cnxn.cursor()

# Start up excel
xl = Dispatch("Excel.Application")

print 'Scraping data...'

# Start scrapin'!
for folder in os.listdir(rootdir):
    testreq, date = ('',)*2
    paths = []
    paths.append(rootdir + '\\' + folder + '\\' + folder + '.xlsx')
    paths.append(rootdir + '\\' + folder + '\\' + folder + '_shared.xlsx')
    paths.append(rootdir + '\\' + folder + '\\' + folder + '.xlsm')
    paths.append(rootdir + '\\' + folder + '\\' + folder + '_shared.xlsm')
    for path in paths:
        if os.path.isfile(path):
            try:
                testreq = folder
                print testreq
                wbk = xl.Workbooks.Open(path)
    
                # Go into 'Slurry Data' worksheet
                sh = wbk.Worksheets('Slurry Data')
                
                # date
                date = sh.Range('K3').Value
                date = re.sub(' 00:00:00', '', str(date))
                if testreq == '100067': date = '12/5/13' # someone entered this date as '12/5/14'
    #            if date == None:
    #                date = sh.Range('K4').Value
    
                # Look in rows 9 through 49 for more data.
                for i in xrange(9, 49):
                    mean, stddev, active, percent_solids, batch, conductive_additive, electrolyte, instructions, cond_add_percent_C, cond_add_percent_A, C_material, A_material, m = ('',)*13
                    if sh.Range('B' + str(i)).Value != None and sh.Range('B' + str(i)).Value != "Type":
                        A_or_C = sh.Range('B' + str(i)).Value
                        # Batch no: cols B through I
                        for j in xrange(1, 8):
                            range_arg = chr(j + ord('A')) + str(i) # e.g. H10
                            res = str(sh.Range(range_arg))
                            res = re.sub('\.0', '', res) # Remove ".0". Excel numbers were coming out as e.g. "13.0"
                            if res == 'None':
                                res = '_'
                            batch += res
                        number = str(sh.Range('I' + str(i)))
                        number = re.sub('\.0', '', number)
                        number = number.zfill(4)
                        batch += number
                        
                        # electrolyte, mean, stddev
                        electrolyte = str(sh.Range('BA' + str(i)).Value)
                        mean = sh.Range('BV' + str(i)).Value
                        stddev = sh.Range('BW' + str(i)).Value
            
                        #instructions = sh.Range('BI' + str(i)).Value # unused
                        
                        # active ingredient
                        active_material = str(sh.Range('K' + str(i)).Value)
                        if 'LFP' in active_material:
                            active = 'LFP'
                        elif ('NMC' in active_material):
                            active = 'NMC'
                        elif ('MGPA' in active_material or 'MGP-A' in active_material or 'GPA' in active_material or 'MPGA' in active_material):
                            active = 'MGPA'
                        else:
                            active = active_material
                                                
                        # conductive additive
                        if 'C45' in active_material:
                            conductive_additive = 'C45'
                        else:
                            conductive_additive = str(sh.Range('Y' + str(i)).Value)
                            conductive_additive = re.sub('C-nergy C45', 'C45', conductive_additive)
    #                        if '50' in active_material:
    #                            percent_solids = '50'
    #                        elif '45' in active_material:
    #                            percent_solids = '45'
    #                        else:
    #                            percent_solids = '0'
                            
                        # Go into Slurry Request sheet
                        req_sh = wbk.Sheets('Slurry Request')
                        
                        # percent solids, conductive additive percent
                        if A_or_C == 'C':
                            percent_solids = roundto5(req_sh.Range('F6').Value)
                            conductive_additive_percent = req_sh.Range('F8').Value
                        elif A_or_C == 'A':
                            percent_solids = roundto5(req_sh.Range('F18').Value)
                            if '/2' in active_material:
                                conductive_additive_percent = 2
                            else:
                                conductive_additive_percent = req_sh.Range('F20').Value                                
                            
                        if conductive_additive_percent == 'None':
                            C_material = req_sh.Range('C39').Value
                            A_material = req_sh.Range('C18').Value
                            if A_or_C == 'C':
                                m = re.search('[0-9][0-9]/(?P=<percent>[0-9]\.[0-9])', C_material)
                                if m:
                                    conductive_additive_percent = m.group('percent')
                            elif A_or_C == 'A':
                                m = re.search('[0-9][0-9]/(?P=<percent>[0-9](\.[0-9][0-9])?)', A_material)
                                if m:
                                    conductive_additive_percent = m.group('percent')
                                
                        if is_number(mean) and is_number(stddev) and date != 'None':
                            t = (testreq, mean, stddev, date, active, percent_solids, batch, conductive_additive, conductive_additive_percent, electrolyte, A_or_C)
                            for item in t:
                                if item == None:
                                    item = ''
                            dp = DataPt(testreq, mean, stddev, date, active, percent_solids, batch, conductive_additive, conductive_additive_percent, electrolyte, A_or_C)
                            data.append(dp)
                wbk.Save()
                wbk.Close()
            except:
                # If open fails, move on to the next one.
                sys.stdout.write('open failed on file ' + folder)
                pass

# sort data by date
data.sort(key=lambda t: datetime.datetime.strptime(t.date, '%m/%d/%y'))

# Separate data into cathodes and anodes
cathodes, anodes = [], []
for t in data:
    if t.A_or_C == 'C':
        cathodes.append(t)
    elif t.A_or_C == 'A':
        anodes.append(t)

#print 'cathodes:'
#for cathode in cathodes:
#    print cathode.batch
#print 'anodes:'
#for anode in anodes:
#    print anode.batch
        
################# GENERATE RUNCHARTS ####################
        
# generate new compression charts from the template.
book = xl.Workbooks.Open(rootdir + '\\xl_compression_data_template.xls')
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
    if (t.active == 'LFP' and t.percent_solids == 50):
        lfp50.append(t)
        t_date = datetime.datetime.strptime(t.date, '%m/%d/%y')
        from_date = datetime.date.today()-datetime.timedelta(days=14)
        if (t_date.date() >= from_date):
            lfp50_recent.append(t)
    elif (t.active == 'LFP' and t.percent_solids == 45):
        lfp45.append(t)
fillTemplate(lfp50_sh, lfp50)
fillTemplate(lfp45_sh, lfp45)
fillTemplate(lfp50_recent_sh, lfp50_recent, True)

mgpa50, mgpa45, mgpa50_recent = [], [], []
for t in anodes:
    if (t.active == 'MGPA' and t.percent_solids == 50):
        mgpa50.append(t)
        t_date = datetime.datetime.strptime(t.date, '%m/%d/%y')
        from_date = datetime.date.today()-datetime.timedelta(days=14)
        if (t_date.date() >= from_date):
            mgpa50_recent.append(t)
    elif (t.active == 'MGPA' and t.percent_solids == 45):
        mgpa45.append(t)
fillTemplate(mgpa50_sh, mgpa50)
fillTemplate(mgpa45_sh, mgpa45)
fillTemplate(mgpa50_recent_sh, mgpa50_recent, True)

book.Save() # Template file should always be the same as the most recent file.
book.SaveAs(rootdir + '\\' + timeStamped('xl_compression_data.xls'))
    
#################### ADD TO DB ##################

# Delete everything and reset "auto-increment" counters
cursor.execute("""
DELETE FROM StiffnessData;
DELETE FROM SlurryBatch;
DBCC CHECKIDENT (SlurryBatch, RESEED, 0);
DBCC CHECKIDENT (StiffnessData, RESEED, 0);
""")

# Populate SlurryBatch table
print 'Populating SlurryBatch table'
for t in data:
    # Repeated tests on the same slurry do not get put in the SlurryBatch table
    
#        cursor.execute("""
#        merge SlurryBatch as T
#        using (select ?, ?, ?, ?, ?, ?, ?, ?, ?) as S (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
#        on S.Batch = T.Batch and S.Active = T.Active and S.TestRequest = T.TestRequest
#        when not matched then insert (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
#        values (S.TestRequest, S.Batch, S.Active, S.ActiveVolumePercent, S.CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent);
#        """, test_req_number, t.batch, t.active, t.percent_solids, t.A_or_C, t.conductive_additive, t.percent_solids, None, None)
    
    cursor.execute("""
    insert into SlurryBatch (TestRequest, Batch, Active, ActiveVolumePercent, CathodeOrAnode, ConductiveAdditive, ConductiveAdditivePercent, ConductiveAdditive2, ConductiveAdditive2Percent)
    values (?,?,?,?,?,?,?,?,?);
    """, t.testreq, t.batch, t.active, t.percent_solids, t.A_or_C, t.conductive_additive, t.percent_solids, None, None)
    
# Populate StiffnessData table
print 'Populating StiffnessData table'
for t in data:
    # Determine SlurryBatchUID
    slurry_batch_uid = None
    row = cursor.execute("""
    select UID from SlurryBatch
    where Batch = ?
    """, t.batch).fetchone()
    if row:
        slurry_batch_uid = row[0]
    
    cursor.execute("""
    merge StiffnessData as T
    using (select ?, ?, ?, ?) as S (SlurryBatchUID, Mean, StdDev, Date)
    on S.SlurryBatchUID = T.SlurryBatchUID
    when not matched then insert(SlurryBatchUID, Mean, StdDev, Date)
    values (S.SlurryBatchUID, S.Mean, S.StdDev, S.Date);
    """, slurry_batch_uid, t.mean, t.stddev, t.date)
    
cnxn.commit()

#close up shop
cursor.close()
del cursor
cnxn.close()