# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 15:17:11 2014

@author: Bob Voorheis, David Wolf
"""

from os import listdir
from xlrd import open_workbook, XL_CELL_ERROR, error_text_from_code
import re
from time import clock

def qs_indices(sheet):
    index = [None]*45
    top = sheet.row(13)
    bottom = sheet.row(14)
    for i in range(0,10):
        if re.search('On',top[i].value):
            index[0] = i
        elif re.search('ISBN',top[i].value):
            index[1] = i
        elif re.search('Title',bottom[i].value):
            index[2] = i
        elif re.search('Req.',bottom[i].value):
            index[3] = i
        elif re.search('Cond.',bottom[i].value):
            index[4] = i
    for i in range(5,25):
        if re.search('STOCK',top[i].value):
            for j in range(0,10):
                if re.match('N',bottom[i+j].value): 
                    index[5] = i+j
                    index[6] = i+j+1
                elif re.search('LN',bottom[i+j].value): 
                    index[7] = i+j
                    index[8] = i+j+1
                elif re.search('VG',bottom[i+j].value): 
                    index[9] = i+j
                    index[10] = i+j+1
                elif re.search('G',bottom[i+j].value): 
                    index[11] = i+j
                    index[12] = i+j+1
                elif re.search('A',bottom[i+j].value): 
                    index[13] = i+j
                    index[14] = i+j+1
                elif index[13] == None:
                    if re.search('B',bottom[i+j].value):
                        index[13] = i+j
                        index[14] = i+j+1
                    elif re.search('RB',bottom[i+j].value):
                        index[13] = i+j
                        index[14] = i+j+1
        elif re.search('Total',top[i].value):
            index[15] = i
        elif re.search('Quality',top[i].value):
            index[16] = i
        elif re.search('Base',top[i].value):
            index[17] = i
    for i in range(20,30):
        if re.search('FOLLETT',top[i].value):
            for j in range(0,10):
                if re.search('Used \$',bottom[i+j].value):
                    index[18] = i+j
                elif re.search('U Qty',bottom[i+j].value):
                    index[19] = i+j
                elif re.search('New \$',bottom[i+j].value):
                    index[20] = i+j
                elif re.search('N Qty',bottom[i+j].value):
                    index[21] = i+j
                elif re.search('Prem \$',bottom[i+j].value):
                    index[22] = i+j
                elif re.search('P Qty',bottom[i+j].value):
                    index[23] = i+j
            break
    for i in range(25,40):
        if re.search('Publisher',top[i].value):
            for j in range(0,10):
                if re.search('Name',bottom[i+j].value):
                    index[24] = i+j
                elif re.search('Year',bottom[i+j].value):
                    index[25] = i+j
                elif re.search('lbs',bottom[i+j].value):
                    index[26] = i+j
                elif re.search('oz',bottom[i+j].value):
                    index[27] = i+j
                elif re.search('School \$',bottom[i+j].value):
                    index[28] = i+j
                elif re.search('Our \$',bottom[i+j].value):
                    index[29] = i+j
            break
    for i in range(35,40):
        if re.search('MARKET',top[i].value):
            for j in range(0,10):
                if re.search('Low',bottom[i+j].value):
                    index[30] = i+j
                elif re.search('High',bottom[i+j].value):
                    index[31] = i+j
                elif re.search('Average',bottom[i+j].value):
                    index[32] = i+j
                elif re.search('Qty',bottom[i+j].value):
                    index[33] = i+j
            break
    for i in range(20,45):
        if re.search('Vintage',top[i].value):
            for j in range(0,5):
                if re.search('V \$',bottom[i+j].value):
                    index[34] = i+j
                    index[35] = i+j+1
                elif re.search('Z \$',bottom[i+j].value):
                    index[36] = i+j
                    index[37] = i+j+1
            if index[37] is not None:
                break
        elif re.search('Amazon',top[i].value):
            if re.search('Z \$',bottom[i].value):
                index[36] = i
                index[37] = i+1
    for i in range(40,50):
        if re.search('QTY',top[i].value):
            index[38] = i
        elif re.search('Price',top[i].value):
            index[39] = i
        elif re.search('Amount',top[i].value):
            index[40] = i
    for i in range(0,len(bottom)):
        if re.search('Discount',bottom[i].value):
            index[41] = i
        elif re.search('UMRP',bottom[i].value):
            index[42] = i
        elif re.search('PRINT',bottom[i].value):
            index[43] = i
        elif re.search('Ship',top[i].value):
            index[44] = i
    return index

def os_indices(sheet):
    index = [None]*7
    top = sheet.row(13)
    bottom = sheet.row(14)
    for i in range(0,3):
        if re.search('Quote',top[i].value) and re.search('Row',bottom[i].value):
            index[0] = i
    for i in range(10,20):
        if re.search('Source 1',top[i].value):
            index[1] = i
            index[2] = i+1
        elif re.search('Source 2',bottom[i].value):
            index[3] = i
            index[4] = i+1
        elif re.search('Source 3',bottom[i].value):
            index[5] = i
            index[6] = i+1
    for i in range(15,25):
        if re.search('Target',top[i].value):
            index[7] = i
        elif re.search('%',top[i].value):
            index[8] = i
        elif re.search('QTY',top[i].value):
            index[9] = i
        elif re.search('ETA',top[i].value):
            index[10] = i
        elif re.search('Received',top[i].value):
            index[11] = i
        elif re.search('Non',top[i].value):
            index[12] = i
        elif re.search('Need to',top[i].value):
            index[13] = i
    for i in range(20,30):
        if re.search('In House',top[i].value):
            for j in range(0,10):
                if re.search('Extra',bottom[i+j].value):
                    index[14] = i+j
                elif re.search('Reject',bottom[i+j].value):
                    index[15] = i+j
                elif re.search('Clean',bottom[i+j].value):
                    index[16] = i+j
                elif re.search('Refurb',bottom[i+j].value):
                    index[17] = i+j
                elif re.search('Reface',bottom[i+j].value):
                    index[18] = i+j
                elif re.search('Rebind',bottom[i+j].value):
                    index[19] = i+j
                elif re.search('Boxed',bottom[i+j].value):
                    index[20] = i+j
            break
    for i in range(30,40):
        if re.search('Heck',top[i].value):
            index[21] = i
        if re.search('B//O',top[i].value):
            index[21] = i
        if re.search('Shipped',bottom[i].value):
            index[21] = i
        if re.search('Total',bottom[i].value):
            index[21] = i
        if re.search('M-I-A',bottom[i].value):
            index[21] = i

strip_unicode = re.compile("([^-_a-zA-Z0-9!@#%&=,/'\";:~`\$\^\*\(\)\+\[\]\.\{\}\|\?\<\>\\]+|[^\s]+)")

rootdirec = 'c:\\PythonDirectory\\'
#direcs = ['Sales\\','Sales\\Orders\\','Sales\\Orders\\Complete\\','Sales\\Orders\\Cancelled\\']
direcs = ['Direct Sales Orders\\']

f = open('quotes.txt','w')
f.write('Folder\tFilename\tQuote\tSchool\tDate\tRow\tOn Order\tISBN\tTitle\tQty Req\t' +
        'Condition\tN\t$\tLN\t$\tVG\t$\tG\t$\tA\t$\tTotal in Stock\t' +
        'Quality Price\tBase Price\tUsed $\tU Qty\tNew $\tN Qty\tPrem $\tP Qty' +
        '\tPublisher\tCYear\tlbs\toz\tSchool Price\tOur Price\tMarket Low\t' +
        'Market High\tMarket Avg\tMarket Qty\tVintage Price\tVintage Qty\t' +
        'Amazon Price\tAmazon Qty\tQty Quoted\tQuoted Price\tQuoted Amount\t' +
        'Publisher Discount\tUMRP\tIn Print?\tEst Ship\n')

start = clock()
errors = {}
count = 0.0
for direc in direcs:
    count = 0.0
    files = listdir(rootdirec + direc)
    #files = ['D153594 - Mt. Diablo USD.xlsx','D153663 - Mission Cons ISD.xlsx','D153674 - Hesperia USD.xlsx','D153691 - Anne Arundel County BOE.xlsx','D153691a - Anne Arundel County BOE.xlsx','D153713b - Tim Peters.xlsx','D153715 - The Classical Academy.xlsx','D153729 - Silver Stage HS.xlsx','D153736 - Long Beach USD.xlsx','D153738 - Kingman USD #20.xlsx','D160049 - Anthony White.xlsx','D160077 - Ben Gamla Charter School.xlsx','D160174 - South Knox ES.xlsx','D160193 - Shadow Mountain HS.xlsx','D160269 - Long Beach USD.xlsx','D160275 - K-12 Textbook Solutions.xlsx','D160280 - Pinnacle HS.xlsx','D160283 - Elgin SD U-46 - Sycamore 2.xlsx','D160292 -Elgin SD U-46 - Horizon 1.xlsx','D160311 - Eisenhower HS.xlsx','D160315 - Elgin SD U-46 - Coleman 2.xlsx','D160329 - West Bend SD.xlsx','D160331 - Orchard View HS.xlsx','D160461 - Tim Peters.xlsx','D160495a - Floyd County SD.xlsx']
#    files = []
#    reader = csv.reader(open('C://PythonDirectory/list.csv'),delimiter=',')
#    for row in reader:
#        try:
#            files.append(row[0])
#        except:
#            continue
    for filename in files:
        skip = False
        if '.xls' not in filename:
            continue
    #    try:
    #        if not(2000 > int(re.search('(D15)(\d{4})', filename).group(2)) ):
    #            continue
    #    except AttributeError:
    #        continue
        for string in ['bid sheet','lookup sheet','quote sheet','worksheet']:
            if string in filename.lower():
                skip = True
                break
        if skip == True:
            print(string + '; ' + filename)
            count += 1
            continue
    
                
        percent = round(100*count/len(files),2)
        print(str(percent) + "% done")
        count += 1
        orderlines = 0
        try:
            wb = open_workbook(rootdirec + direc + filename,on_demand=True)
        except IOError:
            f.write(direc + '\t' + filename + '\t__Error opening file\n')
            continue
        try:
            ordersheet = wb.sheet_by_name('Order Worksheet')
            quotesheet = wb.sheet_by_name('Quote Worksheet')
            summasheet = wb.sheet_by_name(    'Summary'    )
        except:
            f.write(direc + '\t' + filename + '\t__Error finding worksheet(s)\n')
            continue
        index = qs_indices(quotesheet)
#        if None in index:
#            f.write(direc + '\t' + filename + '\t__Error finding necessary headers on quote sheet\n')
#            continue
        try:
            quote = strip_unicode.sub('',quotesheet.cell_value(7,3)).replace('\n','').replace('\t','')
        except:
            f.write(direc + '\t' + filename + '\t__Error retrieving quote number\n')
            continue
        try:
            school = strip_unicode.sub('',quotesheet.cell_value(1,3)).replace('\n','').replace('\t','').strip()
        except:
            f.write(direc + '\t' + filename + '\t\t__Error retrieving school name\n')
            continue
        try:
            date = int(quotesheet.cell_value(8,3))
        except:
            if str(quotesheet.cell_value(8,3)).lower().strip('()') == 'required':
                date = ''
            else:
                f.write(direc + '\t' + filename + '\t\t\t__Error retrieving date\n')
                continue
        for row in range(15,quotesheet.nrows):
            try:
                cell_type = quotesheet.cell_type(row,index[1])
                if cell_type == 2:
                    isbn = str(int(float(quotesheet.cell_value(row,index[1]))))
                else:
                    isbn = strip_unicode.sub('',str(quotesheet.cell_value(row,index[1]))).replace('\n','').replace('\t','')
                if isbn == '' or isbn == ' ':
                    continue
            except:
                f.write(direc + '\t' 
                              + filename + '\t' 
                              + quote + '\t' 
                              + school + '\t' 
                              + str(date) + '\t' 
                              + str(row+1) + '\t\t__Error retrieving ISBN\n')
                continue
            
            try:
                if quotesheet.cell_value(row,index[0]) == '' or quotesheet.cell_value(row,index[0]) == 'Quote':
                    ordered_bool = 0
                elif quotesheet.cell_value(row,index[0]) == 'Order':
                    ordered_bool = 1
                else:
                    ordered_bool = int(quotesheet.cell_value(row,index[0]))
                
                if ordered_bool == 1:
                    orderlines += 1
                    qty_ordered = ordersheet.cell_value(14 + orderlines,7)
                elif ordered_bool == 0:
                    qty_ordered = 0
                else:
                    raise Exception('Checkbox appears to be neither checked nor unchecked')
            except:
                f.write(direc + '\t' 
                        + filename + '\t' 
                        + quote + '\t' 
                        + school + '\t' 
                        + str(date)+ '\t'
                        + str(row + 1) 
                        + '\t__Error retrieving quantity information\t'
                        + isbn + '\n')
                continue
                
    #            cell_hidden = quotesheet.rowinfo_map[row].hidden
            
            f.write(direc + '\t' 
                          + filename + '\t' 
                          + quote + '\t' 
                          + school + '\t' 
                          + str(date) + '\t' 
                          + str(row + 1) + '\t' 
                          + str(ordered_bool) + '\t' 
                          + isbn + '\t')
            
            for ind in index[2:len(index)]:
                try:
                    if ind == None:
                        f.write('\t')
                        continue
                    try:
                        if quotesheet.cell_type(row,ind) == XL_CELL_ERROR:
                            d = error_text_from_code[quotesheet.cell_value(row,ind)]
                        else:
                            d = strip_unicode.sub('',quotesheet.cell_value(row,ind)).replace('\n','').replace('\t','')
                    except TypeError:
                        d = str(quotesheet.cell_value(row,ind))
                    
                    f.write(d)
                    f.write('\t')
                except:
                    raise
                    f.write('__Error retrieving value\t')
                    continue
    
            f.write(str(qty_ordered))
            f.write('\t')
            f.write(str(cell_type))
            f.write('\t')
    #        f.write(str(cell_hidden))
            f.write('\n')
    #    except Exception, e:
    #        errors[filename] = str(e.__class__.__name__ + ': ' + str(e.message))

print('100.0% done')
print(clock()-start)
f.close()


#for filename in d:
#    if "D13" in filename and "bid sheet" not in filename:
#        wb = open_workbook('Z:\\2013\\2013 Direct Sales\\Sales\\' + filename,on_demand=True)
#        values=[]
#        for col in range(0,wb.sheets()[2].ncols):
#            values.append(wb.sheets()[2].cell(13,col).value)
#    print len(values)