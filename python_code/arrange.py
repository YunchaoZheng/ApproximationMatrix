import xlrd
import xlsxwriter
import string

ExcelFile = xlrd.open_workbook(r'..\total.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols
rowBuffer = sheet.col_values(0)

workbook = xlsxwriter.Workbook(r'..\total_a.xls')
worksheet = workbook.add_worksheet(u'sheet1')
#Key Words
for i in range(1,nrows-1):
    buffer = str(rowBuffer[i]).split()
    if(buffer !=[]):
        for br in range(0,len(buffer)):
            worksheet.write('%s'%chr(73+br)+'%s'%str(i+1),buffer[br])
    else:    
        continue
#Source
for cl in range(0,7):
    rowBuffer = sheet.col_values(cl)
    for i in range(0,nrows-1):
        worksheet.write('%s'%chr(65+cl)+'%s'%str(i+1),rowBuffer[i])
workbook.close()

    
print('----arrange complete----')
