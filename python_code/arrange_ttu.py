import xlrd
import xlsxwriter
import string
import synonyms

ExcelFile = xlrd.open_workbook(r'..\data2\up\ttu.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols
rowBuffer = sheet.col_values(1)

workbook = xlsxwriter.Workbook(r'..\data2\up\ttu_a.xls')
worksheet = workbook.add_worksheet(u'sheet1')
#Key Words
for i in range(1,nrows-1):
    [buffer,empty] = synonyms.seg(str(rowBuffer[i]))
    if(buffer !=[]):
        for br in range(0,len(buffer)):
            worksheet.write('%s'%chr(65+br)+'%s'%str(i+1),buffer[br])
    else:    
        continue
    
#SourceOverwrite,not suggest any more
#for cl in range(1,4):
#    rowBuffer = sheet.col_values(cl)
#    for i in range(0,nrows-1):
#        worksheet.write('%s'%chr(64+cl)+'%s'%str(i+1),rowBuffer[i])
workbook.close()

print('----arrange_ttu complete----')
