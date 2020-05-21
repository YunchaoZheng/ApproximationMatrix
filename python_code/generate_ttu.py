import xlrd
import xlsxwriter
import string
import synonyms

ExcelFile = xlrd.open_workbook(r'..\data2\up\ttu_a.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols

workbook = xlsxwriter.Workbook(r'..\data2\up\ttu_b.xls')
worksheet = workbook.add_worksheet(u'sheet1')
print('----generate_ttu initiated----')
#overwrite,not suggest any more
#rowBuffer = sheet.col_values(5)
#for i in range(0,nrows-1):
#    worksheet.write('A%s'%str(i+1),rowBuffer[i])
#rowBuffer = sheet.col_values(6)
#for i in range(0,nrows-1):
#    worksheet.write('B%s'%str(i+1),rowBuffer[i])
#rowBuffer = sheet.col_values(7)
#for i in range(0,nrows-1):
#    worksheet.write('C%s'%str(i+1),rowBuffer[i])

rowBuffer = sheet.col_values(0)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(65+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
print('----target1 complete----')
rowBuffer = sheet.col_values(1)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(70+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
print('----target2 complete----')
rowBuffer = sheet.col_values(2)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(75+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
print('----target3 complete----')
rowBuffer = sheet.col_values(3)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(80+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
workbook.close()
print('----target4 complete----')

print('----generate_ttn complete----')
