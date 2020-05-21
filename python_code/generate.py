import xlrd
import xlsxwriter
import string
import synonyms

ExcelFile = xlrd.open_workbook(r'..\total_a.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols

workbook = xlsxwriter.Workbook(r'..\total_b.xls')
worksheet = workbook.add_worksheet(u'sheet1')
print('----generate initiated----')
#overwrite
rowBuffer = sheet.col_values(8)
for i in range(0,nrows-1):
    worksheet.write('A%s'%str(i+1),rowBuffer[i])
rowBuffer = sheet.col_values(9)
for i in range(0,nrows-1):
    worksheet.write('B%s'%str(i+1),rowBuffer[i])
rowBuffer = sheet.col_values(10)
for i in range(0,nrows-1):
    worksheet.write('C%s'%str(i+1),rowBuffer[i])
#Generate
rowBuffer = sheet.col_values(8)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(69+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
print('----targer1 complete----')
rowBuffer = sheet.col_values(9)
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
print('----target2 complete----')
rowBuffer = sheet.col_values(10)
for tmp in range(1,nrows-1):
    if(rowBuffer[tmp] !=[]):
        [buffertxt,buffernum] = synonyms.nearby(rowBuffer[tmp])
        if (buffertxt != []):
            for j in range(0,len(buffertxt)-1):
                worksheet.write('%s'%chr(82+j)+'%s'%str(tmp+1),buffertxt[j])
        else:
            continue
    else:
        continue
print('----target3 complete----')
workbook.close()
print('----generate complete----')
