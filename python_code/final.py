#final
import xlrd
import xlsxwriter

print('----final initiated----')
ExcelFile = xlrd.open_workbook(r'..\total.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols
ExcelFile2 = xlrd.open_workbook(r'..\total_g_n.xls')
sheet2 = ExcelFile2.sheet_by_index(0)
nrows2 = sheet2.nrows
ncols2 = sheet2.ncols
ExcelFile3 = xlrd.open_workbook(r'..\total_g_u.xls')
sheet3 = ExcelFile3.sheet_by_index(0)
nrows3 = sheet3.nrows
ncols3 = sheet3.ncols
ExcelFile4 = xlrd.open_workbook(r'..\data2\new\ttn.xls')
sheet4 = ExcelFile4.sheet_by_index(0)
nrows4 = sheet4.nrows
ncols4 = sheet4.ncols
ExcelFile5 = xlrd.open_workbook(r'..\data2\up\ttu.xls')
sheet5 = ExcelFile5.sheet_by_index(0)
nrows5 = sheet5.nrows
ncols5 = sheet5.ncols

workbook = xlsxwriter.Workbook(r'..\final.xls')
worksheet = workbook.add_worksheet(u'sheet1')
worksheet.write('A1','Source')
worksheet.write('B1','Target_n')
worksheet.write('C1','Value_n')
worksheet.write('D1','Target_u')
worksheet.write('E1','Value_u')
worksheet.write('F1','Search+')
wt = 2
for i in range(1,nrows2-1):
    lineBuffer = sheet.row_values(i)
    #total_g_n
    lineBuffer2 = sheet2.row_values(i)
    if(lineBuffer2[1] != '0'):
        worksheet.write('A%s'%str(wt),lineBuffer[0])#source
        worksheet.write('F%s'%str(wt),lineBuffer[4])#search+
        worksheet.write('B%s'%str(wt),lineBuffer2[2])#target_n
        lineBufferTMP = sheet4.row_values(int(lineBuffer2[2]))
        worksheet.write('C%s'%str(wt),lineBufferTMP[2])#value_n

    #total_g_u    
    lineBuffer3 = sheet3.row_values(i)
    if(lineBuffer3[1] != '0'):
        worksheet.write('A%s'%str(wt),lineBuffer[0])#source
        worksheet.write('F%s'%str(wt),lineBuffer[4])#search+
        worksheet.write('D%s'%str(wt),lineBuffer3[2])#target_u
        lineBufferTMP = sheet5.row_values(int(lineBuffer3[2]))
        worksheet.write('E%s'%str(wt),lineBufferTMP[2])#value_u
    if(lineBuffer2[1] != '0' or lineBuffer3[1] != '0'):
        wt = wt+1

workbook.close()
print('----final complete----')
