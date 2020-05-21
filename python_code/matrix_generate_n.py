import xlrd
import xlsxwriter
import string
import synonyms
import numpy

print('----matrix_n_generate initiated----')
ExcelFile = xlrd.open_workbook(r'..\total_b.xls')
sheet = ExcelFile.sheet_by_index(0)
nrows = sheet.nrows
ncols = sheet.ncols
ExcelFile2 = xlrd.open_workbook(r'..\data2\new\ttn_b.xls')
sheet2 = ExcelFile2.sheet_by_index(0)
nrows2 = sheet2.nrows
ncols2 = sheet2.ncols

workbook = xlsxwriter.Workbook(r'..\total_g_n.xls')
worksheet = workbook.add_worksheet(u'sheet1')
worksheet.write('A1','Source')
worksheet.write('B1','Related_Rank_n')
worksheet.write('C1','Related_Target_n')

for i in range(1,nrows-1):
    dear_max = 0
    tgt = 0
    dear = numpy.mat(numpy.zeros((ncols,ncols2)))
    lineBuffer = sheet.row_values(i)
    for tmp in range(1,nrows-1):
        keyBuffer = sheet2.row_values(tmp)
        for p in range(0,len(lineBuffer)-1):
            if(lineBuffer[p] != ''):
                for q in range(0,len(keyBuffer)-1,3):
                    [buffer1,et]=synonyms.nearby(lineBuffer[p])
                    [buffer2,et]=synonyms.nearby(keyBuffer[q])
                    if(buffer1!=[] and buffer2!=[]):
                        dear[p,q%3] = synonyms.compare(lineBuffer[p],keyBuffer[q],False)
        avg = numpy.mean(dear)
        if(avg > dear_max):
            dear_max = avg
            tgt = tmp
    worksheet.write('A%s'%str(i),str(i))
    worksheet.write('B%s'%str(i),str(dear_max))
    worksheet.write('C%s'%str(i),str(tgt))
#percentage display

    percent=float(i)*100/float(nrows-1)
    print("%.1f"%percent);
print("100%!finish!\r");

workbook.close()
print('----matrix_n_generate complete----')
