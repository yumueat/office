import xlrd
import xlwt
import xlwings as xw
# 写入操作
# new_workbook=xlwt.Workbook()
# worksheet=new_workbook.add_sheet('new_test')
# worksheet.write(1,2,'test')
# new_workbook.save('test.xls')
# 读操作
old_workbook=xlrd.open_workbook('答题卡.xls')
old_sheet=old_workbook.sheet_by_index(0)
data=old_sheet.cell_value(5,4)
print(type(data))
print(data)

# wb=xw.Book()#新建一个文档
# wb=xw.Book('test.xls')#打开一个已经有的文档
#
# sht=wb.sheets['Sheet1'] #找到指定sheet
# print(sht.range('A1').value) #读取指定单元格的数据
# sht.range('B1').value=10 #给指定的空白的单元格赋值

# wb.save(r'test.xlsx') #保存excel文档
# wb.close() #退出程序
