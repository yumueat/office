import xlrd
import xlwt
import xlwings as xw
# # new_workbook=xlwt.Workbook()
# # new_workbook_sheet=new_workbook.add_sheet('Sheet1')
#
# old_workbook=app.Book('专业和排名.xlsx')
# old_workbook_sheet=app.sheets['Sheet1']
# new_workbook=app.Book()
# new_workbook_sheet=app.sheets['Sheet1']
# n=0
# for i in range(2,538):
#     if old_workbook_sheet.range('F'+str(i)).value=="软件工程":
#         # print(n)
#         # print(old_workbook_sheet.range('E'+str(i)).value,old_workbook_sheet.range('G'+str(i)).value)
#         new_workbook_sheet.range('A'+str(n)).value=1
#         new_workbook_sheet.range('B'+str(n)).value=1
#         # new_workbook_sheet.write(n,0,old_workbook_sheet.range('E'+str(i)).value)
#         # new_workbook_sheet.write(n,1,old_workbook_sheet.range('G'+str(i)).value)
#         print("over!")
#         n+=1
# # new_workbook.save('软件工程整合.xls')
# # new_workbook_read=xlrd.open_workbook('软件工程整合.xls')
# # new_workbook=xlwt.Workbook('软件工程整合.xls')
# # new_workbook_sheet=new_workbook.sheet_index('Sheet1')
# # new_workbook_sheet_read=new_workbook_read.sheet_by_index(0)
# # for i in range(0,n):
# #     for j in range(i,n):
# #         if new_workbook_sheet_read.cell_value(j,1)>new_workbook_sheet_read.cell_value(j+1,1):
# #             temp=new_workbook_sheet_read.cell_value(j,1)
# #             new_workbook_sheet.write(j,1)=new_workbook_sheet_read.cell_value()
# # print(old_workbook_sheet.range('A1').value)
# new_workbook.save('软件工程整合.xlsx')
# new_workbook.close()
# old_workbook.close()
# # new_workbook.close()
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
wb_old=app.books.open('专业和排名.xlsx')
sht_old=wb_old.sheets[0]
wb_new=app.books.add()
sht_new=wb_new.sheets.add()
n=1
for i in range (2,538):
    if sht_old.range('F'+str(i)).value=="软件工程":
        sht_new.range('A'+str(n)).value=sht_old.range('E'+str(i)).value
        sht_new.range('B'+str(n)).value=sht_old.range('G'+str(i)).value
        n+=1
        # print("over!")
for i in range(1,n):
    for j in range(i,n-1):
        if sht_new.range('B'+str(j)).value>sht_new.range('B'+str(j+1)).value:
            temp=sht_new.range('B'+str(j)).value
            sht_new.range('B' + str(j)).value=sht_new.range('B'+str(j+1)).value
            sht_new.range('B' + str(j + 1)).value=temp
            # print(i,j)
wb_new.save('软件工程.xlsx')
wb_new.close()
wb_old.close()




