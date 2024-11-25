import datetime
import xlwt

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('sheet1')
sheet.write(0, 0, '序号')
sheet.write(0, 1, '状态')
sheet.write(0, 2, '链接')
sheet.write(0, 3, '内容')
sheet.write(0, 4, '异常')



workbook.save(f'./output/{datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xls')

input("Press Enter to exit")