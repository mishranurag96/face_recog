# -*- coding: utf-8 -*-
"""
Created on Mon Mar 26 19:39:54 2018

@author: Anurag Mishra
"""
from xlwt import Workbook , Formula

class Excel(object):
    wb = Workbook()
    sheet2 =wb.add_sheet('Sheet 1')
    sheet2.write(0,0,'ROLL NO.')
    sheet2.write(0,1,'NAME')
    sheet2.write(0,2,'ATTENDANCE')
    sheet2.col(0).width = 6000
    sheet2.col(1).width = 6000
    sheet2.col(2).width = 6000
    wb.save('xlwt attend_sheet.xls')







#to insert formulas below add_sheet
#for i in xrange(10):
#sheet1.write(i,0,i)
#sheet1.write(10,0,Formula('SUM(A1:A10)'))
