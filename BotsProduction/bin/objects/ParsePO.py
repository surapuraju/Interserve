#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Raju.Surapuraju
#
# Created:     05/01/2019
# Copyright:   (c) Raju.Surapuraju 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import xlrd
from xlrd import open_workbook

workbook = xlrd.open_workbook('D:\BDG\jswright\PurchaseOrder\BotsProduction\\bin\data\PurchaseOrderA.xlsx')
print(workbook.nsheets)

# get the first worksheet
FS = workbook.sheet_by_index()[0]

# read a cell
#FAXNum = FS.cell(0,0)
#print(FAXNum)
