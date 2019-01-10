#-------------------------------------------------------------------------------
# Name:        POXLSParser
# Purpose:     JS Wright Purchase Order Parser for Tradex Update
#
# Author:      Raju.Surapuraju
#
# Created:     05/01/2019
# Copyright:   (c) Raju.Surapuraju 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import xlrd
from xlrd import open_workbook

#----------------------------------------------------------------------
poHeaderString = ("Transaction||Supplier||SupplierAddL1||SupplierAddL2||SupplierAddL3||SupplierAddL4||SupplierAddL5||SupplierFAX||SupplierNo||ReferenceNo||Contact||Manufacturer||Type||DeliveryAddL1||DeliveryAddL2||DeliveryAddL3||DeliveryAddL4||DeliveryAddL5||OrderNumber||OrderDate||OrderDescription||DateRequired||PlacedBy||CatDesc1||CatDesc2||CatDesc3||CatNumber||CatQuantity||CatUnitPrice||CatTotal||DateReceived")
#print(poHeaderString)
#poDataString = ""

poParsedFile = "D:\BDG\jswright\PurchaseOrder\BotsProduction\\bin\data\poParsedFile.txt"
poFileHandle = open(poParsedFile, "a")
poFileHandle.write(poHeaderString + '\n')
poFileHandle.close()

def parseXLSX(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)

    # print number of sheets
    #print(book.nsheets)

    # print sheet names
    #print(book.sheet_names())

    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    # read a row
    #print(first_sheet.row_values(0))

    # Read PO Tile
    PO = (first_sheet.cell(0,7))
    PO = PO.value
    poDataString = (str(PO) + "||")
    #print(poDataString)
    #print(PO.value)

    Supplier = first_sheet.cell(1,0)
    Supplier = Supplier.value
    poDataString = (poDataString + str(Supplier) + "||")
    #print(poDataString)
    #print(Supplier.value)

    SuppAddL1 = first_sheet.cell(2,0)
    SuppAddL1 = SuppAddL1.value
    poDataString = (poDataString + str(SuppAddL1) + "||")
    #print(poDataString)
    #print(SuppAddL1.value)

    SuppAddL2 = first_sheet.cell(3,0)
    SuppAddL2 = SuppAddL2.value
    poDataString = (poDataString + str(SuppAddL2) + "||")
    #print(poDataString)
    #print(SuppAddL2.value)

    SuppAddL3 = first_sheet.cell(4,0)
    SuppAddL3 = SuppAddL3.value
    poDataString = (poDataString + str(SuppAddL3) + "||")
    #print(poDataString)
    #print(SuppAddL3.value)

    SuppAddL4 = first_sheet.cell(5,0)
    SuppAddL4 = SuppAddL4.value
    poDataString = (poDataString + str(SuppAddL4) + "||")
    #print(poDataString)
    #print(SuppAddL4.value)

    SuppAddL5 = first_sheet.cell(6,0)
    SuppAddL5 = SuppAddL5.value
    poDataString = (poDataString + str(SuppAddL5) + "||")
    #print(poDataString)
    #print(SuppAddL5.value)

    FAXNum = first_sheet.cell(8,0)
    FAXNum = FAXNum.value
    poDataString = (poDataString + str(FAXNum) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    SupplierNo = first_sheet.cell(9,0)
    SupplierNo = SupplierNo.value
    poDataString = (poDataString + str(SupplierNo) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    #ReferenceNo||Contact||Manufacturer||Type

    ReferenceNo = first_sheet.cell(11,1)
    ReferenceNo = ReferenceNo.value
    poDataString = (poDataString + str(ReferenceNo) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    Contact = first_sheet.cell(13,1)
    SupplierNo = Contact.value
    poDataString = (poDataString + str(Contact) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    Manufacturer = first_sheet.cell(16,1)
    SupplierNo = Manufacturer.value
    poDataString = (poDataString + str(Manufacturer) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    Type = first_sheet.cell(18,1)
    Type = Type.value
    poDataString = (poDataString + str(Type) + "||")
    #print(poDataString)
    #print(FAXNum.value)

    DeliveryAdd = first_sheet.cell(1,8)
    DeliveryAdd = DeliveryAdd.value
    poDataString = (poDataString + str(DeliveryAdd) + "||")
    #print(poDataString)
    #print(DeliveryAdd.value)

    DeliveryAddL1 = first_sheet.cell(2,8)
    DeliveryAddL1 = DeliveryAddL1.value
    poDataString = (poDataString + str(DeliveryAddL1) + "||")
    #print(poDataString)
    #print(DeliveryAddL1.value)

    DeliveryAddL2 = first_sheet.cell(3,8)
    DeliveryAddL2 = DeliveryAddL2.value
    poDataString = (poDataString + str(DeliveryAddL2) + "||")
    #print(poDataString)
    #print(DeliveryAddL2.value)

    DeliveryAddL3 = first_sheet.cell(4,8)
    DeliveryAddL3 = DeliveryAddL3.value
    poDataString = (poDataString + str(DeliveryAddL3) + "||")
    #print(poDataString)
    #print(DeliveryAddL3.value)

    DeliveryAddL4 = first_sheet.cell(5,8)
    DeliveryAddL4 = DeliveryAddL4.value
    poDataString = (poDataString + str(DeliveryAddL4) + "||")
    #print(poDataString)
    #print(DeliveryAddL4.value)

    DeliveryAddL5 = first_sheet.cell(6,8)
    DeliveryAddL5 = DeliveryAddL5.value
    poDataString = (poDataString + str(DeliveryAddL5) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)

    OrderNo = first_sheet.cell(9,10)
    OrderNo = OrderNo.value
    poDataString = (poDataString + str(OrderNo) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)

    OrderDate = first_sheet.cell(11,10)
    OrderDate = OrderDate.value
    poDataString = (poDataString + str(OrderDate) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)

    OrderDesc = first_sheet.cell(13,10)
    OrderDesc = OrderDesc.value
    poDataString = (poDataString + str(OrderDesc) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)

    DateReq = first_sheet.cell(16,10)
    OrderDesc = DateReq.value
    poDataString = (poDataString + str(DateReq) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)

    PlacedBy = first_sheet.cell(18,10)
    OrderDesc = PlacedBy.value
    poDataString = (poDataString + str(PlacedBy) + "||")
    #print(poDataString)
    #print(DeliveryAddL5.value)



    # read a row slice
    #print(first_sheet.row_slice(rowx=23, start_colx=0, end_colx=18))
    #print((first_sheet.row_slice(rowx=23, start_colx=0, end_colx=18).value())

    #print(first_sheet.row_values(23))
    #print(first_sheet.row_values(24))
    #print(first_sheet.row_values(25))
    #print(first_sheet.row_values(26))
    #print(first_sheet.row_values(27))

    poDataLineString = ""
    for sheet in book.sheets():
        for rowidx in range(sheet.nrows):
            row = sheet.row(rowidx)
            for colidx, cell in enumerate(row):
                if cell.value == "Description" :
                    print(sheet.name)
                    print(rowidx+2)
                    print(colidx)
                    rows = first_sheet.nrows
                    columns = first_sheet.ncols
                    #print(rows)
                    #print(columns)

                    num_rows = first_sheet.nrows - 1
                    #curr_row = -1
                    curr_row = (rowidx+1)

                    cnt = 1
                    while curr_row < num_rows:
                            curr_row += 1
                            row = first_sheet.row(curr_row)
                            #print(row)
                            poCatelogueCell = first_sheet.cell_value(curr_row,0)
                            poCType = first_sheet.cell(curr_row,0).ctype
                            #print (poCType)

                            if poCType == 0:
                                print("Change Catelogue Line")
                                nextLine = 'Yes'
                                cnt = 1

                            else:
                                #CatDesc1||CatDesc2||CatDesc3||CatNumber||CatQuantity||CatUnitPrice||CatTotal||DateReceived

                                CatNumberCType = first_sheet.cell(curr_row,2).ctype
                                if CatNumberCType != 0:
                                    CatDesc1 = first_sheet.cell(curr_row,0)
                                    CatDesc1 = CatDesc1.value
                                    poDataLineString = (poDataLineString + str(CatDesc1) + "||")
                                    #print(poCatelogueCell)
                                    #print(poDataString)
                                    if cnt == 2:
                                        CatDesc2 = first_sheet.cell(curr_row,0)
                                        CatDesc2 = CatDesc2.value
                                        poDataLineString = (poDataLineString + str(CatDesc2) + "||")

                                    if cnt == 3:
                                        CatDesc3 = first_sheet.cell(curr_row,0)
                                        CatDesc3 = CatDesc3.value
                                        poDataLineString = (poDataLineString + str(CatDesc3) + "||")

                                    CatNumber = first_sheet.cell(curr_row,2)
                                    CatNumber = CatNumber.value
                                    poDataLineString = (poDataLineString + str(CatNumber) + "||")
                                    #print(poCatelogueCell)
                                    #print(poDataString)

                                    CatQuantity = first_sheet.cell(curr_row,4)
                                    CatDesc1 = CatQuantity.value
                                    poDataLineString = (poDataLineString + str(CatQuantity) + "||")
                                    #print(poCatelogueCell)
                                    #print(poDataString)

                                    CatUnitPrice = first_sheet.cell(curr_row,11)
                                    CatDesc1 = CatUnitPrice.value
                                    poDataLineString = (poDataLineString + str(CatUnitPrice) + "||")
                                    #print(poCatelogueCell)
                                    #print(poDataString)

                                    CatTotal = first_sheet.cell(curr_row,15)
                                    CatDesc1 = CatTotal.value
                                    poDataLineString = (poDataLineString + str(CatTotal) + "||")
                                    #print(poCatelogueCell)
                                    #print(poDataString)

                                    DateReceived = first_sheet.cell(curr_row,17)
                                    CatDesc1 = DateReceived.value
                                    poDataLineString = (poDataLineString + str(DateReceived))
                                    #print(poCatelogueCell)
                                    #print(poDataString)
                                    cnt = cnt + 1

                                    finalDataString = poDataString + poDataLineString
                                    poFileHandle = open(poParsedFile, "a")
                                    poFileHandle.write(finalDataString + '\n')
                                    poFileHandle.close()
                                    poDataLineString = ""

    # First row will have all details along with Description1
    # Second row will have Descripotion2
    # Third row will have Description3 - this is optional

#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "D:\BDG\jswright\PurchaseOrder\BotsProduction\\bin\data\PurchaseOrderB.xlsx"
    parseXLSX(path

)

