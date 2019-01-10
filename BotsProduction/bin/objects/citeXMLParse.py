#-------------------------------------------------------------------------------
# Name:        POXLSParser
# Purpose:     JS Wright Purchase Order Parser for Tradex Update
#
# Author:      Raju.Surapuraju
#
# Created:     04/01/2019
# Copyright:   (c) Raju.Surapuraju 2018
# Licence:     Causeway Technologies Limited
#-------------------------------------------------------------------------------

import re
import os
import configparser

settings = configparser.ConfigParser()
settings._interpolation = configparser.ExtendedInterpolation()
settings.read('C:\\config\\config.ini')
varciteXMLInput = settings.get('InputData', 'citeXMLInput')
varciteCSV = settings.get('InputData', 'citeCSV')

citeXML = varciteXMLInput
citeCSV = varciteCSV

if os.path.exists(citeCSV):
  os.remove(citeCSV)
else:
  print("citeCSV file does not exist")

def parseXML():
    with open(citeXML) as cxFH:
        while True:
            line = cxFH.readline()
            line = str(line)
            invoiceNumStr = re.search('DocNumber="(.+?)" DocType=', line)
            catalogueNumberStr = re.search('E.7140.ItemNumber.D="(.+?)"/', line)
            catalogueLineNumberStr = re.search('E.1082.LineItemNumber.D="(.+?)">', line)
            catalogueLineQuantityStr = re.search('E.6060.Quantity.D="(.+?)"', line)
            if invoiceNumStr:
                invoiceNum = invoiceNumStr.group(1)
                print ("Invoice Number:" + invoiceNum)

            if catalogueNumberStr:
                catalogueNumber = catalogueNumberStr.group(1)
                print (" Catelogue Number:" + catalogueNumber)

            if catalogueLineNumberStr:
                catalogueLineNumber = catalogueLineNumberStr.group(1)
                print (" Catalogue Line Number:" + catalogueLineNumber)

            if catalogueLineQuantityStr:
                catalogueLineQuantity = catalogueLineQuantityStr.group(1)
                print (" Catalogue Line Quantity:" + catalogueLineQuantity)
                print ()

                csvString =  (invoiceNum + ',' +  catalogueNumber + ','  + catalogueLineNumber + ',' + catalogueLineQuantity)
                csvFH = open(citeCSV, "a")
                csvFH.write(csvString + '\n')
                csvFH.close()

            if not line:
                break
        cxFH.close()

def main():
    parseXML()

if __name__ == '__main__':
    main()

