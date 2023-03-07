import sys
import os
import csv

def convertTextToData(inputFile):
    rows = []
    with open(inputFile, 'r') as in_file:
        for line in in_file:
            row = []
            transactDate = line[0:8].strip()
            row.append(transactDate)
            customerName = line[8:48].strip()
            row.append(customerName)
            accountNumber = line[48:68].strip()
            row.append(accountNumber)
            amount = line[68:81].strip()
            row.append(amount)
            depositNubmer = line[81:96].strip()
            row.append(depositNubmer)
            serviceNumber = line[96:116].strip()
            row.append(serviceNumber)
            bankCode = line[116:126].strip()
            row.append(bankCode)
            bankAccountNumber = line[126:146].strip()
            row.append(bankAccountNumber)
            terminalNumber = line[146:161].strip()
            row.append(terminalNumber)
            referenceNumber = line[161:186].strip()
            row.append(referenceNumber)
            debitCreditIndicator = line[186:187].strip()
            row.append(debitCreditIndicator)
            cashCheckIndicator = line[187:188].strip()
            row.append(cashCheckIndicator)
            paymentMode = line[188:190].strip()
            row.append(paymentMode)
            applicationDate = line[190:198].strip()
            row.append(applicationDate)
            creditCardApprovalNo = line[198:213].strip()
            row.append(creditCardApprovalNo)
            creditCardExpirationDate = line[213:221].strip()
            row.append(creditCardExpirationDate)
            filler = line[221:231].strip()
            row.append(filler)
            rows.append(row)

    print "[INF]Generated " + str(len(rows)) + " accounts"
    return rows

def createCsvFile(outputFile, rows):
    with open(outputFile, 'wb') as out_file:
        writer = csv.writer(out_file)
        writer.writerows(rows)

if __name__=="__main__":
    if 3 > len(sys.argv):
        print "\n"
        print "*******************************************************************"
        print "Usage : python <Python Path> <Input Text File> <Output CSV Path>"
        print "*******************************************************************"
        print "\n"
        sys.exit(-1)

    inputFile = sys.argv[1]
    outputFile = sys.argv[2]
    rows = convertTextToData(inputFile)
    createCsvFile(outputFile, rows)