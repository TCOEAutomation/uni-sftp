import sys
import os
import glob
import xlrd

from collections import OrderedDict

class createHTMLreport():
    testsuites = OrderedDict()
    totalTC = 0
    failingTC = 0
    passingTC = 0
    totalCounter = 0

    def __init__(self, htmlPath='Results.html'):
        print("HTML Report will be generated in " + htmlPath)
        self.htmlPath = htmlPath

        self.file = open(self.htmlPath, 'w')

    def setExcelFile(self, excelPath):
        self.totalTC = 0
        self.failingTC = 0
        self.passingTC = 0
        self.testsuites.clear()
        self.excelName = os.path.basename(excelPath)
        try:
            xlrd.open_workbook(excelPath)
        except:
            print('No Excel "%s" Found' % excelPath)
            return
        self.wb = xlrd.open_workbook(excelPath)
        self.readExcel()

    def readExcel(self, excelSheet='Results'):
        try:
            self.wb.sheet_by_name(excelSheet)
        except xlrd.biffh.XLRDError:
            print('No Sheet "%s" Found' % excelSheet)
            return

        sheet = self.wb.sheet_by_name(excelSheet)
        for row in range(1, sheet.nrows):
            if "" != sheet.cell(row, 0).value:
                currentTestCase = sheet.cell(row, 0).value + " - " + sheet.cell(row, 2).value
                self.totalTC += 1

                if "Failed" == sheet.cell(row, 7).value:
                    self.testsuites[currentTestCase] = "Failed"
                    self.failingTC += 1

                if "Passed" == sheet.cell(row, 7).value:
                    self.testsuites[currentTestCase] = "Passed"
                    self.passingTC += 1
                self.totalCounter += 1

    def createIntro(self):
        print("Creating HTML Report")
        self.file.write('<html>\n<body  style="color: black">\n')
        self.file.write('Dear All,<br>')
        self.file.write('Please find the results for the execution. <br><br>')

    def create(self):
        self.createIntro()
        self.createTable()
        self.createSigniture()

    def createSigniture(self):
        self.file.write('<br/><br/>')

        self.file.write('Best Regards, <br>')
        self.file.write('TCoE Automation Team <br>')
        self.file.write('</body>\n</html>')
        self.file.close()

        print("HTML Report Creation Done")

    def createTable(self):
        self.file.write('<b>File Validation for ' + self.excelName + '</b><br>')
        self.file.write('Total Test Cases: ' + str(self.totalTC) + '\tPassed : ' + str(self.passingTC) + '\tFailed : ' + str(self.failingTC) + '<br>\n')

        if 0 != len(self.testsuites):
            self.file.write('<table style="border: 1px solid black;border-collapse: collapse;">\n')
            self.file.write('<th style="border: 1px solid black;border-collapse: collapse;background-color:#b2b2b2">\n');
            self.file.write('   Test case\n');
            self.file.write('</th>\n');
            self.file.write('<th style="border: 1px solid black;border-collapse: collapse;background-color:#b2b2b2">\n');
            self.file.write('   Status\n');
            self.file.write('</th>\n');
            self.createTestCase()
            self.file.write('</table><br>\n')


    def createTestCase(self):
        for testCase in self.testsuites:
            self.file.write('<tr style="border: 1px solid black;border-collapse: collapse;">\n')
            self.file.write('   <td style="border: 1px solid black;border-collapse: collapse;padding: 5px">')
            self.file.write(testCase)
            self.file.write('   </td>\n')

            self.file.write('   <td style="border: 1px solid black;border-collapse: collapse;padding: 5px">')

            if "Passed" == self.testsuites[testCase]:
                self.file.write('<b style="color: green">')

            else:
                self.file.write('<b style="color: red">')
            self.file.write(self.testsuites[testCase] + '</b>\n')
            self.file.write('   </td>\n')

            self.file.write('</tr>')

if __name__=="__main__":
    htmlReport = createHTMLreport(htmlPath="Results\\Results.html")
    files = glob.glob('Results' + '/**/*.xlsx')
    htmlReport.createIntro()
    for f in files:
        htmlReport.setExcelFile(f)
        htmlReport.createTable()
    htmlReport.createSigniture()
