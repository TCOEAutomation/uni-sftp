import sys
import os
import xlrd

from collections import OrderedDict

class createJUnitReport:
	testsuites = OrderedDict()

	def __init__(self, projectName, className, excelPath, jsonPath='Results.xml'):
		self.jsonPath    = jsonPath
		self.projectName = projectName
		self.className   = className
		self.stackTrace  = ""
		self.testsuites.clear()

		print("JUnit Report for " + projectName + " will be generated in " + jsonPath)

		try:
			xlrd.open_workbook(excelPath)
		except:
			print('No Excel "%s" Found' % excelPath)
		self.wb = xlrd.open_workbook(excelPath)
		self.readExcel(excelPath)

		self.file = open(self.jsonPath, 'w')

	def readExcel(self, excelPath, excelSheet='Results'):
		try:
			self.wb.sheet_by_name(excelSheet)
		except xlrd.biffh.XLRDError:
			print('No Sheet "%s" Found' % excelSheet)
			return

		sheet = self.wb.sheet_by_name(excelSheet)
		for row in range(1, sheet.nrows):
			testStatus = ""
			if "" != sheet.cell(row, 0).value:
				currentTestCase = sheet.cell(row, 0).value.encode('ascii', 'ignore')
				currentTestCaseName = sheet.cell(row, 2).value.encode('ascii', 'ignore')
				if "Passed" != sheet.cell(row, 7).value:
					self.testsuites[currentTestCase] = "[Failed] " + currentTestCase + " - " + currentTestCaseName + "\n"
				else:
					self.testsuites[currentTestCase] = "Passed"

			if "" != sheet.cell(row, 1).value and "Passed" != self.testsuites[currentTestCase]:
				testStep = sheet.cell(row, 1).value.encode('ascii', 'ignore')
				testStepDescription = sheet.cell(row, 6).value.encode('ascii', 'ignore')
				if "Failed" == sheet.cell(row, 7).value:
					self.testsuites[currentTestCase] = self.testsuites[currentTestCase] + "\t[Failed] Step " + testStep + " : " + testStepDescription + "\n"
				if "Passed" == sheet.cell(row, 7).value:
					self.testsuites[currentTestCase] = self.testsuites[currentTestCase] + "\t[Passed] Step " + testStep + " : " + testStepDescription + "\n"


	def create(self):
		if 0 == len(self.testsuites):
			print ("No Test Cases parsed.")
			return

		print("Creating JUnit Report")

		self.file.write('<testsuite>\n')
		self.createTestCase()
		self.file.write('</testsuite>\n')
		self.file.close()

		print("JUnit Report Creation Done")

	def createTestCase(self):
		for testCase in self.testsuites:
			if "Passed" == self.testsuites[testCase]:
				self.file.write('\t<testcase classname="' + self.projectName + '.' + self.className + '" name="' + testCase + '"/>\n')
			else:
				self.file.write('\t<testcase classname="' + self.projectName + '.' + self.className + '" name="' + testCase + '">\n')
				self.file.write('\t\t<failure type="fail"> ' + str(self.testsuites[testCase]) + ' </failure>\n')
				self.file.write('\t</testcase>\n')

	def getFailingTestcaseCount(self):
		return self.failingTCs

if __name__=="__main__":

	if 4 > len(sys.argv):
	    print "\n"
	    print "*******************************************************************************************************"
	    print "Usage : python <Python Path> <Project Name> <Class Name> <Input Excel Path> <Output JunitReport Name>"
	    print "*******************************************************************************************************"
	    print "\n"
	    sys.exit(-1)

	projectName = sys.argv[1]
	className = sys.argv[2]
	excelPath = sys.argv[3]
	jsonPath = sys.argv[4]

	JUnitReport = createJUnitReport(projectName=projectName, className=className, excelPath=excelPath, jsonPath=jsonPath)
	JUnitReport.create()