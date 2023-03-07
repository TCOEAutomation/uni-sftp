import Report


Report.InitializeReporting()

Report.WriteTestCase("Sample Scenario","Just a dummy description")
Report.WriteTestStep("Sample desc","Should be ok","is ok","Pass")
Report.WriteTestStep("Sample failed desc","Should be ok","is ok","Fail")
Report.evaluateIfTestCaseIsPassOrFail()
Report.GeneratePDFReport()