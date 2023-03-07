@echo off
set list="AR GLOBE" "AR Bayan Zip with password" "US Innove" "US Bayan Zip with password" "US Bayan Zip with no password" "AR Innove"
(for %%a in (%list%) do (
   python -u run.py %%a
   python -u scripts\createJUnitReport.py UNI SFTP Results\Result_%%a\Report_%%a.xlsx Results\Results_%%a.xml
   copy Results\Result_%%a\Report_%%a.pdfReport.docx .
))
python -u scripts\createHTMLreport.py