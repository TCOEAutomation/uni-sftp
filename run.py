import os, sys
import Config
from adjustDates import adjustDates
import createFile
import DumpFilesToServer
import validations

file = os.path.join(Config.assetsPath, sys.argv[1] + ".xlsx")
print "[INFO] Running test scripts for " + file
AdjustDates = adjustDates(file)
AdjustDates.run()
createFile.run(file)
DumpFilesToServer.run(file)
validations.run(file)