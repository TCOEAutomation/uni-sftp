import os
assetsPath=os.path.join(os.getcwd(),"uniAssets")
ExcelPath = assetsPath+"\UNI.xlsx"


sheetName='Input file for automation'
rowNumberWhereFieldsArePlaced=6;
sftpServerHost="vpce-0291a48b913751631-102boboj.server.transfer.ap-southeast-1.vpce.amazonaws.com"
sftpServerUserName="zecsunga"
sshServerPort=22
privateKeyOpenSshFileName=assetsPath+"/privateKey"
privateKeyOpenSshPassword=""

cronFileLocation="/unfs3dv02/uploads/triggers"

cronFileNames=["createLastMonthExpectedFileEntries","getAndProcessScheduledChannels"]

