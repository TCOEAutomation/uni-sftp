import os
import paramiko
import time
import Config
from stat import S_ISDIR
import traceback
import pandas as pd
import customLibraries.Report as Report

username = Config.sftpServerUserName
mykey = paramiko.RSAKey.from_private_key_file(Config.privateKeyOpenSshFileName, password=Config.privateKeyOpenSshPassword)


def isdir(path):
    #returns true/false based on if the path exists or not, and if the path is a directory
  try:
    return S_ISDIR(sftp.stat(path).st_mode)
  except IOError:
    #Path does not exist, so by definition not a directory
    return False

def uploadFile(fileToUpload,remoteFileNameAbsolute):
    global sftp
    print "[ Uploading file ]"
    sftp.put(fileToUpload, remoteFileNameAbsolute)
    print "[ Verifying if file was uploaded successfully. Filename :  ] : "+remoteFileNameAbsolute
    fileUploadStatus= checkIfFileExists(remoteFileNameAbsolute)

    if fileUploadStatus:
        print "[ File was uploaded successfully ]"
        return True
    else:
        print "[ Uploaded File was not detected on the server, Terminating Execution ]"
        sys.exit(-1)


def goToPath(pathAbsolute):
    global sftp
    print "Absolute path to navigate : ",pathAbsolute
    sftp.chdir(pathAbsolute)

def checkIfFileExists(absPath):
    global sftp

    retries = 0
    isFileExisting = False

    while(not(isFileExisting) and retries != 20):
        try:
            print(sftp.stat(absPath))
            print "[INFO] File exists : ", absPath
            return True
        except IOError as e:
            retries = retries + 1
            print "[WARN] Unable to locate file in the server. Retries: " + str(retries)
            time.sleep(60)
    print "[ERR] Unable to locate file in the server after 20 minutes"
    return False

def downloadFile(remoteFileToDownload,localLocation):
    global sftp
    print "[ Downloading file : {0} to  {1} ]".format(remoteFileToDownload,localLocation)
    if checkIfFileExists(remoteFileToDownload):
        print "File exists, will download file [{0}] to [{1}] ".format(remoteFileToDownload,localLocation)
        try:
            sftp.get(remoteFileToDownload, localLocation, callback=None)
            retries = 0
            while(not(os.path.exists(localLocation)) and retries != 60):
                time.sleep(1)
                retries = retries + 1
            print "[INFO] File Downloaded. Retries: " + str(retries)
            return True
        except Exception,e:
            traceback.print_exc()
    else:
        print "Could not find file on the server : ",remoteFileToDownload
    return False

def Connect():
    global sftp
    global transport

    retries = 0
    isConnected = False

    while(not(isConnected) and retries != 10):
        try:
            transport = paramiko.Transport((Config.sftpServerHost, Config.sshServerPort))
            transport.connect(username=username, pkey=mykey)
            sftp = paramiko.SFTPClient.from_transport(transport)
            print "[INFO] Successfully connected to the server"
            return True
        except Exception as e:
            retries = retries + 1
            print "[WARN] Unable to connect to server. Retries: " + str(retries)
            time.sleep(15)
    print "[ERR] Unable to connect to server after 10 retries"
    return False

def listFiles():
    global sftp
    print sftp.listdir()


def getListOfFilesAfterSpecificTimestamp(absPath,timestamp):
    global sftp
    #files = str(pd.DataFrame([attr.__dict__ for attr in sftp.listdir_attr()]).sort_values("st_mtime", ascending=False))
    #files = [attr.__dict__ for attr in sftp.listdir_attr()].sort_values("st_mtime", ascending=False)

    curDir=sftp.getcwd()
    try:
        goToPath(absPath)
    except Exception as e:
        Report.WriteTestStep('Navigate to server path : {0}'.format(absPath),'Should be able to navigate and path should exist','Unable to navigate. Please check if path is present and user has access to it. Exception : {0}'.format(e),'Failed')
        return []

    files=[]
    timestamp=int(timestamp)

    for attr in sftp.listdir_attr():
        try:
            curFile=(str(attr).split(" ")[-1]).strip()

            if not isdir(curFile):
                fileTimestamp=int(str(attr.__dict__["st_mtime"]).strip())

                if fileTimestamp>timestamp:
                    files.append(curFile)
        except Exception as e:
                traceback.print_exc()
                Report.WriteTestStep('Exception while Get files after specific timestamp','','Exception : {0}'.format(e),'Failed')
    goToPath(curDir)
    return files

def closeConnection():
    global sftp
    global transport
    sftp.close()
    transport.close()
    print "Closed connection."