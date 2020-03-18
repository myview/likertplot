#!/usr/bin/python3

from   optparse import OptionParser
from   datetime import datetime
import pandas as pd
import pysftp
import os

"""
Install:
 1) https://www.python.org/ftp/python/3.8.2/python-3.8.2-amd64.exe
 2) pip3 install pysftp
 3) pip3 install pandas

"""

class MasterExcel:

    def __init__(self, filename, outputfile):
        self.df = pd.read_excel(filename)
        self.outputfile = outputfile

    def addDateToFilename(self):
        (name, ending) = os.path.splitext(self.outputfile)
        date = datetime.today().strftime('%Y%m%d')
        self.outputfile = name + date + ending
        
    def toInteger(self, column):
        self.df[column] = self.df[column].astype('Int64')

    def toDateTime(self, column):
        self.df[column] = self.df[column].astype('datetime64')

    def dropEmptyRows(self, column):
        self.df = self.df.dropna(subset=[column])

class AdImportFile(MasterExcel):

    def __init__(self, filename, options):
        MasterExcel.__init__(self, filename, "ADexterne.csv")

        # Rename existing columns
        self.df = self.df.rename(columns={
            'Mitarbeiter-Nummer'    : 'employeeID',
            'Erster Arbeitstag'     : 'entrydate',
            'Vertragsende'          : 'exitdate',
            'Nachname'              : 'givenName',
            'Vorname'               : 'sn',
            'Vorgesetzer'           : 'manager',
            })

        # Copy columns and set Picking and Packign according to rule
        self.df['lastworkday']                      = self.df.exitdate
        self.df['team']                             = self.df.manager
        self.df.loc[self.df.team != 633, 'team']    = 'Packing'
        self.df.loc[self.df.team == 633, 'team']    = 'Picking'
        self.df['group']                            = self.df.team

        # Re-Index to new output columns
        self.df = self.df.reindex([
            'employeeID',
            'entrydate',
            'exitdate',
            'lastworkday',
            'givenName',
            'sn',
            'mail',
            'mailPrivate',
            'mobilePhone_country',
            'mobile',
            'businessPhone_country',
            'telephoneNumber',
            'homePhone_country',
            'homePhone',
            'phone_modify',
            'title',
            'company',
            'department',
            'department_short',
            'team',
            'group',
            'co',
            'c',
            'CountryCode',
            'physicalDeliveryOfficeName',
            'preferredLanguage',
            'manager',
            'matrixrole1',
            'matrixrole2',
            'matrixrole3',
            'leaderrole',
            'departmentrole'
            ], axis=1)

        # Set default values
        self.df.title                       = 'ExternalWarehouseWorker'
        self.df.company                     = 'Digitec Galaxus AG'
        self.df.department                  = 'Supply Chain Management'
        self.df.department_short            = 'SCM'
        self.df.co                          = 'Switzerland'
        self.df.c                           = 'CH'
        self.df.CountryCode                 = '756'
        self.df.physicalDeliveryOfficeName  = 'Wohlen'
        self.df.preferredLanguage           = 'de_CH'
        self.df.leaderrole                  = 'Employee (FS1)'
        self.df.departmentrole              = 'SCM Employee'

        # Drop empty rows
        self.dropEmptyRows('employeeID')

        # Make real integers where required
        self.toInteger('employeeID')
        self.toInteger('manager')

        # Add current date to filename
        self.addDateToFilename()
        
        # Store to output file
        self.df.to_csv(self.outputfile,
                       index=False,
                       sep=';')



class EcAsesEmployeeData(MasterExcel):

    def __init__(self, filename, options):
        MasterExcel.__init__(self, filename, "EC_ASES_EMPLOYEE_DATA.csv")

        # Drop empty rows
        self.dropEmptyRows('Mitarbeiter-Nummer')

        # Set sepecific cell types
        self.toInteger('Mitarbeiter-Nummer')
        self.toInteger('Vorgesetzer')
        self.toInteger('Gruppe.1')
        self.toInteger('Badgenummer')
        self.toInteger('Arbeitstage pro woche')
        self.toDateTime('Eintrittsdatum')
        self.toDateTime('Eintrittsdatum.1')

        # Store to output file
        self.df.to_csv(self.outputfile,
                       index=False,
                       header=False,
                       date_format='%Y-%m-%d',
                       sep=';')


def upload(options, host, remoteFilePath, localFilePath):
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None 
    with pysftp.Connection(host     = host,
                           username = options.user,
                           password = options.pw,
                           cnopts   = cnopts
                           ) as sftp:
        print ("Connection succesfully stablished ... ")
        sftp.put(localFilePath, remoteFilePath)
        print ("Upload done")

def main():
    parser = OptionParser("usage: process.py SOURCE [options]")

    parser.add_option("-m", "--mode",
                      default="ALL",
                      help="Export mode: (AD) ActieDirectory,"
                      "(EC) Employee Central or (ALL)"
                      "[default: %default]")

    parser.add_option("-u", "--user",
                      action="store",
                      dest="user",
                      help="SFTP user name")

    parser.add_option("-p", "--pw",
                      action="store",
                      dest="pw",
                      help="SFTP password")

    parser.add_option("-x", "--upload",
                      action="store_true",
                      dest="sftp",
                      default=False,
                      help="SFTP upload file")


    (options, args) = parser.parse_args()

    if len(args) != 1:
        parser.error("incorrect number of arguments")


    if (options.mode == "ALL" or options.mode == "EC"):
        run = EcAsesEmployeeData(args[0], options)
        print(f'File written: {run.outputfile}')
        if (options.sftp):
            upload(options          = options,
                   host             = 'ftp.digitecgalaxus.ch',
                   remoteFilePath   = '/',
                   localFilePath    = run.outputfile)

    if (options.mode == "ALL" or options.mode == "AD"):
        run = AdImportFile(args[0], options)
        print(f'File written: {run.outputfile}')
        if (options.sftp):
            upload(options          = options,
                   host             = 'sftp012.successfactors.eu',
                   remoteFilePath   = f'/incoming/ADExport/{run.outputfile}',
                   localFilePath    = f'./{run.outputfile}')

if __name__ == "__main__":
    main()



