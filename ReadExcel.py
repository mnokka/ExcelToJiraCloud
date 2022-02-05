#encoding=latin1

# Tool to read Excel data and create issues to Jira
#
# Author mika.nokka1@gmail.com 4.2.2022 
#
# python3
#
#from __future__ import unicode_literals

import openpyxl 
import sys, logging
import argparse
#import re
from collections import defaultdict
#from author import Authenticate  # no need to use as external command
#from author import DoJIRAStuff
import glob
import json # for json dumo
#from sqlalchemy.sql.expression import false
import re
import time
import os
import datetime
from IssueCreator import CreateIssue


######################################################################################
#CONFIGURATIONS
__version__ = "0.5"
#
# FOR CODE EMBEDDED  SETTINGS ====>  SEE Parse FUNCTION
#
logging.basicConfig(level=logging.DEBUG) # INFO IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out

#######################################################################################

start = datetime.datetime.now()


def main(argv):
    
    JIRASERVICE=""
    JIRAPROJECT=""
    PSWD=''
    USER=''
  
    logging.debug ("--Python starting Excel reading --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    parser.add_argument('-f','--filepath', help='<Path to Excel file directory>')
    parser.add_argument('-n','--filename', help='<Excel filename>')
   # parser.add_argument('-m','--subfilename', help='<Subtasks Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password or token>')
    parser.add_argument('-u','--user', help='<JIRA user or email address if Jira cloud token used>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    parser.add_argument('-p','--project', help='<JIRA project key>')
   
    #parser.add_argument('-a','--attachemnts', help='<Attachment directory>')
        
    args = parser.parse_args()
    
    if args.version:
        print ("Tool version: {0}".format(__version__))
        sys.exit(2)    
           
    filepath = args.filepath or ''
    filename = args.filename or ''
    #subfilename=args.subfilename or ''
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    #ATTACHDIR=args.attachemnts or ''
    
    # quick old-school way to check needed parameters
    if (filepath=='' or  filename=='' or JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER=='' ):
        parser.print_help()
        sys.exit(2)
        


    Parse(filepath, filename,JIRASERVICE,JIRAPROJECT,PSWD,USER)


############################################################################################################################################
# Parse excel and create dictionary of
# Jiraissue data
#  
# Uses hardcoded sheet/column value,change accordingly

def Parse(filepath, filename,JIRASERVICE,JIRAPROJECT,PSWD,USER):
    

    ####################################################################################
    # CONFIGURATIONS ####
    PROD=True # False / True   #false skips issue creation and other jira operations
    ENV="PROD" # or "PROD" or "DEV", sets the custom field IDs 
    AUTH="False" # True / False ,for  jira authorizations
    DRY="on" # on==dry run, dont do   off=do everything THIS IS THE ONE FLAG TO RULE THEM ALL
    # END OF CONFIGURATIONS ############################################################
    
    # flag to indicate whether issue under operations have been already created to Jira
    IMPORT=False
    
    logging.info ("Filepath: %s     Filename:%s" %(filepath ,filename))
    files=filepath+"/"+filename
    logging.info ("Excel file:{0}".format(files))
   
    Issues=defaultdict(dict) 
    MainSheet="Sheet1" 

    try:
        wb= openpyxl.load_workbook(files)
    except: 
      logging.error ("Can't open excel. EXITING")
      sys.exit(5)
       
    types=type(wb)
    logging.debug ("Type:{0}".format(types))
    sheets=wb.sheetnames
    logging.debug ("Sheets:{0}".format(sheets))
   
    CurrentSheet=wb[MainSheet] 
    logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("Column A, row:{0}".format(CurrentSheet['A4'].value))




    ##############################################################################
    #CONFIGURATIONS AND EXCEL COLUMN MAPPINGS
    #
    DATASTARTSROW=2 # data section starting line in the excel sheet
    A=1 #key-id
    B=2 #toka
    C=3 #kolmas
    D=4 #neljas
    E=5 #viides
    
    
    
    ####################################################################################################################
    # Go through main excel sheet for main issue keys (and contents findings)
    # Create dictionary structure
    # NOTE: Uses hardcoded sheet/column values. A is assumed to hold KEY, used to create dictionary entry
    # NOTE: As this handles first sheet, using used row/cell reading (buggy, works only for first sheet) 
    #
    i=DATASTARTSROW # brute force row indexing
    ENDROW=(CurrentSheet.max_row) # to prevent off-by-one in the end of sheet, also excel needs deleting of empty end line1
    logging.debug ("STARTROW:{0} ENDROW:{1}".format(DATASTARTSROW,ENDROW))

    for row in CurrentSheet[('A{}:A{}'.format(DATASTARTSROW,ENDROW))]:  # go trough all column A (KEY) rows
        for mycell in row:
            KEY=mycell.value # column A
            logging.debug  ("KEY:{0} data in ROW:{1}".format(KEY,i))
            Issues[KEY]={} # add KEY to dictionary as master key 
            
            #Hardcoded ovalue picking and dictionary settings operations
        
            COLUMNB=(CurrentSheet.cell(row=i, column=B).value)
            if not COLUMNB:
                COLUMNB="not defined"
            Issues[KEY]["COLUMNB"] = COLUMNB
            
            COLUMNC=(CurrentSheet.cell(row=i, column=C).value)
            if not COLUMNC:
                COLUMNC="not defined"
            Issues[KEY]["COLUMNC"] = COLUMNC
            
            COLUMND=(CurrentSheet.cell(row=i, column=D).value)
            if not COLUMND:
                COLUMND="not defined"
            Issues[KEY]["COLUMND"] = COLUMND
            
            COLUMNE=(CurrentSheet.cell(row=i, column=E).value)
            if not COLUMNE:
                COLUMNE="not defined"
            Issues[KEY]["COLUMNE"] = COLUMNE
           
            
            logging.debug("---------------------------------------------------")
            i=i+1
            
          
    logging.debug (Issues)
    logging.debug (Issues.items()) 
    logging.debug((json.dumps(Issues, indent=4, sort_keys=True)))
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value

    

    ##########################################################################################################################
    # Create main issues
    
    if (AUTH==True):    
        Authenticate(JIRASERVICE,PSWD,USER)
        jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)
    else:
        print ("Simulated execution only")

    #create Jira issues
    i=1
    for key, value in Issues.items() :
        #if (i>20):
        #    print "EXIT DUE i"
        #    sys.exit(1)
        KEYVALUE=(key,value)
        KEY=key
        print ("ORIGINAL ISSUE KEY:{0}\nVALUE:{1}".format(KEY,KEYVALUE))
        
        # TODO: check if issue has been imported in previous import attempts
        #if (ENV=="DEV"):
        #    JQLQuery="project = {0}  and cf[12900]  ~ {1}".format(JIRAPROJECT,key) # check key in Jira
        #    results=jira.search_issues(JQLQuery, maxResults=3000)
        #elif (ENV=="PROD"):
            #print "NOT IMPLEMENTED PROD CODE"
            #sys.exit(1)
        #    JQLQuery="project = {0}  and cf[12900]  ~ {1}".format(JIRAPROJECT,key) # check key in Jira
        #    results=jira.search_issues(JQLQuery, maxResults=3000)   
        #else:
        #    print ("ENV SET WRONG")
        #    sys.exit(1)
               
        #if (len(results) > 0):
        #    print ("Key:{0} exists in Jira already".format(key))
        #    print ("Query result:{0}".format(results))
        #    IMPORT=False
        #else:
        #    print ("Key:{0} is a NEW Key,going to import".format(key))
        #    print ("Query result:{0}".format(results))    
        #    IMPORT=True
             
            
        ####### left old ones as examples of data formatting ######################
        # ISSUETYPENW=str((Issues[key]["ISSUE_TYPENW"]).encode('utf-8'))  # str casting needed      
        # JIRASUMMARY=JIRASUMMARY.replace("\n", " ") # Perl used to have chomp, this was only Python way to do this

        #DECKNW=(Issues[key]["DECKNW"])
        #if (DECKNW is None):
        #    DECKNW=Issues[key]["DECKNW"] #to keep None object??
        #    DECKNW="1" # just set some random default value
        #else: 
        #    DECKNW=str((Issues[key]["DECKNW"]))  # str casting needed
        #print ("DECKNW:{0}".format(DECKNW)) 
               
        #temp settings to test 
        ISSUETYPE="Epic"
        PRIORITY="Low"
        JIRASUMMARY="This is summary text"
        DESCRIPTION="This is description text"
        
        JIRASUMMARY=JIRASUMMARY[:254] ## summary max length is 255
        
        IMPORT=True # temp setting
        #IssueID="SHIP-1826" #temp ID
        if (PROD==True):
            if (IMPORT==True):
                if (DRY=="off"):
                   IssueID=CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,ISSUETYPE,PRIORITY,DESCRIPTION)
                   print ("Created issue:{0}  OK".format(IssueID))
                   print ("-----------------------------------------------------------")
                   time.sleep(0.1) 
                

                
                elif (DRY=="on"):       
                    print ("Dryrun mode: I would have Created issue ")
                    print ("JIRAPROJECT:{0},JIRASUMMARY:{1},ISSUETYPE:{2},PRIORITY:{3},DESCRIPTION:{4})".format(JIRAPROJECT,JIRASUMMARY,ISSUETYPE,PRIORITY,DESCRIPTION)) 
                    print ("-----------------------------------------------------------")


            else:
                print ("Issue exists in Jira. Did nothing")
        else:
           print ("--> SKIPPED ISSUE CREATION") 
      

        
        i=i+1    
    
      
      
   # end = time.clock()
   # totaltime=end-start
   # print "Time taken:{0} seconds".format(totaltime) 
        
#############################################################################


    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 