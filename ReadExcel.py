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
# 1) Jira main issue data
# 2) Jira subtask(s) (remark(s)) data for main issue
# 3) Info of attachment for main issue (to be added using inhouse tooling
#  
#NOTE: Uses hardcoded sheet/column value

def Parse(filepath, filename,JIRASERVICE,JIRAPROJECT,PSWD,USER):
    

    ####################################################################################
    # CONFIGURATIONS ####
    PROD=True # False #True   #false skips issue creation and other jira operations
    ENV="PROD" # or "PROD" or "DEV", sets the custom field IDs 
    AUTH=True # so jira authorizations
    DRY="off" # on==dont do, just tell   off=do everything THIS IS THE ONE FLAG TO RULE THEM ALL
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



    print ("FORCEEXIT")
    sys.exit(5)


    ############################################################################################################################
    # Check any remarks (subtasks) for main issue
    # NOTE: Uses hardcoded sheet/column values
    #
    #removed currently dfue excel changes

    
    print ("Checking all subtasks now")
    print ("Subtasks file:{0}".format(subfilename))

    
    i=DATASTARTSROWSUB # brute force row indexing
    for row in SubCurrentSheet[('B{}:B{}'.format(DATASTARTSROWSUB,SubCurrentSheet.max_row))]:  # go trough all column B (KEY) rows
        for submycell in row:
            PARENTKEY=submycell.value
            #logging.debug("SUBROW:{0} Original PARENT ID:{1}".format(i,PARENTKEY))
            #Issues[KEY]={} # add to dictionary as master key (KEY)
            
            #Just hardocode operations, POC is one off

            if PARENTKEY in Issues:
                logging.debug( "Subtask has a known parent {0}".format(PARENTKEY))
                #REMARKKEY=SubCurrentSheet['J{0}'.format(i)].value  # column J holds Task-ID NW
                REMARKKEY=(SubCurrentSheet.cell(row=i, column=B).value) #parent key value
                SUBORIGINALREMARKEY=REMARKKEY # record old parent key for storing to remark
                REMARKKEY=str(REMARKKEY)+"_"+str(i)  # add _ROWNUBER to create really unique key 
                #print "CREATED REMARKKEY:{0}".format(REMARKKEY)
                #Issues[KEY]["REMARKS"]={}
                Issues[PARENTKEY]["REMARKS"][REMARKKEY] = {}
                
                
                # Just hardcode operattions, POC is one off
                #DECK=SubCurrentSheet['AA{0}'.format(i)].value  # column AA holds DECK
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUBORIGINALREMARKEY"] = SUBORIGINALREMARKEY
                
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUBKEY"] = REMARKKEY
                
                SUBSUMMARY=(SubCurrentSheet.cell(row=i, column=SUB_C).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUMMARY"] = SUBSUMMARY
                
                SUBISSUE_TYPENW=(SubCurrentSheet.cell(row=i, column=SUB_D).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["ISSUE_TYPENW"] = SUBISSUE_TYPENW
                
                
                
                # NOT IN EXCEL    
                #SUBISSUE_TYPE=(SubCurrentSheet.cell(row=i, column=SUB_E).value)
                PARENT_ISSUE_TYPE=Issues[PARENTKEY]["ISSUE_TYPE"] 
                
                if (PARENT_ISSUE_TYPE=="Outfitting Inspection"):
                   SUBISSUE_TYPE = "Outfitting Remark"
                elif (PARENT_ISSUE_TYPE=="Hull Inspection"):
                   SUBISSUE_TYPE = "Hull Remark"
                else:
                    logging.error ("ERROR: No type match for subissue:{0}. Forcing type!!".format(REMARKKEY))
                    SUBISSUE_TYPE = "Hull Remark"
              
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["ISSUE_TYPE"] = SUBISSUE_TYPE
                
                
                
                
                
                
                     
            
                SUBSTATUSNW=(SubCurrentSheet.cell(row=i, column=SUB_F).value)
                if (SUBSTATUSNW is None):
                    SUBSTATUSNW="open"
                elif (SUBSTATUSNW =="done"): #remap changed Jira states (drop down values)                   
                     SUBSTATUSNW="closed"   
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["STATUSNW"] = SUBSTATUSNW
                
                # All Open in excel
                #SUBSTATUS=(SubCurrentSheet.cell(row=i, column=SUB_G).value)
                SUBSTATUS="Open"
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["STATUS"] = SUBSTATUS
                
                
                SUBREPORTERNW=(SubCurrentSheet.cell(row=i, column=SUB_H).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["REPORTERNW"] = SUBREPORTERNW
                
                SUBCREATED=(SubCurrentSheet.cell(row=i, column=SUB_I).value) #Inspection date
                # ISO 8601 conversion to Exceli time
                subtime2=SUBCREATED.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
                #print "CREATED SUBTASK ISOFORMAT TIME2:{0}".format(subtime2)
                SUBCREATED=subtime2
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUBCREATED"] = SUBCREATED
                
                SUBDESCRIPTION=(SubCurrentSheet.cell(row=i, column=SUB_J).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DESCRIPTION"] = SUBDESCRIPTION
                
                SUBSHIPNUMBER=(SubCurrentSheet.cell(row=i, column=SUB_K).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SHIPNUMBER"] = SUBSHIPNUMBER
                
                SUBSYSTEMNUMBERNW=(SubCurrentSheet.cell(row=i, column=SUB_L).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SYSTEMNUMBERNW"] = SUBSYSTEMNUMBERNW
                
                SUBPERFORMER=(SubCurrentSheet.cell(row=i, column=SUB_M).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["PERFORMER"] = SUBPERFORMER
                
                SUBRESPONSIBLENW=(SubCurrentSheet.cell(row=i, column=SUB_N).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["RESPONSIBLENW"] = SUBRESPONSIBLENW
                
                SUBASSIGNEE=(SubCurrentSheet.cell(row=i, column=SUB_O).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["ASSIGNEE"] = SUBASSIGNEE
           
                SUBINSPECTION=(SubCurrentSheet.cell(row=i, column=SUB_R).value)
                #ISO 8601 conversion to Exceli time
                #SUBINSPECTION=SUBINSPECTION.to_datetime(SUBINSPECTION)
                subtime3=SUBINSPECTION.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
                #subtime3=SUBINSPECTION.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
                
                #print "CREATED SUBTASK ISOFORMAT TIME3:{0}".format(subtime3)
                SUBINSPECTION=subtime3
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["SUBINSPECTION"] = SUBINSPECTION
           
           
                SUBDEPARTMENTNW=(SubCurrentSheet.cell(row=i, column=SUB_S).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DEPARTMENTNW"] = SUBDEPARTMENTNW
                
                # NO LOGIC TO SET, SETTING ONE RANDOM VALUE
                #SUBDEPARTMENT=(SubCurrentSheet.cell(row=i, column=SUB_T).value)
                SUBDEPARTMENT="500 - Outfitting Common"
                logging.error ("FORCING: Forcesetting subissue:{0}. Forcing deparmtment!!".format(SUBDEPARTMENT))
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DEPARTMENT"] = SUBDEPARTMENT
                
                
                SUBBLOCKNW=(SubCurrentSheet.cell(row=i, column=SUB_U).value)
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["BLOCKNW"] = SUBBLOCKNW
                
                SUBDECKNW=(SubCurrentSheet.cell(row=i, column=SUB_V).value)
                if (SUBDECKNW is None):
                    SUBDECKNW="1"
                Issues[PARENTKEY]["REMARKS"][REMARKKEY]["DECKNW"] = SUBDECKNW
                #SUBTASKID=REMARKKEY
            
            else:
                    print ("ERROR: Unknown parent found --> originazl key: {0}".format(PARENTKEY))
            logging.debug( "---------------------------------------------------------------------------")
            i=i+1
    
    
 
    print(json.dumps(Issues, indent=4, sort_keys=True))
    
   # print "EXITING NOW ALL DONE"
   # sys.exit(5)

    ##########################################################################################################################
    # Create main issues
    
    if (AUTH==True):    
        Authenticate(JIRASERVICE,PSWD,USER)
        jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)
    else:
        print ("Simulated execution only")

    #create main issues
    i=1
    for key, value in Issues.iteritems() :
        #if (i>20):
        #    print "EXIT DUE i"
        #    sys.exit(1)
        KEYVALUE=(key,value)
        KEY=key
        print ("ORIGINAL ISSUE KEY:{0}\nVALUE:{1}".format(KEY,KEYVALUE))
        
        # check if issue has been imported in previous import attempts
        if (ENV=="DEV"):
            JQLQuery="project = {0}  and cf[12900]  ~ {1}".format(JIRAPROJECT,key) # check key in Jira
            results=jira.search_issues(JQLQuery, maxResults=3000)
        elif (ENV=="PROD"):
            #print "NOT IMPLEMENTED PROD CODE"
            #sys.exit(1)
            JQLQuery="project = {0}  and cf[12900]  ~ {1}".format(JIRAPROJECT,key) # check key in Jira
            results=jira.search_issues(JQLQuery, maxResults=3000)
            
        else:
            print ("ENV SET WRONG")
            sys.exit(1)
               
        if (len(results) > 0):
            print ("Key:{0} exists in Jira already".format(key))
            print ("Query result:{0}".format(results))
            IMPORT=False
        else:
            print ("Key:{0} is a NEW Key,going to import".format(key))
            print ("Query result:{0}".format(results))    
            IMPORT=True
             
            
        
        REMARKS=Issues[key]["REMARKS"]
        print ("REMARKS:{0}".format(REMARKS))
        
        ISSUETYPE=((Issues[key]["ISSUE_TYPE"]).encode('utf-8'))
        # excel is full of typos, fix them here
        if (ISSUETYPE.lower()=="Outfitting Inspection".lower()):
                ISSUETYPE="Outfitting Inspection"
        elif (ISSUETYPE.lower()=="Hull Inspection".lower()):
                ISSUETYPE="Hull Inspection"
        else:
            print("Totally lost main task issuetype casting. HELP!")
        print ("JIRA ISSUE_TYPE:{0}".format(ISSUETYPE)) 
        
        ISSUETYPENW=(Issues[key]["ISSUE_TYPENW"])
        if (ISSUETYPENW is None):
             ISSUETYPENW=(Issues[key]["ISSUE_TYPENW"]) #to keep None object??
        elif (ISSUETYPENW=="HVAC"):
             ISSUETYPENW="Steel"
        elif (ISSUETYPENW=="LNG"):
              ISSUETYPENW="Steel"
        elif (str(ISSUETYPENW)=="Preservation"):
             ISSUTYPENW="Steel"
                
        else: 
            ISSUETYPENW=str((Issues[key]["ISSUE_TYPENW"]).encode('utf-8'))  # str casting needed
       
            
            
        print ("ORIGINAL ISSUE_TYPE:{0}".format(ISSUETYPENW))  
        
        STATUS=Issues[key]["STATUS"]  
        print ("JIRA STATUS:{0}".format(STATUS))  
        STATUSNW=Issues[key]["STATUSNW"]
        print ("ORIGINAL STATUS:{0}".format(STATUSNW))  
        PRIORITY=Issues[key]["PRIORITY"]
        print ("JIRA PRIORITY:{0}".format(PRIORITY))  
        RESPONSIBLENW=str(((Issues[key]["RESPONSIBLENW"]).encode('utf8')))  
        print ("ORIGINAL RESPONSIBLE:{0}".format(RESPONSIBLENW))    
        RESPONSIBLE=(Issues[key]["RESPONSIBLE"])
        print ("JIRA RESPONSIBLE:{0}".format(RESPONSIBLE))    
        INSPECTEDTIME= Issues[key]["INSPECTED"]
        print ("ORIGINAL CREATED TIME:{0}".format(INSPECTEDTIME))
        SHIP=Issues[key]["SHIPNUMBER"]       
        print ("SHIP NUMBER:{0}".format(SHIP))  

        SYSTEM= Issues[key]["SYSTEM"]
        if (SYSTEM is None):
            SYSTEM=Issues[key]["SYSTEM"] #to keep None object??
        else: 
            SYSTEM=str((Issues[key]["SYSTEM"]))  # str casting needed
        print ("SYSTEM:{0}".format(SYSTEM)) 
        
        SYSTEMNUMBERNW= Issues[key]["SYSTEMNUMBERNW"]
        if (SYSTEMNUMBERNW is None):
            SYSTEMNUMBERNW=Issues[key]["SYSTEMNUMBERNW"] #to keep None object??
        else: 
            SYSTEMNUMBERNW=str((Issues[key]["SYSTEMNUMBERNW"]))  # str casting needed
        print ("SYSTEMNUMBERNW:{0}".format(SYSTEMNUMBERNW)) 
        
        PERFORMERNW=(Issues[key]["PERFORMERNW"]).encode('utf8')
        print ("ORIGINAL PERFOMER:{0}".format(PERFORMERNW))   
        DEPARTMENTNW=(Issues[key]["DEPARTMENTNW"])
        print ("ORIGINAL DEPARTMENT:{0}".format(DEPARTMENTNW)) 
        DEPARTMENT=(Issues[key]["DEPARTMENT"])
        print ("DEPARTMENT:{0}".format(DEPARTMENT)) 
        DESCRIPTION=(Issues[key]["DESCRIPTION"])
        print ("DESCPTION + TOPOLOGY:{0}".format(DESCRIPTION)) 

        JIRASUMMARY=(Issues[key]["SUMMARY"]).encode('utf-8')          
        JIRASUMMARY=JIRASUMMARY.replace("\n", " ") # Perl used to have chomp, this was only Python way to do this
        JIRASUMMARY=JIRASUMMARY[:254] ## summary max length is 255
        print ("SUMMARY:{0}".format(JIRASUMMARY))
       
        AREA=(Issues[key]["AREA"])
        print ("AREA:{0}".format(AREA)) 
        
        SURVEYOR=(Issues[key]["SURVEYOR"])
        print ("SURVEYOR:{0}".format(SURVEYOR)) 
        
        DECKNW=(Issues[key]["DECKNW"])
        if (DECKNW is None):
            DECKNW=Issues[key]["DECKNW"] #to keep None object??
            DECKNW="1" # just set some random default value
        else: 
            DECKNW=str((Issues[key]["DECKNW"]))  # str casting needed
        print ("DECKNW:{0}".format(DECKNW)) 
        
        BLOCKNW=Issues[key]["BLOCKNW"]
        if (BLOCKNW is None):
            BLOCKNW=Issues[key]["BLOCKNW"] #to keep None object??
        else: 
            BLOCKNW=str((Issues[key]["BLOCKNW"]))  # str casting needed
        print ("BLOCKNW:{0}".format(BLOCKNW)) 
        
        FIREZONENW=Issues[key]["FIREZONENW"]
        if (FIREZONENW is None):
            FIREZONENW=Issues[key]["FIREZONENW"] #to keep None object??
        else: 
            FIREZONENW=str((Issues[key]["FIREZONENW"]))  # str casting needed
        print ("FIREZONENW:{0}".format(FIREZONENW)) 
        
        
        #IssueID="SHIP-1826" #temp ID
        if (PROD==True):
            if (IMPORT==True):
                if (DRY=="off"):
                   IssueID=CreateIssue(ENV,jira,JIRAPROJECT,JIRASUMMARY,KEY,ISSUETYPE,ISSUETYPENW,STATUS,STATUSNW,PRIORITY,RESPONSIBLENW,RESPONSIBLE,INSPECTEDTIME,SHIP,SYSTEMNUMBERNW,SYSTEM,PERFORMERNW,DEPARTMENTNW,DEPARTMENT,DESCRIPTION,AREA,SURVEYOR,DECKNW,BLOCKNW,FIREZONENW)
                   print ("Created issue:{0}  OK".format(IssueID))
                   print ("-----------------------------------------------------------")
                   time.sleep(0.1) 
                
                
                   if (ATTACHMENTS==True):
                       DRY="off"
                       HandleAttachemnts(filepath,key,ATTACHDIR,IssueID,jira,DRY)
                
                elif (DRY=="on"):       
                    print ("Dryrun mode: I would have Created issue ")
                    IssueID="NOTREAL-007"
                    DRY="on"
                    HandleAttachemnts(filepath,key,ATTACHDIR,IssueID,jira,DRY)  # dangezone!!
                    print ("-----------------------------------------------------------")
                else:
                    print ("CONFUSED: Skipped attachments operation, check internal configs")
        
            #sys.exit(1) 
            #print "IssueKey:{0}".format(IssueID.key)
            else:
                print ("Issue exists in Jira. Did nothing")
        else:
           print ("--> SKIPPED ISSUE CREATION") 
        
        #filesx=filepath+"/*{0}*".format(key)
        #print "filesx:{0}".format(filesx)
        
        
        
        
        Remarks=Issues[key]["REMARKS"] # take a copy of remarks and use it
        
        #print "-----------------------------------------------------------------------------------------------------------------"
        if (PROD==True and IMPORT==True):
            PARENT=IssueID
        #create subtask(s) under one parent
        # custom ids in comments: 1) dev 2) production
        for subkey , subvalue in Remarks.iteritems():
            
            SUBKEYVALUE=(subkey,subvalue)
            SUBKEY=subkey.encode('utf-8')
            
            ParentCheck = re.search( r"(\d*)(_)(\d*)", SUBKEY) # remove unique _ROWNUJMBER identifier
            if ParentCheck:
                CurrentGroups=ParentCheck.groups()    
                #print ("Group 1: %s" % CurrentGroups[0]) 
                #print ("Group 2: %s" % CurrentGroups[1]) 
                SUBPARENTKEY=CurrentGroups[0] #logical key (parent original key, used to tell teh parent for this subtask), dictionary key is the subkey 
            else:
                log.error("Subtask Parent parsing failure")
            print ("SUBTASK PARENT'S ORIGINAL KEY:{0}\nVALUE:{1}".format(SUBPARENTKEY,SUBKEYVALUE))
       
            
            SUBKEY=Remarks[subkey]["SUBKEY"] 
            SUBORIGINALREMARKEY=Remarks[subkey]["SUBORIGINALREMARKEY"] 
            SUBSUMMARY=Remarks[subkey]["SUMMARY"] 
            SUBSUMMARY=SUBSUMMARY.replace("\n", "")
            SUBSUMMARY=SUBSUMMARY[:254]    ## summary max length is 255
            SUBSUMMARY=(SUBSUMMARY.encode('utf-8')) 
            print ("SUBSUMMARY:{0}".format(SUBSUMMARY))
            
        
            
            SUBISSUTYPENW=Remarks[subkey]["ISSUE_TYPENW"] 
            if (SUBISSUTYPENW=="Preservation"):
                SUBISSUTYPENW="Steel"
                print ("Forcing Subissutype Preservation as Steel")
            elif (SUBISSUTYPENW=="HVAC"):
                SUBISSUTYPENW="Steel"
                print ("Forcing Subissutype HVAC as Steel")
            elif (SUBISSUTYPENW=="Pipes"):
                SUBISSUTYPENW="Steel"
                print ("Forcing Subissutype Pipes as Steel")         
            print ("SUBISSUTYPENW:{0}".format(SUBISSUTYPENW))
           
            
            SUBISSUTYPE=Remarks[subkey]["ISSUE_TYPE"] 
           
            
            SUBSTATUSNW=Remarks[subkey]["STATUSNW"] 
           
            
            SUBSTATUS=Remarks[subkey]["STATUS"] 
           
            
            SUBREPORTERNW=Remarks[subkey]["REPORTERNW"].encode('utf-8') 
            
            
            SUBCREATED=Remarks[subkey]["SUBCREATED"] 
          
            
        
        
            
            if (DESCRIPTION is None):
                SUBDESCRIPTION=Remarks[subkey]["DESCRIPTION"] 
            else:
                SUBDESCRIPTION=Remarks[subkey]["DESCRIPTION"] #.string  #.encode('utf-8')  
                SUBDESCRIPTION=unicode(SUBDESCRIPTION).encode('utf-8') 
                
            
            SUBSHIPNUMBER=Remarks[subkey]["SHIPNUMBER"] 
            
            
            SUBSYSTEMNUMBERNW=Remarks[subkey]["SYSTEMNUMBERNW"] 
        
            
            SUBPERFORMER=Remarks[subkey]["PERFORMER"].encode('utf-8') 
           
            
            SUBRESPONSIBLENW=Remarks[subkey]["RESPONSIBLENW"].encode('utf-8') 
       
            
            SUBASSIGNEE=Remarks[subkey]["ASSIGNEE"] 
            
            
            SUBINSPECTION=Remarks[subkey]["SUBINSPECTION"] 

            
            SUBDEPARTMENTNW=Remarks[subkey]["DEPARTMENTNW"] 
           
            
            SUBDEPARTMENT=Remarks[subkey]["DEPARTMENT"] 
           
            
            SUBBLOCKNW=Remarks[subkey]["BLOCKNW"] 
           
            
            SUBDECKNW=Remarks[subkey]["DECKNW"] 
        
            

            if (PROD==True):
                if (IMPORT==True):
                   if (DRY=="off"):
                   #if (IMPORT==True):
                       SubIssueID=CreateSubTask(ENV,jira,JIRAPROJECT,PARENT,SUBORIGINALREMARKEY,SUBSUMMARY,SUBISSUTYPENW,SUBISSUTYPE,SUBSTATUSNW,SUBSTATUS,SUBREPORTERNW,SUBCREATED,SUBDESCRIPTION,SUBSHIPNUMBER,SUBSYSTEMNUMBERNW,SUBPERFORMER,SUBRESPONSIBLENW,SUBASSIGNEE,SUBINSPECTION,SUBDEPARTMENTNW,SUBDEPARTMENT,SUBBLOCKNW,SUBDECKNW)
                       
                       time.sleep(0.1)
                   #print "SKIPPED SUBTASK OPERATIONS. SHOULD HAVE CREATED"                  
                   elif (DRY=="on"):
                       print ("DRYRUN mode: I would have created subtask")
                   else:
                       print ("Confused: Is this DRY run or not???")    
                else:   
                   print ("9Issue exists in Jira. Did no subtask operations")
            else:
                print ("Skipped subtask creation")
        
        i=i+1    
    
      
      
    end = time.clock()
    totaltime=end-start
   # print "Time taken:{0} seconds".format(totaltime) 
        
#############################################################################

def HandleAttachemnts(filepath,key,ATTACHDIR,IssueID,jira,DRY):
        
        print ("*****************************************")

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 