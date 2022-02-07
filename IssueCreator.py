# Create Issue to given JIRA
#
# 5.2.2022 mika.nokka1@gmail.com 
# 
# Use vie importing only

import datetime 
import time
import argparse
import sys
import netrc
import requests, os
from requests.auth import HTTPBasicAuth
# We don't want InsecureRequest warnings:
import requests
requests.packages.urllib3.disable_warnings()
import itertools, re, sys
from jira import JIRA, JIRAError
import random

from Authorization import Authenticate  # no need to use as external command
from Authorization import DoJIRAStuff

__version__ = "0.5"
thisFile = __file__

    

####################################################################################
# https://zzzzz.atlassian.net/rest/api/2/field to fiend Epic name custom field ID
# in used example Jira, it was customfield_10004

def CreateIssue(jira,JIRAPROJECT,JIRASUMMARY,ISSUETYPE,PRIORITY,DESCRIPTION,Reference,Domain,ExploitStatus,Base,Impact,Exploitability,Vector,Authencity,Integrity,NonRepudiability,Confidentiality,Availability,Authorization):

    jiraobj=jira
    project=JIRAPROJECT

    
   
    print ("Creating issue for JIRA project: {0}".format(project))
    
    
    #creat issue dict created based on example excel sheet, cunstom fields created to jira and IDs added

    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': JIRASUMMARY,
    'customfield_10011':JIRASUMMARY,   # only needed for Epic issuetype, "Epic Name"

    'customfield_10068':Reference,  
    'customfield_10069':Vector,
    
    'customfield_10072':Base,
    'customfield_10073':Impact,
    'customfield_10074':Exploitability,
    
    'description': DESCRIPTION,
    'issuetype': {'name': ISSUETYPE},
    'priority': {'name': PRIORITY},
    #'customfield_14613' if (ENV =="DEV") else 'customfield_14615' : str(SYSTEM),
    }

    print ("Using issue dict:{0}".format(issue_dict))

    try:
        new_issue = jiraobj.create_issue(fields=issue_dict)
        print ("Issue created OK")
        print ("Updating now all selection custom fields")        
           # if (AREA is None):
           #     new_issue.update(notify=False,fields={"customfield_10007":[ {"id": "-1"}]})  # multiple selection, see https://developer.atlassian.com/server/jira/platform/jira-rest-api-examples/
           # else:
           #     new_issue.update(notify=False,fields={"customfield_10007": [{"value": AREA}]}) 
            
        CustomFieldSetter(new_issue,"customfield_10062" ,ExploitStatus)
        CustomFieldSetter(new_issue,"customfield_10063" ,Authencity) 
        CustomFieldSetter(new_issue,"customfield_10064" ,Integrity) 
        CustomFieldSetter(new_issue,"customfield_10065" ,Confidentiality) 
        CustomFieldSetter(new_issue,"customfield_10066" ,Authorization) 
        CustomFieldSetter(new_issue,"customfield_10070" ,NonRepudiability) 
        CustomFieldSetter(new_issue,"customfield_10071" ,Availability) 

        print ("Selection custom fields update ok")          
               
    except JIRAError as e: 
        print("Failed to create/use JIRA object, error: %s" % e)
        #print "Issue was:{0}".format(new_issue)
        sys.exit(1)
    return new_issue 

##################################################################################
# used only selection custom fields, create first and update value then

def CustomFieldSetter(new_issue,CUSTOMFIELDNAME,CUSTOMFIELDVALUE):
    
    try:
    
        print ("Trying update issue:{0}, field:{1}, value:{2}".format(new_issue,CUSTOMFIELDNAME,CUSTOMFIELDVALUE))
        if (CUSTOMFIELDVALUE is None or (not CUSTOMFIELDVALUE)): # None or "nothing" cases
            new_issue.update(notify=False,fields={CUSTOMFIELDNAME: {"id": "-1"}})
            print ("Customfieldsetter: setting -1")
            print ("----------------------------------------------------------------")
        else:    
            new_issue.update(notify=False,fields={CUSTOMFIELDNAME: {'value': CUSTOMFIELDVALUE}})            
        print ("Issue:{0} field:{1} updated ok (value:{2})".format(new_issue,CUSTOMFIELDNAME,CUSTOMFIELDVALUE))  
        print ("----------------------------------------------------------------")  

    except JIRAError as e: 
        print("Failed to UPDATE JIRA object, error: %s" % e)
        print ("Issue was:{0}".format(new_issue))
        sys.exit(1)


    return new_issue 

########################################################################################
# test creating issue with multiple selection list custom field
def CreateSimpleIssue(jira,JIRAPROJECT,JIRASUMMARY,JIRADESCRIPTION):
    #jiraobj=jira
    project=JIRAPROJECT
    
    
    #lottery = random.randint(1,3)
    #if (lottery==1):
    #    TASKTYPE="Steal"
    #elif (lottery>1):
    #    TASKTYPE="Outfitting"
    #else:
    #    TASKTYPE="Task"
    
    #TASKTYPE="Hull Inspection NW"
    TASKTYPE="Task"
    print ("Creating issue for JIRA project: {0}".format(project))
    
    issue_dict = {
    'project': {'key': JIRAPROJECT},
    'summary': str(JIRASUMMARY),
    'description': str(JIRADESCRIPTION),
    'issuetype': {'name': TASKTYPE},
    #'customfield_14600' : [{'value': str("cat")},{'value': str("bear")}] ,
    }

    try:
        new_issue = jira.create_issue(fields=issue_dict)
        print ("Issue created OK")
    except JIRAError as e: 
        print("Failed to create JIRA object, error: %s" % e)
        sys.exit(1)
    return new_issue 



        
if __name__ == "__main__":
        main(sys.argv[1:])
        
        
        
        
        