# POC to create Jira Cloud issue
#
# 7.9.2020,3.2.2022 mika.nokka1@gmail.com 
# 
# python3

import sys, logging
import argparse
from collections import defaultdict
from Authorization import Authenticate  # no need to use as external command
from Authorization import DoJIRAStuff

import glob
import re
import os
import datetime
import unidecode
from jira import JIRA, JIRAError
from collections import defaultdict
import math

start = datetime.datetime.now()
__version__ = u"0.5"

# should pass via  parameters
#ENV="demo"
ENV=u"PROD"

logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    JIRASERVICE=u""
    JIRAPROJECT=u""
    PSWD=u''
    USER=u''
    SUMMARY=u''
    DESCRIPTION=u''
  


    parser = argparse.ArgumentParser(description="Get given Jira instance and issue data",
    
    
    epilog="""
    
    EXAMPLE:
    
    CreateIssue.py  -u MYUSERNAME -w MYPASSWORD(TOKEN) -s JIRASERVICE -y ISSUESUMMARY -d ISSUEDECRIPTION -k PROJECTKEY"""  
    )

    parser.add_argument('-v', help='Show version&author and exit', action='version',version="Version:{0}   mika.nokka1@gmail.com ,  MIT licenced ".format(__version__) )
    
    parser.add_argument("-w",help='<JIRA Cloud token>',metavar="password")
    parser.add_argument('-u', help='<JIRA user account>',metavar="user")
    parser.add_argument('-s', help='<JIRA service>',metavar="server_address")
    parser.add_argument('-y', help='<JIRA issue summary>',metavar="IssueSummary")
    parser.add_argument('-r', help='<DryRun - do nothing but emulate. Off by default>',metavar="on|off",default="off")
    parser.add_argument('-d', help='<JIRA issue desciption>',metavar="IssueDescription")
    parser.add_argument('-k', help='<JIRA project key>',metavar="project_key")

    args = parser.parse_args()
       
    JIRASERVICE = args.s or ''
    PSWD= args.w or ''
    USER= args.u or ''
    SUMMARY=args.y or ''
    DESCRIPTION=args.d or ''
    PROJECT=args.k or ''
    if (args.r=="on"):
        SKIP=1
    else:
        SKIP=0    

    
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or  PSWD=='' or USER=='' or  SUMMARY=='' or  DESCRIPTION=='' or PROJECT=='' ):
        logging.error("\n---> MISSING ARGUMENTS!!\n ")
        parser.print_help()
        sys.exit(2)
        
     
    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)
    
    Parse(JIRASERVICE,PSWD,USER,ENV,jira,SKIP,SUMMARY,DESCRIPTION,PROJECT)



############################################################################################################################################
# Parse args and create Jira Cloud issue. Using fixed task issuetype
#
def Parse(JIRASERVICE,PSWD,USER,ENV,jira,SKIP,SUMMARY,DESCRIPTION,PROJECT):


    try:    

            newissue=jira.create_issue(fields={
            'project': {'key': PROJECT},
            'issuetype': {
                "name": "Task"
            },
            'summary': SUMMARY,
            'description': DESCRIPTION,
            })
    
    except JIRAError as e: 
            logging.error(" ********** JIRA ERROR DETECTED: ***********")
            #logging.error("Tried create issue:{0}".format(newissue))
            
            logging.error(" ********** Statuscode:{0}    Statustext:{1} ************".format(e.status_code,e.text))
            if (e.status_code==400):
                logging.error("400 error dedected") 
    else:
        logging.info("All OK")
        logging.info("Issue created:{0}".format(newissue))
  
        
    end = datetime.datetime.now()
    totaltime=end-start
    seconds=totaltime.total_seconds()
    print ("Time taken:{0} seconds".format(totaltime))
       
            
    print ("*************************************************************************")
    
logging.debug ("--Python exiting--")


###########################################################################################################################


if __name__ == "__main__":
    main(sys.argv[1:]) 