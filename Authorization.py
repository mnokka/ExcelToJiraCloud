# Jira auhentications
#
# 3.2.2022  mika.nokka1@gmail.com
# 
# To be used via importing only
# 


import argparse
import sys
import requests, os
from requests.auth import HTTPBasicAuth
# We don't want InsecureRequest warnings:
import requests
requests.packages.urllib3.disable_warnings()
import itertools, re, sys
from jira import JIRA
import random


__version__ = "0.5"
thisFile = __file__


####################################################################################################   

def Authenticate(JIRASERVICE,PSWD,USER):
    host=JIRASERVICE
    user=USER
    PASSWORD=PSWD
        
    f = requests.get(host,auth=(user, PASSWORD))
       
    status=f.status_code
    print ("STATUS CODE: %s" % status)
    
    header=str(f.headers)
    if (status == 200): #401 is failure code
         print ("Authentication OK \nHEADER: {0}".format(header))
         
    else:
        print ("--> ERROR: Apparantly user authentication gone wrong")   
        print ("HEADER: {0}".format(header))
        print ("EXITING!!!!")
        
        sys.exit(1)  
    print ("---------------------------------------------------------")
    return user,PASSWORD

###################################################################################    
def DoJIRAStuff(user,PASSWORD,JIRASERVICE):
 jira_server=JIRASERVICE
 try:
     print("Connecting to JIRA: {0}".format(jira_server))
     jira_options = {'server': jira_server}
     jira = JIRA(options=jira_options,basic_auth=(user,PASSWORD))
     print ("JIRA Authorization OK")
 except Exception as e:
    print("Failed to connect to JIRA: %s" % e)
 return jira   
    