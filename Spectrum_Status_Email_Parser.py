import win32com.client;
from datetime import datetime, timedelta;
import pandas as pd
import re;
import AdvancedHTMLParser;
import collections;
import html2text;

#==================================================================================================#
#   Building Outlook App Objects. 
#        outApp - Outlook Application Object
#        allowedFolders - List of Folders that the email parser will loop through
#        items_pTag_RexEx - Key Items found in Emails that are to be parsed out or used as landmarks in email
#        subjectTermsRexEx - RegEx Expression of Subjects for which emails will be gathered
#        yDate - Stores yesterday's date and will be used to Determine start and end dates of Default Fiscal
#        sDate_default_date - stores the date format of the starting period
#        sDate_default_txt - stores the string format of the sDate_default_date
#        sDate - StartDate for which emails will be gathered. Items before this will be ingored
#        eDate - EndDate for which emails will be gathered. Items after this will be ingored
#        eDate_default_date store the date format of the ending period
#        eDate_default_txt - store the string format fo the eDte_default_date
#        accounts - Initial List of DataFiles Associated with the PC
#        neededEmails - Starts off as an Empty List for which emails that meet requirements are added too
#==================================================================================================#

outApp = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
allowedFolders = ["Inbox"]
items_pTag_RexEx = ".*(Yellow|Red) Status: "
subjectTermsRexEx = ".*Spectrum.*Video.*(Red|Yellow).*Status"
headerSectionTags = ["b","o:p"]
prefHeaderSectionTags = ["b"]
yDate = datetime.today() + timedelta(days=-1)
#Setting yDate to End of Current Fiscal
if yDate.day < 29:
    yDate = yDate + timedelta(days=(28 - yDate.day))
else:
    yDate = yDate + pd.DateOffset(months=1) + timedelta(days=(-1*(yDate.day - 28)))

# Setting the Default Fiscal to the Last Completed Fiscal #
# Default Start and End Dates can be used by just Hitting Enter at the Prompt #
sDate_default_date = (yDate + pd.DateOffset(months=-2) + timedelta(days=1)).date()
sDate_default_txt = sDate_default_date.strftime("%m/%d/%Y")
sDate = input("Input the Start Date of the Search (MM/dd/yyyyy) or Press Enter for Default ("+sDate_default_txt+") :")
#If no input was entered, the default will be assinged otherwise the input will be parsed
if len(sDate) > 0:
    sDate = datetime.strptime(sDate,"%m/%d/%Y").date()
else:
    sDate = datetime.strptime(sDate_default_txt,"%m/%d/%Y").date()

eDate_default_date = (yDate + pd.DateOffset(months=-1)).date()
eDate_default_text = eDate_default_date.strftime("%m/%d/%Y")
# Accepting Input from User on Start and End Date Ranges #
eDate = input("Input the End Date of the Search (MM/dd/yyyyy) or Press Enter for Default ("+eDate_default_text+") :")
#If no input was entered, the default will be assinged otherwise the input will be parsed
if len(eDate) > 0 :
    eDate = datetime.strptime(eDate,"%m/%d/%Y").date()
else:
    eDate = datetime.strptime(eDate_default_text,"%m/%d/%Y").date()

accounts = outApp.Folders
neededEmails = []
cleanEmails = []

# Custom HTML Parser to Look through Spectrum Update Automated Emails and only Extract Data Needed #
def parseHTML(email_item):
    html_body = email_item.HTMLbody
    parser = AdvancedHTMLParser.AdvancedHTMLParser()
    parser.parseStr(html_body)

    wordSection1 = parser.getElementsByClassName("WordSection1")

    header = ""
    pref_header = ""
    pref_body = ""
    action_header = ""
    action_body = ""

    #Status Verbiage Node is a P Tag thats has one b tag, and one o:p tag
    pTags = wordSection1.getElementsByTagName(tagName="p")
    # Grabbing Header Data from DOM #
    for ele in pTags:
        #Need to test for 2 elements: one b tag, one o:p tag
        cNodeCount = len(ele.childNodes)
        # Build up List of Child Tags in the Element #
        cNodeTagNames = []
        for Cele in ele.childNodes:
            cNodeTagNames.append(Cele.tagName)
        # If the Element has 2 ChildNodes and both are that of what is in the headerSectionTags list #
        if cNodeCount == 2 and collections.Counter(cNodeTagNames) == collections.Counter(headerSectionTags):
            if(re.search("(Yellow|Red).*Status:.*",ele.firstChild.innerText)):
                header = ele.firstChild.innerText

        # If the Element has 1 ChildNodes and both are that of what is in the prefHeaderSectionTags list #
        if cNodeCount == 1 and collections.Counter(cNodeTagNames) == collections.Counter(prefHeaderSectionTags):
            # Performance Header and Action Header are nested with in a span tag in a b tag #
            firstEle = ele.firstChild
            typeHeader = ""
            #Check if it is a Performance or Action Header
            eleInnerText = firstEle.innerText
            if(re.search("(Performance|Action).*(Summary|Recomm).*",eleInnerText)):
                if(re.search("Performance.*",eleInnerText)):
                    pref_header = eleInnerText
                else:
                    action_header = eleInnerText

                #Looping over the next Siblings to find the first UL List after the Header 
                ul_exists = False
                ul_exists_breakInt = 0
                nextSiblingEle = ele.nextElementSibling
                while not ul_exists and ul_exists_breakInt <= 10:

                    if nextSiblingEle.tagName == "ul":
                        ul_exists = True
                        subParser = AdvancedHTMLParser.AdvancedHTMLParser()

                        if(re.search("Performance.*",eleInnerText)):
                            subParser.parseStr(nextSiblingEle.innerHTML)
                            pref_body = html2text.html2text(subParser.getFormattedHTML())
                        else:
                            subParser.parseStr(nextSiblingEle.innerHTML)
                            action_body = html2text.html2text(subParser.getFormattedHTML())
                    else:
                        ul_exists = False

                    nextSiblingEle = nextSiblingEle.nextElementSibling
                    ul_exists_breakInt = ul_exists_breakInt + 1
    if len(header) == 0:
        header = email_item.Subject       

    recv_time = email_item.ReceivedTime.strftime("%m/%d/%Y %H:%M:%S") 
    #Append the Cleaned Up Email to the cleanEmail List
    cleanEmails.append([header,pref_header,pref_body,action_header,action_body,recv_time])

#Looping through Each Account
for accts in accounts:
    print("Looking in Account: " + accts.Name)
    #Looping through Each Mail Folder in the Account
    for accts_folder in accts.Folders:
        mail_folderName = accts_folder.Name
        #Only Looping through folders in the allowedFolders List
        #For each Mail Folder
        if mail_folderName in allowedFolders:

            #Looping Over Emails that Exist Directly in the Main Folder
            mail_folderItems = accts_folder.Items
            for mail_item in mail_folderItems:
                #If its an Email, it'll have a RecievedTime
                if(hasattr(mail_item,"ReceivedTime")):
                    if(sDate < mail_item.ReceivedTime.date() < eDate):
                        if(re.search(subjectTermsRexEx,mail_item.Subject)):
                           parseHTML(mail_item)

            # Now Going 1 SubFolder Deep from the Inital Allowed List #
            mail_folderSubFolders = accts_folder.Folders
            for sub_folder in mail_folderSubFolders:
                #Looping Over Emails that Exist in the Sub Folder
                subFolder_items = sub_folder.Items
                for subFolder_item in subFolder_items:
                    #Check to see if its an acutal RecievedEmail by checking if it has  ReceivedTime Attribute
                    if(hasattr(subFolder_item,"ReceivedTime")):
                        #Check to see if the Email was Recieved During the Period Selected
                        if(sDate < subFolder_item.ReceivedTime.date() < eDate):
                            #Check if it contains any of the Subjects we are Looking For
                            if(re.search(subjectTermsRexEx,subFolder_item.Subject)):
                                parseHTML(subFolder_item)

                # Now Going 1 More SubFolder Deep From the Inital SubFolder #
                subFolder_folders = sub_folder.Folders
                for sub_folder2 in subFolder_folders:
                    sub_folder2_items = sub_folder2.Items
                    for sub_folder2_item in sub_folder2_items:
                        #Check to see if its an acutal RecievedEmail by checking if it has  ReceivedTime Attribute
                        if(hasattr(sub_folder2_item,"ReceivedTime")):
                            #Check to see if the Email was Recieved During the Period Selected
                            if(sDate < sub_folder2_item.ReceivedTime.date() < eDate):
                                #Check if it contains any of the Subjects we are Looking For
                                if(re.search(subjectTermsRexEx,sub_folder2_item.Subject)):
                                    parseHTML(sub_folder2_item)

# Build The Text File to house the emails #
txtFile = open('EmailSummaryTxt.txt','w', encoding="utf-8")
txtFile.close()

for email in cleanEmails:

    if len(email) > 0:
        #Grabbing Email Context Info
        e_heading = email[0]
        e_pHeading = email[1]
        e_pBody = email[2]
        e_aHeading = email[3]
        e_aBody = email[4]
        e_time = email[5]

    
    e_pBody = re.sub(' +',' ',e_pBody)
    e_aBody = re.sub(' +',' ',e_aBody)
    e_pBody = e_pBody.replace('\n','')
    e_pBody = e_pBody.replace('\t','')
    e_aBody = e_aBody.replace('\n','')
    e_aBody = e_aBody.replace('\t','')
    e_pBody = e_pBody.replace('*','\n*')
    e_aBody = e_aBody.replace('*','\n*')
    e_aBody = re.sub('[*]\s\n[*]','',e_aBody)
    e_pBody = re.sub('[*]\s\n[*]','',e_pBody)
    e_aBody = re.sub('[*]\n[*]','',e_aBody)
    e_pBody = re.sub('[*]\n[*]','',e_pBody)
    e_aBody = re.sub('\n\n','',e_aBody)
    e_pBody = re.sub('\n\n','',e_pBody)

    txtfile = open('EmailSummaryTxt.txt','a')
    txtfile.write(e_time.strip() + " - " + e_heading.strip()+"\n")
    txtfile.write(e_pHeading.strip()+"\n")
    txtfile.write(e_pBody.strip()+"\n")
    txtfile.write(e_aHeading.strip()+"\n")
    txtfile.write(e_aBody.strip()+"\n\n\n")
    txtfile.close()
