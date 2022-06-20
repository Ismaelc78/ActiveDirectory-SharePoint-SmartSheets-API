## PowerShell Script 
De-Identified method used to compare data from different sites for training and access requirments from users.( SharePoint Lists, SmartSheet data, Active Directory Groups, DocuSign)

Pulls required data through REst APIs/PowerShell :  
   Active Directory users by group 
   Google Admin Active users 
   Pull CGCP requests for AD/GCP accounts from SharePoint
   Qurey Staff List from SmartSheets 

Format and Merge Files: 
  Sort data in tabulated format
  Rename columns for convenience 
  Format columns and remove redundancies 
  Merges tables 
  Query for necessary data in Audits

Saving/Upload: 
  Creates Excel Workbooks: GCP Audit, AD Audit, Training Audit
  Uploads completed audit files to desired SharePoint folder 


## Process / Organization
1. SharePoint:
      a. Prompts user for SharePoint credentials
      b. Connect/Authenticate with SharePointOnline Admin access via REst API
      c. Query SharePoint Online for desired SharePoint List data.
      d. Saves queried data as .csv files in SharePoint folder (TrainingLog.csv, AccessRequests.csv)
      e. Notifes User when complete.
       
       
2. AuditScript.ps1 :

     a. Calls Staff.ps1
        1. Prompts user for SmartSheet credentials
        2. Connect/Authenticate SmartSheets access via REst API
        3. Pulls Data from a worksheet in SmartSheets as StaffList
        4. Imports GCP TrainingLog
        5. Imports DocuSign spreadsheet
        6. Groups/Sorts items in files to only keep most recent training and certs completed.
        7. Compares GCP TrainingLog vs Staff List
        8. Checks and marks Expiration Status (< 30 Days,   < 5 Days,   Expired)
        9. Saves data as tables in Excel Workbooks (Staff.xlsx , TrainingAudit(date).xlsx)
      
     b. Calls AD_script.ps1
        1. Connects to Active Directory
        2. Exports per-user data based on a pre-defined list of AD Groups
        3. Looks up GUIDs in AD. 
        4. Gets List of all member GUIDs- Outputs UserGUID and GroupNames.
        5. Creates list of Users and AD Groups they belong to.
        6. Get AD info for eachunique  member
        7. Save data as "AD_Audit.xlsx"
      
     c. Imports Staff List 
     d. Imports GCP Access Requests List
     e. Imports Google Admin Active Users Spreadsheet 

     f. Conducts GCP Audt
         1. Compares GCP User Account Requests vs Google Admin Active Users
         2. Adds "User Roles" from Staff List onto GAdmin data.
         4. Checks Account Expiration Status in GAdmin data.

     g. Conducts AD Audit
         1. Compares Active Directory Groups/Users against GCP Access Request List
         2. Format Table as requied

     h. Save Files in desired SharePoint folder.
   
       
   
