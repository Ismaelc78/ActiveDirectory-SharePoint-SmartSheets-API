## PowerShell Script 
De-Identified method used to compare data from different sites for training and access requirments from users.( SharePoint Lists, SmartSheet data, Active Directory Groups, DocuSign)

Pulls required data through REst APIs/PowerShell :  
* Active Directory users by group 
* Google Admin Active users 
* Pull CGCP requests for AD/GCP accounts from SharePoint
* Qurey Staff List from SmartSheets 

Format and Merge Files: 
* Sort data in tabulated format
* Rename columns for convenience 
* Format columns and remove redundancies 
* Merges tables 
* Query for necessary data in Audits

Saving/Upload: 
* Creates Excel Workbooks: GCP Audit, AD Audit, Training Audit
* Uploads completed audit files to desired SharePoint folder 


## Process / Organization
#### 1. SharePoint.ps1:
   * Prompts user for SharePoint credentials
   * Connect/Authenticate with SharePointOnline Admin access via REst API
   * Query SharePoint Online for desired SharePoint List data.
   * Saves queried data as .csv files in SharePoint folder (TrainingLog.csv, AccessRequests.csv)
   * Notifes User when complete.
   * *TrainingLog.csv* 
   * <img src="https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/Tlog.png" width="600" height="300">
   
   * *AccessRequest.csv* ![alt text](https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/AR.png?raw=true) 
       
#### 2. AuditScript.ps1 :
   * Calls Staff.ps1
     * Prompts user for SmartSheet credentials
     * Connect/Authenticate SmartSheets access via REst API
     * Pulls Data from a worksheet in SmartSheets as StaffList
     * Imports GCP TrainingLog
     * Imports DocuSign spreadsheet
     * <img src="https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/Doc.png" width="600" height="300">
       
     * Groups/Sorts items in files to only keep most recent training and certs completed.
     * Compares GCP TrainingLog vs Staff List
     * Checks and marks Expiration Status (< 30 Days,   < 5 Days,   Expired)
     * Saves data as tables in Excel Workbooks (Staff.xlsx , TrainingAudit(date).xlsx)
     * *TrainingAudit(date).xlsx* ![alt text](https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/Staff.png?raw=true)
     
     
   * Calls AD_script.ps1
     * Connects to Active Directory
     * Exports per-user data based on a pre-defined list of AD Groups
     * Looks up GUIDs in AD. 
     * Gets List of all member GUIDs- Outputs UserGUID and GroupNames.
     * Creates list of Users and AD Groups they belong to.
     * Get AD info for eachunique  member
     * Save data as "AD_Audit(date).xlsx"
     
     
   * Imports Staff List 
   
   * Imports GCP Access Requests List
   
   * Imports Google Admin Active Users Spreadsheet 
     * *GAdmin.csv*  ![alt text](https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/GAdmin.png?raw=true)
   
   * Conducts GCP Audt
     * Compares GCP User Account Requests vs Google Admin Active Users
     * Adds "User Roles" from Staff List onto GAdmin data.
     * Checks Account Expiration Status in GAdmin data.
     * *GCP_Audit(date).xlsx*   ![alt text](https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/GCPAudit.png?raw=true)   
   * Conducts AD Audit
     * Compares Active Directory Groups/Users against GCP Access Request List
     * Format Table as requied
     * *AD_Audit(date).xlsx*   ![alt text](https://github.com/Ismaelc78/ActiveDirectory-SharePoint-SmartSheets-API/blob/main/Audit.png?raw=true)   
      
   * Save Files in desired SharePoint folder.
   
       
   
