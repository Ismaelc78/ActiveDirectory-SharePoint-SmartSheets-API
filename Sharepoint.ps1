# Author: Ismael Contreras 11/4/21
# Description: Connects to SharePoint Online. Pull SharePoint Lists and Sorts/Formats into CSV files. 

#Import the required DLL
Import-Module 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
#OR
Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'

#Mysite URL
$site = 'https://SomeSite.sharepoint.com/site/team'

#Admin User Principal Name
$admin = Read-Host 'Enter SharePoint Email Log In' 

#Get Password as secure String
$password = Read-Host 'Enter Password' -AsSecureString

#Get the Client Context and Bind the Site Collection
$context = New-Object Microsoft.SharePoint.Client.ClientContext($site)

#Authenticate/ Load Lists
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
$context.Credentials = $credentials
$TrListName = 'CGCP Training Log'
$ARListName = 'CGCP Access Requests'
$TrainingList = $context.Web.Lists.GetByTitle($TrListName)
$ARList = $context.Web.Lists.GetByTitle($ARListName)
$TrainingListItems = $TrainingList.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$ARListItems = $ARList.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$context.Load($ARListItems)
$context.Load($TrainingListItems)
$context.ExecuteQuery()      
"SharePoint Status: Connected"
"Data Retrieval: Success"


## Set up Access Requests Table
$AccessRequests = $ARListItems.TypedObject | 
Select-Object -Property  @{label='Name';expression={$($_.FieldValues.User_x0020_Account_x0020_Holder)}},
    @{label = "GCP_Email";expression={$($_.FieldValues.Description_x0020_of_x0020_Acces)}},
    @{label = "UC_Email";expression={$($_.FieldValues.User_x0020_Email_x0020_Address)}},
    @{label="SubType";expression={$($_.FieldValues.Account_x0020_Type )}},
    @{label="Category";expression={$($_.FieldValues.Category )}},
    @{label="Creation-Date";expression={$($_.FieldValues.Account_x0020_Creation_x002f__x0 )}},
    @{label="Expiration-Date";expression={$($_.FieldValues.Account_x0020_Expiration_x0020_D)}}

#Group By email and sub-type type. Only keep latest submission
$AccessRequests = $AccessRequests |   
Sort-Object -Property @{Expression = "GCP_Email"; Descending = $true},
    @{Expression = "Creation-Date"; Descending = $true} | 
Group-Object {$_.'GCP_Email'}, {$_.'SubType'} |
ForEach-Object {$_ |Select-Object -ExpandProperty Group |
    Sort-Object -Property $_.'Creation-Date' -Descending |
    Select-Object -First 1
}

"Access Requests Count: " + $AccessRequests.Count



## Set up Training List Table
$Training = $TrainingListItems.TypedObject |  
Select-Object -Property  @{label='Name';expression={$($_.FieldValues.Completed_x0020_By.LookupValue)}},
    @{label='Email';expression={$($_.FieldValues.Completed_x0020_By.Email)}},
    @{label='Training';expression={$($_.FieldValues.Training.GetValue(0))}},
    @{label='Created Date';expression={($($_.FieldValues.Completion.Date)).GetDateTimeFormats()[2]}} 

#Group By email and training type. Only Keeps latest submission
$Training = $Training | 
Where-Object {($_.'Training' -like '*Skillsoft*') -or ($_.'Training' -like '*HIPS*')} |  
Sort-Object -Property @{Expression = "Email"; Descending = $true},
    @{Expression = "Created Date"; Descending = $true} | 
Group-Object {$_.'Email'}, {$_.'Training'} |
ForEach-Object {$_ |Select-Object -ExpandProperty Group |
    Sort-Object -Property $_.'Created Date' -Descending |
    Select-Object -First 1
}

"Training Log Count: " + $Training.Count


# Export tables

$Training | Export-Csv "$PSScriptRoot\Script Audits\TrainingLog.csv" -NoTypeInformation
$AccessRequests | Export-Csv "$PSScriptRoot\Script Audits\AccessRequests.csv" -NoTypeInformation

"Uploaded Files: Success"
"SharePoint Status: Disconnected"
"Program Successful..."
""
""
Read-Host -Prompt "Please enter to exit"
""
""
