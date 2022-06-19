# Author: Ismael Contreras 11/4/21
# Description: Pulls the following data: GCP Access Requests, Google Admind Active Users, Staff List(SmartSheets)
# AD Groups. Cleans, merges, and exports audit data in two seperete files (AD_Audit_Month & CGCP_Audit_Month).
# Calls other scripts in directory to accomplish tasks.


if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist"
}





#
#========= CALL OTHER SCRIPTS, IMPORT STAFF LIST AND ACCESS REQUESTS LIST==================#
#==================================================================================================#

## Calls the StaffList download script & AD Script
"Loading Staff Script..."
& $PSScriptRoot\Staff.ps1

"Success`r`n"

"Loading AD Script..."
& $PSScriptRoot\AD_script.ps1
"Success`r`n"

## Import Staff as Staff
"Importing Staff List.."
$Staff = Import-Excel  "$PSScriptRoot\Script Audits\Staff.xlsx"

## Import GCP Access Request as GAR, Only keep required columns. 
"Importing GCP Access Request List.."
$GAR = Import-Csv  "$PSScriptRoot\Script Audits\AccessRequests.csv" | 
    Select-Object -Property  Name, GCP_Email, UC_Email , SubType,Category,Creation-Date,Expiration-Date

## Import Google Admin from spreadsheet. Combine First and Last names
"Importing Google Admin Active Users.."
$GAdmin = Import-Csv  "$PSScriptRoot\Script Audits\GAdmin.csv" | 
    Select-Object -Property  @{label="Name";expression={$($_."First Name [Required]") + " " + $($_."Last Name [Required]")}},
        @{label="GCP_Email";expression={$($_."Email Address [Required]")}},
        @{label="UC_Email";expression={$($_."")}},
        @{label="Status";expression={$($_."Status [READ ONLY]")}},
        @{label="LastSignIn";expression={$($_."Last Sign In [READ ONLY]")}}

$GAdmin = $GAdmin | Select-Object *,'GCP User Account', 'GCP User Account for Eureka Instance', 'GCP Account Creation', 'GCP Account Expiration', 'Expired or Current', 'UserRole'
"Import Complete`r`n"
#///////////////////////////////////////////////////////////////////////////////////////////#
#





# 
#========================CGCP AUDIT AGAINST ACCESS REQUEST LIST================================#
#==============================================================================================#

## Checks for emails associated with GCP User Accounts and Eureka Instance (Google Admin - GCP Access Requests)
"=========================`r`nConducting GCP Audit`r`n========================="
"Comparing GCP User Account Requests with Active Users"
foreach($r in $GAdmin){
    foreach($row in $GAR){
        if ($row.GCP_Email -eq $r.GCP_Email) {    
            $r.UC_Email = $row.UC_Email
            $string = $row.SubType
            if( $string.Contains("GCP User Account")){
                $r.'GCP User Account' = 'TRUE'
                $r.'GCP Account Creation' = $row.'Creation-Date'
                $r.'GCP Account Expiration'= $row.'Expiration-Date'
            }
            if($string.Contains("Eureka")) {
                $r.'GCP User Account for Eureka Instance' = 'TRUE'
            }
        }
    } 
}

# If not TRUE, marks as FALSE
foreach ($item in $GAdmin) {
    if($item.'GCP User Account' -ne 'TRUE') {
        $item.'GCP User Account' = 'FALSE'
    }
    if($item.'GCP User Account for Eureka Instance' -ne 'TRUE'){
        $item.'GCP User Account for Eureka Instance' = 'FALSE' 
    }
}

"Aquiring User Roles from Staff List"
# Acquires User Role from Staff List
foreach ($line in $GAdmin) {
    foreach($l in $Staff){
        if($l.Name -eq $line.Name){
            $line.UserRole = $l.'User Role'
        }
    }
}
"Checking account expiration status"
$Expired = 0
$Missing = 0
$30day = 0
$Current = 0
$vendors = 0
$Today = (Get-Date) -as [datetime]
$Future30 = $Today.AddDays(30)
foreach ($item in $GAdmin){
    if($item.UserRole -like "vendor"){
        $vendors += 1
        $DateChecked = ($item.'GCP Account Expiration') -as [datetime]
        if($DateChecked -gt $Today){
            $item.'Expired or Current' = "Current"
            $Current += 1
            Continue
        }
        if(!$DateChecked){
            $item.'Expired or Current' = "Missing"
            $Missing += 1
            Continue
        }
        if($DateChecked -le $Today){
            $item.'Expired or Current' = "Expired"
            $Expired += 1
            Continue
        }
        elseif($DateChecked -le $Future30){
            $item.'Expired or Current' = "Within 30 Days"
            $30day += 1
            Continue
        }


    }
    
}

"GCP Audit Success`r`n"
#/////////////////////////////////////////////////////////////////////////////////////#
#

$GCPStats = @( @{Expired = $Expired; Within_30_Days = $30day;  Missing = $Missing; Current = $Current; Total = $vendors} ) | 
    % { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru }


# 
#======================ACTIVE DIRECTORY AUDIT AGAINST ACCESS REQUEST LIST=========================#
#=================================================================================================#
"=========================`r`nConducting Active Directory Audit`r`n========================="
$AD = Import-Excel "$PSScriptRoot\Script Audits\AD_Audit.xlsx" | Select-Object FirstLast, Enabled ,OuCn, UserPrincipalName,
    SourceControl, "SourceControl Request", C-BI-Admins, "C-BI-Admins Request", C-BI-Developers,
    "C-BI-Developers Request", C-CIDER-Admin,"C-CIDER-Admin Request",C-CIDER-Dev,
    "C-CIDER-Dev Request", C-ETL-Admins,"C-ETL-Admins Request", C-ETL-Developers,
    "C-ETL-Developers Request", C-ETL-Testers,"C-ETL-Testers Request", CTEAMJB-Admins,
    "CTEAMJB-Admins Request", CTEAMJB-Users,"CTEAMJB-Users Request"
"Comparing Active Directory Users and Groups with GCP Access Request List"
foreach($r in $AD){
    foreach($row in $GAR){
        if ($row.Name -eq $r.FirstLast) {    
            $string = $row.SubType
            if( $string.Contains("SourceControl")){
                $r.'SourceControl Request' = 'TRUE'
            }  
            if($string.Contains("HDC-BI-Admins")) {
                $r.'C-BI-Admins Request' = 'TRUE' 
            } 
            if($string.Contains("HDC-BI-Developers")) {
                $r.'C-BI-Developers Request' = 'TRUE'
            } 
            if($string.Contains("HDC-CIDER-Admin")) {
                $r.'C-CIDER-Admin Request' = 'TRUE'
            } 
            if($string.Contains("HDC-CIDER-Dev")) {
                $r.'C-CIDER-Dev Request' = 'TRUE'
            } 
            if($string.Contains("HDC-ETL-Admins")) {
                $r.'C-ETL-Admins Request' = 'TRUE'
            } 
            if($string.Contains("HDC-ETL-Developers")) {
                $r.'C-ETL-Developers Request' = 'TRUE'
            } 
            if($string.Contains("HDC-ETL-Testers")) {
                $r.'C-ETL-Testers Request' = 'TRUE'
            } 
            if($string.Contains("HDCTEAMJB-Admins")) {
                $r.'CTEAMJB-Admins Request' = 'TRUE'
            } 
            if($string.Contains("HDCTEAMJB-Users")) {
                $r.'CTEAMJB-Users Request' = 'TRUE'
            } 
        }
    } 
}

foreach($r in $AD){
    
    if( $r.'SourceControl Request' -ne 'TRUE'){
        $r.'SourceControl Request' = 'FALSE'
    }  
    if($r.'C-BI-Admins Request' -ne 'TRUE') {
        $r.'C-BI-Admins Request' = 'FALSE' 
       } 
    if($r.'C-BI-Developers Request' -ne 'TRUE') {
        $r.'C-BI-Developers Request' = 'FALSE'
    } 
    if($r.'C-CIDER-Admin Request' -ne 'TRUE') {
        $r.'C-CIDER-Admin Request' = 'FALSE'
    } 
    if($r.'C-CIDER-Dev Request' -ne 'TRUE') {
        $r.'C-CIDER-Dev Request' = 'FALSE'
    } 
    if($r.'H-ETL-Admins Request'-ne 'TRUE') {
        $r.'C-ETL-Admins Request' = 'FALSE'
    } 
    if($r.'C-ETL-Developers Request' -ne 'TRUE') {
        $r.'C-ETL-Developers Request' = 'FALSE'
    } 
    if($r.'C-ETL-Testers Request' -ne 'TRUE') {
        $r.'C-ETL-Testers Request' = 'FALSE'
    } 
    if($r.'CTEAMJB-Admins Request' -ne 'TRUE') {
        $r.'CTEAMJB-Admins Request' = 'FALSE'
    } 
    if ($r.'CTEAMJB-Users Request' -ne 'TRUE') {
        $r.'CTEAMJB-Users Request' = 'FALSE'
    } 
}
"Active Directory Audit Complete`r`n"
#/////////////////////////////////////////////////////////////////////////////////////////////////#
#





#================ SAVE FILES IN EXCEWLL AS TABLES========#
$GAdmin | Export-Excel "$PSScriptRoot\Script Audits\GCP_Audit-$(Get-Date -UFormat %b).xlsx" -WorkSheetname GCP_AUDIT -StartRow 2 -TableName CGCP_Audit_HDC
$GCPStats | Export-Excel "$PSScriptRoot\Script Audits\GCP_Audit-$(Get-Date -UFormat %b).xlsx" -WorkSheetname GCP_VendorStats -StartRow 2 -TableName VendorStats
$AD | Export-Excel "$PSScriptRoot\Script Audits\AD_Audit-$(Get-Date -UFormat %b).xlsx" -StartRow 2 -TableName AD_Audit

"Program Success`r`n"


