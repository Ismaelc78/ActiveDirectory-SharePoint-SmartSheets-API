# Author: Ismael Contreras 11/4/21
# Description: Pulls data from a worksheet on SmartSheets.com, required for the CGCP Audit
# Requirements: User must obtain a token from SmartSheets. Log into smartsheets. 
# Click on your account (bottom left corner), click 'Personal Settings', click 'API Access',
# click 'Generate new access token', enter a name for it. Copy down the token as you won't be able
# to see it again. This token is your credential and is associated with your account. Don't share it.

$global:Exp = 0
$global:Miss = 0
$global:30 = 0
$global:5 = 0
$global:Curr = 0

function CheckExpiration ($DateChecked, $k){

    if(!$DateChecked){
        $global:Miss += 1
        return "Missing"
    }
    $Expiration = (Get-Date).AddYears(-$k);  
    $Under30 = $Expiration.AddDays(30)
    $Under5 = $Expiration.AddDays(5)

    if($DateChecked -le $Expiration){
        $global:Exp += 1
        return "Expired"
    }
    if($DateChecked -le $Under30){
        if($DateChecked -le $Under5){ 
            $global:5 += 1 
            return "Expires under 5 Days"
        }
        else{ 
            $global:30 += 1
            return "Expires under 30 Days"
        }
    }
    $global:Curr += 1
    return "Current"
}

##============================================================
## CONNECT TO SMATSHEETS AND PULL A SHEET
##============================================================
$inp | Read-Host "Enter your Smartsheets Token" -MaskInput 
$Token = $inp | ConvertTo-SecureString -AsPlainText -Force

$header = @{
    Accept = "text/csv"
}

$Param = @{
    Uri =  "https://api.smartsheet.com/2.0/sheets/SMARTSHEETS_ID_HERE"
    Authentication = "Bearer"
    Token = $Token
    Headers = $header
}

$CSL = Invoke-RestMethod  @Param  | ConvertFrom-Csv 



##============================================================
## Import required files.
##============================================================

$CSL = $CSL | Select-Object @{label="Name";expression={$($_."Staff User Names")}}, "Email", "User Role", 
    "User Status", "Decommission Date", "Staff Agreement (valid 1 year)", 
    @{label="CSA Status";expression={$($_."")}}, "HIPAA Agreement (Valid 1 year)", 
    @{label="HIPAA Status";expression={$($_."")}}, "CITI HIPS (Valid 3 years)", 
    @{label="CITI Status";expression={$($_."")}} | Where-Object {$_.Name -ne ""}

$TrainLog = Import-Csv  "$PSScriptRoot\Script Audits\TrainingLog.csv"  | 
    Select-Object -Property  @{label="Name";expression={$($_."Name")}},
    @{label="Email";expression={$($_."Email")}},
    @{label="Training";expression={$($_."Training")}},
    @{label="Created Date";expression={$($_."Created Date")}}


$DocuSign = Import-Csv "$PSScriptRoot\Script Audits\DocuSign.csv" | 
    Select-Object -Property 'Recipient Name', 'Completed On',  
    @{label="Status";expression={$($_."")}}

    

foreach ($item in $DocuSign) {
    $item.'Completed On' = Get-Date($item.'Completed On'.Split(" ")[0])
}    
$DocuSign = $DocuSign | Group-Object -Property 'Recipient Name' | 
    ForEach-Object {$_ |Select-Object -ExpandProperty Group |
    Sort-Object -Property 'Completed On' -Descending |
    Select-Object -First 1
}


"=========================`r`nConducting Training Audit`r`n========================="
"Comparing GCP Training Log with Staff List..."
foreach ($line in $CSL) {
    foreach($row in $TrainLog){
        if($row.'Email' -eq $line.'Email') {
            if($row.'Training' -like '*Skillsoft*'){
                $line.'HIPAA Agreement (Valid 1 year)' = $row.'Created Date'
            }
            elseif ( $row.'Training' -like '*CITI*'){
                $line.'CITI HIPS (Valid 3 years)' = $row.'Created Date'
            }
        }
    }
    foreach($r in $DocuSign){
        if($line.Name -like $r.'Recipient Name'){
            $line.'Staff Agreement (valid 1 year)' = $r.'Completed On'
        }
    }
}



"Checking for the follwing expiration status:`r`nWithin 30 Days   Within 5 Days   Expired"


foreach ($item in $CSL){
    $DateChecked = ($item.'Staff Agreement (valid 1 year)') -as [datetime]
    $k = 1
    $item.'CSA Status' = CheckExpiration $DateChecked $k
    $DateChecked = ($item.'HIPAA Agreement (Valid 1 year)')  -as [datetime]
    $item.'HIPAA Status' = CheckExpiration $DateChecked $k
    $DateChecked = ($item.'CITI HIPS (Valid 3 years)') -as [datetime]
    $k = 3
    $item.'CITI Status' = CheckExpiration $DateChecked $k
}



$Stats = @( @{Expired = $global:Exp; Within_30_Days = $global:30;Within_5_Days = $global:5; Missing = $global:Miss; Current = $global:Curr; Total = $CSL.Count * 3} ) | 
    % { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru }



## Save to excel files as tables.

$CSL | Export-Excel "$PSScriptRoot\Script Audits\Staff.xlsx" -WorkSheetname Staff -StartRow 1 -TableName Staff_List 
$Stats | Export-Excel "$PSScriptRoot\Script Audits\Staff.xlsx" -WorkSheetname Stats -StartRow 1 -TableName Stats 
##$CSL | Export-Excel "C:\Users\user\Audits\Training Audit\Training Audit-$(Get-Date -UFormat %b).xlsx" -WorkSheetname Staff -StartRow 1 -TableName Staff_List 
