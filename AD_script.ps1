# AD Audit (ADA)
# Exports per-user audit information for C based on a pre-defined list of AD groups
# Modified: Ismael Contreras 11/4/21


# A list of AD Group GUIDs to pull membership data from, since group names can\do change
# You can get this with something like this and putting quotes around the output:
# ("GCP_AD_Admins" | Get-AdGroup).ObjectGuid -join "`",`""
$AdaGroupGuids = "d8739920-a******","7a7eeec8-f************"


#Output File
$AdaOutputPath = "$PSScriptRoot\Script Audits\AD_Audit.xlsx"

# Lookup all the group GUIDs in AD, and then sort the group details by name.  This gets the current name for each group we need to audit
$AdaGroupsResolved = $AdaGroupGuids | Get-AdGroup | Sort-Object Name

# Get a list of all member GUIDs (UserGuid) in all groups and output the UserGuid and GroupName for each member
$AdaGroupMembers = ForEach ($AdaGroup in $AdaGroupsResolved) {
    Get-AdGroupMember $AdaGroup -Recursive | Select-Object @{n="UserGuid";e={$_.ObjectGuid}},@{n="GroupName";e={$AdaGroup.Name}}
}

# Deduplicate the list of UserGuids and GroupNames by grouping the list by UserGuid.  This creates a list of users in any group, and the groups they are in.
$AdaGroupsByMember = $AdaGroupMembers | Group-Object -Property UserGuid

# Process each unique member and get the AD information we need for each user
$AdaOutput = ForEach ($AdaMember in $AdaGroupsByMember) {
    Remove-Variable AdaFe*
    $AdaFeUser = Get-AdUser $AdaMember.Name -properties CanonicalName
    [PsCustomObject][Ordered]@{
        FirstLast = "$($AdaFeUser.GivenName) $($AdaFeUser.SurName)"
        Enabled = $AdaFeUser.Enabled
        OuCn = ($AdaFeUser.CanonicalName -split("/") | Select-Object -SkipLast 1) -Join ("/")
        UserPrincipalName = $AdaFeUser.UserPrincipalName
        ObjectGuid = $AdaFeUser.ObjectGuid
        ScriptDate = $AdaDate
    }
}

# Process each group to audit, adding the name of the group as a column to the output variable.
# Then, go through the list of unique members and determine if each member is in the new column's group or not.
# TRUE = in the group.  FALSE = not in the group.
ForEach ($AdaGroup in $AdaGroupsResolved) {
    $AdaOutput | Add-Member -MemberType NoteProperty -Name $AdaGroup.Name -Value $NULL
    $AdaOutput | ForEach-Object {
        $_.$($AdaGroup.Name) = ($AdaGroupsByMember | Where-Object Name -eq $_.ObjectGuid).Group.GroupName -Contains $AdaGroup.Name
    }
}

$AdaOutput | Export-Excel $AdaOutputPath -TableName AD_Audit
