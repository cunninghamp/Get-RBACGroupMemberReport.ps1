<#
.SYNOPSIS
Get-RBACGroupMemberReport.ps1 - Enumerate the membership of Exchange RBAC groups

.DESCRIPTION 
This PowerShell script reports the membership of Exchange RBAC groups.

.OUTPUTS
Results are output to CSV. Each RBAC group is reported in a separate CSV file
with a Summary.csv file also generated.

.EXAMPLE
.\Get-RBACGroupMemberReport.ps1
Generates the CSV files.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 15/07/2015 - Initial version
#>

#requires -version 2


#...................................
# Variables
#...................................

$summary = @()
$summaryFile = "Summary.csv"


#...................................
# Initialize
#...................................

#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (Test-Path $env:ExchangeInstallPath\bin\RemoteExchange.ps1)
{
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}
else
{
    Write-Warning "Exchange Server management tools are not installed on this computer."
    EXIT
}


#...................................
# Script
#...................................

# Set the AD Server Settings to handle multi-domain forests
Set-ADServerSettings -ViewEntireForest $true

# Retrieve AD forest information
$forest = (Get-ADForest).Name
Write-Host "AD Forest: $forest"

$forestDN = (Get-ADDomain $forest).DistinguishedName

# Locate a domain controller in the forest root domain too use for queries
$rootDC = (Get-ADDomainController -Discover -DomainName $forest).HostName[0]
Write-Host "AD DC: $rootDC"

# Get the list of RBAC role groups
$RoleGroups = @(Get-RoleGroup)


# Loop through the list of role groups and extract membership
foreach ($RoleGroup in $RoleGroups)
{
    Write-Host -ForegroundColor White "----------------- Processing $RoleGroup"
    
    $MemberList = @()

    $RoleGroupMembers = @(Get-ADGroupMember $RoleGroup.Name -Server $rootDC -Recursive | Get-ADUser -Properties CanoniCalName,DisplayName,SamAccountName,UserPrincipalName,Enabled,PasswordLastSet)

    if ($RoleGroupMembers.Count -gt 0)
    {
        Write-Host "Getting info about group members"

        foreach ($member in $RoleGroupMembers)
        {
            
            if (!($MemberList.CanoniCalName -icontains $member.CanoniCalName))
            {
                Write-Host -Foreground Green "Adding $($member.DisplayName)"
                $MemberList += $member
            }
            else
            {
                Write-Host -Foreground Cyan "Results already include $($member.DisplayName)"
            }
        }
    
    #Export the membership of the role group to CSV
    $MemberList | Sort CanonicalName| Export-CSV -NoTypeInformation -Path "$($RoleGroup.Name)-Members.csv"

    }
    else
    {
        Write-Host "$RoleGroup contains no members"
    }
    
    # Calculate some stats for the summary CSV
    $totalcount = $MemberList.Count
    $enabledcount = @($MemberList | Where {$_.Enabled -eq $true}).count
    $disabledcount = @($MemberList | Where {$_.Enabled -eq $false}).count

    # Custom object foor the summary CSV data
    $summaryObj = New-Object PSObject
    $summaryObj | Add-Member NoteProperty -Name "Role Group" -Value $RoleGroup.Name
    $summaryObj | Add-Member NoteProperty -Name "Total Members" -Value $totalcount
    $summaryObj | Add-Member NoteProperty -Name "Enabled Accounts" -Value $enabledcount
    $summaryObj | Add-Member NoteProperty -Name "Disabled Accounts" -Value $disabledcount

    $summary += $summaryObj
}

# Generate the summary CSV file
$summary | Export-CSV -NoTypeInformation -Path $summaryFile

Write-Host "Finished."

#...................................
# Finished
#...................................