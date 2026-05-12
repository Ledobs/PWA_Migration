#Version: 16.0.7715.1200
<#
.SYNOPSIS
Removes personal data of a user from PWA instance.

.DESCRIPTION
Invoke-RedactProjectUser.ps1 removes personal data of a user from PWA instance and optionally can update the display name of the user.

.PARAMETER Url
Url of the PWA instance.

.PARAMETER ResourceId
Resource guid of the target user.

.PARAMETER LoginName
Login name of the target user.

.PARAMETER UpdateDisplayName
New display name for the target user

.PARAMETER RedactTimesheet
When $true, update names from timesheet records also.
When $false, delinks the timesheet records.

.PARAMETER Region
The region of the tenant. Can be one of following values: "Default", "ITAR", "Germany" or "China"

.EXAMPLE
.\Invoke-RedactProjectUser.ps1 -Url https://contoso.sharepoint.com/sites/pwa -LoginName joe@contoso.com -UpdateDisplayName "Deleted User" -RedactTimesheet $false
Remove the personal data of a user joe@contoso.com and update display name to "Deleted User" everywhere except timesheet records.

.EXAMPLE
.\Invoke-RedactProjectUser.ps1 -Url https://contoso.sharepoint.com/sites/pwa -LoginName joe@contoso.com
Remove the personal data of a user joe@contoso.com except the display name.

.EXAMPLE
.\Invoke-RedactProjectUser.ps1 -Url https://contoso.sharepoint.com/sites/pwa -ResourceId "203A14CE-105A-472C-8F19-F3896D9AA2FA" -UpdateDisplayName "Deleted User" -RedactTimesheet $true
Remove the personal data of a user with resource guid "203A14CE-105A-472C-8F19-F3896D9AA2FA" and update the display name to "Deleted User" everywhere.

.EXAMPLE
.\Invoke-RedactProjectUser.ps1 -Url https://contoso.sharepoint.de/sites/pwa -LoginName joe@contoso.de -Region Germany
Remove the personal data of a user joe@contoso.de except the display name from Germany region.

.NOTES
You need to be a farm admin, a site collection administrator or should have "Manage Users and Groups" permission on the target PWA site to run this script.
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithoutUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithoutUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithUpdateDisplayName")]
    [string] $Url,

    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithoutUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithUpdateDisplayName")]
    [guid] $ResourceId,

    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithoutUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithUpdateDisplayName")]
    [string] $LoginName,

    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithUpdateDisplayName")]
    [string] $UpdateDisplayName,

    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceIdWithUpdateDisplayName")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginNameWithUpdateDisplayName")]
    [bool] $RedactTimesheet,

    [Parameter(Mandatory=$false, ParameterSetName = "ByResourceIdWithUpdateDisplayName")]
    [Parameter(Mandatory=$false, ParameterSetName = "ByLoginNameWithUpdateDisplayName")]
    [ValidateSet("Default", "ITAR", "Germany", "China")]
    [string] $Region = "Default"
)

# include common.ps1 scripts
. "$PSScriptRoot\Common.ps1"

$OnPrem = $false

function RedactUser ([string] $Url, [guid] $ResourceId, [string] $LoginName, [string] $ResourceIdForDisplay, [string] $UpdateDisplayName, [bool] $RedactTimesheet, [string] $Region, [Boolean] $OnPrem)
{
    $proxy = GetPSIProxy -Url $Url -Region $Region -OnPrem $OnPrem

    try
    {
        if ([System.String]::IsNullOrEmpty($LoginName))
        {
            $newTimesheetResourceUid = $proxy.UserDataRedactUserByResourceId($ResourceId, $UpdateDisplayName, $RedactTimesheet)
        }
        else
        {
            $encodedClaim = GetEncodedClaim -LoginName $LoginName -OnPrem $OnPrem
            $newTimesheetResourceUid = $proxy.UserDataRedactUserByClaimsAccount($encodedClaim, $UpdateDisplayName, $RedactTimesheet)
        }

        if ([System.String]::IsNullOrEmpty($UpdateDisplayName))
        {
            Write-Host "All personal data for resource $($ResourceIdForDisplay) has been removed, except the name of the resource."
        }
        else
        {
            if ($RedactTimesheet)
            {
                Write-Host "All personal data for resource $($ResourceIdForDisplay) has been removed and the name of the resource has been changed to '$($UpdateDisplayName)' everywhere including timesheet records."
            }
            else
            {
                Write-Host "All personal data for resource $($ResourceIdForDisplay) has been removed and the name of the resource has been changed to '$($UpdateDisplayName)', except in timesheet records. Timesheet records that were linked to $($ResourceIdForDisplay) has been linked to a new resource GUID $($newTimesheetResourceUid). The name of the resource on the timesheet records remains the same."
            }
        }
    }
    catch
    {
        ShowExceptionDetail($_.Exception)
    }
}

function ConfirmUser([string] $message)
{
    Write-Host
    Write-Host "Confirm"
    Write-Host $message
    Write-Host -NoNewline "[Y] Yes "
    Write-Host -NoNewLine -ForegroundColor Yellow "[N] No "
    Write-Host -NoNewLine "(default is ""N""): "
    $input = Read-Host

    if ($input -eq "y" -or $input -eq "yes")
    {
        return $true
    }

    return $false
}

# run with confirmation.
if ([System.String]::IsNullOrEmpty($LoginName))
{
    $ResourceIdForDisplay= $ResourceId.ToString()
}
else
{
    $ResourceIdForDisplay = $LoginName
}

if ([System.String]::IsNullOrEmpty($UpdateDisplayName))
{
    if (-not (ConfirmUser("Are you sure you want to remove the personal data of the resource $($ResourceIdForDisplay)?")))
    {
        return
    }
}
else
{
    if ($RedactTimesheet)
    {
        if (-not (ConfirmUser("The personal data for $($ResourceIdForDisplay) will be removed and the name of the resource will be changed to '$($UpdateDisplayName)' everywhere including ALL timesheet records. Do you want to continue?")))
        {
            return
        }
    }
    else
    {
        if (-not (ConfirmUser("The personal data for $($ResourceIdForDisplay) will be removed and the name of the resource will be changed to '$($UpdateDisplayName)' everywhere EXCEPT for timesheet records. Timesheets referring to $($ResourceIdForDisplay) will be updated to a new GUID delinking the original GUID from the name stored on timesheet record. Do you want to continue?")))
        {
            return
        }
    }
}

if ($ResourceId -eq $null)
{
    RedactUser -Url $Url -LoginName $LoginName -ResourceIdForDisplay $ResourceIdForDisplay -UpdateDisplayName $UpdateDisplayName -RedactTimesheet $RedactTimesheet -Region $Region
}
else
{
    RedactUser -Url $Url -ResourceId $ResourceId -ResourceIdForDisplay $ResourceIdForDisplay -UpdateDisplayName $UpdateDisplayName -RedactTimesheet $RedactTimesheet -Region $Region
}
