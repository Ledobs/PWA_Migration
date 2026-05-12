#Version: 16.0.7715.1200
<#
.SYNOPSIS
Exports user related data from PWA instance.

.DESCRIPTION
ExportProjectUserContent.ps1 exports related data for a from a given PWA instance.

.PARAMETER Url
Url of the PWA instance.

.PARAMETER ResourceUid
Resource guid of the target user.

.PARAMETER LoginName
Login name of the target user.

.PARAMETER OutputDirectory
The location where user related data files will be stored.

.PARAMETER Options
This is the list of options of different features that can be exported.
This is an optional parameter and the default is All, but customers can choose which features to export.

.PARAMETER Region
The region of the tenant. Can be one of following values: "Default", "ITAR", "Germany" or "China"

.EXAMPLE
.\ExportProjectUserContent.ps1 -Url https://contoso.sharepoint.com/sites/pwa -LoginName joe@contoso.com -OutputDirectory c:\OutputFolder
Exports related data of a user joe@contoso.com and saves it in the OutputFolder.

.EXAMPLE
.\ExportProjectUserContent.ps1 -Url https://contoso.sharepoint.com/sites/pwa -ResourceUid "203A14CE-105A-472C-8F19-F3896D9AA2FA" -OutputDirectory c:\OutputFolder
Exports related data of a user with resource guid "203A14CE-105A-472C-8F19-F3896D9AA2FA" and saves it in the OutputFolder.

.EXAMPLE
.\ExportProjectUserContent.ps1 -Url https://contoso.sharepoint.com/sites/pwa -LoginName joe@contoso.com -OutputDirectory c:\OutputFolder -Options Timesheets
Exports timesheets related data of a user joe@contoso.com and saves it in the OutputFolder.

.EXAMPLE
.\ExportProjectUserContent.ps1 -Url https://contoso.sharepoint.de/sites/pwa -LoginName joe@contoso.de -OutputDirectory c:\OutputFolder -Options Timesheets -Region Germany
Exports timesheets related data of a user joe@contoso.de and saves it in the OutputFolder from Germany region

.NOTES
You need to be a tenant admin, site collection administrator or should have "Manage Users and Groups" and "Access Project Server Reporting Service" permission on the target PWA site to run this script.
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceUid")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginName")]
    [string] $Url,

    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceUid")]
    [Guid] $ResourceUid,

    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginName")]
    [string] $LoginName,

    [Parameter(Mandatory=$true, ParameterSetName = "ByResourceUid")]
    [Parameter(Mandatory=$true, ParameterSetName = "ByLoginName")]
    [string] $OutputDirectory,

    [Parameter(Mandatory=$false)]
    [ValidateSet(
        "All",
        "Resource",
        "Timesheets",
        "TaskStatus",
        "ServiceSettings",
        "Portfolio",
        "Security",
        "StatusReports",
        "Engagements",
        "ResourcePlans",
        "Projects",
        "Workflows",
        "WorkspaceItems",
        "UserViewSettings"
    )]
    [string] $Options = "All",

    [Parameter(Mandatory=$false)]
    [ValidateSet("Default", "ITAR", "Germany", "China")]
    [string] $Region = "Default",

    [Parameter(Mandatory=$false)]
    [switch] $OnPrem
)

# include common.ps1 scripts
. "$PSScriptRoot\Common.ps1"

# include the export scripts.
. "$PSScriptRoot\ExportDraftAndPublishedAsXML.ps1"

function ValidateDirectory([string] $OutputDirectory)
{
    if (-not (Test-Path $OutputDirectory))
    {
        Write-Host "Output directory does not exist, creating $OutputDirectory"
        mkdir $OutputDirectory > $null
        return $true
    }
    elseif (-not (Test-Path -PathType Container $OutputDirectory))
    {
        Write-Error "Output location specified is not a directory. Please specify a directory."
        return $false
    }
    elseif ((Get-ChildItem $OutputDirectory | Measure-Object).Count -gt 0)
    {
        Write-Error "Output location is not empty"
        return $false
    }

    # Output location is an existing empty directory
    return $true
}

function Export-ProjectUserContent ([string] $Url, [Guid] $ResourceUid, [string] $LoginName, [string] $OutputDirectory, [string] $Region, [bool] $OnPrem)
{
    if (-not (ValidateDirectory $OutputDirectory))
    {
        Write-Error "Error validating output location"
        return $false
    }

    # Instantiate PSI proxy - cannot continue if this fails.
    Write-Host "Connecting..."
    $proxy = GetPSIProxy -Url $Url -Region $Region -OnPrem $OnPrem
    Write-Host "Proxy created"

    # Get $ClaimsAccount when $LoginName is given. Also try to get $ResourceUid from the $ClaimsAccount.
    [string] $ClaimsAccount = $null
    if (-not [System.String]::IsNullOrEmpty($LoginName))
    {
        $ClaimsAccount = GetEncodedClaim -LoginName $LoginName -OnPrem $OnPrem
        $ResourceUid = ContinueOnError -ScriptBlock {
            return $proxy.UserDataGetResourceIdFromClaimsAccount($ClaimsAccount)
        }

        # $ResourceUid is not found by given name - run "WssItems" only
        if ($ResourceUid -eq $null -or $ResourceUid -eq [System.Guid]::Empty) 
        {
            Write-Host -ForegroundColor Yellow "Could not find user by given LoginName from the PWA instance, will get WssItems only."
            $Options = "WssItems"
        }
    }

    # Define functions for fetching data - these may be selectively called, depending on user input
    function GetResourceData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related resource data..."
            $data = $proxy.UserDataExportResource($ResourceUid)
            $file = $OutputDirectory + "\Resource.json"
            $data > $file
            Write-Host "Resource data saved to $file"
        }
        
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting reporting resource data..."
            $data = $proxy.UserDataExportReportingResource($ResourceUid)
            $file = $OutputDirectory + "\ReportingResource.json"
            $data > $file
            Write-Host "Reporting resource data saved to $file"
        }
    }

    function GetWorkflowData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related workflow data..."
            $data = $proxy.UserDataExportWorkflow($ResourceUid)
            $file = $OutputDirectory + "\Workflow.json"
            $data > $file
            Write-Host "Workflow data saved to $file"
        }
    }

    function GetPortfolioData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related business driver data..."
            $data = $proxy.UserDataExportDrivers($ResourceUid)
            $file = $OutputDirectory + "\BusinessDrivers.json"
            $data > $file
            Write-Host "Business driver data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related driver prioritization data..."
            $data = $proxy.UserDataExportPrioritizations($ResourceUid)
            $file = $OutputDirectory + "\DriverPrioritizations.json"
            $data > $file
            Write-Host "Driver prioritization data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related portfolio analysis data..."
            $data = $proxy.UserDataExportAnalyses($ResourceUid)
            $file = $OutputDirectory + "\PortfolioAnalyses.json"
            $data > $file
            Write-Host "Portfolio analysis data saved to $file"
        }
    }

    function GetServerSettingsData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related custom field data..."
            $data = $proxy.UserDataExportCustomFields($ResourceUid)
            $file = $OutputDirectory + "\CustomFields.json"
            $data > $file
            Write-Host "Custom field data saved to $file"
        }

        
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related lookup table data..."
            $data = $proxy.UserDataExportLookupTables($ResourceUid)
            $file = $OutputDirectory + "\LookupTables.json"
            $data > $file
            Write-Host "Lookup table data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related calendar data..."
            $data = $proxy.UserDataExportCalendars($ResourceUid)
            $file = $OutputDirectory + "\Calendars.json"
            $data > $file
            Write-Host "Calendar data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related delegations data..."
            $data = $proxy.UserDataExportDelegations($ResourceUid)
            $file = $OutputDirectory + "\Delegations.json"
            $data > $file
            Write-Host "Delegation data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related queue job data..."
            $data = $proxy.UserDataExportQueueJobs($ResourceUid)
            $file = $OutputDirectory + "\QueueJobs.json"
            $data > $file
            Write-Host "Queue job data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related subscribed reminder data..."
            $data = $proxy.UserDataExportSubscribedReminders($ResourceUid)
            $file = $OutputDirectory + "\SubscribedReminders.json"
            $data > $file
            Write-Host "Subscribed reminders data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related unsubscribed alert data..."
            $data = $proxy.UserDataExportUnsubscribedAlerts($ResourceUid)
            $file = $OutputDirectory + "\UnsubscribedAlerts.json"
            $data > $file
            Write-Host "Unsubscribed alert data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related reminder email data..."
            $data = $proxy.UserDataExportReminderEmails($ResourceUid)
            $file = $OutputDirectory + "\ReminderEmails.json"
            $data > $file
            Write-Host "Reminder emails data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related admin audit data..."
            $data = $proxy.UserDataExportAdminAudit($ResourceUid)
            $file = $OutputDirectory + "\AdminAudit.json"
            $data > $file
            Write-Host "Admin audit data saved to $file"
        }
    }

    function GetSecurityData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related security settings data..."
            $data = $proxy.UserDataExportSecurity($ResourceUid)
            $file = $OutputDirectory + "\Security.json"
            $data > $file
            Write-Host "Security settings data saved to $file"
        }
    }

    function GetStatusReportsData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related status report data..."
            $data = $proxy.UserDataExportStatusReports($ResourceUid)
            $file = $OutputDirectory + "\StatusReports.json"
            $data > $file
            Write-Host "Status reports saved to $file"
        }
    }

    function GetTimesheetData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related timesheet data..."
            $i = 0
            $page = 0
            $list = $proxy.UserDataExportTimesheetList($ResourceUid)
            while ($i -lt $list.count) {
                if ($i+$pagingSize -ge $list.count)
                {
                    $guids = $list[$i..($list.count-1)]
                }
                else
                {
                    $guids = $list[$i..($i+$pagingSize-1)]
                }

                $i += $pagingSize
                $page++
                ContinueOnError -ScriptBlock {
                    $file = $OutputDirectory + "\Timesheets_$($page).json"
                    $data = $proxy.UserDataExportTimesheetsDetails($ResourceUid,$guids)
                    $data > $file
                }
            }
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting related reporting timesheet data..."

            $i = 0
            $page = 0
            $list = $proxy.UserDataExportReportingTimesheetList($ResourceUid)
            while ($i -lt $list.count) {
                if ($i+$pagingSize -ge $list.count)
                {
                    $guids = $list[$i..($list.count-1)]
                }
                else
                {
                    $guids = $list[$i..($i+$pagingSize-1)]
                }

                $i += $pagingSize
                $page++
                ContinueOnError -ScriptBlock {
                    $file = $OutputDirectory + "\Timesheets_Reporting$($page).json"
                    $data = $proxy.UserDataExportReportingTimesheetDetails($ResourceUid,$guids)
                    $data > $file
                }
            }
        }
    }

    function GetDraftProjectData
    {
        ContinueOnError -ScriptBlock {
            $listFile = $OutputDirectory + "\DraftProjectList.xml"
            $dataset = $proxy.UserDataExportDraftProjectList($ResourceUid)
            $dataset.Tables[0].WriteXml($listFile)
            Write-Host "Draft project list saved to $listFile" -Foreground green

            if ($dataset.Tables[0].Rows.Count -gt 0)
            {
                $siteId = $($dataset.Tables[0].Rows[0]["SiteId"])
                Connect-WinProjToProjectServer -pwaUrl $Url -siteId $siteId

                foreach ($row in $dataset.Tables[0].Rows)
                {
                    $projName = $($row["PROJ_NAME"])
                    $projUid = $($row["PROJ_UID"])
                    #write-host "ProjUid: $projUid, ProjName: $projName"

                    ContinueOnError -ScriptBlock {
                        Write-Host "Exporting draft data for project: $projName..."
                        $file = $OutputDirectory + "\Project_$($projName)_draft.json"
                        $data = $proxy.UserDataExportDraftProject($ResourceUid, $projUid)
                        $data > $file
                        #Write-Host "Draft data exported for project: $projName"
                    }

                    ContinueOnError -ScriptBlock {
                        Export-DraftProjectsAsXML -projectName $projName -projectGuid $projUid -exportFolder $OutputDirectory
                    }
                }
            }
        }
    }

    function GetPublishedProjectData
    {
        ContinueOnError -ScriptBlock {
            $listFile = $OutputDirectory + "\PublishedProjectList.xml"
            $dataset = $proxy.UserDataExportPublishedProjectList($ResourceUid)
            $dataset.Tables[0].WriteXml($listFile)
            Write-Host "Published project list saved to $listFile" -Foreground green

            if ($dataset.Tables[0].Rows.Count -gt 0)
            {
                $siteId = $($dataset.Tables[0].Rows[0]["SiteId"])
                Connect-WinProjToProjectServer -pwaUrl $Url -siteId $siteId

                foreach ($row in $dataset.Tables[0].Rows)
                {
                    $projName = $($row["PROJ_NAME"])
                    $projUid = $($row["PROJ_UID"])
                    #write-host "ProjUid: $projUid, ProjName: $projName"

                    ContinueOnError -ScriptBlock {
                        Write-Host "Exporting published data for project: $projName..."
                        $file = $OutputDirectory + "\Project_$($projName)_published.json"
                        $data = $proxy.UserDataExportPubProject($ResourceUid, $projUid)
                        $data > $file
                        #Write-Host "Published data exported for project: $projName"
                    }

                    ContinueOnError -ScriptBlock {
                        Export-PublishedProjectsAsXML -projectName $projName -projectGuid $projUid -exportFolder $OutputDirectory
                    }
                }
            }
        }
    }

    function GetReportingProjectData
    {
        ContinueOnError -ScriptBlock {
            $listFile = $OutputDirectory + "\ReportingProjectList.xml"
            $dataset = $proxy.UserDataExportReportingProjectList($ResourceUid)
            $dataset.Tables[0].WriteXml($listFile)
            Write-Host "Reporting project list saved to $listFile" -Foreground green

            foreach ($row in $dataset.Tables[0].Rows)
            {
                $projName = $($row["ProjectName"])
                $projUid = $($row["ProjectUID"])
                #$siteId = $($row["SiteId"])
                #write-host "SiteId: $siteId, ProjUid: $projUid, ProjName: $projName"

                ContinueOnError -ScriptBlock {
                    Write-Host "Exporting reporting data for project: $projName..."
                    $file = $OutputDirectory + "\Project_$($projName)_reporting.json"
                    $data = $proxy.UserDataExportReportingProject($ResourceUid, $projUid)
                    $data > $file
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting tasks..."
                    $file = $OutputDirectory + "\Project_$($projName)_reporting_Tasks.json"
                    $data = $proxy.UserDataExportReportingProjectTasks($ResourceUid, $projUid)
                    $data > $file
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting assignments..."
                    $file = $OutputDirectory + "\Project_$($projName)_reporting_Assignments.json"
                    $data = $proxy.UserDataExportReportingProjectAssignments($ResourceUid, $projUid)
                    $data > $file
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting resources..."
                    $file = $OutputDirectory + "\Project_$($projName)_reporting_Resources.json"
                    $data = $proxy.UserDataExportReportingProjectResources($ResourceUid, $projUid)
                    $data > $file
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting project baselines..."
                    $file = $OutputDirectory + "\Project_$($projName)_reporting_Baselines.json"
                    $data = $proxy.UserDataExportReportingProjectBaseline($ResourceUid, $projUid)
                    $data > $file
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting task timephased data..."
                    $fileName = "Project_$($projName)_reporting_TaskTimephased.json"
                    SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "TaskTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting assignment timephased data..." 
                    $fileName = "Project_$($projName)_reporting_AssignmentTimephased.json"
                    SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "AssignmentTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting task baseline timephased data..."
                    $fileName = "Project_$($projName)_reporting_TaskBaselineTimephased.json"
                    SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "TaskBaselineTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
                }

                ContinueOnError -ScriptBlock {
                    Write-Host "...Exporting assignment baseline timephased data..."
                    $fileName = "Project_$($projName)_reporting_AssignmentBaselineTimephased.json"
                    SaveODataAllPages -Proxy $proxy -BaseUrl $Url -Request "AssignmentBaselineTimephasedDataSet?`$filter=ProjectId eq guid'$($projUid)'" -OutputDirectory $OutputDirectory -OutputfileName $fileName -OnPrem $OnPrem
                }

                #Write-Host "Reporting data exported for project: $projName"
                Write-Host
            }
        }
    }

    function GetProjectData
    {
        Write-Host "Exporting projects..."

        Write-Host "Exporting draft projects..."
        GetDraftProjectData
        Write-Host

        Write-Host "Exporting published projects..."
        GetPublishedProjectData
        Write-Host

        Write-Host "Exporting reporting projects..."
        GetReportingProjectData
        Write-Host
        Write-Host "Project export complete"
        Write-Host
    }

    function GetWorkspaceItemsData
    {
	if ([System.String]::IsNullOrEmpty($ClaimsAccount))
        {
            ContinueOnError -ScriptBlock {
               Write-Host "Exporting related project workspace items..."
                $data = $proxy.UserDataExportIssuesRisksDeliverablesDocumentsByGuid($ResourceUid)
                $file = $OutputDirectory + "\WorkspaceItems.json"
                $data > $file
                Write-Host "Workspace items saved to $file"
            }
        }
        else
        {
            ContinueOnError -ScriptBlock {
                Write-Host "Exporting related project workspace items..."
                $data = $proxy.UserDataExportIssuesRisksDeliverablesDocumentsByResClaims($ClaimsAccount)
                $file = $OutputDirectory + "\WorkspaceItems.json"
                $data > $file
                Write-Host "Workspace items saved to $file"
            }
        }
    }

    function GetTaskStatusData
    {

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting Rules data..."
            $data = $proxy.UserDataExportRules($ResourceUid)
            $file = $OutputDirectory + "\Rules.json"
            $data > $file
            Write-Host "Rules data saved to $file"
        }

        Write-Host "Exporting task status data..."

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting assignments history..."
            $i = 0
            $page = 0
            $list = $proxy.UserDataExportAssignmentsHistoryList($ResourceUid)
            while ($i -lt $list.count) {
                if ($i+$pagingSize -ge $list.count)
                {
                    $guids = $list[$i..($list.count-1)]
                }
                else
                {
                    $guids = $list[$i..($i+$pagingSize-1)]
                }

                $i += $pagingSize
                $page++
                ContinueOnError -ScriptBlock {
                    $file = $OutputDirectory + "\TaskStatus_AssignmentsHistory_$($page).json"
                    $data = $proxy.UserDataExportAssignmentsHistoryDetails($ResourceUid,$guids)
                    $data > $file
               }
            }
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting saved assignments..."
            $data = $proxy.UserDataExportAssignmentsSaved($ResourceUid)
            $file = $OutputDirectory + "\TaskStatus_AssignmentsSaved.json"
            $data > $file
            Write-Host "Saved assignments data saved to $file"
        }

        ContinueOnError -ScriptBlock {
            Write-Host "Exporting submitted assignments..."
            $data = $proxy.UserDataExportAssignmentsSubmitted($ResourceUid)
            $file = $OutputDirectory + "\TaskStatus_AssignmentsSubmitted.json"
            $data > $file
            Write-Host "Submitted assignments data saved to $file"
        }

        Write-Host "Task status data export complete"
   }

    function GetEngagementsData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting engagements..."
            $i = 0
            $page = 0
            $list = $proxy.UserDataExportEngagementList($ResourceUid)
            while ($i -lt $list.count) {
                if ($i+$pagingSize -ge $list.count)
                {
                    $guids = $list[$i..($list.count-1)]
                }
                else
                {
                    $guids = $list[$i..($i+$pagingSize-1)]
                }

                $i += $pagingSize
                $page++
                ContinueOnError -ScriptBlock {
                    $file = $OutputDirectory + "\Engagements_$($page).json"
                    $data = $proxy.UserDataExportEngagementsDetails($ResourceUid,$guids)
                    $data > $file
                }
           }
        }
    }

    function GetResourcePlanData
    {
        Write-Host "Exporting resource plans..."
        
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting draft resource plan data..."
            $i = 0
            $page = 0
            $list = $proxy.UserDataExportDraftResourcePlanList($ResourceUid)
            while ($i -lt $list.count) {
                if ($i+$pagingSize -ge $list.count)
                {
                    $guids = $list[$i..($list.count-1)]
                }
                else
                {
                    $guids = $list[$i..($i+$pagingSize-1)]
                }

                $i += $pagingSize
                $page++
                ContinueOnError -ScriptBlock {
                    $file = $OutputDirectory + "\ResourcePlans_$($page).json"
                    $data = $proxy.UserDataExportDraftResourcePlansDetails($ResourceUid,$guids)
                    $data > $file
                }
            }
        }
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting reporting resource plan data..."
            $data = $proxy.UserDataExportReportingResourcePlans($ResourceUid)
            $file = $OutputDirectory + "\ReportingResourcePlans.json"
            $data > $file
            Write-Host "Reporting resource plan data saved to $file"
        }
    }

    function GetUserViewSettingsData
    {
        ContinueOnError -ScriptBlock {
            Write-Host "Exporting user view settings..."
            $data = $proxy.UserDataExportUserProperties($ResourceUid)
            $file = $OutputDirectory + "\UserViewSettings.json"
            $data > $file
            Write-Host "User view settings saved to $file"
        }
    }

    # More formatting and error handling to come with later check-ins.
    try
    {
        $pagingSize = 200
        switch ($Options)
        {
            "All"
            {
                # This should contain everything from all of the other options
                GetResourceData
                GetSecurityData
                GetUserViewSettingsData
                GetWorkspaceItemsData
                GetServerSettingsData
                GetStatusReportsData
                GetWorkflowData
                GetPortfolioData
                GetEngagementsData
                GetResourcePlanData
                GetTimesheetData
                GetTaskStatusData
                GetProjectData
            }
            "Resource"
            {
                GetResourceData
            }
            "Timesheets"
            {
                GetTimesheetData
            }
            "TaskStatus"
            {
                GetTaskStatusData
            }
            "ServiceSettings" 
            {
                GetServerSettingsData
            }
            "Portfolio"
            {
                GetPortfolioData
            }
            "StatusReports" 
            {
                GetStatusReportsData
            }
            "Security" 
            {
                GetSecurityData
            }
            "Engagements" 
            {
                GetEngagementsData
            }
            "ResourcePlans" 
            {
                GetResourcePlanData
            }
            "Projects" 
            {
                GetProjectData
            }
            "Workflows"
            {
                GetWorkflowData
            }
            "WorkspaceItems"
            {
                GetWorkspaceItemsData
            }
            "UserViewSettings"
            {
                GetUserViewSettingsData
            }
        }
    }
    catch
    {
        ShowExceptionDetail($_.Exception)
    }
}

if ([System.String]::IsNullOrEmpty($LoginName))
{
    Export-ProjectUserContent -Url $Url -ResourceUid $ResourceUid -OutputDirectory $OutputDirectory -Region $Region -OnPrem $OnPrem
}
else
{
    Export-ProjectUserContent -Url $Url -LoginName $LoginName -OutputDirectory $OutputDirectory -Region $Region -OnPrem $OnPrem
}
