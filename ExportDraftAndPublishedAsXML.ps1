#Version: 16.0.7715.1200
#usage
#   . \Users\myself\Documents\ExportDraftAndPublishedAsXML.ps1
#   Connect-WinProjToProjectServer -pwaUrl http://<machine>/pwa -siteId {B3CF132F-B7E4-44DB-B1F4-280DE3BFB9CD}
#   Export-PublishedProjectsAsXML -projectName project1 -projectGuid A4D3AA6A-C50A-E811-80E8-00155DE9FA01 -exportFolder C:\Users\myself\Documents\
#   Export-DraftProjectsAsXML -projectName project1 -projectGuid A4D3AA6A-C50A-E811-80E8-00155DE9FA01 -exportFolder C:\Users\myself\Documents\

function RunScriptBlockWithRetry ($ScriptBlock)
{
	$retryCount = 0
	$tryAgain = $true
	while ($tryAgain -and $retryCount -le 30)
	{
		try
		{
			Invoke-Command -ScriptBlock $ScriptBlock
			$tryAgain = $false
		}
		catch
		{
			if ($_.ToString().Contains("RPC_E_CALL_REJECTED"))
			{
				# retry on RPC_E_CALL_REJECTED error - with a second sleep
				Start-Sleep -Seconds 1
				$retryCount++
			}
			else
			{
				throw
			}
		}
	}
}

function Connect-WinProjToProjectServer
{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[string]$pwaUrl,

		[Parameter(Mandatory=$true)]
		[string]$siteId
	)

	echo "1. verify existence of interop dll"
	$gac = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.MSProject').GlobalAssemblyCache
	if (-not $gac)
	{
		throw "interop dll not in gac"
	}

	echo "2. get the working directory for WinProj"
	$p = New-Object -ComObject msproject.application
	$path = $p.Path
	$p.Quit([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
	# echo $path

	echo "3. kill existing winproj processes"
	[System.Diagnostics.Process]::GetProcessesByName("winproj") | % { $_.Kill() }

	echo "4. start winproj connected to given site"
	$process = New-Object -TypeName System.Diagnostics.Process
	$process.StartInfo.FileName = "winproj.exe"
	$process.StartInfo.WorkingDirectory = $path
	$process.StartInfo.Arguments = [String]::Format("/s {0} /g {1}", $pwaUrl, $siteId)
	$process.Start() | Out-Null
	[System.Threading.Thread]::Sleep(5000)

	echo "5. initialize the interop dll"
	$WinProjApp = New-Object -ComObject msproject.application
	RunScriptBlockWithRetry -ScriptBlock {
		$WinProjApp.DisplayAlerts = $false
	}

	echo "6. verify connected profile"
	$retPwaUrl = RunScriptBlockWithRetry -ScriptBlock {
		return $WinProjApp.Profiles.ActiveProfile.Server
	};

	if ($retPwaUrl -contains $pwaUrl -or $pwaUrl -contains $retPwaUrl)
	{
		echo "successfully connected to $retPWaUrl"
	}
	else
	{
		echo "Connected to $retPwaUrl "
	}
	
	$global:WinProjApp = $WinProjApp
}

function Close-WinProj()
{
	if ($global:WinProjApp -ne $null)
	{
		RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.Quit([Microsoft.Office.Interop.MSProject.PjSaveType]::pjDoNotSave)
		}

		$global:WinProjApp = $null
	}
}

function Export-DraftProjectsAsXML
{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[string]$projectName,

		[Parameter(Mandatory=$true)]
		[Guid]$projectGuid,

		[Parameter(Mandatory=$true)]
		[string]$exportFolder
	)

	echo "1. make sure initialize the interop dll - $($projectName)"
	if ($global:WinProjApp -eq $null)
	{
		throw "Please call Connect-WinProjToProjectServer first"
	}

	echo "2. open file from draft store - $($projectName)"
	$openResult = RunScriptBlockWithRetry -ScriptBlock {
		$global:WinProjApp.FileOpenEx("<>\" + $projectName, $true)
	}
	if (-not $openResult)
	{
		throw "Failed to open file - $($projectName) - from draft store"
	}

	try
	{
		echo "3. save file as xml - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_draft.xml", $exportFolder, $projectName), "pjMPP", $false, $false, $true, $true, "", "", "", "MSProject.xml")
		}
		if (-not $saveResult)
		{
			throw "Failed to save - $($projectName) as XML"
		}
	
		echo "4. save file as mpp - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_draft.mpp", $exportFolder, $projectName), "pjMPP")
		}
		if (-not $saveResult)
		{
			throw "Failed to save -$($projectName) as MPP"
		}	
	}
	finally
	{
		echo "5. close file - $($projectName)"
		$closeResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileCloseEx('pjDoNotSave')
		}

		if (-not $closeResult)
		{
			echo "Failed to close $($projectName)"
		}
	}
}
function Export-PublishedProjectsAsXML
{
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[string]$projectName,

		[Parameter(Mandatory=$true)]
		[Guid]$projectGuid,

		[Parameter(Mandatory=$true)]
		[string]$exportFolder
	)

	echo "1. make sure initialize the interop dll - $($projectName)"
	if ($global:WinProjApp -eq $null)
	{
		throw "Please call Connect-WinProjToProjectServer first"
	}

	echo "2. open file from published store - $($projectName)"
	$openResult = RunScriptBlockWithRetry -ScriptBlock {
		$global:WinProjApp.FileOpenEx("<1>\" + $projectName, $true)
	}
	if (-not $openResult)
	{
		throw "Failed to open file - $($projectName) - from published store"
	}

	try
	{
		echo "3. save file as xml - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_published.xml", $exportFolder, $projectName), "pjMPP", $false, $false, $true, $true, "", "", "", "MSProject.xml")
		}
		if (-not $saveResult)
		{
			throw "Failed to save - $($projectName) as XML"
		}

		echo "4. save file as mpp - $($projectName)"
		$saveResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileSaveAs([String]::Format("{0}\Project_{1}_published.mpp", $exportFolder, $projectName), "pjMPP")
		}
		if (-not $saveResult)
		{
			throw "Failed to save - $($projectName) as MPP"
		}
	}
	finally
	{
		echo "5. close file - $($projectName)"
		$closeResult = RunScriptBlockWithRetry -ScriptBlock {
			$global:WinProjApp.FileCloseEx('pjDoNotSave')
		}
		if (-not $closeResult)
		{
			echo "Failed to close $($projectName)"
		}
	}
}
