<#
.SYNOPSIS
Gets application and services list from remote computers

.DESCRIPTION
Gets application and services list from remote computers

.PARAMETER ComputerList
Path to text file with computernames. One computername per line.

.PARAMETER Credentials
Credentials for connecting to remote computers

.PARAMETER OutputPath
PAth where reports will be saved

.PARAMETER ReportOwner
Person who created reports

.PARAMETER LogName
Log file name

.EXAMPLE
.\computerReport.ps1 -ComputerList .\pcs.txt -OutputPath .\Reports -LogName 'pclog.log' -Credentials (Get-Credential)


#>
[CmdletBinding()]
param ([parameter(Mandatory=$true)]
        [string]$ComputerList,
        [System.Management.Automation.PSCredential]$Credentials,
        [parameter(Mandatory=$true)]
        [string]$OutputPath,
        [string]$ReportOwner = $env:USERNAME,
        [parameter(Mandatory=$true)]
        [string]$LogName
)
Import-Module $(Join-Path $PSScriptRoot 'computerReport.psm1')
Function Start-ComputerReportBuild
{
    [CmdletBinding()]
    param ([parameter(Mandatory=$true)]
            [string]$Computername
    )
    try
    {
        $HTMLFragments = Get-RemoteHTMLFragments -Computername $Computername -Credentials $Credentials -ErrorAction Stop
        New-HTMLReport -Computername $Computername `
                       -Timestamp $(get-date -Format "yyyy-MM-dd HH:mm:ss") `
                       -AppsHTMLFragment $HTMLFragments.AppsHTML `
                       -ServicesHTMLFragment $HTMLFragments.ServicesHTML `
                       -ReportOwner $ReportOwner `
                       -OutputPath $OutputPath `
                       -ErrorAction Stop
        Write-Output "Succeeded"
    }
    catch
    {
        Write-Output "Failed"
    }
}

try
{
if (Test-Path $ComputerList)
{
    $Computers = Get-Content -Path $ComputerList
    if (!(Test-Path $OutputPath))
    {
        $null = mkdir $OutputPath
    }
    if ([string]::IsNullOrEmpty($Computers))
    {
        Write-Error "$ComputerList is empty"
    }
    else
    {
        $computersHTMLFragment = @()
        $null = New-Item -Path $OutputPath -Name $LogName -ItemType File -Force -ErrorAction Stop
        $logFile = Join-Path $OutputPath $LogName
        $i = 0
        $successCount = 0
        $failedCount = 0
        foreach ($Computer in $Computers)
        {
            $Computer = $Computer.ToUpper()
            Write-Progress -Activity "Gathering Data from $Computer"  -percentComplete ($i / $Computers.count*100)
            $obj = $null
            $obj  = New-Object Object
            $obj  | Add-Member Noteproperty 'Computername' -value $Computer
            $obj  | Add-Member Noteproperty 'Timestamp' -value $(get-date -Format "yyyy-MM-dd HH:mm:ss")
            if (Test-WSMan -ComputerName $Computer -ErrorAction Stop)
            {
                $buildStatus = Start-ComputerReportBuild -Computername $Computer
                $obj  | Add-Member Noteproperty 'Status' -value $buildStatus
            }
            else
            {
                $obj  | Add-Member Noteproperty 'Status' -value "Failed"
                Write-Host "$Computer is offline" -ForegroundColor Yellow
            }
            $computersHTMLFragment += Get-ComputerHTMLFragment -ComputersObject $obj -ErrorAction Stop
            Add-Content -Path $logFile -Value "$($obj.Timestamp) $($obj.Computername) $($obj.Status)" 
            if ($obj.Status -eq "Succeeded") 
            {
                $successCount++
            }
            else
            {
                $failedCount++
            }
            $i++
        }
        New-HomeHTMLReport -ComputersHTMLFragment "$computersHTMLFragment" `
                           -LogName $LogName `
                           -Timestamp $(get-date -Format "yyyy-MM-dd HH:mm:ss") `
                           -ReportOwner $ReportOwner `
                           -SuccessCount $successCount `
                           -FailCount $failedCount `
                           -OutputPath $OutputPath `
                           -ErrorAction Stop
    }
}
else
{
    Write-Error "$ComputerList doesnt exist"
}
}
catch 
{
    throw $Error[0].Exception
}