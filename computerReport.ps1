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
    if (!($OutputPath))
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
        foreach ($Computer in $Computers)
        {
            $Computer = $Computer.ToUpper()
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
        }
        $successCount = ($ComputersObject | ? {$_.Status -eq "Succeeded"}).count
        $failedCount = ($ComputersObject | ? {$_.Status -eq "Failed"}).count
        #$computersHTMLFragment = Get-ComputersHTMLFragment -ComputersObject $ComputersObject -ErrorAction Stop
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