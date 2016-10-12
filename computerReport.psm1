
Function Get-RemoteHTMLFragments
{
    [CmdletBinding()]
    param ([parameter(Mandatory=$true)]
           [string]$Computername,
           [System.Management.Automation.PSCredential]$Credentials)
    try
    {
        If ([string]::IsNullOrEmpty($Credentials))
        {
            $session = New-PSSession -ComputerName $Computername -ErrorAction Stop
        }
        else
        {
            $session = New-PSSession -ComputerName $Computername -Credential $Credentials -ErrorAction Stop
        }
        Invoke-Command -Session $session -ScriptBlock {
            function Get-InstalledApps
            {
                try
                {
                    $installedApps = @()  
                    $UninstallRegKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                    # Get the architecture 32/64 bit
                    if ([IntPtr]::Size -eq 8)
                    {
                        # If 64 bit check both 32 and 64 bit locations in the registry
                        $RegistryViews = @('Registry32','Registry64')
                    } else {
                        # Otherwise only 32 bit
                        $RegistryViews = @('Registry32')
                    }
                    $sudas = @()
                    foreach ( $RegistryView in $RegistryViews )
                    {
                        $HKLM = [microsoft.win32.registrykey]::OpenBaseKey('LocalMachine',$RegistryView)
                        $UninstallRef = $HKLM.OpenSubKey($UninstallRegKey)
                        $Applications = $UninstallRef.GetSubKeyNames()

                        foreach ($App in $Applications)
                        {           
                            $AppRegistryKey = $UninstallRegKey + "\\" + $App
                            $AppDetails = $HKLM.OpenSubKey($AppRegistryKey)
                            $AppGUID = $App
                            $AppDisplayName = $($AppDetails.GetValue("DisplayName"))
                            $AppVersion = $($AppDetails.GetValue("DisplayVersion"))
                            $Publisher = $($AppDetails.GetValue("Publisher"))
                            if(!$AppDisplayName) { continue }


                            $OutputObj = New-Object -TypeName PSobject
                            $OutputObj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $AppDisplayName
                            $OutputObj | Add-Member -MemberType NoteProperty -Name Version -Value $AppVersion
                            $OutputObj | Add-Member -MemberType NoteProperty -Name Publisher -Value $Publisher
                            $installedApps += $OutputObj

                        }
                    }
                    Write-Output $installedApps
                }
                catch            
                {
                    throw $Error[0].Exception
                }
            }
            Function Get-AppsTableBody
            {
                [CmdletBinding()]
                param ()
                try
                {
                $apps = Get-InstalledApps
                $html = @()
                foreach ($app in $apps)
                {
                    if (!($app.Publisher -eq "Microsoft" -and $app.Displayname -like "*KB*"))
                    {
                        $html += @"
                    <tr>
                        <td>$($app.DisplayName)</td>
                        <td>$($app.Version)</td>
                        <td>$($app.Publisher)</td>        
                    </tr>   
"@
                     }
                }
                Write-Output $html
                }
                catch
                {
                    throw $Error[0].Exception
                }

            }
            Function Get-ServicesTableBody
            {
                [CmdletBinding()]
                param ()
                $Services = Get-Service -ErrorAction Stop | ? {($_.StartType -eq "Automatic") -or ($_.status -eq "Running")} | 
                    Select Name, 
                           DisplayName, 
                           Status

                $html = @()
                foreach ($Service in $Services)
                {
               $html += @"
                    <tr>
                        <td>$($Service.Name)</td>
                        <td>$($Service.DisplayName)</td>
                        <td>$($Service.Status)</td>        
                    </tr>   
"@
                }
                Write-Output $html
            }
            try
            {    
                $AppsHTML = Get-AppsTableBody -ErrorAction Stop
                $ServicesHTML = Get-ServicesTableBody -ErrorAction Stop
                $hash = @{"AppsHTML" = $AppsHTML; "ServicesHTML" = $ServicesHTML}
                Write-Output $hash

            }
            catch
            {
                throw $Error[0].Exception
            }
    
    
        } -ErrorAction Stop
        $session | Remove-PSSession -ErrorAction Stop
    }
    catch
    {
        throw $Error[0].Exception
    }
}


Function New-HTMLReport
{
    [CmdletBinding()]
    param ([parameter(Mandatory=$true)]
           [string]$Computername,
           [parameter(Mandatory=$true)]
           [string]$AppsHTMLFragment,
           [parameter(Mandatory=$true)]
           [string]$ServicesHTMLFragment,
           [string]$Timestamp = $(get-date -Format "yyyy-MM-dd HH:mm:ss"),
           [string]$ReportOwner = $env:USERNAME,
           [parameter(Mandatory=$true)]
           [string]$OutputPath)

    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="aartemjevas">
    <title>$Computername - Computer Report</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
	<link href="https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css" rel="stylesheet" integrity="sha384-T8Gy5hrqNKT+hzMclPo118YTQO6cYprQmhrYwIiQ/3axmI1hQomh7Ud2hPOy8SP1" crossorigin="anonymous">
    <link href="https://cdn.datatables.net/1.10.11/css/dataTables.bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/responsive/2.1.0/css/responsive.dataTables.min.css" rel="stylesheet">
	<link href="https://fonts.googleapis.com/css?family=Roboto" rel="stylesheet">   
	<style>
		body {
		font-family: 'Roboto', sans-serif;
		}
		.navbar-inverse {
			background:#34495e;
			border:0;
			
		}
		.navbar-brand .a {
			color:#ecf0f1;
		}
		.panel-success .panel-heading{
			background-color:#27ae60 !important;
			border:0 !important;
			color:#ecf0f1;
		}
		.panel-warning .panel-heading{
			background-color:#e74c3c!important;
			border:0 !important;
			color:#ecf0f1;
		}
		.panel-primary .panel-heading{
			background-color:#3498db!important;
			border:0 !important;
			color:#ecf0f1;
			
		}
		.progress {
			height: 50px;
		}
	</style>
  </head>
  <body>
    <nav class="navbar navbar-inverse navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <a class="navbar-brand" href="index.html">Computer Report</a>
        </div>
      </div>
    </nav>
    <div class="container" style="margin-top:70px">
		<div class=row>	
			<div class="col-md-2 col-sm-2">			
			</div>		
			<div class="col-md-8 col-sm-8">
				<center><h2>$Computername</h2></center>
			</div>
			<div class="col-md-2 col-sm-2">			
			</div>
		</div>
		<center><h3>Services</h3></center>
		<hr/>
		<div class="table-responsive">
			<table id="services" class="display nowrap table" cellspacing="0" width="100%">
				<thead>
					<tr>
						<th>Name</th>
						<th>DisplayName</th>
						<th>Status</th>
					</tr>
				</thead>
				<tbody>
					$ServicesHTMLFragment
				</tbody>
			</table>	
		</div>				
		<center><h3>Installed Applications</h3></center>
		<hr/>
		<div class="table-responsive">
			<table id="applications" class="display nowrap table" cellspacing="0" width="100%" style="cursor:pointer;">
				<thead>
					<tr>
						<th>DisplayName</th>
						<th>Version</th>
						<th>Publisher</th>
					</tr>
				</thead>
				<tbody>
					$AppsHTMLFragment
				</tbody>
			</table>	
		</div>				
    </div><!-- /.container -->
	<div class="container" style="padding-top:30px;margin-top:70px">
		<nav class="navbar navbar-inverse navbar-fixed-bottom">
		  <div class="container" style="color:#ecf0f1;margin-top:10px;margin-bot:10px">
			  <div class="row">
					<div class="col-md-3 col-sm-2"></div>
					<div class="col-md-3 col-sm-4 col-xs-6">
						Generated at: $Timestamp
						<br/>
						Generated by: $ReportOwner		
					</div>
					<div class="col-md-2 col-sm-4 col-xs-6">
						<a href="https://github.com/aartemjevas/computerreport"><i class="fa fa-github fa-3x" aria-hidden="true"></i></a>
					</div>
					<div class="col-md-4 col-sm-2"></div>
			  </div>
			  <br/>
		  </div>
		</nav>	
	</div> 
	<script src="http://code.jquery.com/jquery-1.12.4.min.js"   integrity="sha256-ZosEbRLbNQzLpnKIkEdrPv7lOy9C27hHQ+Xp8a4MxAQ="   crossorigin="anonymous"></script>	
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>  
    <script src="https://cdn.datatables.net/1.10.12/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.10.11/js/dataTables.bootstrap.min.js"></script>
    <script src="https://cdn.datatables.net/select/1.2.0/js/dataTables.select.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.1.0/js/dataTables.responsive.min.js"></script>
   <script>
		`$(document).ready( function () {
			`$('#services').DataTable({	"aLengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
			responsive: true
			});
			`$('#applications').DataTable({	"aLengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
			responsive: true
			});	
		} );   
   </script>
  </body>
</html>

"@
      try
      {
        if (!(Test-Path $OutputPath))
        {
            Write-Verbose "Creating $OutputPath"
            $null = mkdir $OutputPath
        }
        $htmlTemplate | Out-File "$OutputPath\$Computername.html" -Encoding utf8 -Force

      }
      catch
      {
        throw $Error[0].Exception
      }  
}

Function New-HomeHTMLReport
{
    [CmdletBinding()]
    param ([parameter(Mandatory=$true)]
           [string]$ComputersHTMLFragment,
           [parameter(Mandatory=$true)]
           [string]$LogName,
           [string]$Timestamp = $(get-date -Format "yyyy-MM-dd HH:mm:ss"),
           [string]$ReportOwner = $env:USERNAME,
           [parameter(Mandatory=$true)]
           [int]$SuccessCount,
           [parameter(Mandatory=$true)]
           [int]$FailCount,
           [parameter(Mandatory=$true)]
           [string]$OutputPath)


$htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">
    <title>Home - Computer Report</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
	<link href="https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css" rel="stylesheet" integrity="sha384-T8Gy5hrqNKT+hzMclPo118YTQO6cYprQmhrYwIiQ/3axmI1hQomh7Ud2hPOy8SP1" crossorigin="anonymous">
    <link href="https://cdn.datatables.net/1.10.11/css/dataTables.bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.datatables.net/responsive/2.1.0/css/responsive.dataTables.min.css" rel="stylesheet">
	<link href="https://fonts.googleapis.com/css?family=Roboto" rel="stylesheet">    
	<style>
		.disabled {
		   pointer-events: none;
		   cursor: default;
		}
		body {
		font-family: 'Roboto', sans-serif;
		}
		.navbar-inverse {
			background:#34495e;
			border:0;
			
		}
		.panel-success .panel-heading{
			background-color:#27ae60 !important;
			border:0 !important;
			color:#ecf0f1;
		}
		.panel-warning .panel-heading{
			background-color:#e74c3c!important;
			border:0 !important;
			color:#ecf0f1;
		}
		.panel-primary .panel-heading{
			background-color:#3498db!important;
			border:0 !important;
			color:#ecf0f1;
			
		}
		.progress {
			height: 50px;
		}
	</style>
  </head>
  <body>
    <nav class="navbar navbar-inverse navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <a class="navbar-brand disabled" href="#">Computer Report</a>				
        </div>
      </div>
    </nav>
    <div class="container" style="margin-top:70px">
		<div class=row>	
			<div class="col-md-2 col-sm-2">			
			</div>		
			<div class="col-md-4 col-sm-4">
				<div class="panel panel-success">
				  <div class="panel-heading">
					<div class="row">
						<div class="col-md-6 col-sm-6 col-xs-6  fa fa-laptop fa-5x"></div>
						<div class="col-md-6 col-sm-6 col-xs-6 "><p style="font-size:40px" class="center-text">$SuccessCount</p></div>
					</div>
				  </div>
				  <div class=panel-body>
					Successfully collected
				  </div>				  
				</div>
			</div>
			<div class="col-md-4 col-sm-4">
				<div class="panel panel-warning">
				  <div class="panel-heading">
					<div class="row">
						<div class="col-md- col-sm-6 col-xs-6 fa fa-laptop fa-5x"></div>
						<div class="col-md-6 col-sm-6 col-xs-6 "><p style="font-size:40px" class="center-text">$FailCount</p></div>
					</div>				  
				  </div>
				  <div class=panel-body>
					Failed to collect
				  </div>				  
				</div>			
			</div>
			<div class="col-md-2 col-sm-2">			
			</div>
		</div>
		<div class="row">
			<div class="col-md-2"></div>
			<div class="col-md-8 col-sm-8 col-xs-12">
			<hr/>
		<div class="table-responsive">
		<table id="computers" class="display nowrap table" cellspacing="0" width="100%">
			<thead>
				<tr>
					<th>Computername</th>
					<th>Timestamp</th>
					<th>Status</th>
				</tr>
			</thead>
			<tbody>
				$ComputersHTMLFragment
			</tbody>
		</table>	
		</div>	
			<div class="pull-right">
			<br/>
				<a href="$LogName" class="btn btn-info" role="button">Open log file</a>
			</div>
			</div>
			<div class="col-md-2"></div>
		</div>
	
    </div><!-- /.container -->
	<div class="container" style="padding-top:30px;margin-top:70px">
    <nav class="navbar navbar-inverse navbar-fixed-bottom">
      <div class="container" style="color:#ecf0f1;margin-top:10px;margin-bot:10px">
	  <div class="row">
		<div class="col-md-3 col-sm-2"></div>
		<div class="col-md-3 col-sm-4 col-xs-6">
		Generated at: $Timestamp<br/>
		Generated by: $ReportOwner		
		</div>
		<div class="col-md-2 col-sm-4 col-xs-6">
			<a href="https://github.com/aartemjevas/computerreport"><i class="fa fa-github fa-3x" aria-hidden="true"></i></a>
		</div>
		<div class="col-md-4 col-sm-2"></div>
	  </div>
	  <br/>
      </div>
    </nav>	
	</div>
	<script src="http://code.jquery.com/jquery-1.12.4.min.js"   integrity="sha256-ZosEbRLbNQzLpnKIkEdrPv7lOy9C27hHQ+Xp8a4MxAQ="   crossorigin="anonymous"></script>	
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>  
    <script src="https://cdn.datatables.net/1.10.12/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.10.11/js/dataTables.bootstrap.min.js"></script>
    <script src="https://cdn.datatables.net/select/1.2.0/js/dataTables.select.min.js"></script>
    <script src="https://cdn.datatables.net/responsive/2.1.0/js/dataTables.responsive.min.js"></script>
   <script>
		`$(document).ready( function () {
			`$('#computers').DataTable({
			"aLengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
			responsive: true
			});
		} );   
   </script>
  </body>
</html>
"@
      try
      {
        if (!(Test-Path $OutputPath))
        {
            Write-Verbose "Creating $OutputPath"
            $null = mkdir $OutputPath
        }
        $htmlTemplate | Out-File "$OutputPath\index.html" -Encoding utf8 -Force

      }
      catch
      {
        throw $Error[0].Exception
      } 
}

Function Get-ComputerHTMLFragment
{
    [CmdletBinding()]
    param ([parameter(Mandatory=$true)]
           [Object]$ComputersObject
)

    $htmlFragment = @()       
    if ($ComputersObject.Status -like "Succeeded")
    {
        $computerCell += "<a href='$($ComputersObject.Computername).html'>$($ComputersObject.Computername)</a>"
    }
    else
    {
        $computerCell += "$($ComputersObject.Computername)"
    }
    $htmlFragment = @"
            <tr>
                <td>$computerCell</td>
                <td>$($ComputersObject.Timestamp)</td>
                <td>$($ComputersObject.Status)</td>
            </tr>
"@
    Write-Output $htmlFragment   
}