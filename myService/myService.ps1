<#
.Synopsis
Get-MyService is an extend mode for Get-Service

. DESCRIPTION
The MyService will help you to get not only the state of the service but
also the Process ID, the start mode and Required / dependencies services.


.EXAMPLE
Get-MyService -ServiceName net -status Running | Format-Table -Autosize

Name       DisplayName                     ProcesSId StartMode Started State  
----       -----------                     --------- --------- ------- -----  
lmhosts    TCP/IP NetBIOS Helper               12248 Manual       True Running
NcbService Network Connection Broker            1464 Manual       True Running
netprofm   Network List Service                 2972 Manual       True Running
NlaSvc     Network Location Awareness           2584 Auto         True Running
nsi        Network Store Interface Service      1840 Auto         True Running

.EXAMPLE
Get-MyService -ServiceName "xbox" -status Stopped | Format-Table -Autosize

Name          DisplayName                       ProcesSId StartMode Started State  
----          -----------                       --------- --------- ------- -----  
XboxGipSvc    Xbox Accessory Management Service         0 Manual      False Stopped
XboxNetApiSvc Xbox Live Networking Service              0 Manual      False Stopped
.EXAMPLE
Get-MyService -DisplayName "Network L" -status "AllStatus" | Format-Table -Autosize

Name     DisplayName                ProcesSId StartMode Started State  
----     -----------                --------- --------- ------- -----  
netprofm Network List Service            2972 Manual       True Running
NlaSvc   Network Location Awareness      2584 Auto         True Running

.EXAMPLE
Get-MyService -ServiceName L -Status AllStatus | ? {$_.startMode -eq 'Disabled'} | Format-Table -AutoSize

Name       DisplayName            ProcesSId StartMode Started State  
----       -----------            --------- --------- ------- -----  
AppVClient Microsoft App-V Client         0 Disabled    False Stopped

.EXAMPLE
Get-MyService -DisplayName search -Status Running | ? {$_.startMode -eq "Auto"}

Name        : WSearch
DisplayName : Windows Search
ProcesSId   : 8256
StartMode   : Auto
Started     : True
State       : Running

REMARKS
    To see the examples, type: "get-help myService -examples".
    For more information, type: "get-help myService -detailed".
    
.NOTES
Author: Badrane DERBAZI
Version: 1.0
Purpose: Extend the Capabilities of the Get-Service Cmdlet
#>

#region functions

# Searching service(s) with a partial name and list Process ID, State, Start Mode 
Function Get-MyService{
  [cmdletBinding(DefaultParameterSetName='ByServiceName')]
 
  param(
  [Parameter(ParameterSetName='ByServiceName')]
  [string]$ServiceName,
    
  [parameter(ParameterSetName='ByDisplayName')]
  [string]$DisplayName,
    
  [Parameter(Mandatory)]
  [Validateset('Running','Stopped','AllStatus')]
  [string]$Status 
  )
     
   if($ServiceName){
     $theName = $ServiceName
     $searchBy = "Name"
   }elseif($DisplayName){
     $theName = $DisplayName
     $searchBy = "DisplayName"
   }else{
     $theName = ""
     $searchBy = "Name"
 
   }
    
  $s = Get-WmiObject -Class Win32_Service -Filter "$SearchBy like '%$theName%'" | Select-Object Name, DisplayName, ProcesdID, StartMode, Started, State
  $r = $s |  Where-Object {$_.State -eq $Status}     
 
   if($status -in ("Running","Stopped")){
   
     if($r){#return Write-Host "Are You sure about the service?" -ForegroundColor Yellow -BackgroundColor Red
      Add-Type -AssemblyName PresentationFramework
      Add-Type -AssemblyName WindowsBase
      return [windows.MessageBox]::Show("Wrong Service's or Display Name. Please repeat...","Get-myService","OK","Error")
     
     }  
     
     if(-not $s.Name -or -not $s.DisplayName){
      return Write-Host "Wrong Service's or Display Name. Please repeat..." -ForegroundColor Yellow -BackgroundColor Red
     
     }else{
     return $r
     }
    
   }elseif($status -like "AllStatus"){
   return $s 
   }
 } 
 
 # List of Services according to their mode of startup 
 Function Get-StartModeMyService{
  [CmdletBinding()]
      
  param(
  [Parameter(Mandatory)]
  #[Validateset('Auto','Manual','Disabled')]
  [string]$StartMode
  )
      
  $r= Get-WmiObject -Class Win32_Service | Select-Object ProcessID, Name, DisplayName, Started, State, StartMode
    
   switch($StartMode){
     default{
     Write-Host "Start Mode Should be : Auto, Disabled or Manual" -ForegroundColor Yellow -BackgroundColor Red
     break;
     }
 
     'Manual'{
      $r | Where-Object {$_.StartMode -eq "Manual"} 
     break; 
     }
 
     'Auto'{
     $r | Where-Object{$_.StartMode -eq "Auto"}
     break; 
     }
 
     'Disabled'{
     $r | Where-Object {$_.StartMode -eq "Disabled"} 
     break; 
     }
   }
 }
 
 
 # Changing StartMode from one state to another state
 Function Switch-StartModeMyService{
 [CmdletBinding()]
      
  param(
  [Parameter(Mandatory)]
   [string]$ServiceName,
 
  [Parameter(Mandatory)]
  [Validateset('Auto','Manual','Disabled')]
  [string]$StartModeNew
  )
 
  $s = get-wmiobject -class win32_service | where-object {$_.name -eq "$ServiceName"} 
  $r = $s.changestartmode("StartModeNew") 
       if($r.returnvalue -eq 0) {Write-Host  "success" -ForegroundColor Green -BackgroundColor White
       } 
       else{"$($r.returnvalue) was reported" 
       } 
  }
 
 # Get-MyService results exported to HTML Format
 Function Export-StartModeToHTML{
 $myFile ="c:\temp\myService.html"
 $myStyle ="<div align='center'>"
 $myStyle = $myStyle + "<style>BODY{background-color:#eaeaea;}"
 $myStyle = $myStyle + "TABLE{align:center; border: 1px solid black; border-collapse: collapse;}"
 $myStyle = $myStyle + "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
 $mystyle = $myStyle + "TD{border: 1px solid black; padding: 5px;}"
 $mystyle = $myStyle + "</style>"
 $Output = Get-MyService
 $Output | ConvertTo-Html -head $myStyle -body "<H2>Services Status List</H2>"| Out-File $myDestination$myFile
 Invoke-Expression $myDestination$myFile  
 }
    
 Function Export-StartModeToExcel
 {
 # New feature 3  
 }
 
 # List all the running services with Process ID and State
 Function Get-AllMyRunningService{
  $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Running'" 
  $r = ($s |  Select-Object ProcessID, Name, DisplayName, Started, State) 
 return $r
 }
 
 # Display the list of started services
 Function Get-AllMyStoppedService{
  $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Stopped'"
  $r = ($s |  Select-Object Name, DisplayName, Started, State )
 return $r
 }
 
 # Display the list of started services
 Function Get-MyServiceProcessID{
  $s = Get-WmiObject -Class Win32_Service | Where-Object {$_.ProcessID -ne 0}
  $r = ($s |Select-Object ProcessID, Started, StartMode, State, Name, DisplayName | Sort-Object State, Name)
  
 return $r
 }
 
 # Stop/Check a service
 Function Stop-MyService{
  [CmdletBinding()]
 
  param(
  [Parameter(Mandatory)]
  [string]$ServiceName
  )
     
  $s= Get-WmiObject -Class Win32_Service -Filter "Name = '$ServiceName'"
  
   if($s.State -eq "Stopped"){
     Write-Host "The Service $ServiceName already stopped, no need to stop it" -ForegroundColor White -BackgroundColor Green 
     Break
   }
   else{
     $s | Stop-Service 
     Write-Host "The Service $ServiceName stopped" -ForegroundColor Yellow -BackgroundColor Red
   Break
   }
 
 return $r
 }
  
 # Start/Check a service
 Function Start-MyService{
  [CmdletBinding()]
 
  param(
  [Parameter(Mandatory)]
  [string]$ServiceName
  )
     
  $s= Get-WmiObject -Class Win32_Service -Filter "Name = '$ServiceName'"
 
   if($s.State -eq "Running"){
     Write-Host "The Service $ServiceName already started, no need to start it" -ForegroundColor Yellow -BackgroundColor Red
   Break
   }
   else{
     $s | Start-Service 
     Write-Host "The Service $ServiceName started" -ForegroundColor White -BackgroundColor Green
   Break
   }
 return $r
 }
 
 # List of the service or services in which a particular service is required to it/them to run
 Function Get-MyRequiredService{
  [CmdletBinding()]
  
  param(
  [Parameter(Mandatory)]
  [string]$ServiceName
  )
 
  Get-Service $ServiceName -RequiredServices
 }
 
 # get the Service or Services in which a particular service depends to it/them to run
 Function Get-myDependentService{
  [CmdletBinding()]
  
  param(
  [Parameter(Mandatory)]
  [string]$ServiceName
   )
  
  Get-Service $ServiceName -DependentServices
 }
 #endregion