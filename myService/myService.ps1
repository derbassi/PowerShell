<#
.Synopsis
Get-MyService is an extend mode for Get-Service

. DESCRIPTION
The Get-MyService will help you to get not only the state of the service but
also the Process ID and the start mode


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
#>
function Get-MyService{
    [cmdletBinding(DefaultParameterSetName='ByServiceName')]
    
    param(
     [Parameter(Mandatory, ParameterSetName='ByServiceName')]
     [string]$ServiceName,
   
     [parameter(Mandatory, ParameterSetName='ByDisplayName')]
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
     }
   
     $s = Get-WmiObject -Class Win32_Service -Filter "$SearchBy like '%$theName%'" | Select-Object Name, DisplayName, ProcesSId, StartMode, Started, State
     
     if($status -in ("Running","Stopped")){
       $r = $s |  Where-Object {$_.State -eq $Status} 
       
       if(-not $s.Name -or -not $s.DisplayName){
         return Write-Host "Wrong Service's or Display Name. Please repeat..." -ForegroundColor Yellow -BackgroundColor Red
       }else{
          return $r
       }
   
     }elseif($status -like "AllStatus"){
        return $s 
     }
   } 
   
   function Get-StartModeMyService{
   [CmdletBinding()]
     param(
      [Parameter(Mandatory)]
      #[Validateset('Auto','Manual','Disabled')]
      [string]$StartMode
     )
     
     $r= Get-WmiObject -Class Win32_Service | Select-Object ProcesSId, Name, DisplayName, Started, State, StartMode | Sort-Object Name
   
     switch($startMode){
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
   
   function GetMyServiceToHTML{
   
   }
   
   function Get-MyServiceExcel{
   
   }
   
   function Get-AllMyRunningService{
       $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Running'" | Sort-Object State, Name
       $r = ($s |  Select-Object ProcesSId, Name, DisplayName, Started, State) 
   return $r
   }
   
   function Get-AllMyStoppedService{
     $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Stopped'"| Sort-Object State, Name
     $r = ($s |  Select-Object Name, DisplayName, Started, State )
   return $r
   }
   
   function Get-MyServiceProcessID{
     $s = Get-WmiObject -Class Win32_Service | Sort-Object State, Name
     $r = ($s |  Select-Object ProcesSId, Started, StartMode, State, Name, DisplayName) | Sort-Object Name 
   return $r
   }
   
   function Stop-MyService{
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
   
   function Start-MyService{
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
   
   
   
   #$c=Get-Command Get-MyService
   #$c.DefaultParameterSet
   #$c.ParameterSets  | Select Name
   
   #Get-AllMyStoppedService | ft -AutoSize -Wrap