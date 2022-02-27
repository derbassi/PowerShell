$Processes = Get-Process
$RequiredProcesses = "Chrome", "Outloock", "Firefox"
foreach($p in $RequiredProcesses){
if($processes.ProcessName -notcontains $p){
    Write-Host "$p not running " -BackgroundColor DarkGreen -ForegroundColor Yellow
    
 }else{Write-Host "$p is already running." -BackgroundColor DarkCyan -ForegroundColor White
}
}