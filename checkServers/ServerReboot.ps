$servers = GC "C:\temp\test.txt"

# Rebbot Pending State
$PR = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"

# Rebbot Required State
$RR = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"

# Session Manager
$SessionManager ="HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"

function Test-PendingReboot([string]$Server)
{
  if(invoke-command -scriptblock {Get-ChildItem $RP -EA SilentlyContinue} -computername $Server)
  {$1=$true
  }

  if(invoke-command -scriptblock {Get-Item $RR -EA SilentlyContinue} -computername $Server){
      $2=$true
      }

  if(invoke-command -scriptblock {Get-ItemProperty $SessionManager -Name PendingFileRenameOperations -EA SilentlyContinue} -computername $Server){
      $3=$true
      }

  if($1 -or $2 -or $3){
      $RebootRequired=$true
      }
else{
      $RebootRequired=$false
    }

[pscustomobject]@{
 ServerName=$Server;
 RebootRequired=$RebootRequired
}
}

$tmp=foreach ($srv in $servers){
Test-PendingReboot -Server $srv
}

$tmp  | Out-File C:\temp\test2.txt  