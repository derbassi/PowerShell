Function Get-theservice{
    [CmdletBinding()]
    param(
    [string]$serviceName
    )
        $s = Get-WmiObject -Class Win32_Service -Filter "Name='$serviceName'"
        $r = ($s |  Select-Object Name, DisplayName, ProcesSId, Started, State ) | Format-Table -AutoSize -Wrap
        
        if(!($s.Name)){
            return Write-Host "Wrong Service's Name or unavailble..." -ForegroundColor Yellow -BackgroundColor Red
        }
        else{
            return $r
        }
    }
    
    
    Function Get-AllTheService{
        $s = Get-WmiObject -Class Win32_Service | Sort-Object State, Name
        $r = ($s |  Select-Object Name, DisplayName, ProcesSId, Started, State ) | Format-Table -AutoSize -Wrap
    return $r
    }
    
    
    Function Get-AllRunningService{
        $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Running'"| Sort-Object State, Name
        $r = ($s |  Select-Object Name, DisplayName, ProcesSId, Started, State ) | Format-Table -AutoSize -Wrap
    return $r
    }
    
    Function Get-AllStoppedService{
        $s = Get-WmiObject -Class Win32_Service -Filter "State = 'Stopped'"| Sort-Object State, Name
        $r = ($s |  Select-Object Name, DisplayName, ProcesSId, Started, State ) | Format-Table -AutoSize -Wrap
    return $r
    }