# Shared Office session ownership. No process-name killing and no user documents closed.
Set-StrictMode -Version Latest
function Release-OfficeObject($Value) {
    if($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)}catch [Runtime.InteropServices.InvalidComObjectException]{}
    }
}
function New-PresentationSession([string]$Application) {
    if($PSVersionTable.PSVersion.Major -lt 7){throw 'PowerShell 7 required; do not fall back to Windows PowerShell 5.1'}
    $before=@(Get-Process -ErrorAction SilentlyContinue | Where-Object {$_.ProcessName -in @('POWERPNT','wpp','wps')} | ForEach-Object Id)
    $prog=if($Application -eq 'WPS'){'KWPP.Application'}else{'PowerPoint.Application'}
    $app=New-Object -ComObject $prog
    $version=[string]$app.Version
    $initial=[int]$app.Presentations.Count
    $new=@(Get-Process -ErrorAction SilentlyContinue | Where-Object {$_.ProcessName -in @('POWERPNT','wpp','wps') -and $before -notcontains $_.Id})
    $owned=$new.Count -eq 1
    $ownedId=0;$started=$null
    if($owned){$ownedId=$new[0].Id;$started=$new[0].StartTime.ToUniversalTime().ToString('o')}
    return @{App=$app;Application=$Application;Version=$version;InitialCount=$initial;Owned=$owned;ProcessId=$ownedId;StartTime=$started}
}
function Close-PresentationSession($Session,$Presentation) {
    $closed=$false;$quit=$false;$cleanupError=$null
    if($null -ne $Presentation){
        try{$Presentation.Close();$closed=$true}catch{$cleanupError=$_.Exception.Message}
        Release-OfficeObject $Presentation
    }
    if($null -ne $Session -and $null -ne $Session.App){
        try{
            if($Session.Owned -and $Session.InitialCount -eq 0 -and $Session.App.Presentations.Count -eq 0){$Session.App.Quit();$quit=$true}
        }catch{$cleanupError=$_.Exception.Message}
        Release-OfficeObject $Session.App
        $Session.App=$null
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    $exited=$null
    if($null -ne $Session -and $Session.Owned){
        $proc=Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue
        if($null -ne $proc -and $proc.StartTime.ToUniversalTime().ToString('o') -eq $Session.StartTime){
            try{[void]$proc.WaitForExit(5000)}catch{}
        }
        $remaining=Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue
        $exited=$null -eq $remaining -or $remaining.StartTime.ToUniversalTime().ToString('o') -ne $Session.StartTime
    }
    return @{task_document_closed=$closed;quit_sent=$quit;owned_process_exited=$exited;cleanup_error=$cleanupError;owned_pid=$(if($null -ne $Session){$Session.ProcessId}else{0});owned_start_time=$(if($null -ne $Session){$Session.StartTime}else{$null});ownership_verified=$(if($null -ne $Session){$Session.Owned}else{$false})}
}
function Confirm-OfficeWorkerReport([string]$ReportPath) {
    $r=Get-Content -LiteralPath $ReportPath -Raw | ConvertFrom-Json -AsHashtable
    if($r.cleanup.ownership_verified){
        $proc=Get-Process -Id $r.cleanup.owned_pid -ErrorAction SilentlyContinue
        if($null -ne $proc -and $proc.StartTime.ToUniversalTime().ToString('o') -eq $r.cleanup.owned_start_time){
            try{[void]$proc.WaitForExit(5000)}catch{}
        }
        $proc=Get-Process -Id $r.cleanup.owned_pid -ErrorAction SilentlyContinue
        $r.cleanup.owned_process_exited=$null -eq $proc -or $proc.StartTime.ToUniversalTime().ToString('o') -ne $r.cleanup.owned_start_time
        $r.cleanup.verified_after_worker_exit=$true
    }
    $r | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $ReportPath -Encoding utf8
    return $r
}
