<# .SYNOPSIS Read-only PowerPoint render; reuse unchanged slides and inspect only requested pages. #>
param(
 [Parameter(Mandatory=$true)][string]$Source,
 [Parameter(Mandatory=$true)][string]$OutputDir,
 [ValidateSet('PowerPoint')][string]$Application='PowerPoint',
 [int]$Width=1600,[int]$Height=900,[int[]]$Pages=@(),
 [string]$CacheDir,[string]$Python='python',[string]$PagesCsv,[switch]$Detailed,[switch]$OfficeWorker
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Office-Com.ps1')
if(-not $OfficeWorker){
 $exe=(Get-Process -Id $PID).Path
 $destPath=[IO.Path]::GetFullPath($OutputDir)
 if(-not $CacheDir){$CacheDir=Join-Path (Split-Path $destPath) '.powerpoint-preview-cache'}
 $argsList=@('-NoProfile','-File',$PSCommandPath,'-Source',$Source,'-OutputDir',$destPath,'-Application',$Application,'-Width',$Width,'-Height',$Height,'-CacheDir',[IO.Path]::GetFullPath($CacheDir),'-Python',$Python,'-OfficeWorker')
 if($Pages.Count){$argsList+=@('-PagesCsv',($Pages -join ','))}
 $childOutput=& $exe @argsList 2>&1
 $childCode=$LASTEXITCODE
 $reportFile=Join-Path $destPath 'render-report.json'
 if(Test-Path -LiteralPath $reportFile){
  $done=Confirm-OfficeWorkerReport $reportFile
  if($Detailed){$done | ConvertTo-Json -Depth 12}else{
   [ordered]@{application=$done.application;opened=$done.opened;source_sha256=$done.source_sha256;report_path=$reportFile;exported_pages=$done['exported_pages'];reused_pages=$done['reused_pages'];complete_deck=$done['complete_deck'];owned_process_exited=$done.cleanup.owned_process_exited;error=$done['error']} | ConvertTo-Json -Depth 4 -Compress
  }
  if($done.cleanup.ownership_verified -and -not $done.cleanup.owned_process_exited){exit 4}
 }
 if($childCode -ne 0){$childOutput | Write-Output}
 exit $childCode
}
$sourceFull=(Resolve-Path -LiteralPath $Source).Path
if($PagesCsv){$Pages=@($PagesCsv.Split(',') | ForEach-Object {[int]$_})}
$dest=[IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $dest){throw 'Preview directory exists; use a new directory to avoid stale files'}
New-Item -ItemType Directory -Path $dest -Force | Out-Null
$sourceHash=(Get-FileHash -LiteralPath $sourceFull -Algorithm SHA256).Hash.ToLower()
$r=[ordered]@{application=$Application;source=$sourceFull;source_sha256=$sourceHash;opened=$false;slides=@()}
$mutex=[Threading.Mutex]::new($false,'Local\ResearchWhPptOffice')
if(-not $mutex.WaitOne(0)){$mutex.Dispose();throw 'Another presentation operation is running'}
$session=$null;$pres=$null
try{
 $session=New-PresentationSession $Application;$r.application_version=$session.Version
 $pres=$session.App.Presentations.Open($sourceFull,-1,0,0);$r.opened=$true
 $actualHeight=[int][math]::Round($Width*$pres.PageSetup.SlideHeight/$pres.PageSetup.SlideWidth)
 if($Height -ne $actualHeight){$Height=$actualHeight}
 $cacheArgs=@((Join-Path $PSScriptRoot 'render_cache.py'),'plan','--source',$sourceFull,'--output',$dest,'--cache',$CacheDir,'--version',$session.Version,'--width',$Width,'--height',$Height)
 if($Pages.Count){$cacheArgs+=@('--pages')+$Pages}
 & $Python @cacheArgs | Out-Null
 if($LASTEXITCODE -ne 0){throw 'Preview cache planning failed'}
 $plan=Get-Content -LiteralPath (Join-Path $dest 'cache-plan.json') -Raw | ConvertFrom-Json -AsHashtable
 $r.cache_dir=$plan.cache_dir;$r.renderer=$plan.renderer;$r.slide_count=$plan.slide_count;$r.complete_deck=$plan.complete_deck
 foreach($job in $plan.jobs){
  $i=[int]$job.page
  $slide=$pres.Slides.Item($i);$file=('slide-{0:D3}.png' -f $i);$target=Join-Path $dest $file
  if(-not $job.cache_hit){$slide.Export($target,'PNG',$Width,$Height)}
  $job.sha256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLower()
  $r.slides+=,$job
  Release-OfficeObject $slide
 }
 if((Get-FileHash -LiteralPath $sourceFull -Algorithm SHA256).Hash.ToLower() -ne $sourceHash){throw 'Source changed during read-only rendering'}
}catch{$r.error=$_.Exception.Message}finally{
 $r.cleanup=Close-PresentationSession $session $pres;$session=$null;$pres=$null
 $r | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $dest 'render-report.json') -Encoding utf8
 $mutex.ReleaseMutex();$mutex.Dispose()
}
if($r.Contains('error')){exit 1}
& $Python (Join-Path $PSScriptRoot 'render_cache.py') complete --report (Join-Path $dest 'render-report.json')
if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}
