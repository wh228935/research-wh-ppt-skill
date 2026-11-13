<# .SYNOPSIS Read-only PowerPoint/WPS render with source hash and process ownership. #>
param(
 [Parameter(Mandatory=$true)][string]$Source,
 [Parameter(Mandatory=$true)][string]$OutputDir,
 [ValidateSet('PowerPoint','WPS')][string]$Application='PowerPoint',
 [int]$Width=1600,[int]$Height=900,[switch]$OfficeWorker
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Office-Com.ps1')
if(-not $OfficeWorker){
 $exe=(Get-Process -Id $PID).Path
 $destPath=[IO.Path]::GetFullPath($OutputDir)
 $childOutput=& $exe -NoProfile -File $PSCommandPath -Source $Source -OutputDir $destPath -Application $Application -Width $Width -Height $Height -OfficeWorker 2>&1
 $childCode=$LASTEXITCODE
 $reportFile=Join-Path $destPath 'render-report.json'
 if(Test-Path -LiteralPath $reportFile){
  $done=Confirm-OfficeWorkerReport $reportFile;$done | ConvertTo-Json -Depth 12
  if($done.cleanup.ownership_verified -and -not $done.cleanup.owned_process_exited){exit 4}
 }
 if($childCode -ne 0){$childOutput | Write-Output}
 exit $childCode
}
$sourceFull=(Resolve-Path -LiteralPath $Source).Path
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
 for($i=1;$i -le $pres.Slides.Count;$i++){
  $slide=$pres.Slides.Item($i);$file=('slide-{0:D3}.png' -f $i);$target=Join-Path $dest $file
  $slide.Export($target,'PNG',$Width,$Height)
  $r.slides+=@{page=$i;file=$file;sha256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLower()}
  Release-OfficeObject $slide
 }
 if((Get-FileHash -LiteralPath $sourceFull -Algorithm SHA256).Hash.ToLower() -ne $sourceHash){throw 'Source changed during read-only rendering'}
}catch{$r.error=$_.Exception.Message}finally{
 $r.cleanup=Close-PresentationSession $session $pres;$session=$null;$pres=$null
 $r | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $dest 'render-report.json') -Encoding utf8
 $mutex.ReleaseMutex();$mutex.Dispose()
}
$r | ConvertTo-Json -Depth 10
if($r.Contains('error')){exit 1}
