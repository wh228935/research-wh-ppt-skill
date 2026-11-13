<# .SYNOPSIS Initial install of a verified skill folder; refuses existing targets. #>
param([Parameter(Mandatory=$true)][string]$DestinationRoot)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'PowerShell 7 required'}
$source=[IO.Path]::GetFullPath((Split-Path $PSScriptRoot))
$destRoot=[IO.Path]::GetFullPath($DestinationRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
$dest=[IO.Path]::GetFullPath((Join-Path $destRoot 'research-wh-ppt-skil'))
if(-not $dest.StartsWith($destRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Destination outside explicit root'}
if($source -eq $dest -or $dest.StartsWith($source+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Cannot install into source package'}
if(Test-Path -LiteralPath $dest){throw 'Target exists; inspect version and request an explicit update plan. No files replaced.'}
$version=Get-Content -LiteralPath (Join-Path $source 'VERSION.json') -Raw | ConvertFrom-Json
if($version.status -ne 'stable'){throw 'Development package is not an approved stable version; initial install is deferred'}
if(-not (Test-Path -LiteralPath (Join-Path $source 'RELEASE-APPROVAL.json'))){throw 'Missing stable release approval record'}
$approval=Get-Content -LiteralPath (Join-Path $source 'RELEASE-APPROVAL.json') -Raw | ConvertFrom-Json
if(-not $approval.user_quote -or $approval.version -ne $version.version){throw 'Release confirmation/version mismatch'}
if(-not $approval.files -or $approval.files.Count -eq 0){throw 'Release approval must bind package files and hashes'}
$files=@(Get-ChildItem -LiteralPath $source -Recurse -File -Force | Where-Object {
 $rel=[IO.Path]::GetRelativePath($source,$_.FullName)
 $parts=$rel -split '[\\/]'
 -not @($parts | Where-Object {$_ -in @('.git','__pycache__','.venv','work','local','outputs','validation-private')}).Count -and $_.Extension -notin @('.pyc','.log') -and $rel -ne 'RELEASE-APPROVAL.json'
})
if($files.Count -ne $approval.files.Count){throw 'Approval does not match the complete package file set'}
foreach($f in $files){
 $rel=[IO.Path]::GetRelativePath($source,$f.FullName).Replace('\','/')
 $entry=@($approval.files | Where-Object {$_.path -eq $rel})
 if($entry.Count -ne 1 -or (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash.ToLower() -ne $entry[0].sha256){throw "Unapproved package file: $rel"}
}
New-Item -ItemType Directory -Force -Path $destRoot | Out-Null
New-Item -ItemType Directory -Path $dest | Out-Null
foreach($f in $files){
 $target=Join-Path $dest ([IO.Path]::GetRelativePath($source,$f.FullName))
 New-Item -ItemType Directory -Force -Path (Split-Path $target) | Out-Null
 Copy-Item -LiteralPath $f.FullName -Destination $target
}
Copy-Item -LiteralPath (Join-Path $source 'RELEASE-APPROVAL.json') -Destination $dest
Write-Output $dest
