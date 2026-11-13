<#
.SYNOPSIS
Build a native editable PPTX from an approved object plan. New/prototype decks only.
No cropping, no external image links, no arbitrary script execution from plan data.
#>
param(
 [Parameter(Mandatory=$true)][string]$PlanPath,
 [Parameter(Mandatory=$true)][string]$WorkDir,
 [Parameter(Mandatory=$true)][string]$OutputPath,
 [ValidateSet('prototype','full')][string]$Stage='prototype',
 [ValidateSet('PowerPoint','WPS')][string]$Application='PowerPoint',
 [string]$Python='python',
 [switch]$OfficeWorker
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Office-Com.ps1')
if(-not $OfficeWorker){
 $exe=(Get-Process -Id $PID).Path
 $out=[IO.Path]::GetFullPath($OutputPath)
 $childOutput=& $exe -NoProfile -File $PSCommandPath -PlanPath $PlanPath -WorkDir $WorkDir -OutputPath $out -Stage $Stage -Application $Application -Python $Python -OfficeWorker 2>&1
 $childCode=$LASTEXITCODE
 if(Test-Path -LiteralPath ($out+'.build.json')){
  $done=Confirm-OfficeWorkerReport ($out+'.build.json');$done | ConvertTo-Json -Depth 12
  if($done.cleanup.ownership_verified -and -not $done.cleanup.owned_process_exited){exit 4}
 }
 if($childCode -ne 0){$childOutput | Write-Output}
 exit $childCode
}
function Color-Rgb([string]$Hex){
 $h=$Hex.TrimStart('#');if($h -notmatch '^[0-9A-Fa-f]{6}$'){throw "Invalid color: $Hex"}
 return [Convert]::ToInt32($h.Substring(0,2),16)+256*[Convert]::ToInt32($h.Substring(2,2),16)+65536*[Convert]::ToInt32($h.Substring(4,2),16)
}
function Full-Path([string]$Value){if([IO.Path]::IsPathRooted($Value)){return [IO.Path]::GetFullPath($Value)};return [IO.Path]::GetFullPath((Join-Path $script:PlanDir $Value))}
function Set-ShapeText($Shape,[string]$Text,[double]$Size,[string]$Color,[bool]$Bold=$false,[bool]$TableCell=$false){
 if(-not $TableCell){$Shape.TextFrame.AutoSize=0;$Shape.TextFrame.WordWrap=-1}
 $Shape.TextFrame.MarginLeft=0;$Shape.TextFrame.MarginRight=0;$Shape.TextFrame.MarginTop=0;$Shape.TextFrame.MarginBottom=0
 $r=$Shape.TextFrame.TextRange;$r.Text=$Text
 $r.Font.Name=$script:Style.font
 try{$r.Font.NameFarEast=$script:Style.font}catch{}
 $r.Font.Size=$Size;$r.Font.Bold=if($Bold){-1}else{0};$r.Font.Color.RGB=Color-Rgb $Color
 $r.ParagraphFormat.Alignment=1;$r.ParagraphFormat.SpaceBefore=0;$r.ParagraphFormat.SpaceAfter=5
 Release-OfficeObject $r
}
function New-Text($Slide,[string]$Name,[string]$Text,[double[]]$Box,[double]$Size,[string]$Color,[bool]$Bold=$false){
 $s=$Slide.Shapes.AddTextbox(1,$Box[0],$Box[1],$Box[2],$Box[3]);$s.Name=$Name
 Set-ShapeText $s $Text $Size $Color $Bold
 return $s
}
function Add-PlanObject($Slide,$Obj,$Assets,$Shapes){
 $b=@($Obj.box);if($b.Count -ne 4){throw 'Every object requires box:[x,y,width,height] in points'}
 $name=[string]$Obj.id;if(-not $name){throw 'Every object requires a stable id'}
 if($Shapes.ContainsKey($name)){throw "Duplicate object id: $name"}
 $s=$null
 switch($Obj.type){
  'text' {
   $size=if($Obj.Contains('font_size')){[double]$Obj['font_size']}else{[double]$script:Style.body_font_pt}
   $color=if($Obj.Contains('color')){[string]$Obj['color']}else{[string]$script:Style.text}
   $s=New-Text $Slide $name ([string]$Obj.text) $b $size $color ([bool]$Obj['bold'])
   if($Obj['align'] -eq 'center'){$s.TextFrame.TextRange.ParagraphFormat.Alignment=2}
   if($Obj['align'] -eq 'right'){$s.TextFrame.TextRange.ParagraphFormat.Alignment=3}
   foreach($hl in @($Obj['highlights'])){
    if($null -eq $hl){continue};$needle=[string]$hl.text;if(-not $needle){continue};$at=0
    while(($at=([string]$Obj.text).IndexOf($needle,$at,[StringComparison]::Ordinal)) -ge 0){
     $run=$s.TextFrame.TextRange.Characters($at+1,$needle.Length)
     $run.Font.Color.RGB=Color-Rgb ([string]$hl.color);$run.Font.Bold=-1;Release-OfficeObject $run;$at+=$needle.Length
    }
   }
  }
  'rect' {
   $kind=if($Obj['rounded']){5}else{1};$s=$Slide.Shapes.AddShape($kind,$b[0],$b[1],$b[2],$b[3])
   $s.Name=$name;$s.Fill.Solid();$s.Fill.ForeColor.RGB=Color-Rgb $Obj.fill
   if($Obj['line']){$s.Line.ForeColor.RGB=Color-Rgb $Obj['line'];$s.Line.Weight=0.6}else{$s.Line.Visible=0}
  }
  'image' {
   $asset=$Assets[[string]$Obj.asset_id];$imagePath=Full-Path $asset.path
   $s=$Slide.Shapes.AddPicture($imagePath,0,-1,0,0,-1,-1);$s.Name=$name
   $scale=[math]::Min($b[2]/$s.Width,$b[3]/$s.Height)
   $iw=$s.Width*$scale;$ih=$s.Height*$scale;$s.LockAspectRatio=0
   $s.Width=$iw;$s.Height=$ih;$s.Left=$b[0]+($b[2]-$iw)/2;$s.Top=$b[1]+($b[3]-$ih)/2;$s.LockAspectRatio=-1
  }
  'table' {
   $rows=@($Obj.rows);$nc=@($rows[0]).Count
   if($nc -eq 0){throw 'Empty table'}
   foreach($row in $rows){if(@($row).Count -ne $nc){throw 'Unequal table column count'}}
   $s=$Slide.Shapes.AddTable($rows.Count,$nc,$b[0],$b[1],$b[2],$b[3]);$s.Name=$name
   $size=if($Obj['font_size']){[double]$Obj['font_size']}else{[double]$script:Style.body_font_pt}
   for($ri=1;$ri -le $rows.Count;$ri++){
    for($ci=1;$ci -le $nc;$ci++){
     $cell=$s.Table.Cell($ri,$ci).Shape
     Set-ShapeText $cell ([string]$rows[$ri-1][$ci-1]) $size $script:Style.text ($ri -eq 1) $true
     $cell.TextFrame.MarginLeft=5;$cell.TextFrame.MarginTop=3
     $cell.Fill.Solid();$cell.Fill.ForeColor.RGB=Color-Rgb $(if($ri -eq 1){$script:Style.pale}else{'FFFFFF'})
     Release-OfficeObject $cell
    }
   }
  }
  'line' {
   $s=$Slide.Shapes.AddConnector(1,$b[0],$b[1],$b[0]+$b[2],$b[1]+$b[3]);$s.Name=$name
   $s.Line.ForeColor.RGB=Color-Rgb $(if($Obj['color']){$Obj['color']}else{$script:Style.primary})
   $s.Line.Weight=1.2;if($Obj['arrow']){$s.Line.EndArrowheadStyle=2}
   if($Obj['from'] -and $Obj['to']){
    $fromSite=if($Obj['from_site']){[int]$Obj['from_site']}else{4}
    $toSite=if($Obj['to_site']){[int]$Obj['to_site']}else{2}
    $s.ConnectorFormat.BeginConnect($Shapes[[string]$Obj['from']],$fromSite)
    $s.ConnectorFormat.EndConnect($Shapes[[string]$Obj['to']],$toSite)
   }
  }
  default {throw "Unsupported native object type: $($Obj.type)"}
 }
 $Shapes[$name]=$s
}
$planFull=(Resolve-Path -LiteralPath $PlanPath).Path
$script:PlanDir=Split-Path $planFull
$outputFull=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $outputFull){throw 'Output exists; use a new draft filename'}
& $Python (Join-Path $PSScriptRoot 'project.py') gate --workdir $WorkDir --plan $planFull --stage $Stage
if($LASTEXITCODE -ne 0){throw 'Approval/content gate failed; no presentation created'}
$plan=Get-Content -LiteralPath $planFull -Raw | ConvertFrom-Json -AsHashtable
$script:Style=Get-Content -LiteralPath (Full-Path $plan.style_file) -Raw | ConvertFrom-Json -AsHashtable
$assets=@{};foreach($a in @($plan.assets)){if($null -ne $a){$assets[[string]$a.id]=$a}}
$mutex=[Threading.Mutex]::new($false,'Local\ResearchWhPptOffice')
if(-not $mutex.WaitOne(0)){$mutex.Dispose();throw 'Another presentation operation is running'}
$session=$null;$pres=$null;$cleanup=$null
$report=[ordered]@{application=$Application;stage=$Stage;output=$outputFull;source_plan_sha256=(Get-FileHash -LiteralPath $planFull -Algorithm SHA256).Hash.ToLower();text_overflows=@()}
try{
 $session=New-PresentationSession $Application;$report.application_version=$session.Version
 $pres=$session.App.Presentations.Add(0)
 $pres.PageSetup.SlideWidth=[double]$script:Style.canvas_pt[0];$pres.PageSetup.SlideHeight=[double]$script:Style.canvas_pt[1]
 foreach($page in $plan.pages){
  $slide=$pres.Slides.Add($pres.Slides.Count+1,12);$slide.Name=[string]$page.id
  $slide.FollowMasterBackground=0;$slide.Background.Fill.Solid();$slide.Background.Fill.ForeColor.RGB=Color-Rgb $script:Style.background
  $shapes=@{}
  if($page.header -or $page.title){
   $band=$slide.Shapes.AddShape(1,0,0,$script:Style.canvas_pt[0],$script:Style.header_height_pt)
   $band.Fill.Solid();$band.Fill.ForeColor.RGB=Color-Rgb $script:Style.primary;$band.Line.Visible=0;Release-OfficeObject $band
   $mode=if($page.Contains('header_mode')){$page.header_mode}elseif($script:Style.Contains('header_mode')){$script:Style.header_mode}else{'section'}
   if($mode -notin @('section','topic')){throw 'header_mode must be section or topic'}
   if($mode -eq 'topic' -or -not $page.header){
    $t=New-Text $slide 'PageTitle' $page.title @(24,8,900,40) $script:Style.header_font_pt 'FFFFFF' $true;Release-OfficeObject $t
   }else{
    $t=New-Text $slide 'SectionHeader' $page.header @(24,8,900,40) $script:Style.header_font_pt 'FFFFFF' $true;Release-OfficeObject $t
    $t=New-Text $slide 'PageTitle' $page.title @(24,68,900,36) $script:Style.header_font_pt $script:Style.primary $true;Release-OfficeObject $t
   }
  }
  foreach($obj in $page.objects){Add-PlanObject $slide $obj $assets $shapes}
  $notesBody=$null
  foreach($shape in $slide.NotesPage.Shapes){
   try{if($shape.PlaceholderFormat.Type -eq 2){$notesBody=$shape;break}}catch{}
   Release-OfficeObject $shape
  }
  if($null -eq $notesBody){throw 'Notes body placeholder unavailable; cannot claim notes were added'}
  $notesBody.TextFrame.TextRange.Text=[string]$page.notes;Release-OfficeObject $notesBody
  foreach($shape in $slide.Shapes){
   if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
    try{if($shape.TextFrame2.TextRange.BoundHeight -gt $shape.Height+3){$report.text_overflows+=@{page=$page.id;shape=$shape.Name}}}catch{}
   }
   Release-OfficeObject $shape
  }
  foreach($shape in $shapes.Values){Release-OfficeObject $shape};$shapes.Clear()
  Release-OfficeObject $slide
 }
 New-Item -ItemType Directory -Force -Path (Split-Path $outputFull) | Out-Null
 $pres.SaveAs($outputFull,24)
 $report.output_sha256=(Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash.ToLower()
 $report.slide_count=$pres.Slides.Count
}catch{$report.error=$_.Exception.Message;throw}finally{
 $cleanup=Close-PresentationSession $session $pres;$pres=$null;$session=$null
 $report.cleanup=$cleanup
 $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath ($outputFull+'.build.json') -Encoding utf8
 $mutex.ReleaseMutex();$mutex.Dispose()
}
$report | ConvertTo-Json -Depth 10
if($report.text_overflows.Count -gt 0){exit 3}
