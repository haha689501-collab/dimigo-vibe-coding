# 새로 만든 3차시 v2 덱만 PNG 로 렌더링해 확인합니다.
# 안전 규칙: 원본은 복사본만 열고, 이미 켜져 있던 PowerPoint 는 절대 종료하지 않습니다.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root '01_수업PPT\지난버전\디미고_바이브코딩_3차시_PRD를화면으로_v2.pptx'
$work = Join-Path $PSScriptRoot '_work2'
$dest = Join-Path $PSScriptRoot '_preview_l3v2'

if (-not (Test-Path $src)) { Write-Host "원본 없음: $src"; exit 1 }

$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)

if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path $work -Force | Out-Null
$copy = Join-Path $work 'l3v2.pptx'
Copy-Item -LiteralPath $src -Destination $copy -Force

if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory -Path $dest -Force | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($copy, -1, 0, 0)
  try {
    $pres.Export($dest, 'PNG', 1600, 900)
    Write-Host ("렌더링 완료 - 슬라이드 {0}장" -f $pres.Slides.Count)
  } finally { $pres.Close() }
} finally {
  if (-not $wasRunning) { $ppt.Quit() }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
Remove-Item $work -Recurse -Force
Write-Host "-> $dest"
