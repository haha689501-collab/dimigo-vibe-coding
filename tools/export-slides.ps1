# =====================================================================
#  PPTX -> PNG 슬라이드 + PDF 내보내기  (Microsoft PowerPoint COM 사용)
#
#  사용법:  powershell -ExecutionPolicy Bypass -File tools\export-slides.ps1
#  결과:    tools\_png\lesson-0X\*.PNG   (웹 슬라이드 이미지의 원본)
#           tools\_png\lesson-0X.pdf     (다운로드용 PDF)
#
#  [안전 규칙]
#   1) 원본 PPTX 는 절대 열지 않습니다. tools\_work\ 로 "복사" 한 뒤 복사본만 엽니다.
#   2) PowerPoint 가 이미 켜져 있으면 그 창을 절대 닫지 않습니다.
#      (사용자가 작업 중인 발표 자료를 지키기 위해서입니다)
#      이 스크립트가 직접 켠 경우에만 종료합니다.
#   3) 복사본은 창 없이(WithWindow=0) 읽기 전용으로 열고, 작업 후 바로 닫습니다.
# =====================================================================
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$outRoot = Join-Path $PSScriptRoot '_png'
$workDir = Join-Path $PSScriptRoot '_work'

# 회기 번호 <-> 원본 PPTX 찾기 패턴
# (파일명에 한글이 있어도 안전하도록 와일드카드로 찾습니다)
# 원본은 '01_수업PPT' 폴더에서 찾습니다. sub 를 주면 그 아래 폴더에서 찾습니다.
#
#  slides : 내보낼 슬라이드 번호를 순서대로 지정합니다. (생략하면 전체)
#           한 덱이 두 회기를 담을 때 쓰는 옵션입니다. 지금은 7~10차시를
#           tools/split-decks-07-10.py 로 회기별 덱으로 나눠 두어 쓰지 않습니다.
#
$srcDir = Join-Path $root '01_수업PPT'
$map = @(
  @{ id = 1;  pattern = '*_01차시*.pptx' },
  @{ id = 2;  pattern = '*_02차시*.pptx' },
  @{ id = 3;  pattern = '*_03차시*.pptx' },
  @{ id = 4;  pattern = '*_04차시*.pptx' },
  @{ id = 5;  pattern = '*_05차시*.pptx' },
  @{ id = 6;  pattern = '*_06차시*.pptx' },
  @{ id = 7;  pattern = '*_07차시*.pptx' },
  @{ id = 8;  pattern = '*_08차시*.pptx' },
  @{ id = 9;  pattern = '*_09차시*.pptx' },
  @{ id = 10; pattern = '*_10차시*.pptx' }
)

# --- PowerPoint 가 이미 실행 중인지 기억해 둡니다 --------------------
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
if ($wasRunning) {
  Write-Host "[안내] PowerPoint 가 이미 실행 중입니다." -ForegroundColor Yellow
  Write-Host "       열려 있는 창은 건드리지 않고, 작업이 끝나도 종료하지 않습니다."
  Write-Host ""
}

# --- 1단계: 원본을 작업 폴더로 복사 -----------------------------------
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

$jobs = @()
foreach ($m in $map) {
  $dir = if ($m.ContainsKey('sub')) { Join-Path $srcDir $m.sub } else { $srcDir }

  # 같은 패턴에 여러 파일이 있으면 "가장 최근에 수정된" 파일을 사용합니다.
  $src = Get-ChildItem -Path $dir -Filter $m.pattern -File |
         Sort-Object LastWriteTime -Descending |
         Select-Object -First 1
  if ($null -eq $src) { Write-Host ("SKIP  lesson-{0:D2} : 원본 없음" -f $m.id); continue }

  $copy = Join-Path $workDir ("lesson-{0:D2}.pptx" -f $m.id)
  Copy-Item -LiteralPath $src.FullName -Destination $copy -Force
  $pick = if ($m.ContainsKey('slides')) { $m.slides } else { $null }
  $jobs += @{ id = $m.id; copy = $copy; name = $src.Name; slides = $pick }
  $range = if ($null -eq $pick) { '전체' } else { ("{0}장 선택" -f $pick.Count) }
  Write-Host ("복사  lesson-{0:D2}  <- {1}  ({2})" -f $m.id, $src.Name, $range)
}

# --- 2단계: 복사본만 열어서 PNG + PDF 로 내보내기 ---------------------
$ppt = New-Object -ComObject PowerPoint.Application
try {
  foreach ($j in $jobs) {
    $dest = Join-Path $outRoot ("lesson-{0:D2}" -f $j.id)
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    # Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0)
    $pres = $ppt.Presentations.Open($j.copy, -1, 0, 0)
    try {
      if ($null -eq $j.slides) {
        # (1-a) 덱 전체를 원본 16:9 비율 그대로 1920x1080 PNG 로 내보내기
        $pres.Export($dest, 'PNG', 1920, 1080)
        $count = $pres.Slides.Count

        # (2) 다운로드용 PDF 저장 — 덱 전체일 때만
        $pdf = Join-Path $outRoot ("lesson-{0:D2}.pdf" -f $j.id)
        if (Test-Path $pdf) { Remove-Item $pdf -Force }
        try {
          # ppFixedFormatTypePDF = 2, ppFixedFormatIntentPrint = 2 (고화질)
          $pres.ExportAsFixedFormat($pdf, 2, 2)
        } catch {
          # 일부 버전에서는 SaveAs 방식이 필요합니다. (ppSaveAsPDF = 32)
          $pres.SaveCopyAs($pdf, 32)
        }
        Write-Host ("변환  lesson-{0:D2}  슬라이드 {1}장  (PNG + PDF)" -f $j.id, $count)
      } else {
        # (1-b) 지정한 슬라이드만 골라 1번부터 다시 번호를 붙여 내보내기
        $i = 0
        foreach ($no in $j.slides) {
          if ($no -lt 1 -or $no -gt $pres.Slides.Count) {
            Write-Host ("      [건너뜀] {0}번 슬라이드가 없습니다" -f $no) -ForegroundColor Yellow
            continue
          }
          $i++
          $file = Join-Path $dest ("슬라이드{0:D3}.PNG" -f $i)
          $pres.Slides.Item($no).Export($file, 'PNG', 1920, 1080)
        }
        $count = $i
        Write-Host ("변환  lesson-{0:D2}  슬라이드 {1}장  (원본 {2}장 중 선택 · PNG)" -f $j.id, $count, $pres.Slides.Count)
      }
    } finally {
      $pres.Close()
    }
  }
} finally {
  # ★ 이 스크립트가 직접 켠 경우에만 PowerPoint 를 종료합니다.
  if (-not $wasRunning) {
    $ppt.Quit()
  } else {
    Write-Host ""
    Write-Host "[안내] 원래 실행 중이던 PowerPoint 는 그대로 두었습니다." -ForegroundColor Yellow
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

# --- 3단계: 작업용 복사본 정리 ----------------------------------------
Remove-Item $workDir -Recurse -Force

Write-Host ""
Write-Host "내보내기 완료 -> $outRoot"
