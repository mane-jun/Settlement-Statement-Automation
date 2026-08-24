<#
  정산내역서 자동화
  ============================================================
  하나의 프로그램으로 아래 일을 모두 합니다.

    1) 주문 엑셀 파일(또는 폴더)을 정산내역서 맨 아래에 규칙대로 이어붙이기
    2) 보관박스 표시를 금액으로 환산해서 넣기
    3) 이미 글자로 남아 있는 '포장가방' 을 금액으로 바꾸기
    4) 금액 설정 바꾸기

  쓰는 법
    - 주문 파일/폴더를 "정산내역서 자동화.bat" 위로 끌어다 놓으면 바로 추가됩니다.
    - 그냥 더블클릭하면 메뉴가 뜹니다.

  [옮기는 규칙]
    주문일   -> A열 (해당 파일의 행 전체를 세로 병합)
    구분     -> B열 (비움)
    품목     -> C열
    수령인   -> D열 (수령자)
    사이즈   -> E열 (152X203 처럼 대문자 X 로 적힌 것은 152x203 으로 통일)
    단가     -> F열 (비움)
    수량     -> G열
    공급가액 -> H열 (비움)
    VAT      -> I열 (비움)
    택배비   -> J열 (비움)
    합계     -> K열 (비움)
    보관박스 -> L열 (금액으로 환산. 아래 참고)
    비고     -> M열 (비움)

  [보관박스 금액 규칙]
    주문 파일에 보관박스 표시가 있는 행만 금액이 들어갑니다.
      - 가로·세로가 모두 254 보다 작으면      3,850 원
      - 가로·세로 중 하나라도 254 이상이면    4,400 원
      - 여기에 수량을 곱한 값을 넣습니다. (예: 4,400 x 2 = 8,800)
    금액·기준은 "보관박스_금액설정.txt" 에서 바꿉니다.
#>

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths,
    [switch]$Force,
    [switch]$RecalcOnly
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# ============================================================
#  설정
# ============================================================
$SettlementPattern = '*정산내역서*.xlsx'   # 정산내역서 파일 찾는 패턴
$SettingsFileName  = '보관박스_금액설정.txt'

# 정산내역서 열 번호 (A=1, B=2, ...)
$COL_DATE = 1    # 주문일
$COL_ITEM = 3    # 품목
$COL_NAME = 4    # 수령자
$COL_SIZE = 5    # 사이즈
$COL_QTY  = 7    # 수량
$COL_BOX  = 12   # 보관박스
$COL_LAST = 13   # 마지막 열 (M = 비고)

# Excel 상수
$xlUp           = -4162
$xlPasteFormats = -4122
$xlContinuous   = 1
$xlThin         = 2
$xlAutomatic    = -4105
$EDGES          = @(7, 8, 9, 10)   # 왼쪽, 위, 아래, 오른쪽

# 보관박스 금액 기본값 (설정 파일이 없으면 이 값을 씁니다)
$BOX_SIZE_THRESHOLD  = 254
$BOX_PRICE_SMALL     = 3850
$BOX_PRICE_BIG       = 4400
$BOX_MULTIPLY_QTY    = $true
$BOX_RECALC_EXISTING = $true

# ============================================================
#  금액 설정 파일
# ============================================================
$SettingsPath = Join-Path $PSScriptRoot $SettingsFileName

function EnsureBoxSettingsFile {
    if (Test-Path -LiteralPath $SettingsPath) { return }
    $text = @"
# ============================================================
#  보관박스(포장가방) 금액 설정
# ============================================================
#  주문 파일에 "보관박스" 표시가 있는 행만 금액이 들어갑니다.
#
#  가로·세로가 모두 기준사이즈보다 작으면  -> 작은금액
#  가로·세로 중 하나라도 기준사이즈 이상   -> 큰금액
#  (예: 기준 254 일 때, 203x305 는 세로가 305 이므로 큰금액)
#
#  수량곱하기 를 "예" 로 두면 위 금액에 수량을 곱해서 넣습니다.
#  (예: 4400 x 2 = 8800)
#
#  기존글자도변환 을 "예" 로 두면, 정산내역서 보관박스 칸에 글자로 남아 있는
#  "포장가방" 표시도 실행할 때마다 금액으로 바꿔줍니다.
#  (이미 숫자가 들어 있는 칸은 건드리지 않습니다)
#
#  숫자만 고치면 됩니다. # 로 시작하는 줄은 설명이라 무시됩니다.
# ============================================================

기준사이즈 = 254
작은금액 = 3850
큰금액 = 4400
수량곱하기 = 예
기존글자도변환 = 예
"@
    Set-Content -LiteralPath $SettingsPath -Value $text -Encoding UTF8
    Write-Host ("설정 파일을 새로 만들었습니다 : {0}" -f $SettingsFileName) -ForegroundColor DarkGray
}

function LoadBoxSettings {
    if (-not (Test-Path -LiteralPath $SettingsPath)) { return }
    foreach ($line in (Get-Content -LiteralPath $SettingsPath -Encoding UTF8)) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $kv = $t -split '=', 2
        if ($kv.Count -ne 2) { continue }
        $k = ($kv[0].Trim() -replace '\s', '')
        $v = $kv[1].Trim()
        switch ($k) {
            '기준사이즈'     { if ($v -match '^\d+$') { $script:BOX_SIZE_THRESHOLD = [int]$v } }
            '작은금액'       { if ($v -match '^\d+$') { $script:BOX_PRICE_SMALL    = [int]$v } }
            '큰금액'         { if ($v -match '^\d+$') { $script:BOX_PRICE_BIG      = [int]$v } }
            '수량곱하기'     { $script:BOX_MULTIPLY_QTY    = ($v -in @('예', 'Y', 'y', '1', 'true', 'TRUE')) }
            '기존글자도변환' { $script:BOX_RECALC_EXISTING = ($v -in @('예', 'Y', 'y', '1', 'true', 'TRUE')) }
        }
    }
}

function BoxSettingsSummary {
    $qtyNote = '수량 곱함'
    if (-not $BOX_MULTIPLY_QTY) { $qtyNote = '수량 안 곱함' }
    return ("보관박스 금액 : {0} 미만 {1:N0}원 / {0} 이상 {2:N0}원 ({3})" -f `
            $BOX_SIZE_THRESHOLD, $BOX_PRICE_SMALL, $BOX_PRICE_BIG, $qtyNote)
}

# ============================================================
#  도우미 함수
# ============================================================
function CellText {
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ([Math]::Abs($v - [Math]::Round($v)) -lt 1e-9) { return ([string][int][Math]::Round($v)) }
        return ([string]$v)
    }
    return (([string]$v).Trim())
}

function NormalizeSize {
    param($v)
    $s = CellText $v
    if ($s -eq '') { return '' }
    $s = $s -replace '\s', ''
    $s = $s -replace '[Xx*]', 'x'
    return $s
}

function OrderDateText {
    param($v, [string]$FileName)
    $days = @('일', '월', '화', '수', '목', '금', '토')

    if ($v -is [double]) {
        try {
            $dt = [DateTime]::FromOADate($v)
            return ('{0:yyyy-MM-dd}({1})' -f $dt, $days[[int]$dt.DayOfWeek])
        } catch {}
    }
    if ($v -is [DateTime]) {
        return ('{0:yyyy-MM-dd}({1})' -f $v, $days[[int]$v.DayOfWeek])
    }

    $s = CellText $v
    if ($s -ne '') { return $s }

    # 셀이 비어 있으면 파일 이름 앞 6자리(예: 260727)에서 뽑아냅니다.
    if ($FileName -match '^(\d{2})(\d{2})(\d{2})') {
        try {
            $dt = New-Object DateTime -ArgumentList @((2000 + [int]$Matches[1]), [int]$Matches[2], [int]$Matches[3])
            return ('{0:yyyy-MM-dd}({1})' -f $dt, $days[[int]$dt.DayOfWeek])
        } catch {}
    }
    return ''
}

# 셀 값 쓰기.
# PowerShell 이 COM 속성 형식을 캐시해 두는 탓에 "문자 -> 숫자" 순서로 쓰면
# "Double 을 String 으로 변환할 수 없습니다" 오류가 납니다. 늦은 바인딩으로 우회합니다.
function SetValue2 {
    param($Target, $Value)
    try {
        $Target.Value2 = $Value
        return
    } catch {
        $firstError = $_.Exception.Message
    }
    try {
        # @($Value) 로 감싸면 2차원 배열이 풀려버리므로 인자 배열을 직접 만듭니다.
        $callArgs = New-Object 'object[]' 1
        $callArgs[0] = $Value
        [void]$Target.GetType().InvokeMember(
            'Value2',
            [System.Reflection.BindingFlags]::SetProperty,
            $null, $Target, $callArgs)
    } catch {
        throw ("셀에 값을 쓰지 못했습니다. ({0} / {1})" -f $firstError, $_.Exception.Message)
    }
}

function PickColumn {
    param([hashtable]$Map, [string[]]$Names)
    foreach ($n in $Names) { if ($Map.ContainsKey($n)) { return $Map[$n] } }
    return 0
}

function ExpandInputs {
    param([string[]]$InputPaths)
    $out = @()
    foreach ($p in $InputPaths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            $out += @(Get-ChildItem -LiteralPath $p -Filter '*.xlsx' -File |
                      Sort-Object Name | ForEach-Object { $_.FullName })
        } elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $out += (Resolve-Path -LiteralPath $p).Path
        } else {
            Write-Host ("  [건너뜀] 찾을 수 없는 경로: {0}" -f $p) -ForegroundColor Yellow
        }
    }
    $out = @($out | Where-Object {
        $n = [IO.Path]::GetFileName($_)
        ($n -like '*.xlsx') -and ($n -notlike '~$*') -and ($n -notlike '*정산내역서*')
    } | Sort-Object -Unique)
    return $out
}

# ============================================================
#  보관박스 금액 계산
# ============================================================
function BoxValue {
    param([string]$RawBox, [string]$Size, [string]$Item, $Qty)

    # 원본에 보관박스(포장가방) 표시가 없는 행은 그대로 비워둡니다.
    if ([string]::IsNullOrWhiteSpace($RawBox)) { return $null }

    # 사이즈에서 가로·세로 숫자를 뽑습니다. (152x203 -> 152, 203)
    if ($Size -notmatch '^(\d+)\s*x\s*(\d+)$') {
        Write-Host ("    [주의] 사이즈를 읽을 수 없어 보관박스 금액을 비워둡니다: '{0}'" -f $Size) -ForegroundColor Yellow
        return $null
    }
    $w = [int]$Matches[1]
    $h = [int]$Matches[2]

    # 가로·세로 모두 기준보다 작으면 작은 금액, 하나라도 기준 이상이면 큰 금액
    if ($w -lt $BOX_SIZE_THRESHOLD -and $h -lt $BOX_SIZE_THRESHOLD) {
        $unit = $BOX_PRICE_SMALL
    } else {
        $unit = $BOX_PRICE_BIG
    }

    # 수량 곱하기
    $n = 1
    if ($BOX_MULTIPLY_QTY) {
        $q = CellText $Qty
        if ($q -match '^\d+$' -and [int]$q -gt 0) { $n = [int]$q }
    }
    return [double]($unit * $n)
}

# ============================================================
#  주문 파일 한 개 읽기
# ============================================================
function ReadSourceOrders {
    param($Excel, [string]$Path)

    $fileName = [IO.Path]::GetFileName($Path)
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        $ws = $wb.Worksheets.Item(1)
        $ur = $ws.UsedRange
        $nR = [int]$ur.Rows.Count
        $nC = [int]$ur.Columns.Count
        if ($nR -lt 2 -or $nC -lt 2) { throw '내용이 없는 파일입니다.' }
        $vals = $ur.Value2

        # 머리글 행 찾기 ('주문일' 이 있는 행)
        $hdrR = 0
        $limit = [Math]::Min($nR, 15)
        for ($i = 1; $i -le $limit; $i++) {
            for ($j = 1; $j -le $nC; $j++) {
                if ((CellText $vals[$i, $j]) -eq '주문일') { $hdrR = $i; break }
            }
            if ($hdrR -gt 0) { break }
        }
        if ($hdrR -eq 0) { throw "'주문일' 머리글을 찾지 못했습니다." }

        # 머리글 -> 열 번호
        $map = @{}
        for ($j = 1; $j -le $nC; $j++) {
            $t = (CellText $vals[$hdrR, $j]) -replace '\s', ''
            if ($t -ne '' -and -not $map.ContainsKey($t)) { $map[$t] = $j }
        }
        $cDate = PickColumn $map @('주문일')
        $cItem = PickColumn $map @('품목')
        $cName = PickColumn $map @('수령인', '수령자')
        $cSize = PickColumn $map @('사이즈', '싸이즈')
        $cQty  = PickColumn $map @('수량')
        $cBox  = PickColumn $map @('보관박스', '보관박스여부')
        if ($cItem -eq 0 -or $cSize -eq 0) { throw "'품목' 또는 '사이즈' 열을 찾지 못했습니다." }

        $orderDate = ''
        $rows = @()
        for ($i = $hdrR + 1; $i -le $nR; $i++) {
            if ($orderDate -eq '' -and $cDate -gt 0) {
                $d = OrderDateText $vals[$i, $cDate] $fileName
                if ($d -ne '') { $orderDate = $d }
            }

            $item = CellText $vals[$i, $cItem]
            $size = NormalizeSize $vals[$i, $cSize]
            $name = ''
            if ($cName -gt 0) { $name = CellText $vals[$i, $cName] }
            $box = ''
            if ($cBox -gt 0) { $box = CellText $vals[$i, $cBox] }
            $qty = $null
            if ($cQty -gt 0) { $qty = $vals[$i, $cQty] }
            $qtyText = CellText $qty

            if ($item -eq '' -and $name -eq '' -and $size -eq '' -and $qtyText -eq '') { continue }

            $rows += [pscustomobject]@{
                Item = $item
                Name = $name
                Size = $size
                Qty  = $qty
                Box  = $box
            }
        }

        if ($orderDate -eq '') { $orderDate = OrderDateText $null $fileName }
        if ($orderDate -eq '') { throw '주문일을 알아낼 수 없습니다.' }

        return [pscustomobject]@{ Date = $orderDate; Rows = @($rows) }
    } finally {
        $wb.Close($false) | Out-Null
    }
}

# ============================================================
#  정산내역서에 한 파일분 붙이기
# ============================================================
function AppendBlock {
    param($Excel, $Sheet, [string]$OrderDate, $Rows)

    $count = @($Rows).Count
    if ($count -eq 0) { return 0 }

    $lastItemRow = [int]$Sheet.Cells.Item($Sheet.Rows.Count, $COL_ITEM).End($xlUp).Row
    if ($lastItemRow -lt 2) { $lastItemRow = 2 }
    $startRow = $lastItemRow + 1
    $endRow   = $startRow + $count - 1

    # 1) 서식: 기존 데이터 행(B~M)을 새 행 범위에 그대로 입힙니다.
    if ($lastItemRow -ge 3) {
        $tpl = $Sheet.Range($Sheet.Cells.Item($lastItemRow, 2), $Sheet.Cells.Item($lastItemRow, $COL_LAST))
        $tpl.Copy() | Out-Null
        $dst = $Sheet.Range($Sheet.Cells.Item($startRow, 2), $Sheet.Cells.Item($endRow, $COL_LAST))
        $dst.PasteSpecial($xlPasteFormats) | Out-Null
        # 주의: 여기서 $Excel.CutCopyMode 를 건드리면 엑셀 형식(PIA)이 로드되면서
        #       그 뒤로 숫자(수량) 쓰기가 실패합니다. 복사 표시는 저장/종료 때 알아서 사라집니다.
    }

    # 2) 병합 서식을 물려받을 기존 주문일 칸 위치를 미리 찾아둡니다.
    $aTplRow = 0
    for ($r = $lastItemRow; $r -ge 3; $r--) {
        $v = $Sheet.Cells.Item($r, $COL_DATE).Value2
        if ((CellText $v) -ne '') { $aTplRow = $r; break }
    }

    # 3) 값 채우기 (A~M 를 2차원 배열로 한 번에 씁니다)
    $arr = New-Object 'object[,]' -ArgumentList $count, $COL_LAST
    for ($k = 0; $k -lt $count; $k++) {
        $o = $Rows[$k]
        if ($o.Item -ne '') { $arr[$k, ($COL_ITEM - 1)] = $o.Item }
        if ($o.Name -ne '') { $arr[$k, ($COL_NAME - 1)] = $o.Name }
        if ($o.Size -ne '') { $arr[$k, ($COL_SIZE - 1)] = $o.Size }
        if ((CellText $o.Qty) -ne '') { $arr[$k, ($COL_QTY - 1)] = $o.Qty }
        $bv = BoxValue -RawBox $o.Box -Size $o.Size -Item $o.Item -Qty $o.Qty
        if ($null -ne $bv -and (CellText $bv) -ne '') { $arr[$k, ($COL_BOX - 1)] = $bv }
    }
    $arr[0, ($COL_DATE - 1)] = $OrderDate

    $block = $Sheet.Range($Sheet.Cells.Item($startRow, 1), $Sheet.Cells.Item($endRow, $COL_LAST))
    SetValue2 $block $arr

    # 4) 주문일(A열) 세로 병합 + 테두리
    $aRange = $Sheet.Range($Sheet.Cells.Item($startRow, $COL_DATE), $Sheet.Cells.Item($endRow, $COL_DATE))

    if ($aTplRow -gt 0) {
        $t = $Sheet.Cells.Item($aTplRow, $COL_DATE)
        try { $aRange.NumberFormatLocal    = $t.NumberFormatLocal }    catch {}
        try { $aRange.HorizontalAlignment  = $t.HorizontalAlignment }  catch {}
        try { $aRange.VerticalAlignment    = $t.VerticalAlignment }    catch {}
        try { $aRange.Font.Name            = $t.Font.Name }            catch {}
        try { $aRange.Font.Size            = $t.Font.Size }            catch {}
        try { $aRange.Font.Bold            = $t.Font.Bold }            catch {}
    }
    if ($endRow -gt $startRow) { $aRange.Merge() }
    foreach ($e in $EDGES) {
        $b = $aRange.Borders.Item($e)
        $b.LineStyle  = $xlContinuous
        $b.Weight     = $xlThin
        $b.ColorIndex = $xlAutomatic
    }

    Write-Host ("    -> {0}행 ~ {1}행에 {2}건 추가" -f $startRow, $endRow, $count) -ForegroundColor Green
    return $endRow
}

# ============================================================
#  L열에 글자로 남아 있는 보관박스 표시('포장가방')를 금액으로 바꾸기
#  - 이미 숫자가 들어 있는 칸은 건드리지 않습니다.
# ============================================================
function RecalcBoxColumn {
    param($Sheet)

    $lastRow = [int]$Sheet.Cells.Item($Sheet.Rows.Count, $COL_ITEM).End($xlUp).Row
    if ($lastRow -lt 3) { return 0 }

    $changed = 0
    for ($r = 3; $r -le $lastRow; $r++) {
        $raw = $Sheet.Cells.Item($r, $COL_BOX).Value2
        if ($null -eq $raw) { continue }
        if ($raw -is [double]) { continue }        # 이미 금액이면 그대로 둡니다.

        $txt = CellText $raw
        if ($txt -eq '') { continue }

        $size = NormalizeSize $Sheet.Cells.Item($r, $COL_SIZE).Value2
        $qty  = $Sheet.Cells.Item($r, $COL_QTY).Value2
        $bv   = BoxValue -RawBox $txt -Size $size -Item '' -Qty $qty
        if ($null -eq $bv) { continue }

        SetValue2 $Sheet.Cells.Item($r, $COL_BOX) $bv
        Write-Host ("    {0}행  '{1}' ({2} x {3}개) -> {4:N0}" -f `
                    $r, $txt, $size, (CellText $qty), $bv) -ForegroundColor Green
        $changed++
    }
    return $changed
}

# ============================================================
#  정산내역서 파일 찾기 / 주문 파일 고르기
# ============================================================
$TargetMemoPath = Join-Path $PSScriptRoot '정산내역서_경로.txt'
$TargetFile     = $null

function ScanFolderForSettlement {
    param([string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Folder)) { return @() }
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $Folder -Filter $SettlementPattern -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending)
}

function RememberTarget {
    param([string]$Path)
    try { Set-Content -LiteralPath $TargetMemoPath -Value $Path -Encoding UTF8 } catch {}
}

function PickSettlementFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title            = '정산내역서 파일을 고르세요'
    $dlg.Filter           = 'Excel 파일 (*.xlsx)|*.xlsx'
    $dlg.InitialDirectory = $PSScriptRoot
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.FileName
}

function ChooseFromCandidates {
    param($Candidates)
    Write-Host ''
    Write-Host '정산내역서가 여러 개입니다. 번호를 고르세요:' -ForegroundColor Yellow
    for ($i = 0; $i -lt $Candidates.Count; $i++) {
        Write-Host ("  [{0}] {1}" -f ($i + 1), $Candidates[$i].FullName)
    }
    $sel = Read-Host '번호'
    $idx = 0
    if (-not [int]::TryParse($sel, [ref]$idx) -or $idx -lt 1 -or $idx -gt $Candidates.Count) {
        Write-Host '잘못된 번호입니다.' -ForegroundColor Red
        return $null
    }
    return $Candidates[$idx - 1].FullName
}

# 정산내역서 위치를 정합니다.
#   1) 프로그램 폴더  2) 지난번에 쓴 경로  3) 바로 위 폴더  4) 파일 선택창
# 한 번 정하면 "정산내역서_경로.txt" 에 기억해 둡니다.
function ResolveTarget {
    param([bool]$Repick = $false)

    if ($Repick) {
        $p = PickSettlementFile
        if ($p) { $script:TargetFile = $p; RememberTarget $p }
        return $script:TargetFile
    }
    if ($script:TargetFile -and (Test-Path -LiteralPath $script:TargetFile -PathType Leaf)) {
        return $script:TargetFile
    }

    # 1) 프로그램 폴더
    $cands = ScanFolderForSettlement $PSScriptRoot
    if ($cands.Count -eq 1) {
        $script:TargetFile = $cands[0].FullName
        RememberTarget $script:TargetFile
        return $script:TargetFile
    }
    if ($cands.Count -gt 1) {
        $p = ChooseFromCandidates $cands
        if ($p) { $script:TargetFile = $p; RememberTarget $p }
        return $script:TargetFile
    }

    # 2) 지난번에 쓴 경로
    if (Test-Path -LiteralPath $TargetMemoPath) {
        try {
            $memo = (@(Get-Content -LiteralPath $TargetMemoPath -Encoding UTF8) | Select-Object -First 1)
            if ($memo) { $memo = $memo.Trim() }
            if ($memo -and (Test-Path -LiteralPath $memo -PathType Leaf)) {
                $script:TargetFile = $memo
                return $script:TargetFile
            }
        } catch {}
    }

    # 3) 바로 위 폴더
    $cands = ScanFolderForSettlement (Split-Path -Path $PSScriptRoot -Parent)
    if ($cands.Count -eq 1) {
        $script:TargetFile = $cands[0].FullName
        RememberTarget $script:TargetFile
        return $script:TargetFile
    }
    if ($cands.Count -gt 1) {
        $p = ChooseFromCandidates $cands
        if ($p) { $script:TargetFile = $p; RememberTarget $p }
        return $script:TargetFile
    }

    # 4) 파일 선택창
    Write-Host ''
    Write-Host '정산내역서를 자동으로 찾지 못했습니다. 직접 골라주세요.' -ForegroundColor Yellow
    $p = PickSettlementFile
    if ($p) { $script:TargetFile = $p; RememberTarget $p }
    return $script:TargetFile
}

function ChooseSourceFiles {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title            = '정산내역서에 추가할 주문 엑셀 파일을 고르세요 (여러 개 선택 가능)'
    $dlg.Filter           = 'Excel 파일 (*.xlsx)|*.xlsx'
    $dlg.Multiselect      = $true
    $dlg.InitialDirectory = $PSScriptRoot
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return @() }
    return (ExpandInputs $dlg.FileNames)
}

# ============================================================
#  실제 작업 (추가 + 보관박스 금액 반영)
# ============================================================
function RunJob {
    param(
        [string[]]$Sources = @(),
        [bool]$ForceAdd = $false,
        [bool]$RecalcMode = $false
    )

    $target = ResolveTarget
    if ([string]::IsNullOrWhiteSpace($target)) {
        Write-Host '정산내역서를 찾지 못했습니다.' -ForegroundColor Red
        return
    }

    Write-Host ''
    if ($RecalcMode) {
        Write-Host '할 일 : 보관박스 금액만 다시 계산 (주문 추가 없음)'
    } else {
        Write-Host ("추가할 파일 : {0}개" -f $Sources.Count)
        foreach ($s in $Sources) { Write-Host ("   - {0}" -f [IO.Path]::GetFileName($s)) }
    }

    # --- 백업 ---
    $bkDir = Join-Path $PSScriptRoot '백업'
    if (-not (Test-Path -LiteralPath $bkDir)) { New-Item -ItemType Directory -Path $bkDir | Out-Null }
    $bkPath = Join-Path $bkDir ('{0}_{1}.xlsx' -f [IO.Path]::GetFileNameWithoutExtension($target), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Copy-Item -LiteralPath $target -Destination $bkPath

    $excel     = $null
    $wb        = $null
    $added     = 0
    $converted = 0
    $saved     = $false
    $skipped   = @()
    $failed    = @()

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible          = $false
        $excel.DisplayAlerts    = $false
        $excel.AskToUpdateLinks = $false
        $excel.ScreenUpdating   = $false

        $wb = $excel.Workbooks.Open($target, 0, $false)
        if ($wb.ReadOnly) {
            throw ("정산내역서가 이미 열려 있어 저장할 수 없습니다. 엑셀에서 '{0}' 를 닫고 다시 실행해주세요." -f [IO.Path]::GetFileName($target))
        }

        # 월별 시트 고르기 ('26년 8월' 처럼 연/월이 든 첫 시트, 없으면 첫 번째 시트)
        $ws = $null
        foreach ($s in $wb.Worksheets) {
            if ($s.Name -match '\d+\s*년' -and $s.Name -match '\d+\s*월') { $ws = $s; break }
        }
        if ($null -eq $ws) { $ws = $wb.Worksheets.Item(1) }
        Write-Host ("대상 시트 : {0}" -f $ws.Name)
        Write-Host ''

        # 이미 들어가 있는 주문일 모으기
        $existing = @{}
        $lastItemRow = [int]$ws.Cells.Item($ws.Rows.Count, $COL_ITEM).End($xlUp).Row
        if ($lastItemRow -ge 3) {
            for ($r = 3; $r -le $lastItemRow; $r++) {
                $t = CellText $ws.Cells.Item($r, $COL_DATE).Value2
                if ($t -ne '') { $existing[$t] = $true }
            }
        }

        $lastWrittenRow = $lastItemRow
        foreach ($src in $Sources) {
            $name = [IO.Path]::GetFileName($src)
            Write-Host ("[{0}]" -f $name)
            try {
                $data = ReadSourceOrders $excel $src
                if ($data.Rows.Count -eq 0) {
                    Write-Host '    -> 옮길 주문이 없습니다. 건너뜁니다.' -ForegroundColor Yellow
                    $skipped += ("{0} (내용 없음)" -f $name)
                    continue
                }
                if ($existing.ContainsKey($data.Date) -and -not $ForceAdd) {
                    Write-Host ("    -> 주문일 {0} 은(는) 이미 들어가 있습니다. 건너뜁니다." -f $data.Date) -ForegroundColor Yellow
                    $skipped += ("{0} (주문일 {1} 중복)" -f $name, $data.Date)
                    continue
                }
                Write-Host ("    주문일 {0} / {1}건" -f $data.Date, $data.Rows.Count)
                $lastWrittenRow = AppendBlock $excel $ws $data.Date $data.Rows
                $existing[$data.Date] = $true
                $added += $data.Rows.Count
            } catch {
                Write-Host ("    -> 실패: {0}" -f $_.Exception.Message) -ForegroundColor Red
                $failed += ("{0} : {1}" -f $name, $_.Exception.Message)
            }
        }

        # --- 보관박스: 글자로 남아 있는 표시를 금액으로 바꾸기 ---
        if ($BOX_RECALC_EXISTING -or $RecalcMode) {
            Write-Host ''
            Write-Host '[보관박스 금액 계산]'
            $converted = RecalcBoxColumn $ws
            if ($converted -gt 0) {
                Write-Host ("    -> {0}칸을 금액으로 바꿨습니다." -f $converted) -ForegroundColor Green
            } else {
                Write-Host '    -> 바꿀 칸이 없습니다. (이미 모두 금액)' -ForegroundColor DarkGray
            }
        }

        if ($added -gt 0) {
            # 인쇄 영역이 새 데이터보다 짧으면 늘려줍니다.
            try {
                $pa = [string]$ws.PageSetup.PrintArea
                $paEnd = 0
                if ($pa -match ':\$?[A-Z]+\$?(\d+)') { $paEnd = [int]$Matches[1] }
                if ($paEnd -gt 0 -and $paEnd -lt $lastWrittenRow) {
                    $ws.PageSetup.PrintArea = ('$A$1:$M${0}' -f $lastWrittenRow)
                    Write-Host ''
                    Write-Host ("인쇄 영역을 A1:M{0} 로 넓혔습니다." -f $lastWrittenRow) -ForegroundColor DarkGray
                }
            } catch {}
        }

        if ($added -gt 0 -or $converted -gt 0) {
            $wb.Save()
            $saved = $true
            $done = @()
            if ($added -gt 0)     { $done += ("{0}건 추가" -f $added) }
            if ($converted -gt 0) { $done += ("보관박스 {0}칸 금액 반영" -f $converted) }
            Write-Host ''
            Write-Host ("저장 완료! " + ($done -join ', ')) -ForegroundColor Green
            Write-Host ("되돌리려면 : 백업\{0}" -f [IO.Path]::GetFileName($bkPath)) -ForegroundColor DarkGray
        } else {
            Write-Host ''
            Write-Host '바뀐 내용이 없어 저장하지 않았습니다.' -ForegroundColor Yellow
        }
    } catch {
        Write-Host ''
        Write-Host ("오류: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $failed += $_.Exception.Message
    } finally {
        if ($null -ne $wb)    { try { $wb.Close($false) } catch {} }
        if ($null -ne $excel) {
            try { $excel.ScreenUpdating = $true } catch {}
            try { $excel.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # 저장한 게 없으면 원본이 그대로이므로 방금 뜬 백업은 지웁니다. (백업 폴더가 쌓이지 않도록)
    if (-not $saved) {
        try { Remove-Item -LiteralPath $bkPath -Force -ErrorAction Stop } catch {}
    }

    if ($skipped.Count -gt 0) {
        Write-Host ''
        Write-Host '건너뛴 파일:' -ForegroundColor Yellow
        foreach ($s in $skipped) { Write-Host ("   - {0}" -f $s) -ForegroundColor Yellow }
        Write-Host '   (중복이어도 꼭 다시 넣으려면 메뉴 2번을 쓰세요)' -ForegroundColor DarkGray
    }
    if ($failed.Count -gt 0) {
        Write-Host ''
        Write-Host '실패한 항목:' -ForegroundColor Red
        foreach ($f in $failed) { Write-Host ("   - {0}" -f $f) -ForegroundColor Red }
    }
}

# ============================================================
#  본 작업
# ============================================================
function Banner {
    Write-Host ''
    Write-Host '============================================' -ForegroundColor Cyan
    Write-Host '   정산내역서 자동화' -ForegroundColor Cyan
    Write-Host '============================================' -ForegroundColor Cyan
    Write-Host (BoxSettingsSummary) -ForegroundColor DarkGray
}

EnsureBoxSettingsFile
LoadBoxSettings
Banner

$t = ResolveTarget
if ($t) { Write-Host ("정산내역서   : {0}" -f $t) -ForegroundColor DarkGray }

# --- 1) 파일/폴더를 끌어다 놓은 경우: 바로 추가 ---
if ($Paths -and @($Paths).Count -gt 0) {
    $sources = ExpandInputs $Paths
    if ($sources.Count -eq 0) {
        Write-Host '추가할 주문 파일이 없습니다. (.xlsx 파일 또는 폴더를 올려주세요)' -ForegroundColor Red
        exit 1
    }
    RunJob -Sources $sources -ForceAdd $Force.IsPresent
    exit 0
}

# --- 2) -RecalcOnly 로 실행한 경우 ---
if ($RecalcOnly) {
    RunJob -RecalcMode $true
    exit 0
}

# --- 3) 그냥 실행한 경우: 메뉴 ---
$emptyCount = 0
while ($true) {
    Write-Host ''
    Write-Host '--------------------------------------------'
    Write-Host '  1. 주문 파일 추가하기'
    Write-Host '  2. 주문 파일 추가하기 (중복 무시)'
    Write-Host '  3. 보관박스 금액 다시 계산하기'
    Write-Host '  4. 금액 설정 바꾸기'
    Write-Host '  5. 정산내역서 파일 바꾸기'
    Write-Host '  0. 끝내기'
    Write-Host '--------------------------------------------'
    $sel = Read-Host '번호를 고르세요'
    $sel = ([string]$sel).Trim()

    if ($sel -eq '') {
        $emptyCount++
        if ($emptyCount -ge 3) { break }
        continue
    }
    $emptyCount = 0

    switch ($sel) {
        '1' {
            $sources = ChooseSourceFiles
            if ($sources.Count -eq 0) { Write-Host '취소했습니다.' -ForegroundColor Yellow }
            else { RunJob -Sources $sources -ForceAdd $false }
        }
        '2' {
            $sources = ChooseSourceFiles
            if ($sources.Count -eq 0) { Write-Host '취소했습니다.' -ForegroundColor Yellow }
            else { RunJob -Sources $sources -ForceAdd $true }
        }
        '3' {
            RunJob -RecalcMode $true
        }
        '4' {
            Write-Host ''
            Write-Host '메모장이 열립니다. 숫자를 고치고 저장한 뒤 메모장을 닫으세요.' -ForegroundColor DarkGray
            Start-Process notepad.exe -ArgumentList $SettingsPath -Wait
            LoadBoxSettings
            Write-Host (BoxSettingsSummary) -ForegroundColor Green
        }
        '5' {
            $p = ResolveTarget -Repick $true
            if ($p) { Write-Host ("정산내역서 : {0}" -f $p) -ForegroundColor Green }
            else    { Write-Host '취소했습니다.' -ForegroundColor Yellow }
        }
        '0' {
            break
        }
        default {
            Write-Host '1, 2, 3, 4, 5, 0 중에서 골라주세요.' -ForegroundColor Yellow
        }
    }

    if ($sel -eq '0') { break }
}

Write-Host ''
Write-Host '끝냈습니다.' -ForegroundColor DarkGray
exit 0
