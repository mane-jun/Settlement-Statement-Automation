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

  [옮기는 규칙]  (표는 B열에서 시작합니다. A열은 비워 둡니다)
    주문일   -> B열 (해당 파일의 행 전체를 세로 병합)
    구분     -> C열 (비움)
    품목     -> D열
    수령인   -> E열 (수령자)
    사이즈   -> F열 (152X203 처럼 대문자 X 로 적힌 것은 152x203 으로 통일)
    단가     -> G열 (단가표.txt 에서 찾아 넣음)
    수량     -> H열
    공급가액 -> I열 (수식  =G*H)
    VAT      -> J열 (수식  =I*0.1)
    택배비   -> K열 (택배비설정.txt 규칙으로 계산)
    합계     -> L열 (수식  =I+J+K+M)
    보관박스 -> M열 (금액으로 환산. 아래 참고)
    비고     -> N열 (제주·방문·단가확인 같은 표시를 남김)

  [단가 규칙]
    품목과 사이즈로 "단가표.txt" 를 찾아 넣습니다.
    표에 없는 사이즈는 비워 두고 비고에 "단가확인" 을 남깁니다.

  [택배비 규칙]
    긴변 + 짧은변 의 합으로 구간을 정합니다. (면적이 아닙니다)
      ~700 : 3,000 / ~900 : 3,500 / ~1100 : 4,000 / ~1300 : 5,000 / 그 위 : 6,000
      - 수령인이 비어 있는 줄은 윗줄과 같은 배송이라 0원
      - 배송메모에 방문·퀵 이 있으면 0원, 주소가 제주면 +3,000원
      - 종이포스터는 한 단계 아래 금액
      - 수량이 2개 이상인 줄은 비워 둡니다 (상자 수가 그때그때 다름)
    구간·금액은 "택배비설정.txt" 에서 바꿉니다.

  [보관박스 금액 규칙]
    주문 파일에 보관박스 표시가 있는 행만 금액이 들어갑니다.
    금액은 "짧은 변" 으로 정합니다. 긴 변이 아닙니다.
      - 짧은 변 < 254        3,850 원   (203x305 는 여기)
      - 254 <= 짧은 변 < 355  4,400 원
      - 짧은 변 >= 355        7,700 원
      - 여기에 수량을 곱한 값을 넣습니다.
    세 금액 모두 부가세 포함가라 합계에서 VAT 를 다시 곱하지 않습니다.
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
$PriceFileName     = '단가표.txt'
$ShipFileName      = '택배비설정.txt'

# 정산내역서 열 번호 (A=1, B=2, ...)  ※ 표는 B열에서 시작합니다
$COL_DATE   = 2    # B 주문일
$COL_DIV    = 3    # C 구분
$COL_ITEM   = 4    # D 품목
$COL_NAME   = 5    # E 수령자
$COL_SIZE   = 6    # F 사이즈
$COL_PRICE  = 7    # G 단가
$COL_QTY    = 8    # H 수량
$COL_SUPPLY = 9    # I 공급가액
$COL_VAT    = 10   # J VAT
$COL_SHIP   = 11   # K 택배비
$COL_TOTAL  = 12   # L 합계
$COL_BOX    = 13   # M 보관박스
$COL_LAST   = 14   # N 마지막 열 (비고)

# 열 번호 -> 엑셀 열 글자 (수식을 만들 때 씁니다)
function ColLetter {
    param([int]$n)
    $s = ''
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [char](65 + $m) + $s
        $n = [int](($n - $m - 1) / 26)
    }
    return $s
}

# Excel 상수
$xlUp           = -4162
$xlPasteFormats = -4122
$xlContinuous   = 1
$xlThin         = 2
$xlAutomatic    = -4105
$EDGES          = @(7, 8, 9, 10)   # 왼쪽, 위, 아래, 오른쪽

# 보관박스 금액 기본값 (설정 파일이 없으면 이 값을 씁니다)
$BOX_EDGE1           = 254     # 짧은 변 기준 1
$BOX_EDGE2           = 355     # 짧은 변 기준 2
$BOX_PRICE_SMALL     = 3850
$BOX_PRICE_MID       = 4400
$BOX_PRICE_BIG       = 7700
$BOX_MULTIPLY_QTY    = $true
$BOX_RECALC_EXISTING = $true

# 택배비 기본값 (설정 파일이 없으면 이 값을 씁니다)
$SHIP_BRACKETS   = @(@(700, 3000), @(900, 3500), @(1100, 4000), @(1300, 5000), @(99999, 6000))
$SHIP_VISIT_WORDS = @('방문', '방문수령', '퀵', '직접수령')
$SHIP_JEJU_ADD    = 3000
$SHIP_BLANK_QTY2  = $true

# 표준 규격 (사이즈 오타를 ±5mm 안에서 이 값으로 맞춥니다)
$STD_DIMS = @(125, 148, 152, 175, 203, 210, 254, 279, 297, 305, 355, 406, 420, 430, 457, 508, 594, 610, 762, 840)

# 단가표 : "품목|짧은변x긴변[x두께]" -> @( @{Amount=; From=} ... )
$PRICE_TABLE = @{}

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
#  금액은 사이즈의 "짧은 변" 으로 정해집니다. 긴 변이 아닙니다.
#  (203x305 는 짧은 변이 203 이라 작은금액입니다)
#
#     짧은 변 < 기준1              -> 작은금액
#     기준1 <= 짧은 변 < 기준2     -> 중간금액
#     짧은 변 >= 기준2             -> 큰금액
#
#  예) 기준1=254, 기준2=355 일 때
#      152x203 -> 3850    203x305 -> 3850    254x254 -> 4400
#      305x406 -> 4400    355x355 -> 7700    406x406 -> 7700
#
#  세 금액 모두 부가세가 포함된 값입니다.
#  그래서 합계에서 VAT 를 다시 곱하지 않고 그대로 더합니다.
#
#  수량곱하기 를 "예" 로 두면 위 금액에 수량을 곱해서 넣습니다.
#  기존글자도변환 을 "예" 로 두면, 정산내역서 보관박스 칸에 글자로 남아 있는
#  "포장가방" 표시도 실행할 때마다 금액으로 바꿔줍니다.
#  (이미 숫자가 들어 있는 칸은 건드리지 않습니다)
#
#  숫자만 고치면 됩니다. # 로 시작하는 줄은 설명이라 무시됩니다.
# ============================================================

기준1 = 254
기준2 = 355
작은금액 = 3850
중간금액 = 4400
큰금액 = 7700
수량곱하기 = 예
기존글자도변환 = 예
"@
    Set-Content -LiteralPath $SettingsPath -Value $text -Encoding UTF8
    Write-Host ("설정 파일을 새로 만들었습니다 : {0}" -f $SettingsFileName) -ForegroundColor DarkGray
}

function LoadBoxSettings {
    if (-not (Test-Path -LiteralPath $SettingsPath)) { return }
    $sawNew = $false
    foreach ($line in (Get-Content -LiteralPath $SettingsPath -Encoding UTF8)) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $kv = $t -split '=', 2
        if ($kv.Count -ne 2) { continue }
        $k = ($kv[0].Trim() -replace '\s', '')
        $v = $kv[1].Trim()
        switch ($k) {
            '기준1'          { if ($v -match '^\d+$') { $script:BOX_EDGE1       = [int]$v; $sawNew = $true } }
            '기준2'          { if ($v -match '^\d+$') { $script:BOX_EDGE2       = [int]$v; $sawNew = $true } }
            '작은금액'       { if ($v -match '^\d+$') { $script:BOX_PRICE_SMALL = [int]$v } }
            '중간금액'       { if ($v -match '^\d+$') { $script:BOX_PRICE_MID   = [int]$v; $sawNew = $true } }
            '큰금액'         { if ($v -match '^\d+$') { $script:BOX_PRICE_BIG   = [int]$v } }
            '수량곱하기'     { $script:BOX_MULTIPLY_QTY    = ($v -in @('예', 'Y', 'y', '1', 'true', 'TRUE')) }
            '기존글자도변환' { $script:BOX_RECALC_EXISTING = ($v -in @('예', 'Y', 'y', '1', 'true', 'TRUE')) }
        }
    }
    # 옛 형식(기준사이즈/작은금액/큰금액 2단계) 파일이면 새 형식으로 다시 만듭니다.
    if (-not $sawNew) {
        Write-Host '보관박스 설정이 옛 형식(2단계)이라 3단계 형식으로 새로 만듭니다.' -ForegroundColor Yellow
        $script:BOX_EDGE1 = 254; $script:BOX_EDGE2 = 355
        $script:BOX_PRICE_SMALL = 3850; $script:BOX_PRICE_MID = 4400; $script:BOX_PRICE_BIG = 7700
        Remove-Item -LiteralPath $SettingsPath -Force
        EnsureBoxSettingsFile
    }
}

function BoxSettingsSummary {
    $qtyNote = '수량 곱함'
    if (-not $BOX_MULTIPLY_QTY) { $qtyNote = '수량 안 곱함' }
    return ("보관박스 : 짧은변 {0}미만 {1:N0} / {0}~{2} {3:N0} / {2}이상 {4:N0}원 ({5})" -f `
            $BOX_EDGE1, $BOX_PRICE_SMALL, $BOX_EDGE2, $BOX_PRICE_MID, $BOX_PRICE_BIG, $qtyNote)
}

# ============================================================
#  택배비 설정 파일
# ============================================================
$ShipPath = Join-Path $PSScriptRoot $ShipFileName

function LoadShipSettings {
    if (-not (Test-Path -LiteralPath $ShipPath)) { return }
    $br = @()
    foreach ($line in (Get-Content -LiteralPath $ShipPath -Encoding UTF8)) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $kv = $t -split '=', 2
        if ($kv.Count -ne 2) { continue }
        $k = ($kv[0].Trim() -replace '\s', '')
        $v = $kv[1].Trim()
        switch ($k) {
            '구간' {
                if ($v -match '^(\d+)\s*:\s*(\d+)$') { $br += , @([int]$Matches[1], [int]$Matches[2]) }
            }
            '방문낱말'     { $script:SHIP_VISIT_WORDS = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
            '제주추가'     { if ($v -match '^\d+$') { $script:SHIP_JEJU_ADD = [int]$v } }
            '수량2이상비움'{ $script:SHIP_BLANK_QTY2 = ($v -in @('예', 'Y', 'y', '1', 'true', 'TRUE')) }
        }
    }
    if ($br.Count -gt 0) { $script:SHIP_BRACKETS = $br }
}

function ShipSettingsSummary {
    $parts = @()
    foreach ($b in $SHIP_BRACKETS) { $parts += ("~{0}:{1:N0}" -f $b[0], $b[1]) }
    $tail = '수량2이상 비움'
    if (-not $SHIP_BLANK_QTY2) { $tail = '수량2이상도 채움' }
    return ("택배비 : " + ($parts -join ' / ') + " (" + $tail + ")")
}

# ============================================================
#  단가표 파일
# ============================================================
$PricePath = Join-Path $PSScriptRoot $PriceFileName

function LoadPriceTable {
    $script:PRICE_TABLE = @{}
    if (-not (Test-Path -LiteralPath $PricePath)) {
        Write-Host ("[주의] {0} 이 없어 단가를 채우지 못합니다." -f $PriceFileName) -ForegroundColor Yellow
        return
    }
    foreach ($line in (Get-Content -LiteralPath $PricePath -Encoding UTF8)) {
        $t = $line.Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        $f = @($t -split '\|' | ForEach-Object { $_.Trim() })
        if ($f.Count -lt 3) { continue }
        $item = NormalizeItemGroup $f[0]
        $key  = SizeKey $f[1] $item      # 캔버스는 두께까지 열쇠에 들어갑니다
        if ($item -eq '' -or $key -eq '') { continue }
        if ($f[2] -notmatch '^\d+$') { continue }
        $from = $null
        if ($f.Count -ge 4 -and $f[3] -ne '') {
            try { $from = [DateTime]::Parse($f[3]) } catch { $from = $null }
        }
        $k = $item + '|' + $key
        if (-not $PRICE_TABLE.ContainsKey($k)) { $script:PRICE_TABLE[$k] = @() }
        $script:PRICE_TABLE[$k] += , ([pscustomobject]@{ Amount = [int]$f[2]; From = $from })
    }
}

function PriceSettingsSummary {
    return ("단가표 : {0}개 규격" -f $PRICE_TABLE.Count)
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

# 품목 이름을 6개 품목군 중 하나로 맞춥니다.
#  '알루미늄골드(아크릴판)', '알루미늄 화이트 프레임' -> '알루미늄'
function NormalizeItemGroup {
    param([string]$Item)
    $t = ([string]$Item) -replace '[\s ]', ''
    if ($t -eq '') { return '' }
    if ($t -like '*알루미늄*')   { return '알루미늄' }
    if ($t -like '*호두나무*')   { return '호두나무' }
    if ($t -like '*원목클래식*') { return '원목클래식' }
    if ($t -like '*종이포스터*') { return '종이포스터' }
    if ($t -like '*캔버스*')     { return '캔버스' }
    if ($t -like '*아크릴*')     { return '아크릴' }
    return $t
}

# 숫자를 표준 규격(±5mm)으로 맞춥니다. 203 <- 202, 305 <- 304 ...
function SnapDim {
    param([int]$v)
    $best = $v; $bd = 999
    foreach ($s in $STD_DIMS) {
        $d = [Math]::Abs($s - $v)
        if ($d -lt $bd -and $d -le 5) { $bd = $d; $best = $s }
    }
    return $best
}

# 사이즈 문자열 -> @{ Short; Long; Thick }   못 읽으면 $null
#  '203x254', '203X254X30', '203*305', 'A3_297 x 420', 'A2(594X420)' 모두 처리
function ParseSizeDims {
    param([string]$Size)
    $s = ([string]$Size).ToUpper() -replace '[\s ]', ''
    if ($s -eq '') { return $null }

    # A규격은 mm 가 함께 적혀 있어도 A번호를 우선합니다.
    $aw = 0; $ah = 0
    if     ($s -match 'A2') { $aw = 420; $ah = 594 }
    elseif ($s -match 'A3') { $aw = 297; $ah = 420 }
    elseif ($s -match 'A4') { $aw = 210; $ah = 297 }
    elseif ($s -match 'A5') { $aw = 148; $ah = 210 }
    if ($aw -gt 0) {
        return @{ Short = $aw; Long = $ah; Thick = 0 }
    }

    $nums = @([regex]::Matches($s, '\d+') | ForEach-Object { [int]$_.Value })
    if ($nums.Count -lt 2) { return $null }
    $w = SnapDim $nums[0]
    $h = SnapDim $nums[1]
    $short = [Math]::Min($w, $h)
    $long  = [Math]::Max($w, $h)
    $th = 0
    if ($nums.Count -ge 3) { $th = [int]$nums[2] }
    return @{ Short = $short; Long = $long; Thick = $th }
}

# 단가표 조회에 쓰는 열쇠. 캔버스만 두께를 함께 씁니다.
function SizeKey {
    param([string]$Size, [string]$ItemGroup = '')
    $d = ParseSizeDims $Size
    if ($null -eq $d) { return '' }
    if ($ItemGroup -eq '캔버스') {
        $th = $d.Thick
        if ($th -le 0) { $th = 30 }    # 두께 표기가 없으면 30mm 로 봅니다.
        return ('{0}x{1}x{2}' -f $d.Short, $d.Long, $th)
    }
    return ('{0}x{1}' -f $d.Short, $d.Long)
}

# 품목·사이즈·주문일로 단가를 찾습니다. 없으면 $null.
function PriceFor {
    param([string]$Item, [string]$Size, $OrderDate)
    $g = NormalizeItemGroup $Item
    if ($g -eq '') { return $null }
    $key = SizeKey $Size $g
    if ($key -eq '') { return $null }
    $rows = $PRICE_TABLE[$g + '|' + $key]
    if ($null -eq $rows -or @($rows).Count -eq 0) { return $null }

    # 주문일에 맞는 줄 중 적용시작일이 가장 늦은 것을 씁니다.
    $od = $null
    if ($OrderDate -is [DateTime]) { $od = $OrderDate }
    elseif ($OrderDate -is [string] -and $OrderDate -match '(\d{4})-(\d{2})-(\d{2})') {
        try { $od = New-Object DateTime -ArgumentList @([int]$Matches[1], [int]$Matches[2], [int]$Matches[3]) } catch {}
    }

    $best = $null
    foreach ($r in @($rows)) {
        if ($null -eq $r.From) {
            if ($null -eq $best) { $best = $r }
            continue
        }
        if ($null -ne $od -and $od -ge $r.From) {
            if ($null -eq $best -or $null -eq $best.From -or $r.From -gt $best.From) { $best = $r }
        }
    }
    if ($null -eq $best) { return $null }
    return [int]$best.Amount
}

# 택배비를 계산합니다. 비워 둬야 하면 $null.
#  두 번째 값으로 비고에 남길 표시를 함께 돌려줍니다.
function ShipFor {
    param(
        [string]$Item, [string]$Size, $Qty,
        [string]$Name, [string]$Memo, [string]$Addr,
        [bool]$IsFirstOfShipment = $true
    )
    # 수령인이 비어 있으면 윗줄과 같은 배송입니다.
    if (-not $IsFirstOfShipment) { return @{ Value = [double]0; Note = '' } }

    # 방문수령 / 퀵
    $m = ([string]$Memo) -replace '[\s ]', ''
    foreach ($w in $SHIP_VISIT_WORDS) {
        if ($w -ne '' -and $m -like ('*' + $w + '*')) { return @{ Value = [double]0; Note = $w } }
    }

    $d = ParseSizeDims $Size
    if ($null -eq $d) { return @{ Value = $null; Note = '' } }

    $q = 1
    $qt = CellText $Qty
    if ($qt -match '^\d+$' -and [int]$qt -gt 0) { $q = [int]$qt }
    if ($q -ge 2 -and $SHIP_BLANK_QTY2) { return @{ Value = $null; Note = '택배비확인' } }

    $sum = $d.Short + $d.Long
    $idx = -1
    for ($i = 0; $i -lt $SHIP_BRACKETS.Count; $i++) {
        if ($sum -le $SHIP_BRACKETS[$i][0]) { $idx = $i; break }
    }
    if ($idx -lt 0) { $idx = $SHIP_BRACKETS.Count - 1 }

    # 종이포스터는 한 단계 아래 금액
    if ((NormalizeItemGroup $Item) -eq '종이포스터' -and $idx -gt 0) { $idx = $idx - 1 }
    $fee = [int]$SHIP_BRACKETS[$idx][1]

    # 제주는 기본 택배비에 더합니다.
    $note = ''
    $a = ([string]$Addr) -replace '[\s ]', ''
    if ($a -like '제주*') { $fee = $fee + $SHIP_JEJU_ADD; $note = '제주' }

    return @{ Value = [double]$fee; Note = $note }
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

    # 사이즈에서 가로·세로 숫자를 뽑습니다. (203x254x30 처럼 두께가 붙어도 됩니다)
    $d = ParseSizeDims $Size
    if ($null -eq $d) {
        Write-Host ("    [주의] 사이즈를 읽을 수 없어 보관박스 금액을 비워둡니다: '{0}'" -f $Size) -ForegroundColor Yellow
        return $null
    }

    # 금액은 "짧은 변" 으로 정합니다. (203x305 는 짧은 변이 203 이라 작은금액)
    $s = $d.Short
    if     ($s -lt $BOX_EDGE1) { $unit = $BOX_PRICE_SMALL }
    elseif ($s -lt $BOX_EDGE2) { $unit = $BOX_PRICE_MID }
    else                       { $unit = $BOX_PRICE_BIG }

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
        $cMemo = PickColumn $map @('배송메모', '배송요청', '요청사항')
        $cAddr = PickColumn $map @('주소', '주소1', '배송지')
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
            $memo = ''
            if ($cMemo -gt 0) { $memo = CellText $vals[$i, $cMemo] }
            $addr = ''
            if ($cAddr -gt 0) { $addr = CellText $vals[$i, $cAddr] }
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
                Memo = $memo
                Addr = $addr
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

    # 1) 서식: 기존 데이터 행(C~N)을 새 행 범위에 그대로 입힙니다.
    #    (B열 주문일은 아래에서 병합·테두리를 따로 입힙니다)
    if ($lastItemRow -ge 3) {
        $tpl = $Sheet.Range($Sheet.Cells.Item($lastItemRow, ($COL_DATE + 1)), $Sheet.Cells.Item($lastItemRow, $COL_LAST))
        $tpl.Copy() | Out-Null
        $dst = $Sheet.Range($Sheet.Cells.Item($startRow, ($COL_DATE + 1)), $Sheet.Cells.Item($endRow, $COL_LAST))
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

    # 3) 값 채우기 (A~N 을 2차원 배열로 한 번에 씁니다. A열은 비워 둡니다)
    $LP = ColLetter $COL_PRICE      # G 단가
    $LQ = ColLetter $COL_QTY        # H 수량
    $LS = ColLetter $COL_SUPPLY     # I 공급가액
    $LV = ColLetter $COL_VAT        # J VAT
    $LK = ColLetter $COL_SHIP       # K 택배비
    $LB = ColLetter $COL_BOX        # M 보관박스

    $noPrice = 0
    $arr = New-Object 'object[,]' -ArgumentList $count, $COL_LAST
    for ($k = 0; $k -lt $count; $k++) {
        $o = $Rows[$k]
        $r = $startRow + $k
        if ($o.Item -ne '') { $arr[$k, ($COL_ITEM - 1)] = $o.Item }
        if ($o.Name -ne '') { $arr[$k, ($COL_NAME - 1)] = $o.Name }
        if ($o.Size -ne '') { $arr[$k, ($COL_SIZE - 1)] = $o.Size }
        if ((CellText $o.Qty) -ne '') { $arr[$k, ($COL_QTY - 1)] = $o.Qty }

        $notes = @()

        # --- 단가 ---
        $pv = PriceFor -Item $o.Item -Size $o.Size -OrderDate $OrderDate
        if ($null -ne $pv) {
            $arr[$k, ($COL_PRICE - 1)] = [double]$pv
        } else {
            $noPrice++
            $notes += '단가확인'
        }
        # 공급가액·VAT 수식은 단가를 못 찾았을 때도 넣어 둡니다.
        # 나중에 단가만 손으로 채우면 공급가액·VAT·합계가 저절로 따라옵니다.
        $arr[$k, ($COL_SUPPLY - 1)] = ('={0}{1}*{2}{1}' -f $LP, $r, $LQ)
        $arr[$k, ($COL_VAT - 1)]    = ('={0}{1}*0.1'    -f $LS, $r)

        # --- 택배비 ---
        #  수령인이 비어 있는 줄은 윗줄과 같은 배송이라 0원입니다.
        #  다만 블록의 첫 줄은 윗줄이 다른 주문일이므로 항상 자기 배송으로 봅니다.
        $isFirst = ($o.Name -ne '') -or ($k -eq 0)
        $sh = ShipFor -Item $o.Item -Size $o.Size -Qty $o.Qty `
                      -Name $o.Name -Memo $o.Memo -Addr $o.Addr -IsFirstOfShipment $isFirst
        if ($null -ne $sh.Value) { $arr[$k, ($COL_SHIP - 1)] = [double]$sh.Value }
        if ($sh.Note -ne '') { $notes += $sh.Note }

        # --- 보관박스 ---
        $bv = BoxValue -RawBox $o.Box -Size $o.Size -Item $o.Item -Qty $o.Qty
        if ($null -ne $bv -and (CellText $bv) -ne '') { $arr[$k, ($COL_BOX - 1)] = $bv }

        # --- 합계 (수식) ---
        $arr[$k, ($COL_TOTAL - 1)] = ('={0}{1}+{2}{1}+{3}{1}+{4}{1}' -f $LS, $r, $LV, $LK, $LB)

        # --- 비고 ---
        if ($notes.Count -gt 0) { $arr[$k, ($COL_LAST - 1)] = ($notes -join ' ') }
    }
    $arr[0, ($COL_DATE - 1)] = $OrderDate

    if ($noPrice -gt 0) {
        Write-Host ("    [주의] 단가표에 없는 규격 {0}건 -> 단가를 비우고 비고에 '단가확인' 을 남겼습니다." -f $noPrice) -ForegroundColor Yellow
    }

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
                    $first = ColLetter $COL_DATE
                    $last  = ColLetter $COL_LAST
                    $ws.PageSetup.PrintArea = ('${0}$1:${1}${2}' -f $first, $last, $lastWrittenRow)
                    Write-Host ''
                    Write-Host ("인쇄 영역을 {0}1:{1}{2} 로 넓혔습니다." -f $first, $last, $lastWrittenRow) -ForegroundColor DarkGray
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
    Write-Host (PriceSettingsSummary) -ForegroundColor DarkGray
    Write-Host (ShipSettingsSummary)  -ForegroundColor DarkGray
    Write-Host (BoxSettingsSummary)   -ForegroundColor DarkGray
}

function ReloadSettings {
    EnsureBoxSettingsFile
    LoadBoxSettings
    LoadShipSettings
    LoadPriceTable
}

ReloadSettings
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
    Write-Host '  4. 보관박스 금액 설정 바꾸기'
    Write-Host '  5. 단가표 고치기'
    Write-Host '  6. 택배비 설정 바꾸기'
    Write-Host '  7. 정산내역서 파일 바꾸기'
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
            Write-Host ''
            Write-Host '메모장이 열립니다. 단가를 고치고 저장한 뒤 메모장을 닫으세요.' -ForegroundColor DarkGray
            Write-Host '형식 :  품목 | 사이즈 | 단가 | 적용시작일(선택)' -ForegroundColor DarkGray
            Start-Process notepad.exe -ArgumentList $PricePath -Wait
            LoadPriceTable
            Write-Host (PriceSettingsSummary) -ForegroundColor Green
        }
        '6' {
            Write-Host ''
            Write-Host '메모장이 열립니다. 구간과 금액을 고치고 저장한 뒤 메모장을 닫으세요.' -ForegroundColor DarkGray
            Start-Process notepad.exe -ArgumentList $ShipPath -Wait
            LoadShipSettings
            Write-Host (ShipSettingsSummary) -ForegroundColor Green
        }
        '7' {
            $p = ResolveTarget -Repick $true
            if ($p) { Write-Host ("정산내역서 : {0}" -f $p) -ForegroundColor Green }
            else    { Write-Host '취소했습니다.' -ForegroundColor Yellow }
        }
        '0' {
            break
        }
        default {
            Write-Host '0 부터 7 사이에서 골라주세요.' -ForegroundColor Yellow
        }
    }

    if ($sel -eq '0') { break }
}

Write-Host ''
Write-Host '끝냈습니다.' -ForegroundColor DarkGray
exit 0
