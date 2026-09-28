param(
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [Parameter(Mandatory=$true)][string]$ShotsDir
)

$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignParagraphRight = 2
$wdCollapseEnd = 0
$wdPageBreak = 7
$wdFormatXMLDocument = 12
$wdExportFormatPDF = 17
$wdBorderTop = -1
$wdBorderLeft = -2
$wdBorderBottom = -3
$wdBorderRight = -4
$wdBorderHorizontal = -5
$wdBorderVertical = -6
$wdLineStyleSingle = 1
$wdColorAutomatic = -16777216
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleCaption = -35
$wdStyleTitle = -63
$wdStyleSubtitle = -75

function Rgb([int]$r,[int]$g,[int]$b) { return ($r + 256*$g + 65536*$b) }
$navy = Rgb 31 78 121
$blue = Rgb 47 117 181
$lightBlue = Rgb 221 235 247
$paleBlue = Rgb 242 247 252
$orange = Rgb 237 125 49
$paleOrange = Rgb 252 228 214
$green = Rgb 112 173 71
$paleGreen = Rgb 226 239 218
$gray = Rgb 89 89 89
$lightGray = Rgb 242 242 242
$white = Rgb 255 255 255
$black = Rgb 0 0 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$global:word = $word
$global:doc = $doc
$global:ShotsDir = $ShotsDir
$doc.PageSetup.PageWidth = $word.CentimetersToPoints(21)
$doc.PageSetup.PageHeight = $word.CentimetersToPoints(29.7)
$doc.PageSetup.TopMargin = $word.CentimetersToPoints(1.25)
$doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1.2)
$doc.PageSetup.LeftMargin = $word.CentimetersToPoints(1.65)
$doc.PageSetup.RightMargin = $word.CentimetersToPoints(1.65)

$normal = $doc.Styles.Item($wdStyleNormal)
$normal.Font.Name = 'Malgun Gothic'
$normal.Font.NameFarEast = '맑은 고딕'
$normal.Font.Size = 9.8
$normal.Font.Color = $black
$normal.ParagraphFormat.SpaceAfter = 6
$normal.ParagraphFormat.LineSpacingRule = 0

foreach ($styleId in @($wdStyleTitle,$wdStyleSubtitle,$wdStyleHeading1,$wdStyleHeading2,$wdStyleCaption)) {
    try {
        $st = $doc.Styles.Item($styleId)
        $st.Font.Name = 'Malgun Gothic'
        $st.Font.NameFarEast = '맑은 고딕'
    } catch {}
}
$doc.Styles.Item($wdStyleTitle).Font.Size = 28
$doc.Styles.Item($wdStyleTitle).Font.Bold = $true
$doc.Styles.Item($wdStyleTitle).Font.Color = $navy
$doc.Styles.Item($wdStyleTitle).ParagraphFormat.SpaceAfter = 10
$doc.Styles.Item($wdStyleSubtitle).Font.Size = 13
$doc.Styles.Item($wdStyleSubtitle).Font.Color = $gray
$doc.Styles.Item($wdStyleSubtitle).ParagraphFormat.SpaceAfter = 18
$doc.Styles.Item($wdStyleHeading1).Font.Size = 17
$doc.Styles.Item($wdStyleHeading1).Font.Bold = $true
$doc.Styles.Item($wdStyleHeading1).Font.Color = $navy
$doc.Styles.Item($wdStyleHeading1).ParagraphFormat.SpaceBefore = 0
$doc.Styles.Item($wdStyleHeading1).ParagraphFormat.SpaceAfter = 8
$doc.Styles.Item($wdStyleHeading1).ParagraphFormat.KeepWithNext = $true
$doc.Styles.Item($wdStyleHeading1).ParagraphFormat.PageBreakBefore = $false
$doc.Styles.Item($wdStyleHeading2).Font.Size = 11.5
$doc.Styles.Item($wdStyleHeading2).Font.Bold = $true
$doc.Styles.Item($wdStyleHeading2).Font.Color = $blue
$doc.Styles.Item($wdStyleHeading2).ParagraphFormat.SpaceBefore = 8
$doc.Styles.Item($wdStyleHeading2).ParagraphFormat.SpaceAfter = 4
$doc.Styles.Item($wdStyleHeading2).ParagraphFormat.KeepWithNext = $true

function Move-End {
    $global:sel = $global:doc.Range($global:doc.Content.End - 1, $global:doc.Content.End - 1)
}

function Add-Paragraph {
    param([string]$Text,[string]$Style='Normal',[int]$Align=0,[int]$Color=-16777216,[bool]$Bold=$false,[double]$Size=0)
    Move-End
    $styleId = switch ($Style) {
        'Title' { $wdStyleTitle }
        'Subtitle' { $wdStyleSubtitle }
        'Heading 1' { $wdStyleHeading1 }
        'Heading 2' { $wdStyleHeading2 }
        'Caption' { $wdStyleCaption }
        default { $wdStyleNormal }
    }
    $global:sel.Style = $global:doc.Styles.Item($styleId)
    $global:sel.ParagraphFormat.Alignment = $Align
    $global:sel.Font.Color = $Color
    $global:sel.Font.Bold = $Bold
    if ($Size -gt 0) { $global:sel.Font.Size = $Size }
    $global:sel.InsertAfter($Text)
    $global:sel.InsertParagraphAfter()
}

function Add-PageBreak {
    Move-End
    $global:sel.InsertBreak($wdPageBreak)
}

function Add-MajorHeading {
    param([string]$Text)
    Move-End
    $global:sel.Style = $global:doc.Styles.Item($wdStyleHeading1)
    $global:sel.ParagraphFormat.Alignment = $wdAlignParagraphLeft
    $global:sel.ParagraphFormat.PageBreakBefore = -1
    $global:sel.Font.Color = $navy
    $global:sel.Font.Bold = $true
    $global:sel.InsertAfter($Text)
    $global:sel.InsertParagraphAfter()
}

function Set-CellText {
    param($Cell,[string]$Text,[int]$Color=$black,[bool]$Bold=$false,[double]$Size=9.5,[int]$Align=0)
    $r = $Cell.Range
    $r.End = $r.End - 1
    $r.Text = $Text
    $r.Font.Name = 'Malgun Gothic'
    $r.Font.NameFarEast = '맑은 고딕'
    $r.Font.Size = $Size
    $r.Font.Color = $Color
    $r.Font.Bold = $Bold
    $r.ParagraphFormat.Alignment = $Align
    $r.ParagraphFormat.SpaceAfter = 0
    $Cell.VerticalAlignment = 1
    $Cell.TopPadding = 2.5
    $Cell.BottomPadding = 2.5
    $Cell.LeftPadding = 6
    $Cell.RightPadding = 6
}

function Style-TableBorders {
    param($Table,[int]$Color)
    foreach($idx in @($wdBorderTop,$wdBorderLeft,$wdBorderBottom,$wdBorderRight,$wdBorderHorizontal,$wdBorderVertical)) {
        try {
            $b = $Table.Borders.Item($idx)
            $b.LineStyle = $wdLineStyleSingle
            $b.Color = $Color
            $b.LineWidth = 4
        } catch {}
    }
}

function Add-Callout {
    param([string]$Label,[string]$Text,[int]$Fill=$paleBlue,[int]$Accent=$blue)
    Move-End
    $t = $global:doc.Tables.Add($global:sel,1,1)
    $t.AllowAutoFit = $false
    $t.PreferredWidthType = 2
    $t.PreferredWidth = $global:word.CentimetersToPoints(17.4)
    $t.Columns.Item(1).Width = $global:word.CentimetersToPoints(17.4)
    $t.Cell(1,1).Shading.BackgroundPatternColor = $Fill
    Set-CellText $t.Cell(1,1) ($Label + '  ' + $Text) $black $false 9.3 0
    $rr = $t.Cell(1,1).Range
    $rr.End = [Math]::Min($rr.Start + $Label.Length, $rr.End - 1)
    $rr.Font.Bold = $true
    $rr.Font.Color = $Accent
    Style-TableBorders $t $Accent
}

function Add-Image {
    param([string]$File,[string]$Caption,[string]$Alt)
    Move-End
    $global:sel.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $pic = $global:sel.InlineShapes.AddPicture($File,$false,$true)
    $pic.LockAspectRatio = -1
    $pic.Width = $global:word.CentimetersToPoints(16.0)
    $pic.AlternativeText = $Alt
    $pictureParagraph = $pic.Range.Paragraphs.Item(1).Range
    $pictureParagraph.InsertParagraphAfter()
    Add-Paragraph $Caption 'Normal' $wdAlignParagraphCenter $gray $true 8
}

function Add-FormulaTable {
    param([object[]]$Rows,[double[]]$Widths)
    Move-End
    $t = $global:doc.Tables.Add($global:sel,$Rows.Count,3)
    $t.AllowAutoFit = $false
    $t.PreferredWidthType = 2
    $total = 0; foreach($w in $Widths){$total += $w}
    $t.PreferredWidth = $global:word.CentimetersToPoints($total)
    for($c=1;$c -le 3;$c++){ $t.Columns.Item($c).Width = $global:word.CentimetersToPoints($Widths[$c-1]) }
    for($r=1;$r -le $Rows.Count;$r++) {
        $row = $Rows[$r-1]
        for($c=1;$c -le 3;$c++) {
            Set-CellText $t.Cell($r,$c) ([string]$row[$c-1]) $black ($r -eq 1) 8.6 0
            if($r -eq 1){$t.Cell($r,$c).Shading.BackgroundPatternColor=$navy; $cr=$t.Cell($r,$c).Range; $cr.Font.Color=$white}
            elseif($c -eq 1){$t.Cell($r,$c).Shading.BackgroundPatternColor=$lightBlue; $cr=$t.Cell($r,$c).Range; $cr.Font.Bold=$true; $cr.Font.Color=$navy}
        }
    }
    Style-TableBorders $t (Rgb 191 191 191)
    Move-End
    $global:sel.InsertParagraphAfter()
}

function Add-StepTable {
    param([object[]]$Steps)
    Move-End
    $t = $global:doc.Tables.Add($global:sel,$Steps.Count,2)
    $t.AllowAutoFit = $false
    $t.Columns.Item(1).Width = $global:word.CentimetersToPoints(1.25)
    $t.Columns.Item(2).Width = $global:word.CentimetersToPoints(16.15)
    for($r=1;$r -le $Steps.Count;$r++) {
        Set-CellText $t.Cell($r,1) ([string]$r) $white $true 10 $wdAlignParagraphCenter
        $t.Cell($r,1).Shading.BackgroundPatternColor = $blue
        Set-CellText $t.Cell($r,2) ([string]$Steps[$r-1]) $black $false 8.8 0
        if(($r % 2) -eq 0){$t.Cell($r,2).Shading.BackgroundPatternColor=$paleBlue}
    }
    Style-TableBorders $t (Rgb 217 217 217)
    Move-End
    $global:sel.InsertParagraphAfter()
}

function Add-FunctionPage {
    param(
        [string]$Title,[string]$Lead,[string]$Shot,[string]$Caption,
        [object[]]$FormulaRows,[object[]]$Steps,[string]$Checkpoint
    )
    Add-MajorHeading $Title
    Add-Callout '핵심' $Lead $paleBlue $blue
    Add-Image (Join-Path $global:ShotsDir $Shot) $Caption ($Title + ' 실제 Excel 화면')
    Add-Paragraph '수식을 네 자리로 읽기' 'Heading 2'
    Add-FormulaTable $FormulaRows @(2.6,5.3,9.5)
    Add-Paragraph '실제 입력 순서' 'Heading 2'
    Add-StepTable $Steps
    Add-Callout '시험 체크' $Checkpoint $paleOrange $orange
}

# Header/footer
$header = $doc.Sections.Item(1).Headers.Item(1).Range
$header.Text = '컴퓨터활용능력 2급 실기  |  찾기와 참조 함수'
$header.Font.Name = 'Malgun Gothic'
$header.Font.NameFarEast = '맑은 고딕'
$header.Font.Size = 8.5
$header.Font.Color = $gray
$header.ParagraphFormat.Alignment = $wdAlignParagraphRight
$footer = $doc.Sections.Item(1).Footers.Item(1)
$footer.Range.Text = '실제 Excel 실습 화면 기준 개정판'
$footer.Range.Font.Name = 'Malgun Gothic'
$footer.Range.Font.NameFarEast = '맑은 고딕'
$footer.Range.Font.Size = 8
$footer.Range.Font.Color = $gray
$footer.Range.ParagraphFormat.Alignment = $wdAlignParagraphLeft
$null = $footer.PageNumbers.Add(2,$true)

# Cover
Add-Paragraph '컴퓨터활용능력 2급 실기' 'Subtitle' $wdAlignParagraphCenter $gray
Add-Paragraph '찾기와 참조 함수' 'Title' $wdAlignParagraphCenter $navy
Add-Paragraph 'VLOOKUP · HLOOKUP · CHOOSE · INDEX + MATCH' 'Subtitle' $wdAlignParagraphCenter $blue
Add-Callout '이 개정판의 기준' '실제 Microsoft Excel에서 실습 파일을 열어 캡처했습니다. 모든 설명의 셀 주소·표 범위·결과 셀이 화면과 정확히 일치합니다.' $paleGreen $green
Add-Paragraph '' 'Normal'
Add-Paragraph '이 교재를 보는 방법' 'Heading 1'
Add-StepTable @(
    '화면에서 초록색 테두리로 선택된 결과 셀을 먼저 확인합니다.',
    '위쪽 수식 입력줄에서 완성된 수식을 확인합니다.',
    '수식 안의 색과 같은 색 테두리가 둘러진 참조 셀·범위를 찾습니다.',
    '표 아래의 “네 자리로 읽기”에서 각 인수의 역할을 확인합니다.',
    '실습 파일에서 같은 시트와 셀을 직접 클릭하여 수식을 다시 입력합니다.'
)
Add-Callout '중요' '캡처는 수식 편집 상태(F2)입니다. 그래서 결과 셀 안에 계산 결과 대신 수식이 보이고, 참조 범위가 색 테두리로 표시됩니다. Enter를 누르면 결과값으로 돌아옵니다.' $paleOrange $orange
Add-Paragraph '함수 선택 10초 판단' 'Heading 2'
Add-FormulaTable @(
    @('표 모양 / 목적','함수','판단 문장'),
    @('세로 표','VLOOKUP','첫 열에서 찾아 오른쪽 값을 가져온다.'),
    @('가로 표','HLOOKUP','첫 행에서 찾아 아래쪽 값을 가져온다.'),
    @('번호로 선택','CHOOSE','1, 2, 3번에 따라 하나를 고른다.'),
    @('찾는 열과 반환 열 분리','INDEX + MATCH','MATCH가 위치, INDEX가 값이다.')
) @(3.3,3.2,10.9)

Add-FunctionPage '1. VLOOKUP — 세로 표에서 정확히 찾기' `
    '기준값을 선택 범위의 첫 번째 열에서 찾고, 같은 행의 지정한 열 값을 가져옵니다.' `
    '01_vlookup_exact.png' `
    '그림 1. [01_VLOOKUP] 시트 F7을 F2로 편집한 실제 Excel 화면' `
    @(
        @('자리','화면의 값','뜻'),
        @('기준값','F4','찾을 제품코드 P103'),
        @('표 범위','$A$5:$C$9','제품코드가 첫 열인 전체 표'),
        @('열 번호','3','선택 범위의 세 번째 열인 단가'),
        @('일치 방법','FALSE','P103과 정확히 같은 코드만 찾기')
    ) `
    @(
        '결과를 표시할 F7 셀을 클릭합니다.',
        '=VLOOKUP( 을 입력하고 기준값 F4를 클릭합니다.',
        '쉼표 뒤에 A5:C9를 드래그한 다음 F4 키를 눌러 $A$5:$C$9로 고정합니다.',
        '완성 수식 =VLOOKUP(F4,$A$5:$C$9,3,FALSE) 를 확인한 뒤 Enter를 누릅니다.',
        '결과가 210,000원인지 확인합니다.'
    ) `
    '열 번호 3은 워크시트의 C열이라서가 아니라, 선택 범위 A:C 안에서 C가 세 번째이기 때문입니다.'

Add-FunctionPage '2. VLOOKUP — 점수 구간을 근사값으로 찾기' `
    '점수가 표에 정확히 없어도 “그 점수 이하에서 가장 큰 기준점수”를 찾아 등급을 반환합니다.' `
    '02_vlookup_approx.png' `
    '그림 2. [02_VLOOKUP_근사값] 시트 E7의 실제 Excel 수식 편집 화면' `
    @(
        @('자리','화면의 값','뜻'),
        @('기준값','E4','학생 점수 86'),
        @('표 범위','$A$5:$B$9','최저 점수와 등급표'),
        @('열 번호','2','등급이 있는 두 번째 열'),
        @('일치 방법','TRUE','86 이하 최대 기준점수 80을 사용')
    ) `
    @(
        '결과 셀 E7을 클릭합니다.',
        '=VLOOKUP(E4,$A$5:$B$9,2,TRUE) 를 입력합니다.',
        'Enter를 누르면 86 이하의 기준점수 중 가장 큰 80을 찾습니다.',
        '80행의 두 번째 열 값 B가 결과로 표시되는지 확인합니다.'
    ) `
    'TRUE를 쓸 때 기준표의 첫 열은 반드시 오름차순이어야 합니다. 0, 60, 70, 80, 90 순서를 확인하세요.'

Add-FunctionPage '3. HLOOKUP — 가로 표에서 찾기' `
    '기준값을 선택 범위의 첫 번째 행에서 찾고, 같은 열의 지정한 행 값을 가져옵니다.' `
    '03_hlookup.png' `
    '그림 3. [03_HLOOKUP] 시트 I7의 실제 Excel 수식 편집 화면' `
    @(
        @('자리','화면의 값','뜻'),
        @('기준값','I4','찾을 월 4월'),
        @('표 범위','$B$4:$F$7','1월~5월이 첫 행인 가로 표'),
        @('행 번호','4','선택 범위의 네 번째 행인 이익'),
        @('일치 방법','FALSE','4월과 정확히 같은 항목 찾기')
    ) `
    @(
        '결과 셀 I7을 클릭합니다.',
        '=HLOOKUP(I4,$B$4:$F$7,4,FALSE) 를 입력합니다.',
        '첫 행에서 4월을 찾고, 그 열에서 범위의 네 번째 행을 내려갑니다.',
        'Enter를 눌러 이익 660이 표시되는지 확인합니다.'
    ) `
    '행 번호 4는 워크시트 7행이 아니라, 선택 범위 B4:F7 안에서 이익 행이 네 번째라는 뜻입니다.'

Add-FunctionPage '4. CHOOSE — 번호에 따라 하나 선택하기' `
    '선택 번호가 1이면 첫 번째 값, 2이면 두 번째 값, 3이면 세 번째 값을 반환합니다.' `
    '04_choose.png' `
    '그림 4. [04_CHOOSE] 시트 E7의 실제 Excel 수식 편집 화면' `
    @(
        @('자리','화면의 값','뜻'),
        @('선택 번호','E4','현재 번호는 3'),
        @('값 1','"월요일"','번호 1일 때 반환'),
        @('값 2','"화요일"','번호 2일 때 반환'),
        @('값 3','"수요일"','번호 3일 때 반환')
    ) `
    @(
        '결과 셀 E7을 클릭합니다.',
        '=CHOOSE(E4,"월요일","화요일","수요일","목요일") 를 입력합니다.',
        'E4의 값 3이 목록의 세 번째 항목을 고릅니다.',
        'Enter를 눌러 수요일이 표시되는지 확인합니다.'
    ) `
    '왼쪽 A:B의 요일표는 이해를 위한 보기일 뿐 수식 범위가 아닙니다. CHOOSE는 선택 값을 수식 안에 직접 적습니다.'

Add-FunctionPage '5. INDEX + MATCH — 위치를 찾고 값 꺼내기' `
    'MATCH가 찾는 값의 상대 위치를 구하고, INDEX가 반환 범위에서 같은 위치의 값을 꺼냅니다.' `
    '05_index_match.png' `
    '그림 5. [05_INDEX_MATCH] 시트 F8의 실제 Excel 수식 편집 화면' `
    @(
        @('역할','화면의 식','뜻'),
        @('찾을 값','F4','최유진'),
        @('MATCH','$B$5:$B$9, 0','이름 범위에서 정확히 일치하는 위치 4'),
        @('INDEX','$C$5:$C$9','부서 범위의 네 번째 값 반환'),
        @('최종 결과','회계팀','두 범위의 행 순서가 같아서 연결됨')
    ) `
    @(
        '먼저 =MATCH(F4,$B$5:$B$9,0) 을 확인하면 결과는 4입니다.',
        '다음 =INDEX($C$5:$C$9,4) 를 확인하면 결과는 회계팀입니다.',
        'MATCH 식을 INDEX의 행 번호 자리에 넣습니다.',
        '=INDEX($C$5:$C$9,MATCH(F4,$B$5:$B$9,0)) 를 입력하고 Enter를 누릅니다.'
    ) `
    'MATCH 범위 B5:B9와 INDEX 범위 C5:C9는 시작 행과 끝 행이 같아야 위치 4가 같은 사람을 가리킵니다.'

# Input workflow page
Add-MajorHeading '6. 실제 Excel에서 수식 입력하는 공통 순서'
Add-Paragraph '한 줄 흐름  결과 셀 클릭 → = 입력 → 함수명 → 인수 입력 → 범위 고정 → Enter → 결과 검산' 'Normal' $wdAlignParagraphLeft $green $true 9.5
Add-StepTable @(
    '결과가 들어갈 셀을 한 번 클릭합니다. 이름 상자에서 셀 주소를 확인합니다.',
    '= 기호와 함수 이름을 입력합니다. 예: =VLOOKUP(',
    '기준값 셀을 클릭하고 쉼표를 입력합니다.',
    '표 범위를 마우스로 드래그합니다. 복사할 수식이면 F4 키로 $를 붙입니다.',
    '열/행 번호와 FALSE·TRUE·0을 넣고 닫는 괄호를 입력합니다.',
    'Enter를 누른 뒤 값이 맞는지 원본 표에서 직접 대조합니다.',
    '수식을 다시 볼 때는 결과 셀을 선택하고 F2를 누릅니다.'
)
Add-Paragraph '절대참조 F4 키' 'Heading 2'
Add-FormulaTable @(
    @('입력 상태','F4 키를 누른 결과','사용'),
    @('A5:C9','$A$5:$C$9','행과 열 모두 고정'),
    @('$A$5:$C$9','A$5:C$9','상황에 따라 고정 형태 순환'),
    @('시험 기본','대부분 표 범위를 절대참조','아래로 복사해도 표가 움직이지 않음')
) @(3.2,5.4,8.8)
Add-Callout '오류 점검' '#N/A는 못 찾음, #REF!는 열·행 번호 초과, #VALUE!는 인수 형식 오류입니다. 먼저 기준값·첫 열/첫 행·일치 옵션을 확인하세요.' $paleOrange $orange

# Practice page
Add-MajorHeading '7. 시험 연습 — 실제 시트에서 직접 입력하기'
Add-Callout '실습 위치' '7.찾기와_참조_함수_실습.xlsx 파일의 [06_시험연습] 시트입니다. J5:J9가 정답 입력 칸입니다.' $paleBlue $blue
Add-Image (Join-Path $ShotsDir '06_practice.png') '그림 6. [06_시험연습] 시트 J5를 선택한 실제 Excel 화면' '시험 연습 시트 실제 Excel 화면'
Add-StepTable @(
    'J5: G5의 A104를 상품표에서 찾아 상품명을 반환합니다.',
    'J6: G6의 A105를 상품표에서 찾아 단가를 반환합니다.',
    'J7: G7의 76을 점수표에서 근사값으로 찾아 등급을 반환합니다.',
    'J8: G8의 2를 CHOOSE의 선택 번호로 사용합니다.',
    'J9: G9의 물병 위치를 MATCH로 찾고 INDEX로 분류를 반환합니다.'
)
Add-Callout '풀이 원칙' '먼저 어느 함수인지 결정한 다음, 화면의 열 제목을 보고 기준값·범위·반환 위치·일치 방법을 하나씩 채우세요.' $paleGreen $green

# Answers page
Add-MajorHeading '8. 정답 확인 — 수식과 결과를 함께 비교'
Add-Image (Join-Path $ShotsDir '07_answers.png') '그림 7. [07_정답] 시트 B5를 선택한 실제 Excel 화면' '시험 연습 정답 시트 실제 Excel 화면'
Add-FormulaTable @(
    @('문제','정답 수식','결과'),
    @('1','=VLOOKUP(G5,$A$5:$D$9,2,FALSE)','우산'),
    @('2','=VLOOKUP(G6,$A$5:$D$9,4,FALSE)','35,000원'),
    @('3','=VLOOKUP(G7,$A$13:$B$17,2,TRUE)','C'),
    @('4','=CHOOSE(G8,"월요일","화요일","수요일")','화요일'),
    @('5','=INDEX($C$5:$C$9,MATCH(G9,$B$5:$B$9,0))','생활')
) @(1.5,12.7,3.2)
Add-Callout '채점 기준' '결과값만 보지 말고 표 범위의 $ 고정, VLOOKUP/HLOOKUP의 열·행 번호, FALSE·TRUE·MATCH의 0까지 확인해야 합니다.' $paleOrange $orange

# Final memory page
Add-MajorHeading '9. 시험 직전 1분 암기'
Add-FormulaTable @(
    @('함수','기본 형식','암기 문장'),
    @('VLOOKUP','=VLOOKUP(값,범위,열,FALSE)','첫 열에서 찾아 옆으로'),
    @('HLOOKUP','=HLOOKUP(값,범위,행,FALSE)','첫 행에서 찾아 아래로'),
    @('CHOOSE','=CHOOSE(번호,값1,값2,...)','번호대로 하나 선택'),
    @('MATCH','=MATCH(값,한 줄 범위,0)','몇 번째인지 반환'),
    @('INDEX','=INDEX(범위,행,[열])','그 위치의 값 반환')
) @(2.8,8.8,5.8)
Add-Callout '마지막 확인' '정확히 찾기 = FALSE 또는 0 / 구간 찾기 = TRUE 또는 1 / 근사값 기준표 = 오름차순' $paleGreen $green
Add-Paragraph '실습 파일과 함께 열어 놓고, 각 그림의 시트명·셀 주소를 그대로 따라가면 됩니다.' 'Normal' $wdAlignParagraphCenter $gray $true 11

# Keep tables and images together where possible.
foreach($t in $doc.Tables) {
    try { $t.Rows.AllowBreakAcrossPages = 0 } catch {}
}

$outDir = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$doc.SaveAs2($OutputPath,$wdFormatXMLDocument)
$pdfPath = [System.IO.Path]::ChangeExtension($OutputPath,'.pdf')
$doc.ExportAsFixedFormat($pdfPath,$wdExportFormatPDF)
$pageCount = $doc.ComputeStatistics(2)
$doc.Close($false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
"DOCX=$OutputPath"
"PDF=$pdfPath"
"PAGES=$pageCount"
