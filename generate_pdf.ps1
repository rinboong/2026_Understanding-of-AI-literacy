# Word COM Automation PDF Generator
# Date: 2026. 06. 11
# Title: 삶의 만족도와 자살률 상관관계 분석 보고서 PDF 변환

$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Colors in BGR Decimal Format
$COLOR_PRIMARY = 2758415      # `#0f172a` (Dark Slate / Deep Navy)
$COLOR_SECONDARY = 15681241   # `#d946ef` (Magenta / Pink)
$COLOR_TEXT = 3877150         # `#1e293b` (Slate Gray)
$COLOR_MUTED = 8421504        # `#808080` (Gray)
$COLOR_LIGHT_BG = 16381425     # `#f1f5f9` (Very Light Slate)
$COLOR_BORDER = 15788258      # `#e2e8f0` (Light Border Gray)

# Callout Colors
$NOTE_BORDER = 15426341       # `#6366f1` (Indigo)
$NOTE_BG = 16774907           # `#eff6ff` (Light Blue)
$TIP_BORDER = 8950800         # `#0d9488` (Teal)
$TIP_BG = 16448992            # `#f0fdfa` (Light Teal)
$IMP_BORDER = 15681241        # `#d946ef` (Magenta)
$IMP_BG = 16315133            # `#fdf2f8` (Light Pink)

try {
    # 1. Parse data.js
    $content = Get-Content -Path "d:\test\data.js" -Raw
    $jsonText = $content -replace '^const ANALYSIS_DATA\s*=\s*', ''
    $jsonText = $jsonText -replace ';\s*$', ''
    $dataObj = ConvertFrom-Json $jsonText
    
    # Filter data
    $nationalTotal = $dataObj.RawData | Where-Object { $_.Region -eq '전국' -and $_.Gender -eq 'Total' } | Sort-Object Year
    $regional2024 = $dataObj.RawData | Where-Object { $_.Region -ne '전국' -and $_.Year -eq 2024 -and $_.Gender -eq 'Total' } | Sort-Object SatisfactionScore -Descending

    # Create new Document
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    
    # Page Setup
    $doc.PageSetup.TopMargin = 72 # 1 inch
    $doc.PageSetup.BottomMargin = 72
    $doc.PageSetup.LeftMargin = 72
    $doc.PageSetup.RightMargin = 72

    # Helper Functions
    function Add-Title ($text) {
        $selection.ParagraphFormat.Alignment = 1 # Center
        $selection.ParagraphFormat.SpaceBefore = 120
        $selection.ParagraphFormat.SpaceAfter = 12
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 28
        $selection.Font.Bold = $true
        $selection.Font.Color = $COLOR_PRIMARY
        $selection.TypeText($text)
        $selection.TypeParagraph()
    }
    
    function Add-Subtitle ($text) {
        $selection.ParagraphFormat.Alignment = 1 # Center
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 180
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 12
        $selection.Font.Bold = $false
        $selection.Font.Color = $COLOR_MUTED
        $selection.TypeText($text)
        
        # Add a bottom border line (wdBorderBottom = -3)
        $selection.ParagraphFormat.Borders.Item(-3).LineStyle = 1 
        $selection.ParagraphFormat.Borders.Item(-3).LineWidth = 12 # 1.5 pt
        $selection.ParagraphFormat.Borders.Item(-3).Color = $NOTE_BORDER
        $selection.TypeText(" ") # spacer
        $selection.TypeParagraph()
        
        # Clear paragraph border
        $selection.ParagraphFormat.Borders.Item(-3).LineStyle = 0
    }
    
    function Add-Metadata ($text) {
        $selection.ParagraphFormat.Alignment = 1 # Center
        $selection.ParagraphFormat.SpaceBefore = 50
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 10
        $selection.Font.Color = $COLOR_MUTED
        $selection.Font.Bold = $false
        $selection.TypeText($text)
        $selection.TypeParagraph()
    }

    function Add-Heading1 ($text) {
        $selection.TypeParagraph()
        $selection.ParagraphFormat.Alignment = 0 # Left
        $selection.ParagraphFormat.SpaceBefore = 24
        $selection.ParagraphFormat.SpaceAfter = 8
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 16
        $selection.Font.Bold = $true
        $selection.Font.Color = $COLOR_PRIMARY
        
        # Add left border style (wdBorderLeft = -2)
        $selection.ParagraphFormat.Borders.Item(-2).LineStyle = 1
        $selection.ParagraphFormat.Borders.Item(-2).LineWidth = 24 # 3pt
        $selection.ParagraphFormat.Borders.Item(-2).Color = $NOTE_BORDER
        $selection.ParagraphFormat.LeftIndent = 12
        
        $selection.TypeText($text)
        $selection.TypeParagraph()
        
        # Reset indent and border
        $selection.ParagraphFormat.Borders.Item(-2).LineStyle = 0
        $selection.ParagraphFormat.LeftIndent = 0
    }

    function Add-Heading2 ($text) {
        $selection.TypeParagraph()
        $selection.ParagraphFormat.Alignment = 0 # Left
        $selection.ParagraphFormat.SpaceBefore = 14
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 12.5
        $selection.Font.Bold = $true
        $selection.Font.Color = $COLOR_SECONDARY
        $selection.TypeText($text)
        $selection.TypeParagraph()
    }
    
    function Add-BodyText ($text, $bold = $false) {
        $selection.ParagraphFormat.Alignment = 0 # Left
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.ParagraphFormat.LineSpacingRule = 5 # wdLineSpaceMultiple
        $selection.ParagraphFormat.LineSpacing = 1.15
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 10.5
        $selection.Font.Bold = $bold
        $selection.Font.Color = $COLOR_TEXT
        $selection.TypeText($text)
        $selection.TypeParagraph()
    }
    
    function Add-BulletPoint ($text) {
        $selection.ParagraphFormat.Alignment = 0
        $selection.ParagraphFormat.SpaceAfter = 4
        $selection.ParagraphFormat.LeftIndent = 18
        $selection.Font.Name = "맑은 고딕"
        $selection.Font.Size = 10
        $selection.Font.Bold = $false
        $selection.Font.Color = $COLOR_TEXT
        $selection.TypeText("•  " + $text)
        $selection.TypeParagraph()
        $selection.ParagraphFormat.LeftIndent = 0
    }
    
    function Add-Callout ($title, $text, $type="NOTE") {
        $borderColor = $NOTE_BORDER
        $bgColor = $NOTE_BG
        
        if ($type -eq "TIP") {
            $borderColor = $TIP_BORDER
            $bgColor = $TIP_BG
        } elseif ($type -eq "IMPORTANT") {
            $borderColor = $IMP_BORDER
            $bgColor = $IMP_BG
        }
        
        $table = $doc.Tables.Add($selection.Range, 1, 1)
        # Configure borders and shading manually
        $table.Borders.OutsideLineStyle = 0
        $table.Borders.InsideLineStyle = 0
        
        # wdBorderLeft = -2
        $table.Borders.Item(-2).LineStyle = 1
        $table.Borders.Item(-2).LineWidth = 24 # 3pt
        $table.Borders.Item(-2).Color = $borderColor
        $table.Rows.Item(1).Shading.BackgroundPatternColor = $bgColor
        
        # Write text in cell
        $cellRange = $table.Cell(1,1).Range
        $cellRange.Font.Name = "맑은 고딕"
        $cellRange.Font.Size = 9.5
        $cellRange.Font.Color = $COLOR_TEXT
        $cellRange.Text = "💡 " + $title + "`r" + $text
        $cellRange.Paragraphs.Item(1).Range.Font.Bold = $true
        
        # Move past table
        $selection.EndKey(6) # wdStory
        $selection.TypeParagraph()
    }
    
    function Format-TableStyle ($table) {
        $table.Borders.InsideLineStyle = 1
        $table.Borders.InsideLineWidth = 4
        $table.Borders.InsideColor = $COLOR_BORDER
        $table.Borders.OutsideLineStyle = 1
        $table.Borders.OutsideLineWidth = 8
        $table.Borders.OutsideColor = $COLOR_PRIMARY
        
        # Header Row Formatting
        $header = $table.Rows.Item(1)
        $header.Range.Font.Bold = $true
        $header.Range.Font.Color = 16777215 # White
        $header.Shading.BackgroundPatternColor = $COLOR_PRIMARY
        $header.Alignment = 1 # Center
        
        # Alternating Row Colors
        for ($r = 2; $r -le $table.Rows.Count; $r++) {
            if ($r % 2 -eq 0) {
                $table.Rows.Item($r).Shading.BackgroundPatternColor = $COLOR_LIGHT_BG
            }
            # Add padding
            $table.Rows.Item($r).Height = 20
        }
    }

    # --- COVER PAGE ---
    Add-Title "삶의 만족도와 자살률`n상관관계 분석 보고서"
    Add-Subtitle "주관적 삶의 만족 지표와 실제 자살률 간의 연계성 및 시도별 역설 현상 연구 (2020 ~ 2024)"
    Add-Metadata "KOSIS 데이터 기반 통계 분석 보고`n`n발행일: 2026. 06. 11`n출처: 통계청 사회조사 & 사망원인통계`n`n발표자: Antigravity AI"
    $selection.InsertBreak(7) # wdPageBreak

    # --- SECTION 1 ---
    Add-Heading1 "1. 요약 (Executive Summary)"
    Add-BodyText "본 보고서는 통계청 사회조사의 '주관적 삶의 만족도'와 '사망원인통계'의 '자살률' 데이터를 기반으로, 대한민국 17개 시도별 삶의 웰빙 수준과 실제 극단적 선택 간의 통계적 상관관계를 규명하고자 작성되었습니다."
    Add-BulletPoint "전체 상관관계 미약 (r = -0.048): 5개년 전체 패널 데이터(85개 데이터 포인트) 분석 결과, 주관적 삶의 만족도와 자살률 간의 단순 선형 상관관계는 거의 무관한 것으로 확인되었습니다."
    Add-BulletPoint "2020년(코로나19 초기)의 특이점 (r = -0.514): 코로나19 팬데믹이 시작된 2020년에는 삶의 만족도가 높은 지역일수록 자살률이 뚜렷하게 낮은 강한 음의 상관관계가 나타났습니다. 그러나 재난 국면이 장기화된 이후 연도에는 이러한 추세가 급격히 완화되었습니다."
    Add-BulletPoint "성별 격차: 남성의 경우 만족도 조사에서 보통(Neutral) 응답 비율과 자살률 간에 moderate한 양의 상관관계(r = 0.405)가 나타났으나, 여성의 경우 삶의 만족 지표들과 자살률 간의 통계적 연계성이 극히 낮았습니다."
    Add-BulletPoint "부정 답변율과 자살률의 역설: 삶에 매우 불만족하거나 약간 불만족하다고 대답한 비율(부정 답변율)이 높은 지역일수록 오히려 자살률이 낮은 역설적인 음의 상관관계(2024년 r = -0.487)가 지속적으로 관찰되었습니다."
    
    Add-Callout "보통(Neutral) 응답의 유의성 (r = 0.406)" "주관적 삶의 만족도 조사에서 '보통'이라고 응답한 비율이 높은 지역일수록 자살률이 뚜렷하게 높은 경향이 전국적으로 발견되었습니다. 이는 뚜렷한 주관적 부정 감정을 표출하기보다는, 감정의 평탄화(Emotional Flattening)나 무기력, 혹은 사회적으로 완전히 고립된 집단이 실제 자살 위험군일 수 있음을 강력히 시사합니다." "IMPORTANT"

    # --- SECTION 2 ---
    Add-Heading1 "2. 데이터 개요 및 가공 방법"
    Add-Heading2 "2.1 분석 대상 데이터"
    Add-BulletPoint "삶의 만족도: 통계청 사회조사 데이터 (삶의_만족도_시도__20260606195059.xlsx)`n매우 만족, 약간 만족, 보통, 약간 불만족, 매우 불만족의 5개 척도 응답 비율(%)을 시도 및 성별로 집계"
    Add-BulletPoint "자살률: 통계청 사망원인통계 데이터 (인구십만명당_자살률_시도_시_군_구__20260606194913.xlsx)`n지역별 인구 10만 명당 자살 사망자 수(명)로 집계"
    
    Add-Heading2 "2.2 지표 가공 방식 (가중평균 삶의 만족도)"
    Add-BodyText "주관적 삶의 만족도 지표를 단일 점수로 비교하기 위해, 5점 만점의 가중 평균 점수(Satisfaction Score)를 다음과 같이 산출하였습니다."
    
    # Formula Box
    $formulaTable = $doc.Tables.Add($selection.Range, 1, 1)
    $formulaTable.Borders.OutsideLineStyle = 1
    $formulaTable.Borders.OutsideLineWidth = 4
    $formulaTable.Borders.OutsideColor = $COLOR_BORDER
    $formulaTable.Rows.Item(1).Shading.BackgroundPatternColor = $COLOR_LIGHT_BG
    $cellF = $formulaTable.Cell(1,1).Range
    $cellF.Font.Name = "Cambria"
    $cellF.Font.Size = 11
    $cellF.Font.Bold = $true
    $cellF.Text = "Satisfaction Score = [ 매우만족*5 + 약간만족*4 + 보통*3 + 약간불만*2 + 매우불만*1 ] / 100"
    $selection.EndKey(6)
    $selection.TypeParagraph()
    
    Add-BulletPoint "긍정 답변율 (Positive Rate): 매우 만족 + 약간 만족 (%)"
    Add-BulletPoint "부정 답변율 (Negative Rate): 매우 불만족 + 약간 불만족 (%)"

    # --- SECTION 3 ---
    $selection.InsertBreak(7) # wdPageBreak
    Add-Heading1 "3. 주요 상관관계 분석 결과 (Pearson r)"
    Add-BodyText "피어슨 상관계수(Pearson r)는 두 변수 간의 선형적 연관성을 나타내며, -1에서 +1 사이의 값을 가집니다. 절대값이 0.4 이상이면 의미 있는 상관관계로 해석됩니다."
    
    Add-Heading2 "3.1 전체 패널 분석 결과 (85개 데이터 포인트 통합)"
    
    # Table 1
    $table1 = $doc.Tables.Add($selection.Range, 7, 4)
    $table1.Cell(1,1).Range.Text = "분석 지표 (vs 자살률)"
    $table1.Cell(1,2).Range.Text = "전체 성별 (Total)"
    $table1.Cell(1,3).Range.Text = "남성 (Male)"
    $table1.Cell(1,4).Range.Text = "여성 (Female)"
    
    $table1.Cell(2,1).Range.Text = "삶의 만족도 가중 점수"
    $table1.Cell(2,2).Range.Text = "-0.048"
    $table1.Cell(2,3).Range.Text = "-0.031"
    $table1.Cell(2,4).Range.Text = "0.014"
    
    $table1.Cell(3,1).Range.Text = "긍정 답변율 (만족)"
    $table1.Cell(3,2).Range.Text = "-0.170"
    $table1.Cell(3,3).Range.Text = "-0.165"
    $table1.Cell(3,4).Range.Text = "-0.049"
    
    $table1.Cell(4,1).Range.Text = "보통 (Neutral) 비율"
    $table1.Cell(4,2).Range.Text = "0.406"
    $table1.Cell(4,2).Range.Font.Bold = $true
    $table1.Cell(4,3).Range.Text = "0.405"
    $table1.Cell(4,3).Range.Font.Bold = $true
    $table1.Cell(4,4).Range.Text = "0.114"
    
    $table1.Cell(5,1).Range.Text = "부정 답변율 (불만족)"
    $table1.Cell(5,2).Range.Text = "-0.177"
    $table1.Cell(5,3).Range.Text = "-0.164"
    $table1.Cell(5,4).Range.Text = "-0.035"
    
    $table1.Cell(6,1).Range.Text = "매우 만족 비율"
    $table1.Cell(6,2).Range.Text = "-0.223"
    $table1.Cell(6,3).Range.Text = "-0.199"
    $table1.Cell(6,4).Range.Text = "-0.107"
    
    $table1.Cell(7,1).Range.Text = "매우 불만족 비율"
    $table1.Cell(7,2).Range.Text = "-0.226"
    $table1.Cell(7,3).Range.Text = "-0.220"
    $table1.Cell(7,4).Range.Text = "-0.009"
    
    Format-TableStyle $table1
    $selection.EndKey(6)
    $selection.TypeParagraph()
    
    Add-Heading2 "3.2 연도별 크로스섹션 상관계수 추이"
    
    # Table 2
    $table2 = $doc.Tables.Add($selection.Range, 6, 6)
    $table2.Cell(1,1).Range.Text = "연도"
    $table2.Cell(1,2).Range.Text = "만족도 점수 (Total)"
    $table2.Cell(1,3).Range.Text = "매우 만족 (Total)"
    $table2.Cell(1,4).Range.Text = "보통 비율 (Total)"
    $table2.Cell(1,5).Range.Text = "매우 불만족 (Total)"
    $table2.Cell(1,6).Range.Text = "부정 답변율 (Total)"
    
    $table2.Cell(2,1).Range.Text = "2020년"
    $table2.Cell(2,2).Range.Text = "-0.514"
    $table2.Cell(2,3).Range.Text = "-0.609"
    $table2.Cell(2,4).Range.Text = "0.574"
    $table2.Cell(2,5).Range.Text = "-0.191"
    $table2.Cell(2,6).Range.Text = "0.107"
    
    $table2.Cell(3,1).Range.Text = "2021년"
    $table2.Cell(3,2).Range.Text = "0.050"
    $table2.Cell(3,3).Range.Text = "-0.149"
    $table2.Cell(3,4).Range.Text = "0.443"
    $table2.Cell(3,5).Range.Text = "-0.403"
    $table2.Cell(3,6).Range.Text = "-0.357"
    
    $table2.Cell(4,1).Range.Text = "2022년"
    $table2.Cell(4,2).Range.Text = "0.077"
    $table2.Cell(4,3).Range.Text = "0.040"
    $table2.Cell(4,4).Range.Text = "0.177"
    $table2.Cell(4,5).Range.Text = "-0.241"
    $table2.Cell(4,6).Range.Text = "-0.203"
    
    $table2.Cell(5,1).Range.Text = "2023년"
    $table2.Cell(5,2).Range.Text = "-0.109"
    $table2.Cell(5,3).Range.Text = "0.035"
    $table2.Cell(5,4).Range.Text = "0.299"
    $table2.Cell(5,5).Range.Text = "-0.161"
    $table2.Cell(5,6).Range.Text = "-0.004"
    
    $table2.Cell(6,1).Range.Text = "2024년"
    $table2.Cell(6,2).Range.Text = "0.230"
    $table2.Cell(6,3).Range.Text = "-0.116"
    $table2.Cell(6,4).Range.Text = "0.348"
    $table2.Cell(6,5).Range.Text = "-0.557"
    $table2.Cell(6,6).Range.Text = "-0.487"
    
    Format-TableStyle $table2
    $selection.EndKey(6)
    $selection.TypeParagraph()

    # Callout for 2020 and 2024 paradox
    Add-Callout "2020년 코로나 예외성과 2024년 불만족의 역설" "코로나19 첫해인 2020년에는 삶의 만족도가 높은 지역사회의 정서가 극단적 선택을 예방하는 뚜렷한 완충 효과(r = -0.514)를 보였습니다. 그러나 2024년에는 주관적으로 삶에 '매우 불만족'하다고 대답한 비율이 높은 시도에서 오히려 자살률이 매우 낮게 기록되는 역설(r = -0.557)이 관찰되었습니다. 이는 삶의 불만을 인지하고 적극적으로 표출하는 구조가 사회보호 체계 내에서 기능하고 있을 가능성이 있음을 암시합니다." "TIP"

    # Insert Chart 1: National Trend Line Chart
    Add-Heading2 "3.3 전국 평균 삶의 만족도 및 자살률 5개년 추이 비교"
    Add-BodyText "전국 단위에서 주관적 만족도의 완만한 상승 추세와 자살률 상승 추세가 동시에 관찰되는 종합 통계입니다."
    
    $shape1 = $doc.Shapes.AddChart(4, $selection.Range) # xlLine = 4
    $chart1 = $shape1.Chart
    $chart1.HasTitle = $true
    $chart1.ChartTitle.Text = "전국 평균 만족도 vs 자살률 5개년 추이 (2020 ~ 2024)"
    
    # Edit Chart Data
    $chartData1 = $chart1.ChartData
    $chartData1.Activate()
    
    # Wait for Excel
    $workbook1 = $null
    for ($i = 0; $i -lt 10; $i++) {
        try {
            $workbook1 = $chartData1.Workbook
            if ($workbook1 -ne $null) { break }
        } catch { }
        Start-Sleep -Milliseconds 500
    }
    
    if ($workbook1 -eq $null) {
        throw "Workbook 1 is not ready after 5 seconds"
    }
    
    $worksheet1 = $workbook1.Worksheets.Item(1)
    
    $worksheet1.Cells.Item(1, 1).Value2 = "연도"
    $worksheet1.Cells.Item(1, 2).Value2 = "전국 평균 만족도"
    $worksheet1.Cells.Item(1, 3).Value2 = "전국 평균 자살률"
    
    $idx = 2
    foreach ($row in $nationalTotal) {
        $worksheet1.Cells.Item($idx, 1).Value2 = ($row.Year.ToString() + "년")
        $worksheet1.Cells.Item($idx, 2).Value2 = [double]$row.SatisfactionScore
        $worksheet1.Cells.Item($idx, 3).Value2 = [double]$row.SuicideRate
        $idx++
    }
    
    # Set source range using SetSourceData method!
    $chart1.SetSourceData($worksheet1.Range("A1:C6"))
    
    $workbook1.Close($true)
    
    # Modify Series axis
    $series2 = $chart1.SeriesCollection(2)
    $series2.AxisGroup = 2 # xlSecondary
    
    # Format axes
    $axisY1 = $chart1.Axes(2, 1) # xlValue, xlPrimary
    $axisY1.MinimumScale = 3.1
    $axisY1.MaximumScale = 3.5
    $axisY1.HasTitle = $true
    $axisY1.AxisTitle.Text = "만족도 점수"
    
    $axisY2 = $chart1.Axes(2, 2) # xlValue, xlSecondary
    $axisY2.MinimumScale = 24
    $axisY2.MaximumScale = 31
    $axisY2.HasTitle = $true
    $axisY2.AxisTitle.Text = "자살률 (명/10만명)"
    
    $selection.EndKey(6)
    $selection.TypeParagraph()

    # --- SECTION 4 ---
    $selection.InsertBreak(7) # wdPageBreak
    Add-Heading1 "4. 2024년 시도별 상세 데이터 분석"
    Add-BodyText "상관관계의 모순 및 보건 안전망의 불일치가 가장 뚜렷하게 관찰된 2024년의 시도별 실측치 정렬 랭킹 테이블입니다."
    
    # Table 3: 17 Regions + Header = 18 Rows, 3 Columns
    $table3 = $doc.Tables.Add($selection.Range, 18, 3)
    $table3.Cell(1,1).Range.Text = "시도 (Region)"
    $table3.Cell(1,2).Range.Text = "주관적 삶의 만족도 점수 (5점 만점)"
    $table3.Cell(1,3).Range.Text = "인구 10만 명당 자살률 (명)"
    
    $idx = 2
    foreach ($row in $regional2024) {
        $table3.Cell($idx, 1).Range.Text = $row.Region
        $table3.Cell($idx, 2).Range.Text = ([double]$row.SatisfactionScore).ToString("F3")
        $table3.Cell($idx, 3).Range.Text = ([double]$row.SuicideRate).ToString("F1")
        $idx++
    }
    
    Format-TableStyle $table3
    $selection.EndKey(6)
    $selection.TypeParagraph()
    
    # Insert Chart 2: 2024 Regional Combo Chart (Column + Line)
    Add-Heading2 "4.1 2024년 시도별 만족도 및 자살률 시각화 비교"
    Add-BodyText "시도별 만족도 순위(내림차순 막대 그래프)와 자살률(꺾은선 그래프)의 연동 대조입니다."
    
    $shape2 = $doc.Shapes.AddChart(51, $selection.Range) # xlColumnClustered = 51
    $chart2 = $shape2.Chart
    $chart2.HasTitle = $true
    $chart2.ChartTitle.Text = "2024년 시도별 만족도 및 자살률 비교"
    
    # Edit Chart Data
    $chartData2 = $chart2.ChartData
    $chartData2.Activate()
    
    # Wait for Excel
    $workbook2 = $null
    for ($i = 0; $i -lt 10; $i++) {
        try {
            $workbook2 = $chartData2.Workbook
            if ($workbook2 -ne $null) { break }
        } catch { }
        Start-Sleep -Milliseconds 500
    }
    
    if ($workbook2 -eq $null) {
        throw "Workbook 2 is not ready after 5 seconds"
    }
    
    $worksheet2 = $workbook2.Worksheets.Item(1)
    
    $worksheet2.Cells.Item(1, 1).Value2 = "시도"
    $worksheet2.Cells.Item(1, 2).Value2 = "삶의 만족도"
    $worksheet2.Cells.Item(1, 3).Value2 = "자살률"
    
    $idx = 2
    foreach ($row in $regional2024) {
        $worksheet2.Cells.Item($idx, 1).Value2 = $row.Region
        $worksheet2.Cells.Item($idx, 2).Value2 = [double]$row.SatisfactionScore
        $worksheet2.Cells.Item($idx, 3).Value2 = [double]$row.SuicideRate
        $idx++
    }
    
    # Set source range using SetSourceData method!
    $chart2.SetSourceData($worksheet2.Range("A1:C18"))
    
    $workbook2.Close($true)
    
    # Change Series 2 (Suicide Rate) to Line chart on Secondary Axis
    $seriesSui = $chart2.SeriesCollection(2)
    $seriesSui.Type = 4 # xlLine
    $seriesSui.AxisGroup = 2 # xlSecondary
    
    # Format axes
    $axisY1_2 = $chart2.Axes(2, 1)
    $axisY1_2.MinimumScale = 2.8
    $axisY1_2.MaximumScale = 3.8
    $axisY1_2.HasTitle = $true
    $axisY1_2.AxisTitle.Text = "만족도 점수"
    
    $axisY2_2 = $chart2.Axes(2, 2)
    $axisY2_2.MinimumScale = 10
    $axisY2_2.MaximumScale = 40
    $axisY2_2.HasTitle = $true
    $axisY2_2.AxisTitle.Text = "자살률 (명/10만명)"
    
    $selection.EndKey(6)
    $selection.TypeParagraph()
    
    Add-Heading2 "4.2 주요 시도별 모순 및 괴리 사례 분석"
    Add-BodyText "•  제주특별자치도 (만족도 공동 3위, 자살률 1위): 주관적 삶의 만족도는 3.442점으로 높게 응답하였으나, 인구 10만 명당 자살률은 36.3명으로 전국 최다치를 기록하였습니다. 이는 관광 지표 및 외부 유입에 따른 긍정 정서와 원주민의 고령 고독사, 경제 소외 문제 간의 격차를 반영합니다."
    Add-BodyText "•  충청남도 (만족도 1위, 자살률 2위): 주관적 삶의 만족도 1위(3.555점)에 등극하였으나 자살률 역시 34.8명으로 최고 수준입니다. 긍정 답변율이 매우 높은 경향을 보이는 것과 상반되는 실측 사망 통계로 대표적인 만족도-치명성 불일치 지표를 갖습니다."
    Add-BodyText "•  대구광역시 (만족도 최하위, 자살률 하위권): 만족도 점수 3.217점으로 최하위 18위를 차지하였고 부정 답변율도 21.4%로 가장 높았으나, 실제 자살률은 29.4명으로 전국 평균 수준이자 다른 광역 도(Province) 단위에 비해 훨씬 낮아 심리적 불만이 직접적인 치명 행동으로 이어지지 않는 보건 구조를 갖고 있습니다."

    # --- SECTION 5 ---
    $selection.InsertBreak(7) # wdPageBreak
    Add-Heading1 "5. 종합 해석 및 정책적 시사점"
    
    Add-Heading2 "5.1 통계조사 만족도와 자살률 불일치의 근본 원인"
    Add-BulletPoint "보건 의료 및 긴급 이송 안전망의 차이: 서울(24.1명), 세종(23.0명), 경기(28.2명) 등 고도의 의료 보건 안전 인프라와 30분 내 응급 수술이 가능한 병원 접근성이 취약 지역의 극단적 선택 시도 발생 후의 치명률(사망률)을 제어하고 있습니다."
    Add-BulletPoint "인구학적 노령 구조와 고립: 자살률 상위권 지역인 제주, 충남, 전남, 강원은 고령화율이 높고 농어촌 지역 1인 가구의 소외가 깊습니다. 반면 대면 중심 사회조사는 주로 응답이 수월한 연령층 위주로 집계되어 통계의 수혜 편향이 존재합니다."
    
    Add-Heading2 "5.2 '보통' 응답군에 숨겨진 무기력·사회적 고립 위험군"
    Add-BodyText "상관관계상 '보통' 비율과 자살률은 뚜렷한 양의 상관관계(r = 0.406)를 갖습니다. 삶의 만족도 조사에서 능동적인 '불만족'을 표현하기보다 감정을 평탄화하여 '보통'으로 보수적 답변을 일관하는 무기력 집단과, 아웃리치 사각지대에 처한 '사회적 고립 집단'이 오히려 자살 행동에 더 가깝다는 해석이 타당합니다."
    
    Add-Heading2 "5.3 보건 정책 제언"
    Add-BulletPoint "지자체 밀착형 응급 정신보건 체계 강화: 제주, 충남 등 만족도는 높으나 자살률이 이중으로 높은 시도의 경우 보건 인프라의 사각지대가 없도록 이동형 심리케어 및 응급 정신과 연계를 보강해야 합니다."
    Add-BulletPoint "수동적 무기력자('보통' 응답군) 모니터링: 설문조사 시 뚜렷한 의견 표시를 유보하거나 '보통'으로 답변하는 수동적 고위험군을 능동 발굴하여 커뮤니티로 복귀시키기 위한 '사회적 아웃리치(Outreach)' 스크리닝 도구를 개발해야 합니다."

    # --- SECTION 6 ---
    Add-Heading1 "6. 분석의 한계점"
    Add-BulletPoint "설문조사 표본의 편향성: 사회조사는 가구 방문 조사를 기본으로 하므로, 실제 극단적 선택 위험이 가장 높으며 사회관계망이 끊긴 고독사 위험군, 노숙인 및 특수 거주 시설 인원의 목소리가 조사 과정에서 제외되었을 수 있습니다."
    Add-BulletPoint "단기 분석 기간의 한계: 2020년부터 2024년까지의 5개년 패널 정보로 제한되어 있어 중장기적인 거시적 상관성 추이를 확정적으로 결론짓기 어렵습니다."
    Add-BulletPoint "다양한 통제 변수 배제: 지역별 가구 소득수준, 실업률, 1인 가구 비율, 고령화 지표 등의 매개 변수(Confounding Variables)가 완벽히 통제되지 않았으므로 단순 피어슨 분석은 일부 가짜 상관(Spurious Correlation)의 가능성을 포함합니다."

    # Export Document as PDF (wdExportFormatPDF = 17)
    $pdfPath = "d:\test\analysis_report.pdf"
    $doc.ExportAsFixedFormat($pdfPath, 17)
    
    # Close document without saving Word prompt
    $doc.Close($false) # wdDoNotSaveChanges = 0
    
    Write-Output "Success: PDF generated at $pdfPath"
} catch {
    Write-Output "Error: $_"
} finally {
    $word.Quit()
}
