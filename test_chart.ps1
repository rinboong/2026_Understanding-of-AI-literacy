$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    
    # 51 is xlColumnClustered
    $shape = $doc.Shapes.AddChart(51, $selection.Range)
    $chart = $shape.Chart
    
    # Get Chart Data
    $chartData = $chart.ChartData
    $chartData.Activate()
    
    # Poll for Workbook
    $workbook = $null
    for ($i = 0; $i -lt 10; $i++) {
        try {
            $workbook = $chartData.Workbook
            if ($workbook -ne $null) { break }
        } catch { }
        Start-Sleep -Milliseconds 500
    }
    
    $worksheet = $workbook.Worksheets.Item(1)
    
    # Write headers
    $worksheet.Cells.Item(1, 1).Value2 = "Category"
    $worksheet.Cells.Item(1, 2).Value2 = "Value 1"
    $worksheet.Cells.Item(1, 3).Value2 = "Value 2"
    
    # Write 10 rows
    for ($r = 2; $r -le 11; $r++) {
        $worksheet.Cells.Item($r, 1).Value2 = "Cat $r"
        $worksheet.Cells.Item($r, 2).Value2 = [double](10 + $r)
        $worksheet.Cells.Item($r, 3).Value2 = [double](20 + $r)
    }
    
    # Construct source string
    $sheetName = $worksheet.Name
    $address = $worksheet.Range("A1:C11").Address()
    $sourceStr = "='${sheetName}'!$address"
    
    Write-Output "Source String: $sourceStr"
    $chart.SetSourceData($sourceStr)
    
    # Close Excel workbook and save
    $workbook.Close($true) # Save changes
    
    # Access series collection to verify it works
    $series2 = $chart.SeriesCollection(2)
    Write-Output "Series 2 Name: $($series2.Name)"
    $series2.AxisGroup = 2
    
    # Save Word document
    $doc.SaveAs("d:\test\test_chart.docx")
    $doc.Close()
    Write-Output "Success"
} catch {
    Write-Output "Error: $_"
} finally {
    $word.Quit()
}
