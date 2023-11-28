# IMPORTANT: This script shows different examples, and it's best to run each part separately. For detailed instructions, check out the corresponding YouTube video.

# ------------------------------------------------------------------------
# Demo 1: Use the raw data to generate the Excel file with the first sheet
# ------------------------------------------------------------------------

$SalesGenerated = ".\salesgenerated.xlsx"
Import-Csv rawsales.csv | Export-Excel -Path $SalesGenerated -WorksheetName "Raw Data"


# Open the generated excel file
Invoke-Item $SalesGenerated

# ----------------------------------------------------------------------------------
# Demo 2: Convert csv to json and add an additional sheet to the existing Excel file
# ----------------------------------------------------------------------------------

# Convert the existing .csv to json format and write to disk
Import-Csv .\estimation.csv | ConvertTo-Json | Set-Content .\estimation.json

# Import JSON and write to estimation tab
Get-Content .\estimation.json | ConvertFrom-Json | Export-Excel -Path $SalesGenerated -WorksheetName "Estimation"


# -----------------------------------------------------------------
# Demo 3: Add Excel Formula and Colors (Mimic the sheet Estimation)
# -----------------------------------------------------------------

# Open the Excel package
$excelPackage = Open-ExcelPackage -Path $SalesGenerated

# Access the Estimation worksheet
$worksheet = $excelPackage.Workbook.Worksheets["Estimation"]

# Add the formula to cell E2 through to E5
2..5 | ForEach-Object {
    $worksheet.Cells["E$_"].Formula = "B$_*(1-C$_-D$_)"
}

# Add the sum to E6
$worksheet.Cells["E6"].Formula = "SUM(E2:E5)*3.99"

# Let's add some color
$color = [System.Drawing.Color]::Yellow
Set-ExcelRange -Worksheet $worksheet -Range "E1" -BackgroundColor $color
Set-ExcelRange -Worksheet $worksheet -Range "E6" -BackgroundColor $color

# Save and close the Excel package
Close-ExcelPackage -ExcelPackage $excelPackage -Show


# ----------------------------------------------------------------
### Demo 4: Resize columns to autofit (Mimic the sheet Estimation)
# ----------------------------------------------------------------


# Open the Excel package
$excelPackage = Open-ExcelPackage -Path $SalesGenerated 

# Access the Estimation worksheet
$worksheet = $excelPackage.Workbook.Worksheets["Estimation"]

# Auto size columns for the entire worksheet
$worksheet.Dimension.Columns | ForEach-Object {
    Set-ExcelColumn -Worksheet $worksheet -Column $_ -AutoFit
}

# Save and close the Excel package
Close-ExcelPackage -ExcelPackage $excelPackage -Show

# ------------------------------------------------
# Demo 5: Add Excel Formular (Mimic Forcast Sheet)
# ------------------------------------------------

# Open the Excel package
$excelPackage = Open-ExcelPackage -Path $SalesGenerated

# Access the Estimation worksheet
$worksheet = $excelPackage.Workbook.Worksheets["Raw Data"]

#Set-ExcelColumn -Worksheet $worksheet -Column 12  "Weight"
Set-ExcelColumn -Worksheet $worksheet -Column 12 -Heading "Weight"  -AutoSize


# Add Weights
$worksheet.Cells["L2"].Value = 0.1
$worksheet.Cells["L3"].Value = 0.2
$worksheet.Cells["L4"].Value = 0.7



# Add the formula to cell E2 through to E5
5..7 | ForEach-Object {
    $worksheet.Cells["G$_"].Formula = "AVERAGE(B$($_-3):B$($_-1))"
     $worksheet.Cells["H$_"].Formula = "SUMPRODUCT(B$($_-3):B$($_-1),`$L2:`$L4)"
    $worksheet.Cells["I$_"].Formula = "_xlfn.FORECAST.ETS(A$_,B$($_-3):B$($_-1),A$($_-3):A$($_-1))"
    
}

# Tip: To troubleshoot, one way is to see the properties of the cell:
#$worksheet.Cells["I5"]
#$worksheet.Cells["I6"]



# Save and close the Excel package
Close-ExcelPackage -ExcelPackage $excelPackage -Show

# -----------------------------------------
# Demo 6:  Add charts (Mimic Forcast Sheet)
# -----------------------------------------

# Open the Excel package
$excelPackage = Open-ExcelPackage -Path $SalesGenerated


# This is why we splat, we don't like:
$Chart1 = New-ExcelChartDefinition -ChartType 'line' -XRange "Date" -YRange "Sales", "Forecast_1" -Title "Forecast Method 1: Moving Average Sales Of The Last 3 Days" -TitleBold $true -Width 800 -Row 1 -Column 14 -LegendPosition 'Bottom' -SeriesHeader "Sales", "Forecast_1"


# We also don't like this:
$Chart1 = New-ExcelChartDefinition -ChartType line -XRange "Date" -YRange "Sales" -Title "Forecast Method 1: Moving Average Sales Of The Last 3 Days" -TitleBold -Width 800 -Row 3 -Column 14
Export-Excel -Path $SalesGenerated -WorksheetName "Raw Data" -ExcelChartDefinition $Chart1 -Show -AutoNameRange


# And we definitely don't like this
$Chart1 = New-ExcelChartDefinition -ChartType 'line'`
    -XRange "Date"`
    -YRange "Sales", "Forecast_1"`
    -Title "Forecast Method 1: Moving Average Sales Of The Last 3 Days"`
    -TitleBold`
    -Width 800`
    -Row 1`
    -Column 14`
    -LegendPosition 'Bottom'`
    -SeriesHeader "Sales", "Forecast_1"

# This is why we splat, we like this
$Chart1Splat = @{
        ChartType = 'line'
        XRange    = "Date"
        YRange    = "Sales", "Forecast_1"
        Title     = "Forecast Method 1: Moving Average Sales Of The Last 3 Days"
        TitleBold = $true
        Width     = 800
        Row       = 1
        Column    = 14
        LegendPosition = 'Bottom'
        SeriesHeader = "Sales", "Forecast_1"
    }
$Chart2Splat = @{
        ChartType = 'line'
        XRange    = "Date"
        YRange    = "Sales", "Forecast_2"
        Title     = "Forecast Method 2: Weighted Moving Average Sales Of The Last 3 Days"
        TitleBold = $true
        Width     = 800
        Row       = 21
        Column    = 14
        LegendPosition = 'Bottom'
        SeriesHeader = "Sales", "Forecast_2"
    }
$Chart3Splat = @{
        ChartType = 'line'
        XRange    = "Date"
        YRange    = "Sales", "Forecast_3"
        Title     = "Forecast Method 3: Holt-Winters Method"
        TitleBold = $true
        Width     = 800
        Row       = 41
        Column    = 14
        LegendPosition = 'Bottom'
        SeriesHeader = "Sales", "Forecast_3"
    }



$Chart1 = New-ExcelChartDefinition @Chart1Splat
$Chart2 = New-ExcelChartDefinition @Chart2Splat
$Chart3 = New-ExcelChartDefinition @Chart3Splat

# And let's export it
Export-Excel -Path $SalesGenerated -WorksheetName "Raw Data" -ExcelChartDefinition $Chart1,$Chart2,$Chart3 -Show -AutoNameRange


# Same sample different way of using splatting, leading into how this way of working gets you to creating functions
$ChartSplat = @{
    ChartType = 'line'
    XRange    = "Date"
    YRange    = "Sales", "Forecast_1"
    Title     = "Forecast Method 1: Moving Average Sales Of The Last 3 Days"
    TitleBold = $true
    Width     = 800
    Row       = 1
    Column    = 14
    LegendPosition = 'Bottom'
    SeriesHeader = "Sales", "Forecast_1"
}

$Chart1 = New-ExcelChartDefinition @ChartSplat

$ChartSplat.YRange       = "Sales", "Forecast_2"
$ChartSplat.Title        = "Forecast Method 2: Weighted Moving Average Sales Of The Last 3 Days"
$ChartSplat.Row          = 21
$ChartSplat.SeriesHeader = "Sales", "Forecast_2"

$Chart2 = New-ExcelChartDefinition @ChartSplat

$ChartSplat.YRange       = "Sales", "Forecast_3"
$ChartSplat.Title        = "Forecast Method 3: Holt-Winters Method"
$ChartSplat.Row          = 41
$ChartSplat.SeriesHeader = "Sales", "Forecast_3"


$Chart3 = New-ExcelChartDefinition @ChartSplat

Export-Excel -Path $SalesGenerated -WorksheetName "Raw Data" -ExcelChartDefinition $Chart1,$Chart2,$Chart3 -Show -AutoNameRange


# Save and close the Excel package
Close-ExcelPackage -ExcelPackage $excelPackage -Show

<#
    Next steps:
    - Make minor adjustments like select the right ranges and being able to manipulate the ranges
    - Other graph modifications
#>

