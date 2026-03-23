param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot 'plan.xlsx'),
    # Change this default to target a specific source worksheet.
    # Example: [string]$SourceSheetName = 'MainData'
    [string]$SourceSheetName = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function From-CodePoints {
    param([int[]]$Points)

    return (-join ($Points | ForEach-Object { [char]$_ }))
}

$dashboardName = From-CodePoints @(36914,25431,12480,12483,12471,12517,12508,12540,12489)
$sheetLabel = From-CodePoints @(12527,12540,12463,12471,12540,12488)

$titleText = $dashboardName
$labelVisibleCount = From-CodePoints @(34920,31034,20214,25968)
$labelInProgressCount = From-CodePoints @(36914,34892,20013,20214,25968)
$labelCompleted = From-CodePoints @(23436,20102,20214,25968)
$labelDelayCount = From-CodePoints @(36933,24310,20214,25968)
$labelAverageGap = From-CodePoints @(24179,22343,20054,38626,40,26085,41)
$labelFilteredOverallProgress = From-CodePoints @(12501,12451,12523,12479,24460,20840,20307,36914,25431,40,37,41)
$labelCode = From-CodePoints @(12467,12540,12489)
$labelTestCase = From-CodePoints @(12486,12473,12488,12465,12540,12473)
$labelDesign = From-CodePoints @(35443,32048,35373,35336)
$labelUnitTest = From-CodePoints @(21336,20307,12486,12473,12488)
$labelListSection = From-CodePoints @(31227,34892,23550,35937,27231,33021,19968,35239)
$labelFilterNote = From-CodePoints @(12501,12451,12523,12479,32080,26524,12395,36899,21205,12375,12390,12289,20214,25968,12539,36914,25431,29575,12539,24179,22343,20054,38626,12364,26356,26032,12373,12428,12414,12377,12290)

$labelProgramId = From-CodePoints @(12503,12525,12464,12521,12512,73,68)
$labelFunctionName = From-CodePoints @(27231,33021,21517)
$labelSubsystem = From-CodePoints @(12469,12502,12471,12473,12486,12512)
$labelBlock = From-CodePoints @(12502,12525,12483,12463)
$labelGroup = From-CodePoints @(25285,24403,71,82,79,85,80)
$labelCurrentOwner = From-CodePoints @(29694,25285,24403,32773)
$labelCurrentStatus = From-CodePoints @(29694,22312,29366,27841)
$labelOverallProgress = From-CodePoints @(20840,20307,36914,25431,40,37,41)
$labelCodePercent = From-CodePoints @(12467,12540,12489,40,37,41)
$labelTestCasePercent = From-CodePoints @(12486,12473,12488,12465,12540,12473,40,37,41)
$labelDesignPercent = From-CodePoints @(35443,32048,35373,35336,40,37,41)
$labelUnitTestPercent = From-CodePoints @(21336,20307,12486,12473,12488,40,37,41)
$labelPlannedFinish = From-CodePoints @(23436,20102,20104,23450,26085)
$labelActualDate = From-CodePoints @(23455,32318,26085)
$labelGapDays = From-CodePoints @(20054,38626,40,26085,41)
$labelRemarks = From-CodePoints @(20633,32771)

$relationshipNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$sheetType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
$worksheetContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'

$styleTitle = 5
$styleNote = 6
$styleCardLabel = 7
$styleCardValueCount = 8
$styleCardValuePercent = 9
$styleCardValueGap = 10
$styleCardValueDelay = 11
$styleSection = 12
$styleTableHeader = 13
$styleTableText = 14
$styleTableCenter = 15
$styleTablePercent = 16
$styleTableDate = 17
$styleTableInteger = 18
$styleTableWrap = 19
$styleSummaryRibbon = 20

$columnMap = @{
    ProgramId = 'H'
    FunctionName = 'F'
    Subsystem = 'E'
    Block = 'G'
    Group = 'I'
    CurrentOwner = 'P'
    CurrentStatus = 'Q'
    CodeProgress = 'V'
    TestCaseProgress = 'AG'
    DesignProgress = 'AR'
    UnitTestProgress = 'BC'
    CodePlanEnd = 'S'
    TestCasePlanEnd = 'AD'
    DesignPlanEnd = 'AO'
    UnitTestPlanEnd = 'AZ'
    CodeActualEnd = 'U'
    TestCaseActualEnd = 'AF'
    DesignActualEnd = 'AQ'
    UnitTestActualEnd = 'BB'
    UnitTestCompleted = 'BI'
    Remarks = 'BJ'
}

function Get-EntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) {
        throw "Entry not found: $EntryName"
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

function Set-EntryText {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Text
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($existing) {
        $existing.Delete()
    }

    $entry = $Zip.CreateEntry($EntryName)
    $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try {
        $writer.Write($Text)
    } finally {
        $writer.Dispose()
    }
}

function ConvertTo-Utf8XmlText {
    param([xml]$Document)

    $stream = [System.IO.MemoryStream]::new()
    try {
        $settings = [System.Xml.XmlWriterSettings]::new()
        $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
        $settings.Indent = $false
        $settings.OmitXmlDeclaration = $false

        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try {
            $Document.Save($writer)
        } finally {
            $writer.Dispose()
        }

        return [System.Text.Encoding]::UTF8.GetString($stream.ToArray())
    } finally {
        $stream.Dispose()
    }
}

function Escape-Xml {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return [System.Security.SecurityElement]::Escape($Value)
}

function New-InlineStringCell {
    param(
        [string]$Reference,
        [string]$Text,
        [int]$Style = 0
    )

    return '<c r="' + $Reference + '" s="' + $Style + '" t="inlineStr"><is><t>' + (Escape-Xml $Text) + '</t></is></c>'
}

function New-FormulaCell {
    param(
        [string]$Reference,
        [string]$Formula,
        [int]$Style = 0,
        [switch]$StringResult
    )

    $typeAttribute = ''
    if ($StringResult) {
        $typeAttribute = ' t="str"'
    }

    return '<c r="' + $Reference + '" s="' + $Style + '"' + $typeAttribute + '><f>' + (Escape-Xml $Formula) + '</f></c>'
}

function New-RowXml {
    param(
        [int]$RowNumber,
        [string[]]$Cells,
        [double]$Height = 0
    )

    $heightAttributes = ''
    if ($Height -gt 0) {
        $heightAttributes = ' ht="' + $Height + '" customHeight="1"'
    }

    return '<row r="' + $RowNumber + '"' + $heightAttributes + '>' + ($Cells -join '') + '</row>'
}

function SheetCellRef {
    param(
        [string]$Key,
        [int]$RowNumber
    )

    if (-not $columnMap.ContainsKey($Key)) {
        throw "Unknown column map key: $Key"
    }

    return $script:sourceSheetFormulaName + '!$' + $columnMap[$Key] + '$' + $RowNumber
}

function ConvertTo-WorksheetFormulaName {
    param([string]$WorksheetName)

    return "'" + $WorksheetName.Replace("'", "''") + "'"
}

function New-StylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <numFmts count="2">
    <numFmt numFmtId="179" formatCode="0_);[Red]\(0\)"/>
    <numFmt numFmtId="180" formatCode="0.0"/>
  </numFmts>
  <fonts count="9" x14ac:knownFonts="1">
    <font><sz val="11"/><color theme="1"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>
    <font><sz val="6"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>
    <font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><sz val="10"/><color rgb="FF516072"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FF0F172A"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FFBC1B06"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FFB45309"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="8" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7FAFC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF334155"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A5F"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD8E1EB"/></left><right style="thin"><color rgb="FFD8E1EB"/></right><top style="thin"><color rgb="FFD8E1EB"/></top><bottom style="thin"><color rgb="FFD8E1EB"/></bottom><diagonal/></border>
    <border><left style="thin"><color rgb="FFD8E1EB"/></left><right style="thin"><color rgb="FFD8E1EB"/></right><top style="thin"><color rgb="FFD8E1EB"/></top><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf>
  </cellStyleXfs>
  <cellXfs count="21">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf>
    <xf numFmtId="179" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="180" fontId="5" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="180" fontId="8" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="180" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="14" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
  <extLst>
    <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext>
    <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext>
  </extLst>
</styleSheet>
'@
}

function New-DashboardWorksheetXml {
    $builder = [System.Text.StringBuilder]::new()
    $dataStartRow = 10
    $dataEndRow = 1009

    $visibleCountFormula = 'SUBTOTAL(9,T' + $dataStartRow + ':T' + $dataEndRow + ')'
    $inProgressCountFormula = 'SUBTOTAL(9,Q' + $dataStartRow + ':Q' + $dataEndRow + ')'
    $completedCountFormula = 'SUBTOTAL(9,R' + $dataStartRow + ':R' + $dataEndRow + ')'
    $delayCountFormula = 'SUBTOTAL(9,S' + $dataStartRow + ':S' + $dataEndRow + ')'
    $averageGapFormula = 'IFERROR(ROUND(SUBTOTAL(1,J' + $dataStartRow + ':J' + $dataEndRow + '),1),0)'
    $overallAverageFormula = 'IFERROR(SUBTOTAL(1,K' + $dataStartRow + ':K' + $dataEndRow + '),0)'
    $codeAverageFormula = 'IFERROR(SUBTOTAL(1,L' + $dataStartRow + ':L' + $dataEndRow + '),0)'
    $testCaseAverageFormula = 'IFERROR(SUBTOTAL(1,M' + $dataStartRow + ':M' + $dataEndRow + '),0)'
    $designAverageFormula = 'IFERROR(SUBTOTAL(1,N' + $dataStartRow + ':N' + $dataEndRow + '),0)'
    $unitAverageFormula = 'IFERROR(SUBTOTAL(1,O' + $dataStartRow + ':O' + $dataEndRow + '),0)'
    $summaryBandFormula = '"' + $labelFilteredOverallProgress + ' " & TEXT(' + $overallAverageFormula + ',"0.0") & "% / ' + $labelAverageGap + ' " & TEXT(' + $averageGapFormula + ',"0.0")'

    $mergeRefs = @(
        'A1:P1',
        'A2:P2',
        'A3:D3', 'E3:H3', 'I3:L3', 'M3:P3',
        'A4:D4', 'E4:H4', 'I4:L4', 'M4:P4',
        'A5:D5', 'E5:H5', 'I5:L5', 'M5:P5',
        'A6:D6', 'E6:H6', 'I6:L6', 'M6:P6',
        'A7:P7',
        'A8:P8'
    )

    $null = $builder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $null = $builder.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    $null = $builder.Append('<dimension ref="A1:T1009"/>')
    $null = $builder.Append('<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="90"><pane ySplit="9" topLeftCell="A10" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A10" sqref="A10"/></sheetView></sheetViews>')
    $null = $builder.Append('<sheetFormatPr defaultRowHeight="19"/>')
    $null = $builder.Append('<cols>')
    $null = $builder.Append('<col min="1" max="1" width="13" customWidth="1"/>')
    $null = $builder.Append('<col min="2" max="2" width="17" customWidth="1"/>')
    $null = $builder.Append('<col min="3" max="3" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="4" max="4" width="10" customWidth="1"/>')
    $null = $builder.Append('<col min="5" max="6" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="7" max="7" width="13" customWidth="1"/>')
    $null = $builder.Append('<col min="8" max="12" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="13" max="14" width="11.5" customWidth="1"/>')
    $null = $builder.Append('<col min="15" max="15" width="10" customWidth="1"/>')
    $null = $builder.Append('<col min="16" max="16" width="17" customWidth="1"/>')
    $null = $builder.Append('<col min="17" max="20" width="2" hidden="1" customWidth="1"/>')
    $null = $builder.Append('</cols>')
    $null = $builder.Append('<sheetData>')

    $null = $builder.Append((New-RowXml 1 @(
        (New-InlineStringCell 'A1' $titleText $styleTitle)
    ) 34))

    $null = $builder.Append((New-RowXml 2 @(
        (New-InlineStringCell 'A2' $labelFilterNote $styleNote)
    ) 18))

    $null = $builder.Append((New-RowXml 3 @(
        (New-InlineStringCell 'A3' $labelVisibleCount $styleCardLabel),
        (New-InlineStringCell 'E3' $labelInProgressCount $styleCardLabel),
        (New-InlineStringCell 'I3' $labelCompleted $styleCardLabel),
        (New-InlineStringCell 'M3' $labelDelayCount $styleCardLabel)
    ) 18))

    $null = $builder.Append((New-RowXml 4 @(
        (New-FormulaCell 'A4' $visibleCountFormula $styleCardValueCount),
        (New-FormulaCell 'E4' $inProgressCountFormula $styleCardValueCount),
        (New-FormulaCell 'I4' $completedCountFormula $styleCardValueCount),
        (New-FormulaCell 'M4' $delayCountFormula $styleCardValueDelay)
    ) 30))

    $null = $builder.Append((New-RowXml 5 @(
        (New-InlineStringCell 'A5' $labelCode $styleCardLabel),
        (New-InlineStringCell 'E5' $labelTestCase $styleCardLabel),
        (New-InlineStringCell 'I5' $labelDesign $styleCardLabel),
        (New-InlineStringCell 'M5' $labelUnitTest $styleCardLabel)
    ) 18))

    $null = $builder.Append((New-RowXml 6 @(
        (New-FormulaCell 'A6' $codeAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'E6' $testCaseAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'I6' $designAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'M6' $unitAverageFormula $styleCardValuePercent)
    ) 30))

    $null = $builder.Append((New-RowXml 7 @(
        (New-FormulaCell 'A7' $summaryBandFormula $styleSummaryRibbon -StringResult)
    ) 20))

    $null = $builder.Append((New-RowXml 8 @(
        (New-InlineStringCell 'A8' $labelListSection $styleSection)
    ) 22))

    $null = $builder.Append((New-RowXml 9 @(
        (New-InlineStringCell 'A9' $labelSubsystem $styleTableHeader),
        (New-InlineStringCell 'B9' $labelFunctionName $styleTableHeader),
        (New-InlineStringCell 'C9' $labelBlock $styleTableHeader),
        (New-InlineStringCell 'D9' $labelProgramId $styleTableHeader),
        (New-InlineStringCell 'E9' $labelGroup $styleTableHeader),
        (New-InlineStringCell 'F9' $labelCurrentOwner $styleTableHeader),
        (New-InlineStringCell 'G9' $labelCurrentStatus $styleTableHeader),
        (New-InlineStringCell 'H9' $labelPlannedFinish $styleTableHeader),
        (New-InlineStringCell 'I9' $labelActualDate $styleTableHeader),
        (New-InlineStringCell 'J9' $labelGapDays $styleTableHeader),
        (New-InlineStringCell 'K9' $labelOverallProgress $styleTableHeader),
        (New-InlineStringCell 'L9' $labelCodePercent $styleTableHeader),
        (New-InlineStringCell 'M9' $labelTestCasePercent $styleTableHeader),
        (New-InlineStringCell 'N9' $labelDesignPercent $styleTableHeader),
        (New-InlineStringCell 'O9' $labelUnitTestPercent $styleTableHeader),
        (New-InlineStringCell 'P9' $labelRemarks $styleTableHeader)
    ) 24))

    for ($dashboardRow = $dataStartRow; $dashboardRow -le $dataEndRow; $dashboardRow++) {
        $sourceRow = $dashboardRow - 5

        $programIdRef = SheetCellRef 'ProgramId' $sourceRow
        $functionNameRef = SheetCellRef 'FunctionName' $sourceRow
        $subsystemRef = SheetCellRef 'Subsystem' $sourceRow
        $blockRef = SheetCellRef 'Block' $sourceRow
        $groupRef = SheetCellRef 'Group' $sourceRow
        $currentOwnerRef = SheetCellRef 'CurrentOwner' $sourceRow
        $currentStatusRef = SheetCellRef 'CurrentStatus' $sourceRow
        $codeProgressRef = SheetCellRef 'CodeProgress' $sourceRow
        $testCaseProgressRef = SheetCellRef 'TestCaseProgress' $sourceRow
        $designProgressRef = SheetCellRef 'DesignProgress' $sourceRow
        $unitTestProgressRef = SheetCellRef 'UnitTestProgress' $sourceRow
        $remarksRef = SheetCellRef 'Remarks' $sourceRow

        $progressFormula = 'ROUND((N(' + $codeProgressRef + ')+N(' + $testCaseProgressRef + ')+N(' + $designProgressRef + ')+N(' + $unitTestProgressRef + '))/4,0)'
        $planMaxFormula = 'MAX(N(' + (SheetCellRef 'CodePlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'TestCasePlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'DesignPlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestPlanEnd' $sourceRow) + '))'
        $actualMaxFormula = 'MAX(N(' + (SheetCellRef 'CodeActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'TestCaseActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'DesignActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestCompleted' $sourceRow) + '))'

        $null = $builder.Append((New-RowXml $dashboardRow @(
            (New-FormulaCell ('A' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $subsystemRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('B' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $functionNameRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('C' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $blockRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('D' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $programIdRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('E' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $groupRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('F' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $currentOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('G' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $currentStatusRef + ')') $styleTableCenter -StringResult),
            (New-FormulaCell ('H' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(' + $planMaxFormula + '=0,"",' + $planMaxFormula + '))') $styleTableDate),
            (New-FormulaCell ('I' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(K' + $dashboardRow + '>=100,IF(' + $actualMaxFormula + '=0,"",' + $actualMaxFormula + '),""))') $styleTableDate),
            (New-FormulaCell ('J' + $dashboardRow) ('IF(H' + $dashboardRow + '="","",IF(I' + $dashboardRow + '<>"",ROUND(I' + $dashboardRow + '-H' + $dashboardRow + ',0),IF(K' + $dashboardRow + '>0,ROUND(TODAY()-H' + $dashboardRow + ',0),"")))') $styleTableInteger),
            (New-FormulaCell ('K' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $progressFormula + ')') $styleTablePercent),
            (New-FormulaCell ('L' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $codeProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('M' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $testCaseProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('N' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $designProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('O' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $unitTestProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('P' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $remarksRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('Q' + $dashboardRow) ('IF(D' + $dashboardRow + '="","",IF(AND(K' + $dashboardRow + '>0,K' + $dashboardRow + '<100),1,0))') $styleTableInteger),
            (New-FormulaCell ('R' + $dashboardRow) ('IF(D' + $dashboardRow + '="","",IF(K' + $dashboardRow + '>=100,1,0))') $styleTableInteger),
            (New-FormulaCell ('S' + $dashboardRow) ('IF(D' + $dashboardRow + '="","",IF(AND(J' + $dashboardRow + '<>"",J' + $dashboardRow + '>0),1,0))') $styleTableInteger),
            (New-FormulaCell ('T' + $dashboardRow) ('IF(D' + $dashboardRow + '="","",1)') $styleTableInteger)
        )))
    }

    $null = $builder.Append('</sheetData>')
    $null = $builder.Append('<autoFilter ref="A9:P1009"/>')
    $null = $builder.Append('<mergeCells count="' + $mergeRefs.Count + '">')
    foreach ($mergeRef in $mergeRefs) {
        $null = $builder.Append('<mergeCell ref="' + $mergeRef + '"/>')
    }
    $null = $builder.Append('</mergeCells>')
    $null = $builder.Append('<conditionalFormatting sqref="K10:O1009"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="100"/><color rgb="FF2563EB"/></dataBar></cfRule></conditionalFormatting>')
    $null = $builder.Append('<conditionalFormatting sqref="J10:J1009"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="num" val="-10"/><cfvo type="num" val="0"/><cfvo type="num" val="10"/><color rgb="FF16A34A"/><color rgb="FFFDE68A"/><color rgb="FFDC2626"/></colorScale></cfRule></conditionalFormatting>')
    $null = $builder.Append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    $null = $builder.Append('</worksheet>')

    return $builder.ToString()
}

if (-not (Test-Path $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

$workbookDirectory = Split-Path -Parent $WorkbookPath
$workbookBaseName = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
$backupName = $workbookBaseName + '.backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx'
$backupPath = Join-Path $workbookDirectory $backupName
Copy-Item -Path $WorkbookPath -Destination $backupPath -Force

$zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)

try {
    [xml]$workbookXml = Get-EntryText -Zip $zip -EntryName 'xl/workbook.xml'
    [xml]$relationshipsXml = Get-EntryText -Zip $zip -EntryName 'xl/_rels/workbook.xml.rels'
    [xml]$contentTypesXml = Get-EntryText -Zip $zip -EntryName '[Content_Types].xml'

    $allSheetNames = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [string]$_.name })
    # If you want to lock the dashboard to a specific worksheet in source,
    # set $SourceSheetName above or pass -SourceSheetName when running the script.
    if ([string]::IsNullOrWhiteSpace($SourceSheetName)) {
        $candidateSheetNames = @($allSheetNames | Where-Object { $_ -ne $dashboardName })
        if ($candidateSheetNames.Count -eq 0) {
            throw 'No source worksheet found.'
        }

        $SourceSheetName = $candidateSheetNames[0]
    }

    if ($allSheetNames -notcontains $SourceSheetName) {
        throw "Source worksheet not found: $SourceSheetName"
    }

    if ($SourceSheetName -eq $dashboardName) {
        throw 'Dashboard worksheet cannot be used as the source worksheet.'
    }

    $script:sourceSheetFormulaName = ConvertTo-WorksheetFormulaName $SourceSheetName

    $existingSheet = @($workbookXml.workbook.sheets.sheet | Where-Object { $_.name -eq $dashboardName }) | Select-Object -First 1
    $sheetRelationshipId = $null
    $sheetTarget = $null

    if ($existingSheet) {
        $sheetRelationshipId = $existingSheet.GetAttribute('id', $relationshipNs)
        $existingRelationship = @($relationshipsXml.Relationships.Relationship | Where-Object { $_.Id -eq $sheetRelationshipId }) | Select-Object -First 1
        if (-not $existingRelationship) {
            throw "Relationship not found for existing dashboard sheet: $sheetRelationshipId"
        }

        $sheetTarget = [string]$existingRelationship.Target
    } else {
        $sheetIds = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [int]$_.sheetId })
        $nextSheetId = (($sheetIds | Measure-Object -Maximum).Maximum) + 1

        $relationshipNumbers = @(
            $relationshipsXml.Relationships.Relationship |
            ForEach-Object {
                if ($_.Id -match '^rId(\d+)$') {
                    [int]$matches[1]
                }
            }
        )

        $nextRelationshipId = 'rId' + ((($relationshipNumbers | Measure-Object -Maximum).Maximum) + 1)

        $worksheetNumbers = @(
            $relationshipsXml.Relationships.Relationship |
            Where-Object { $_.Type -eq $sheetType } |
            ForEach-Object {
                if ($_.Target -match 'worksheets/sheet(\d+)\.xml$') {
                    [int]$matches[1]
                }
            }
        )

        $nextWorksheetNumber = (($worksheetNumbers | Measure-Object -Maximum).Maximum) + 1
        $sheetTarget = 'worksheets/sheet' + $nextWorksheetNumber + '.xml'
        $sheetRelationshipId = $nextRelationshipId

        $newSheet = $workbookXml.CreateElement('sheet', $workbookXml.workbook.NamespaceURI)
        $null = $newSheet.SetAttribute('name', $dashboardName)
        $null = $newSheet.SetAttribute('sheetId', [string]$nextSheetId)
        $null = $newSheet.SetAttribute('id', $relationshipNs, $sheetRelationshipId)
        $null = $workbookXml.workbook.sheets.AppendChild($newSheet)

        $newRelationship = $relationshipsXml.CreateElement('Relationship', $relationshipsXml.Relationships.NamespaceURI)
        $null = $newRelationship.SetAttribute('Id', $sheetRelationshipId)
        $null = $newRelationship.SetAttribute('Type', $sheetType)
        $null = $newRelationship.SetAttribute('Target', $sheetTarget)
        $null = $relationshipsXml.Relationships.AppendChild($newRelationship)

        $partName = '/' + 'xl/' + $sheetTarget
        $existingOverride = @($contentTypesXml.Types.Override | Where-Object { $_.PartName -eq $partName }) | Select-Object -First 1
        if (-not $existingOverride) {
            $newOverride = $contentTypesXml.CreateElement('Override', $contentTypesXml.Types.NamespaceURI)
            $null = $newOverride.SetAttribute('PartName', $partName)
            $null = $newOverride.SetAttribute('ContentType', $worksheetContentType)
            $null = $contentTypesXml.Types.AppendChild($newOverride)
        }
    }

    if (-not $workbookXml.workbook.calcPr) {
        $calcNode = $workbookXml.CreateElement('calcPr', $workbookXml.workbook.NamespaceURI)
        $null = $workbookXml.workbook.AppendChild($calcNode)
    }

    $workbookXml.workbook.calcPr.SetAttribute('calcMode', 'auto')
    $workbookXml.workbook.calcPr.SetAttribute('fullCalcOnLoad', '1')
    $workbookXml.workbook.calcPr.SetAttribute('forceFullCalc', '1')

    $sheetNames = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [string]$_.name })
    $titleNodes = $sheetNames | ForEach-Object { '<vt:lpstr>' + (Escape-Xml $_) + '</vt:lpstr>' }
    $appXmlText = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        '<Application>Microsoft Excel</Application>' +
        '<DocSecurity>0</DocSecurity>' +
        '<ScaleCrop>false</ScaleCrop>' +
        '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>' + (Escape-Xml $sheetLabel) + '</vt:lpstr></vt:variant><vt:variant><vt:i4>' + $sheetNames.Count + '</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
        '<TitlesOfParts><vt:vector size="' + $sheetNames.Count + '" baseType="lpstr">' + ($titleNodes -join '') + '</vt:vector></TitlesOfParts>' +
        '<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>' +
        '</Properties>'

    $dashboardEntryName = 'xl/' + $sheetTarget

    Set-EntryText -Zip $zip -EntryName 'xl/workbook.xml' -Text (ConvertTo-Utf8XmlText $workbookXml)
    Set-EntryText -Zip $zip -EntryName 'xl/_rels/workbook.xml.rels' -Text (ConvertTo-Utf8XmlText $relationshipsXml)
    Set-EntryText -Zip $zip -EntryName '[Content_Types].xml' -Text (ConvertTo-Utf8XmlText $contentTypesXml)
    Set-EntryText -Zip $zip -EntryName 'docProps/app.xml' -Text $appXmlText
    Set-EntryText -Zip $zip -EntryName 'xl/styles.xml' -Text (New-StylesXml)
    Set-EntryText -Zip $zip -EntryName $dashboardEntryName -Text (New-DashboardWorksheetXml)
} finally {
    $zip.Dispose()
}

Write-Output ('dashboard_sheet=' + $dashboardName)
Write-Output ('source_sheet=' + $SourceSheetName)
Write-Output ('backup_path=' + $backupPath)
Write-Output ('workbook_path=' + $WorkbookPath)
