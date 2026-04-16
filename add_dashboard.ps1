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
$labelClassificationRelated = From-CodePoints @(20998,39006,38306,36899)
$labelSubsystem = From-CodePoints @(12469,12502,12471,12473,12486,12512)
$labelBlock = From-CodePoints @(12502,12525,12483,12463)
$labelGroup = From-CodePoints @(25285,24403,71,82,79,85,80)
$labelOwner = From-CodePoints @(25285,24403,32773)
$labelManager = From-CodePoints @(31649,29702,32773)
$labelCodeReviewer = From-CodePoints @(12467,12540,12489,12524,12499,12517,12450,12540)
$labelCaseReviewer = From-CodePoints @(12465,12540,12473,12524,12499,12517,12450,12540)
$labelDesignReviewer = From-CodePoints @(35373,35336,12524,12499,12517,12450,12540)
$labelTestReviewer = From-CodePoints @(12486,12473,12488,12524,12499,12517,12450,12540)
$labelCurrentOwner = From-CodePoints @(29694,25285,24403,32773)
$labelCurrentStatus = From-CodePoints @(29694,22312,29366,27841)
$labelOverallProgress = From-CodePoints @(20840,20307,36914,25431,40,37,41)
$labelCodePercent = From-CodePoints @(12467,12540,12489,40,37,41)
$labelTestCasePercent = From-CodePoints @(12486,12473,12488,12465,12540,12473,40,37,41)
$labelDesignPercent = From-CodePoints @(35443,32048,35373,35336,40,37,41)
$labelUnitTestPercent = From-CodePoints @(21336,20307,12486,12473,12488,40,37,41)
$labelCodeReview = From-CodePoints @(12467,12540,12489,12524,12499,12517,12540)
$labelCaseReview = From-CodePoints @(12465,12540,12473,12524,12499,12517,12540)
$labelDesignDocumentReview = From-CodePoints @(35373,35336,26360,12524,12499,12517,12540)
$labelTestReview = From-CodePoints @(12486,12473,12488,12524,12499,12517,12540)
$labelStepCount = From-CodePoints @(12473,12486,12483,12503,25968)
$labelStepCountTotal = From-CodePoints @(12473,12486,12483,12503,25968,21512,35336)
$labelPlannedFinish = From-CodePoints @(23436,20102,20104,23450,26085)
$labelActualDate = From-CodePoints @(23455,32318,26085)
$labelGapDays = From-CodePoints @(20054,38626,40,26085,41)
$labelRemarks = From-CodePoints @(20633,32771)

$relationshipNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$sheetType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
$worksheetContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'

# Source worksheet layout:
# - Row 3: headers
# - Row 4 and below: data
# Change these values if the source worksheet layout changes.
$sourceHeaderRow = 3
$sourceDataStartRow = 4
$minimumSourceLastRow = 2000
$dashboardDataStartRow = 10

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
    ClassificationRelated = 'D'
    ProgramId = 'H'
    FunctionName = 'F'
    Subsystem = 'E'
    Block = 'G'
    Group = 'I'
    Owner = 'J'
    Manager = 'K'
    CodeOwner = 'L'
    TestCaseOwner = 'M'
    DesignOwner = 'N'
    UnitTestOwner = 'O'
    CurrentOwner = 'P'
    CurrentStatus = 'Q'
    CodeProgress = 'V'
    TestCaseProgress = 'AG'
    DesignProgress = 'AR'
    UnitTestProgress = 'BC'
    CodeReview = 'Z'
    CaseReview = 'AK'
    DesignDocumentReview = 'AV'
    TestReview = 'BG'
    StepCount = 'BN'
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

function Get-LastSourceDataRow {
    param(
        [string]$WorksheetXmlText,
        [string]$KeyColumn,
        [int]$DataStartRow
    )

    $pattern = '<c r="' + [regex]::Escape($KeyColumn) + '(\d+)"[^>]*>(?s:.*?)</c>'
    $matches = [regex]::Matches($WorksheetXmlText, $pattern)
    $rowNumbers = @(
        $matches |
        ForEach-Object {
            $rowNumber = [int]$_.Groups[1].Value
            if ($rowNumber -ge $DataStartRow) {
                $rowNumber
            }
        }
    )

    if ($rowNumbers.Count -eq 0) {
        return ($DataStartRow - 1)
    }

    return (($rowNumbers | Measure-Object -Maximum).Maximum)
}

function New-StylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <numFmts count="2">
    <numFmt numFmtId="179" formatCode="0_);[Red]\(0\)"/>
    <numFmt numFmtId="180" formatCode="0.0"/>
  </numFmts>
  <fonts count="11" x14ac:knownFonts="1">
    <font><sz val="11"/><color theme="1"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>
    <font><sz val="6"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/><scheme val="minor"/></font>
    <font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><sz val="10"/><color rgb="FF516072"/><name val="Yu Gothic"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FF0F172A"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FFBC1B06"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FFB45309"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
    <font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Yu Gothic UI Semibold"/><family val="2"/><charset val="128"/></font>
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
    <xf numFmtId="0" fontId="10" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="6" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="180" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="14" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="9" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
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
    param(
        [int]$SourceDataStartRow,
        [int]$SourceLastRow
    )

    $builder = [System.Text.StringBuilder]::new()
    $dataStartRow = $dashboardDataStartRow
    $dataRowCount = [math]::Max(1, ($SourceLastRow - $SourceDataStartRow + 1))
    $dataEndRow = $dataStartRow + $dataRowCount - 1
    $lastWorksheetRow = [math]::Max($dataEndRow, 9)

    $visibleCountFormula = 'SUBTOTAL(9,AF' + $dataStartRow + ':AF' + $dataEndRow + ')'
    $inProgressCountFormula = 'SUBTOTAL(9,AC' + $dataStartRow + ':AC' + $dataEndRow + ')'
    $completedCountFormula = 'SUBTOTAL(9,AD' + $dataStartRow + ':AD' + $dataEndRow + ')'
    $delayCountFormula = 'SUBTOTAL(9,AE' + $dataStartRow + ':AE' + $dataEndRow + ')'
    $averageGapFormula = 'IFERROR(ROUND(SUBTOTAL(1,Q' + $dataStartRow + ':Q' + $dataEndRow + '),1),0)'
    $overallAverageFormula = 'IFERROR(SUBTOTAL(1,R' + $dataStartRow + ':R' + $dataEndRow + '),0)'
    $codeAverageFormula = 'IFERROR(SUBTOTAL(1,S' + $dataStartRow + ':S' + $dataEndRow + '),0)'
    $testCaseAverageFormula = 'IFERROR(SUBTOTAL(1,T' + $dataStartRow + ':T' + $dataEndRow + '),0)'
    $designAverageFormula = 'IFERROR(SUBTOTAL(1,U' + $dataStartRow + ':U' + $dataEndRow + '),0)'
    $unitAverageFormula = 'IFERROR(SUBTOTAL(1,V' + $dataStartRow + ':V' + $dataEndRow + '),0)'
    $stepTotalFormula = 'SUBTOTAL(9,AA' + $dataStartRow + ':AA' + $dataEndRow + ')'
    $summaryBandFormula = '"' + $labelFilteredOverallProgress + ' " & TEXT(' + $overallAverageFormula + ',"0.0") & "% / ' + $labelAverageGap + ' " & TEXT(' + $averageGapFormula + ',"0.0") & " / ' + $labelStepCountTotal + ' " & TEXT(' + $stepTotalFormula + ',"0")'

    $mergeRefs = @(
        'A1:AB1',
        'A2:AB2',
        'A3:G3', 'H3:N3', 'O3:U3', 'V3:AB3',
        'A4:G4', 'H4:N4', 'O4:U4', 'V4:AB4',
        'A5:G5', 'H5:N5', 'O5:U5', 'V5:AB5',
        'A6:G6', 'H6:N6', 'O6:U6', 'V6:AB6',
        'A7:AB7',
        'A8:AB8'
    )

    $null = $builder.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $null = $builder.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    $null = $builder.Append('<dimension ref="A1:AF' + $lastWorksheetRow + '"/>')
    $null = $builder.Append('<sheetViews><sheetView workbookViewId="0" showGridLines="0" zoomScale="90"><pane ySplit="9" topLeftCell="A10" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A10" sqref="A10"/></sheetView></sheetViews>')
    $null = $builder.Append('<sheetFormatPr defaultRowHeight="19"/>')
    $null = $builder.Append('<cols>')
    $null = $builder.Append('<col min="1" max="1" width="14" customWidth="1"/>')
    $null = $builder.Append('<col min="2" max="2" width="13" customWidth="1"/>')
    $null = $builder.Append('<col min="3" max="3" width="17" customWidth="1"/>')
    $null = $builder.Append('<col min="4" max="4" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="5" max="5" width="10" customWidth="1"/>')
    $null = $builder.Append('<col min="6" max="6" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="7" max="13" width="10.5" customWidth="1"/>')
    $null = $builder.Append('<col min="14" max="14" width="13" customWidth="1"/>')
    $null = $builder.Append('<col min="15" max="16" width="11.5" customWidth="1"/>')
    $null = $builder.Append('<col min="17" max="17" width="10" customWidth="1"/>')
    $null = $builder.Append('<col min="18" max="26" width="12" customWidth="1"/>')
    $null = $builder.Append('<col min="27" max="27" width="10" customWidth="1"/>')
    $null = $builder.Append('<col min="28" max="28" width="17" customWidth="1"/>')
    $null = $builder.Append('<col min="29" max="32" width="2" hidden="1" customWidth="1"/>')
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
        (New-InlineStringCell 'H3' $labelInProgressCount $styleCardLabel),
        (New-InlineStringCell 'O3' $labelCompleted $styleCardLabel),
        (New-InlineStringCell 'V3' $labelDelayCount $styleCardLabel)
    ) 18))

    $null = $builder.Append((New-RowXml 4 @(
        (New-FormulaCell 'A4' $visibleCountFormula $styleCardValueCount),
        (New-FormulaCell 'H4' $inProgressCountFormula $styleCardValueCount),
        (New-FormulaCell 'O4' $completedCountFormula $styleCardValueCount),
        (New-FormulaCell 'V4' $delayCountFormula $styleCardValueDelay)
    ) 30))

    $null = $builder.Append((New-RowXml 5 @(
        (New-InlineStringCell 'A5' $labelCode $styleCardLabel),
        (New-InlineStringCell 'H5' $labelTestCase $styleCardLabel),
        (New-InlineStringCell 'O5' $labelDesign $styleCardLabel),
        (New-InlineStringCell 'V5' $labelUnitTest $styleCardLabel)
    ) 18))

    $null = $builder.Append((New-RowXml 6 @(
        (New-FormulaCell 'A6' $codeAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'H6' $testCaseAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'O6' $designAverageFormula $styleCardValuePercent),
        (New-FormulaCell 'V6' $unitAverageFormula $styleCardValuePercent)
    ) 30))

    $null = $builder.Append((New-RowXml 7 @(
        (New-FormulaCell 'A7' $summaryBandFormula $styleSummaryRibbon -StringResult)
    ) 30))

    $null = $builder.Append((New-RowXml 8 @(
        (New-InlineStringCell 'A8' $labelListSection $styleSection)
    ) 26))

    $null = $builder.Append((New-RowXml 9 @(
        (New-InlineStringCell 'A9' $labelClassificationRelated $styleTableHeader),
        (New-InlineStringCell 'B9' $labelSubsystem $styleTableHeader),
        (New-InlineStringCell 'C9' $labelFunctionName $styleTableHeader),
        (New-InlineStringCell 'D9' $labelBlock $styleTableHeader),
        (New-InlineStringCell 'E9' $labelProgramId $styleTableHeader),
        (New-InlineStringCell 'F9' $labelGroup $styleTableHeader),
        (New-InlineStringCell 'G9' $labelOwner $styleTableHeader),
        (New-InlineStringCell 'H9' $labelManager $styleTableHeader),
        (New-InlineStringCell 'I9' $labelCodeReviewer $styleTableHeader),
        (New-InlineStringCell 'J9' $labelCaseReviewer $styleTableHeader),
        (New-InlineStringCell 'K9' $labelDesignReviewer $styleTableHeader),
        (New-InlineStringCell 'L9' $labelTestReviewer $styleTableHeader),
        (New-InlineStringCell 'M9' $labelCurrentOwner $styleTableHeader),
        (New-InlineStringCell 'N9' $labelCurrentStatus $styleTableHeader),
        (New-InlineStringCell 'O9' $labelPlannedFinish $styleTableHeader),
        (New-InlineStringCell 'P9' $labelActualDate $styleTableHeader),
        (New-InlineStringCell 'Q9' $labelGapDays $styleTableHeader),
        (New-InlineStringCell 'R9' $labelOverallProgress $styleTableHeader),
        (New-InlineStringCell 'S9' $labelCodePercent $styleTableHeader),
        (New-InlineStringCell 'T9' $labelTestCasePercent $styleTableHeader),
        (New-InlineStringCell 'U9' $labelDesignPercent $styleTableHeader),
        (New-InlineStringCell 'V9' $labelUnitTestPercent $styleTableHeader),
        (New-InlineStringCell 'W9' $labelCodeReview $styleTableHeader),
        (New-InlineStringCell 'X9' $labelCaseReview $styleTableHeader),
        (New-InlineStringCell 'Y9' $labelDesignDocumentReview $styleTableHeader),
        (New-InlineStringCell 'Z9' $labelTestReview $styleTableHeader),
        (New-InlineStringCell 'AA9' $labelStepCount $styleTableHeader),
        (New-InlineStringCell 'AB9' $labelRemarks $styleTableHeader)
    ) 24))

    for ($dashboardRow = $dataStartRow; $dashboardRow -le $dataEndRow; $dashboardRow++) {
        $sourceRow = $SourceDataStartRow + ($dashboardRow - $dataStartRow)

        $classificationRelatedRef = SheetCellRef 'ClassificationRelated' $sourceRow
        $programIdRef = SheetCellRef 'ProgramId' $sourceRow
        $functionNameRef = SheetCellRef 'FunctionName' $sourceRow
        $subsystemRef = SheetCellRef 'Subsystem' $sourceRow
        $blockRef = SheetCellRef 'Block' $sourceRow
        $groupRef = SheetCellRef 'Group' $sourceRow
        $ownerRef = SheetCellRef 'Owner' $sourceRow
        $managerRef = SheetCellRef 'Manager' $sourceRow
        $codeOwnerRef = SheetCellRef 'CodeOwner' $sourceRow
        $testCaseOwnerRef = SheetCellRef 'TestCaseOwner' $sourceRow
        $designOwnerRef = SheetCellRef 'DesignOwner' $sourceRow
        $unitTestOwnerRef = SheetCellRef 'UnitTestOwner' $sourceRow
        $currentOwnerRef = SheetCellRef 'CurrentOwner' $sourceRow
        $currentStatusRef = SheetCellRef 'CurrentStatus' $sourceRow
        $codeProgressRef = SheetCellRef 'CodeProgress' $sourceRow
        $testCaseProgressRef = SheetCellRef 'TestCaseProgress' $sourceRow
        $designProgressRef = SheetCellRef 'DesignProgress' $sourceRow
        $unitTestProgressRef = SheetCellRef 'UnitTestProgress' $sourceRow
        $codeReviewRef = SheetCellRef 'CodeReview' $sourceRow
        $caseReviewRef = SheetCellRef 'CaseReview' $sourceRow
        $designDocumentReviewRef = SheetCellRef 'DesignDocumentReview' $sourceRow
        $testReviewRef = SheetCellRef 'TestReview' $sourceRow
        $stepCountRef = SheetCellRef 'StepCount' $sourceRow
        $remarksRef = SheetCellRef 'Remarks' $sourceRow

        $progressFormula = 'ROUND((N(' + $codeProgressRef + ')+N(' + $testCaseProgressRef + ')+N(' + $designProgressRef + ')+N(' + $unitTestProgressRef + '))/4,0)'
        $planMaxFormula = 'MAX(N(' + (SheetCellRef 'CodePlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'TestCasePlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'DesignPlanEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestPlanEnd' $sourceRow) + '))'
        $actualMaxFormula = 'MAX(N(' + (SheetCellRef 'CodeActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'TestCaseActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'DesignActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestActualEnd' $sourceRow) + '),N(' + (SheetCellRef 'UnitTestCompleted' $sourceRow) + '))'

        $null = $builder.Append((New-RowXml $dashboardRow @(
            (New-FormulaCell ('A' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $classificationRelatedRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('B' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $subsystemRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('C' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $functionNameRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('D' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $blockRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('E' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $programIdRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('F' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $groupRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('G' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $ownerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('H' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $managerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('I' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $codeOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('J' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $testCaseOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('K' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $designOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('L' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $unitTestOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('M' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $currentOwnerRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('N' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $currentStatusRef + ')') $styleTableCenter -StringResult),
            (New-FormulaCell ('O' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(' + $planMaxFormula + '=0,"",' + $planMaxFormula + '))') $styleTableDate),
            (New-FormulaCell ('P' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(R' + $dashboardRow + '>=100,IF(' + $actualMaxFormula + '=0,"",' + $actualMaxFormula + '),""))') $styleTableDate),
            (New-FormulaCell ('Q' + $dashboardRow) ('IF(O' + $dashboardRow + '="","",IF(P' + $dashboardRow + '<>"",ROUND(P' + $dashboardRow + '-O' + $dashboardRow + ',0),IF(R' + $dashboardRow + '>0,ROUND(TODAY()-O' + $dashboardRow + ',0),"")))') $styleTableInteger),
            (New-FormulaCell ('R' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $progressFormula + ')') $styleTablePercent),
            (New-FormulaCell ('S' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $codeProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('T' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $testCaseProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('U' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $designProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('V' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $unitTestProgressRef + ')') $styleTablePercent),
            (New-FormulaCell ('W' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(LEN(' + $codeReviewRef + ')=0,"",' + $codeReviewRef + '))') $styleTableCenter -StringResult),
            (New-FormulaCell ('X' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(LEN(' + $caseReviewRef + ')=0,"",' + $caseReviewRef + '))') $styleTableCenter -StringResult),
            (New-FormulaCell ('Y' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(LEN(' + $designDocumentReviewRef + ')=0,"",' + $designDocumentReviewRef + '))') $styleTableCenter -StringResult),
            (New-FormulaCell ('Z' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(LEN(' + $testReviewRef + ')=0,"",' + $testReviewRef + '))') $styleTableCenter -StringResult),
            (New-FormulaCell ('AA' + $dashboardRow) ('IF(' + $programIdRef + '="","",IF(LEN(' + $stepCountRef + ')=0,"",' + $stepCountRef + '))') $styleTableInteger),
            (New-FormulaCell ('AB' + $dashboardRow) ('IF(' + $programIdRef + '="","",' + $remarksRef + ')') $styleTableText -StringResult),
            (New-FormulaCell ('AC' + $dashboardRow) ('IF(E' + $dashboardRow + '="","",IF(AND(R' + $dashboardRow + '>0,R' + $dashboardRow + '<100),1,0))') $styleTableInteger),
            (New-FormulaCell ('AD' + $dashboardRow) ('IF(E' + $dashboardRow + '="","",IF(R' + $dashboardRow + '>=100,1,0))') $styleTableInteger),
            (New-FormulaCell ('AE' + $dashboardRow) ('IF(E' + $dashboardRow + '="","",IF(AND(Q' + $dashboardRow + '<>"",Q' + $dashboardRow + '>0),1,0))') $styleTableInteger),
            (New-FormulaCell ('AF' + $dashboardRow) ('IF(E' + $dashboardRow + '="","",1)') $styleTableInteger)
        )))
    }

    $null = $builder.Append('</sheetData>')
    $null = $builder.Append('<autoFilter ref="A9:AB' + $lastWorksheetRow + '"/>')
    $null = $builder.Append('<mergeCells count="' + $mergeRefs.Count + '">')
    foreach ($mergeRef in $mergeRefs) {
        $null = $builder.Append('<mergeCell ref="' + $mergeRef + '"/>')
    }
    $null = $builder.Append('</mergeCells>')
    $null = $builder.Append('<conditionalFormatting sqref="R' + $dataStartRow + ':V' + $dataEndRow + '"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="100"/><color rgb="FF2563EB"/></dataBar></cfRule></conditionalFormatting>')
    $null = $builder.Append('<conditionalFormatting sqref="Q' + $dataStartRow + ':Q' + $dataEndRow + '"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="num" val="-10"/><cfvo type="num" val="0"/><cfvo type="num" val="10"/><color rgb="FF16A34A"/><color rgb="FFFDE68A"/><color rgb="FFDC2626"/></colorScale></cfRule></conditionalFormatting>')
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
    $sourceSheet = @($workbookXml.workbook.sheets.sheet | Where-Object { $_.name -eq $SourceSheetName }) | Select-Object -First 1
    if (-not $sourceSheet) {
        throw "Source worksheet node not found: $SourceSheetName"
    }

    $sourceSheetRelationshipId = $sourceSheet.GetAttribute('id', $relationshipNs)
    $sourceRelationship = @($relationshipsXml.Relationships.Relationship | Where-Object { $_.Id -eq $sourceSheetRelationshipId }) | Select-Object -First 1
    if (-not $sourceRelationship) {
        throw "Relationship not found for source worksheet: $sourceSheetRelationshipId"
    }

    $sourceSheetEntryName = 'xl/' + [string]$sourceRelationship.Target
    $sourceSheetXmlText = Get-EntryText -Zip $zip -EntryName $sourceSheetEntryName
    $sourceLastRow = [math]::Max(
        (Get-LastSourceDataRow -WorksheetXmlText $sourceSheetXmlText -KeyColumn $columnMap.ProgramId -DataStartRow $sourceDataStartRow),
        $minimumSourceLastRow
    )

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
    Set-EntryText -Zip $zip -EntryName $dashboardEntryName -Text (New-DashboardWorksheetXml -SourceDataStartRow $sourceDataStartRow -SourceLastRow $sourceLastRow)
} finally {
    $zip.Dispose()
}

Write-Output ('dashboard_sheet=' + $dashboardName)
Write-Output ('source_sheet=' + $SourceSheetName)
Write-Output ('source_header_row=' + $sourceHeaderRow)
Write-Output ('source_data_start_row=' + $sourceDataStartRow)
Write-Output ('minimum_source_last_row=' + $minimumSourceLastRow)
Write-Output ('source_last_row=' + $sourceLastRow)
Write-Output ('backup_path=' + $backupPath)
Write-Output ('workbook_path=' + $WorkbookPath)
