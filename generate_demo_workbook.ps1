param(
    [string]$SourcePath = (Join-Path $PSScriptRoot 'plan.xlsx'),
    [string]$DestinationPath = (Join-Path $PSScriptRoot 'plan.demo.xlsx'),
    [int]$AddRows = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Escape-Xml {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return [System.Security.SecurityElement]::Escape($Value)
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

function New-TextCell {
    param(
        [string]$Reference,
        [AllowNull()][string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    return '<c r="' + $Reference + '" t="inlineStr"><is><t>' + (Escape-Xml $Text) + '</t></is></c>'
}

function New-NumberCell {
    param(
        [string]$Reference,
        [AllowNull()][object]$Value,
        [int]$Style = 0
    )

    if ($null -eq $Value -or $Value -eq '') {
        return ''
    }

    $styleAttribute = ''
    if ($Style -gt 0) {
        $styleAttribute = ' s="' + $Style + '"'
    }

    return '<c r="' + $Reference + '"' + $styleAttribute + '><v>' + $Value + '</v></c>'
}

function New-DateCell {
    param(
        [string]$Reference,
        [AllowNull()][object]$Date
    )

    if ($null -eq $Date) {
        return ''
    }

    $dateValue = [datetime]$Date
    $serial = [int][math]::Floor($dateValue.ToOADate())
    return New-NumberCell -Reference $Reference -Value $serial -Style 3
}

function New-PercentCell {
    param(
        [string]$Reference,
        [AllowNull()][int]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return New-NumberCell -Reference $Reference -Value $Value -Style 4
}

function New-CountCell {
    param(
        [string]$Reference,
        [AllowNull()][int]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return New-NumberCell -Reference $Reference -Value $Value -Style 4
}

function Get-PhaseData {
    param(
        [datetime]$PlanStart,
        [int]$Duration,
        [int]$Progress,
        [int]$SlipDays,
        [int]$IssueSeed
    )

    $planEnd = $PlanStart.AddDays($Duration)
    $actualStart = $null
    $actualEnd = $null
    $reviewRequest = $null
    $reviewStart = $null
    $reviewEnd = $null
    $reviewStatus = ''
    $issues = $null
    $completed = $null

    if ($Progress -gt 0) {
        $actualStart = $PlanStart.AddDays($SlipDays)
    }

    if ($Progress -ge 100) {
        $actualEnd = $actualStart.AddDays([math]::Max(1, $Duration - 1))
        $reviewRequest = $actualEnd
        $reviewStart = $actualEnd
        $reviewEnd = $actualEnd.AddDays(1)
        $reviewStatus = 'completed'
        $issues = $IssueSeed
        $completed = $reviewEnd.AddDays(1)
    }

    return @{
        PlanStart = $PlanStart
        PlanEnd = $planEnd
        ActualStart = $actualStart
        ActualEnd = $actualEnd
        Progress = $Progress
        ReviewRequest = $reviewRequest
        ReviewStart = $reviewStart
        ReviewEnd = $reviewEnd
        ReviewStatus = $reviewStatus
        Issues = $issues
        Completed = $completed
    }
}

function Get-Profile {
    param([int]$Index)

    switch ($Index % 7) {
        0 { return @{ Code = 0; TestCase = 0; Design = 0; Unit = 0; Status = 'not started'; Note = 'backlog'; Delay = 0 } }
        1 { return @{ Code = 35; TestCase = 0; Design = 0; Unit = 0; Status = 'coding'; Note = 'coding in progress'; Delay = 0 } }
        2 { return @{ Code = 100; TestCase = 45; Design = 0; Unit = 0; Status = 'testcase'; Note = 'testcase in progress'; Delay = 0 } }
        3 { return @{ Code = 100; TestCase = 100; Design = 60; Unit = 0; Status = 'design'; Note = 'design in progress'; Delay = 0 } }
        4 { return @{ Code = 100; TestCase = 100; Design = 100; Unit = 40; Status = 'unit test'; Note = 'unit test in progress'; Delay = 0 } }
        5 { return @{ Code = 100; TestCase = 100; Design = 100; Unit = 100; Status = 'completed'; Note = 'ready to close'; Delay = 0 } }
        default { return @{ Code = 100; TestCase = 100; Design = 55; Unit = 0; Status = 'delayed'; Note = 'delay watch'; Delay = 3 } }
    }
}

function New-DemoRowXml {
    param(
        [int]$RowNumber,
        [int]$Sequence
    )

    $subsystems = @('Payroll', 'Attendance', 'HR', 'Expense', 'Sales', 'Inventory')
    $groups = @('Team-A', 'Team-B', 'Team-C')
    $owners = @('Yamada', 'Sato', 'Suzuki', 'Tanaka', 'Ito', 'Kobayashi')
    $managers = @('Aoki', 'Kato', 'Nakamura')
    $reviewers = @('Ota', 'Shimizu', 'Mori', 'Inoue', 'Hayashi')
    $blocks = @('salary', 'master', 'report', 'batch', 'api')
    $features = @('Calc', 'Import', 'Export', 'Summary', 'Validation', 'Approval')

    $profile = Get-Profile -Index $Sequence
    $baseDate = [datetime]'2026-04-01'
    $planStart = $baseDate.AddDays($Sequence * 2 + $profile.Delay)

    $code = Get-PhaseData -PlanStart $planStart -Duration 3 -Progress $profile.Code -SlipDays ($Sequence % 2) -IssueSeed (($Sequence % 4) + 1)
    $testCase = Get-PhaseData -PlanStart $planStart.AddDays(2) -Duration 3 -Progress $profile.TestCase -SlipDays ($Sequence % 3) -IssueSeed (($Sequence % 5) + 1)
    $design = Get-PhaseData -PlanStart $planStart.AddDays(4) -Duration 3 -Progress $profile.Design -SlipDays ($Sequence % 2) -IssueSeed (($Sequence % 3) + 1)
    $unit = Get-PhaseData -PlanStart $planStart.AddDays(7) -Duration 4 -Progress $profile.Unit -SlipDays ($Sequence % 4) -IssueSeed (($Sequence % 6) + 1)

    $subsystem = $subsystems[$Sequence % $subsystems.Count]
    $group = $groups[$Sequence % $groups.Count]
    $owner = $owners[$Sequence % $owners.Count]
    $manager = $managers[$Sequence % $managers.Count]
    $reviewer1 = $reviewers[$Sequence % $reviewers.Count]
    $reviewer2 = $reviewers[($Sequence + 1) % $reviewers.Count]
    $reviewer3 = $reviewers[($Sequence + 2) % $reviewers.Count]
    $reviewer4 = $reviewers[($Sequence + 3) % $reviewers.Count]
    $block = $blocks[$Sequence % $blocks.Count]
    $feature = $subsystem + $features[$Sequence % $features.Count]
    $programId = ('module_{0:000}.php' -f $Sequence)
    $note = $profile.Note + ' / batch-' + ('{0:00}' -f (($Sequence % 12) + 1))

    $cells = @(
        (New-NumberCell ('A' + $RowNumber) (($Sequence % 9) + 10)),
        (New-NumberCell ('B' + $RowNumber) (100 + ($Sequence % 20))),
        (New-NumberCell ('C' + $RowNumber) (($Sequence % 7) + 1)),
        (New-TextCell ('D' + $RowNumber) ('Wave-' + ('{0:00}' -f (($Sequence % 6) + 1)))),
        (New-TextCell ('E' + $RowNumber) $subsystem),
        (New-TextCell ('F' + $RowNumber) $feature),
        (New-TextCell ('G' + $RowNumber) $block),
        (New-TextCell ('H' + $RowNumber) $programId),
        (New-TextCell ('I' + $RowNumber) $group),
        (New-TextCell ('J' + $RowNumber) $owner),
        (New-TextCell ('K' + $RowNumber) $manager),
        (New-TextCell ('L' + $RowNumber) $reviewer1),
        (New-TextCell ('M' + $RowNumber) $reviewer2),
        (New-TextCell ('N' + $RowNumber) $reviewer3),
        (New-TextCell ('O' + $RowNumber) $reviewer4),
        (New-TextCell ('P' + $RowNumber) $owner),
        (New-TextCell ('Q' + $RowNumber) $profile.Status),
        (New-DateCell ('R' + $RowNumber) $code.PlanStart),
        (New-DateCell ('S' + $RowNumber) $code.PlanEnd),
        (New-DateCell ('T' + $RowNumber) $code.ActualStart),
        (New-DateCell ('U' + $RowNumber) $code.ActualEnd),
        (New-PercentCell ('V' + $RowNumber) $code.Progress),
        (New-DateCell ('W' + $RowNumber) $code.ReviewRequest),
        (New-DateCell ('X' + $RowNumber) $code.ReviewStart),
        (New-DateCell ('Y' + $RowNumber) $code.ReviewEnd),
        (New-TextCell ('Z' + $RowNumber) $code.ReviewStatus),
        (New-CountCell ('AA' + $RowNumber) $code.Issues),
        (New-DateCell ('AB' + $RowNumber) $code.Completed),
        (New-DateCell ('AC' + $RowNumber) $testCase.PlanStart),
        (New-DateCell ('AD' + $RowNumber) $testCase.PlanEnd),
        (New-DateCell ('AE' + $RowNumber) $testCase.ActualStart),
        (New-DateCell ('AF' + $RowNumber) $testCase.ActualEnd),
        (New-PercentCell ('AG' + $RowNumber) $testCase.Progress),
        (New-DateCell ('AH' + $RowNumber) $testCase.ReviewRequest),
        (New-DateCell ('AI' + $RowNumber) $testCase.ReviewStart),
        (New-DateCell ('AJ' + $RowNumber) $testCase.ReviewEnd),
        (New-TextCell ('AK' + $RowNumber) $testCase.ReviewStatus),
        (New-CountCell ('AL' + $RowNumber) $testCase.Issues),
        (New-DateCell ('AM' + $RowNumber) $testCase.Completed),
        (New-DateCell ('AN' + $RowNumber) $design.PlanStart),
        (New-DateCell ('AO' + $RowNumber) $design.PlanEnd),
        (New-DateCell ('AP' + $RowNumber) $design.ActualStart),
        (New-DateCell ('AQ' + $RowNumber) $design.ActualEnd),
        (New-PercentCell ('AR' + $RowNumber) $design.Progress),
        (New-DateCell ('AS' + $RowNumber) $design.ReviewRequest),
        (New-DateCell ('AT' + $RowNumber) $design.ReviewStart),
        (New-DateCell ('AU' + $RowNumber) $design.ReviewEnd),
        (New-TextCell ('AV' + $RowNumber) $design.ReviewStatus),
        (New-CountCell ('AW' + $RowNumber) $design.Issues),
        (New-DateCell ('AX' + $RowNumber) $design.Completed),
        (New-DateCell ('AY' + $RowNumber) $unit.PlanStart),
        (New-DateCell ('AZ' + $RowNumber) $unit.PlanEnd),
        (New-DateCell ('BA' + $RowNumber) $unit.ActualStart),
        (New-DateCell ('BB' + $RowNumber) $unit.ActualEnd),
        (New-PercentCell ('BC' + $RowNumber) $unit.Progress),
        (New-DateCell ('BD' + $RowNumber) $unit.ReviewRequest),
        (New-DateCell ('BE' + $RowNumber) $unit.ReviewStart),
        (New-DateCell ('BF' + $RowNumber) $unit.ReviewEnd),
        (New-TextCell ('BG' + $RowNumber) $unit.ReviewStatus),
        (New-CountCell ('BH' + $RowNumber) $unit.Issues),
        (New-DateCell ('BI' + $RowNumber) $unit.Completed),
        (New-TextCell ('BJ' + $RowNumber) $note)
    ) | Where-Object { $_ -ne '' }

    return '<row r="' + $RowNumber + '" spans="1:62">' + ($cells -join '') + '</row>'
}

if (-not (Test-Path $SourcePath)) {
    throw "Source workbook not found: $SourcePath"
}

if ($AddRows -lt 1) {
    throw 'AddRows must be 1 or greater.'
}

Copy-Item -Path $SourcePath -Destination $DestinationPath -Force

$zip = [System.IO.Compression.ZipFile]::Open($DestinationPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $sheetText = Get-EntryText -Zip $zip -EntryName 'xl/worksheets/sheet1.xml'
    $rowMatches = [regex]::Matches($sheetText, '<row r="(\d+)"')
    $existingRows = @($rowMatches | ForEach-Object { [int]$_.Groups[1].Value })
    $maxRow = ($existingRows | Measure-Object -Maximum).Maximum
    $startRow = $maxRow + 1

    $newRows = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $AddRows; $i++) {
        $rowNumber = $startRow + $i
        $sequence = ($rowNumber - 4)
        $null = $newRows.Append((New-DemoRowXml -RowNumber $rowNumber -Sequence $sequence))
    }

    $sheetText = $sheetText -replace '</sheetData>', ($newRows.ToString() + '</sheetData>')
    $sheetText = [regex]::Replace($sheetText, '<dimension ref="[^"]+"\/>', '<dimension ref="A1:BJ' + ($startRow + $AddRows - 1) + '"/>', 1)

    Set-EntryText -Zip $zip -EntryName 'xl/worksheets/sheet1.xml' -Text $sheetText
} finally {
    $zip.Dispose()
}

$dashboardScript = Join-Path $PSScriptRoot 'add_dashboard.ps1'
if (-not (Test-Path $dashboardScript)) {
    throw "Dashboard script not found: $dashboardScript"
}

& $dashboardScript -WorkbookPath $DestinationPath | Out-Null

Write-Output ('source_path=' + $SourcePath)
Write-Output ('destination_path=' + $DestinationPath)
Write-Output ('added_rows=' + $AddRows)
