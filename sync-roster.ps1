param(
  [switch]$Preview
)

$ErrorActionPreference = 'Stop'

$workbookPath = Join-Path $PSScriptRoot "Hedith's Roster.xlsm"
$outputPath = Join-Path $PSScriptRoot 'roster-data.js'

if (-not (Test-Path -LiteralPath $workbookPath)) {
  throw "Workbook not found: $workbookPath"
}
if (-not (Test-Path -LiteralPath $outputPath)) {
  throw "Existing roster-data.js not found: $outputPath"
}

function Normalize-Text([object]$Value) {
  if ($null -eq $Value) { return '' }
  return ([regex]::Replace([string]$Value, '\s+', ' ')).Trim()
}

function Get-PersonKey([string]$ClassName, [string]$First, [string]$Last) {
  return ((Normalize-Text $ClassName) + '|' + (Normalize-Text $First) + '|' + (Normalize-Text $Last)).ToLowerInvariant()
}

function ConvertFrom-JavaScriptString([string]$Value) {
  return $Value.Replace('\\', '\').Replace("\'", "'")
}

function Get-CellValue($Sheet, [int]$Row, [int]$Column) {
  return $Sheet.Cells.Item($Row, $Column).Value2
}

function ConvertTo-RosterDate($Value) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }

  if ($Value -is [double] -or $Value -is [int] -or $Value -is [long]) {
    $date = [DateTime]::FromOADate([double]$Value)
  } else {
    $date = [DateTime]::Parse([string]$Value, [Globalization.CultureInfo]::CurrentCulture)
  }

  return $date.ToString('M/d/yy', [Globalization.CultureInfo]::InvariantCulture)
}

$previousData = Get-Content -LiteralPath $outputPath -Raw
$previousResigned = New-Object 'System.Collections.Generic.List[object]'
$previousDates = @{}
$resignedPattern = "\['(?<class>(?:\\.|[^'\\])*)',\s*'(?<first>(?:\\.|[^'\\])*)',\s*'(?<last>(?:\\.|[^'\\])*)',\s*'(?<payroll>\d+)'\]"

if ($previousData -match '(?s)window\.RESIGNED_DATA\s*=\s*\[(?<entries>.*?)\];') {
  foreach ($entry in [regex]::Matches($Matches['entries'], $resignedPattern)) {
    $className = ConvertFrom-JavaScriptString $entry.Groups['class'].Value
    $first = ConvertFrom-JavaScriptString $entry.Groups['first'].Value
    $last = ConvertFrom-JavaScriptString $entry.Groups['last'].Value
    $previousResigned.Add([ordered]@{
      class = $className
      first = $first
      last = $last
      payroll = $entry.Groups['payroll'].Value
    })
  }
}

if ($previousData -match '(?s)window\.RESIGNED_DATES\s*=\s*\{(?<dates>.*?)\};') {
  foreach ($entry in [regex]::Matches($Matches['dates'], "'(?<payroll>\d+)'\s*:\s*'(?<date>[^']*)'")) {
    $previousDates[$entry.Groups['payroll'].Value] = $entry.Groups['date'].Value
  }
}

if ($previousResigned.Count -eq 0) {
  throw 'Could not read existing resigned records from roster-data.js; no files were changed.'
}

$nameAliases = @{}
$nameAliases[(Get-PersonKey '2026 G' 'Blake' 'Sakesenberg')] = @('Blake', 'Saksenberg')
$nameAliases[(Get-PersonKey '2026 G' 'Marier' 'Davalos')] = @('Marifer', 'Davalos')
$nameAliases[(Get-PersonKey '2026 H' 'Efren' 'Rubio-Zavala')] = @('Efren', 'Rubio Zavala')
$nameAliases[(Get-PersonKey '2026 H' 'Laurence' 'Lucus')] = @('Laurence', 'Lucas')
$nameAliases[(Get-PersonKey '2026 H' 'Jecal' 'Gee.')] = @('Jecal', 'Gee')
$nameAliases[(Get-PersonKey '2026 H' 'Resignation' 'SejourNiaave')] = @('Niaave', 'Sejour')
$nameAliases[(Get-PersonKey '2026 H' 'Julian' 'Johnson')] = @('Julien', 'Johnson')

$excel = $null
$workbook = $null
$roster = New-Object 'System.Collections.Generic.List[object]'
$classByName = @{}
$fullRosterByClass = @{}
$peopleByNameByClass = @{}
$peopleByPayrollByClass = @{}
$statusByClass = @{}
$resignedRecords = New-Object 'System.Collections.Generic.List[object]'
$resignedDates = [ordered]@{}
$currentResignedByClass = @{}
$currentResignedPayrollKeys = @{}
$historicalByClass = @{}
$usedAliases = 0

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)

  foreach ($sheet in $workbook.Worksheets) {
    $className = Normalize-Text $sheet.Name
    if ($className -notmatch '^\d{4}\s+[A-Z]$') { continue }

    $fullCadets = New-Object 'System.Collections.Generic.List[object]'
    $peopleByName = @{}
    $peopleByPayroll = @{}
    $usedRange = $sheet.UsedRange
    $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1
    $lastColumn = $usedRange.Column + $usedRange.Columns.Count - 1
    $blockCount = 0

    for ($lastNameColumn = 1; $lastNameColumn -le $lastColumn; $lastNameColumn++) {
      $header = Normalize-Text (Get-CellValue $sheet 5 $lastNameColumn)
      if ($header -notmatch '^Last\s*Name$') { continue }

      $firstHeader = Normalize-Text (Get-CellValue $sheet 5 ($lastNameColumn + 1))
      $payrollHeader = Normalize-Text (Get-CellValue $sheet 5 ($lastNameColumn + 2))
      if ($firstHeader -notmatch '^First\s*Name$' -or $payrollHeader -notmatch '^Payroll\s*#?$') {
        throw "Unexpected roster headers on '$className'; no files were changed."
      }
      $blockCount++

      for ($row = 6; $row -le $lastRow; $row++) {
        $last = Normalize-Text (Get-CellValue $sheet $row $lastNameColumn)
        $first = Normalize-Text (Get-CellValue $sheet $row ($lastNameColumn + 1))
        $payroll = Normalize-Text (Get-CellValue $sheet $row ($lastNameColumn + 2))
        if (-not $first -and -not $last) { continue }
        if (-not $first -or -not $last -or -not $payroll) {
          throw "Incomplete starting-roster row on '$className', row $row; no files were changed."
        }

        $nameKey = Get-PersonKey $className $first $last
        if ($peopleByName.ContainsKey($nameKey) -or $peopleByPayroll.ContainsKey($payroll)) {
          throw "Duplicate name or payroll ID on '$className', row $row; no files were changed."
        }

        $person = [ordered]@{ first = $first; last = $last; payroll = $payroll }
        $fullCadets.Add($person)
        $peopleByName[$nameKey] = $person
        $peopleByPayroll[$payroll] = $person
      }
    }

    if ($blockCount -eq 0 -or $fullCadets.Count -eq 0) {
      throw "No roster blocks found on '$className'; no files were changed."
    }

    $classData = [ordered]@{ name = $className; started = $fullCadets.Count; left = 0; cadets = @() }
    $roster.Add($classData)
    $classByName[$className] = $classData
    $fullRosterByClass[$className] = $fullCadets.ToArray()
    $peopleByNameByClass[$className] = $peopleByName
    $peopleByPayrollByClass[$className] = $peopleByPayroll
  }

  if ($roster.Count -eq 0) {
    throw 'No class worksheets named like "2026 D" were found; no files were changed.'
  }

  $statusSheet = $workbook.Worksheets.Item('Status')
  $statusRange = $statusSheet.UsedRange
  $statusLastRow = $statusRange.Row + $statusRange.Rows.Count - 1
  for ($row = 2; $row -le $statusLastRow; $row++) {
    $className = Normalize-Text (Get-CellValue $statusSheet $row 2)
    if ($className -notmatch '^\d{4}\s+[A-Z]$') { continue }
    $activeCount = Get-CellValue $statusSheet $row 3
    if ($null -eq $activeCount -or -not $classByName.ContainsKey($className)) {
      throw "Missing or unmatched active count for '$className' on Status; no files were changed."
    }
    $statusByClass[$className] = [int]$activeCount
  }

  $resignedSheet = $workbook.Worksheets.Item('Resigned')
  $resignedRange = $resignedSheet.UsedRange
  $resignedLastRow = $resignedRange.Row + $resignedRange.Rows.Count - 1
  $resignedLastColumn = $resignedRange.Column + $resignedRange.Columns.Count - 1

  for ($classColumn = 1; $classColumn -le $resignedLastColumn; $classColumn++) {
    $classLabel = Normalize-Text (Get-CellValue $resignedSheet 2 $classColumn)
    if ($classLabel -notmatch '^(?<year>\d{4})\s*(?<letter>[A-Z])$') { continue }

    $className = $Matches['year'] + ' ' + $Matches['letter']
    if (-not $classByName.ContainsKey($className)) {
      throw "Resigned sheet includes '$className', but no matching class worksheet was found; no files were changed."
    }
    if (-not $currentResignedByClass.ContainsKey($className)) {
      $currentResignedByClass[$className] = New-Object 'System.Collections.Generic.List[object]'
    }

    for ($row = 4; $row -le $resignedLastRow; $row++) {
      $last = Normalize-Text (Get-CellValue $resignedSheet $row ($classColumn - 1))
      $first = Normalize-Text (Get-CellValue $resignedSheet $row $classColumn)
      if (-not $first -and -not $last) { continue }
      if (-not $first -or -not $last) {
        throw "Incomplete resignation row on '$className', row $row; no files were changed."
      }

      $nameKey = Get-PersonKey $className $first $last
      $lookupFirst = $first
      $lookupLast = $last
      if ($nameAliases.ContainsKey($nameKey)) {
        $alias = $nameAliases[$nameKey]
        $lookupFirst = $alias[0]
        $lookupLast = $alias[1]
        $usedAliases++
      }
      $canonicalKey = Get-PersonKey $className $lookupFirst $lookupLast
      if (-not $peopleByNameByClass[$className].ContainsKey($canonicalKey)) {
        throw "Could not match resignation '$first $last' to a person on '$className'; no files were changed."
      }

      $person = $peopleByNameByClass[$className][$canonicalKey]
      $payrollKey = $className + '|' + $person.payroll
      if ($currentResignedPayrollKeys.ContainsKey($payrollKey)) {
        throw "Duplicate resignation for '$first $last' in '$className'; no files were changed."
      }
      $currentResignedPayrollKeys[$payrollKey] = $true
      $currentResignedByClass[$className].Add([ordered]@{
        class = $className
        first = $person.first
        last = $person.last
        payroll = $person.payroll
      })

      $dateValue = Get-CellValue $resignedSheet $row ($classColumn + 1)
      if ($null -ne $dateValue -and -not [string]::IsNullOrWhiteSpace([string]$dateValue)) {
        $resignedDates[$person.payroll] = ConvertTo-RosterDate $dateValue
      }
    }
  }

  $historicalKeys = @{}
  foreach ($previous in $previousResigned) {
    $className = Normalize-Text $previous.class
    if (-not $peopleByPayrollByClass.ContainsKey($className) -or -not $peopleByPayrollByClass[$className].ContainsKey($previous.payroll)) {
      throw "Previously exported resignation '$($previous.first) $($previous.last)' is missing from '$className'; no files were changed."
    }

    $payrollKey = $className + '|' + $previous.payroll
    if ($currentResignedPayrollKeys.ContainsKey($payrollKey) -or $historicalKeys.ContainsKey($payrollKey)) { continue }
    if (-not $historicalByClass.ContainsKey($className)) {
      $historicalByClass[$className] = New-Object 'System.Collections.Generic.List[object]'
    }

    $person = $peopleByPayrollByClass[$className][$previous.payroll]
    $historicalKeys[$payrollKey] = $true
    $historicalByClass[$className].Add([ordered]@{
      class = $className
      first = $person.first
      last = $person.last
      payroll = $person.payroll
    })
  }

  foreach ($classData in $roster) {
    $className = $classData['name']
    if (-not $statusByClass.ContainsKey($className)) {
      throw "No active count found for '$className' on Status; no files were changed."
    }

    $fullCadets = $fullRosterByClass[$className]
    $expectedLeft = $fullCadets.Count - $statusByClass[$className]
    $current = @()
    if ($currentResignedByClass.ContainsKey($className)) { $current = $currentResignedByClass[$className].ToArray() }
    if ($current.Count -gt $expectedLeft) {
      throw "'$className' lists $($current.Count) resignations, but Status implies only $expectedLeft; reconcile those sheets before syncing."
    }

    $selected = New-Object 'System.Collections.Generic.List[object]'
    $selectedPayrolls = @{}
    foreach ($person in $current) {
      $selected.Add($person)
      $selectedPayrolls[$person.payroll] = $true
    }

    $needed = $expectedLeft - $selected.Count
    if ($needed -gt 0 -and $historicalByClass.ContainsKey($className)) {
      foreach ($person in $historicalByClass[$className]) {
        if ($needed -eq 0) { break }
        if ($selectedPayrolls.ContainsKey($person.payroll)) { continue }
        $selected.Add($person)
        $selectedPayrolls[$person.payroll] = $true
        $needed--
        if ($previousDates.ContainsKey($person.payroll)) {
          $resignedDates[$person.payroll] = $previousDates[$person.payroll]
        }
      }
    }
    if ($needed -ne 0) {
      throw "'$className' is missing $needed resignation record(s) needed to match Status. Add them to the Resigned sheet before syncing."
    }

    $activeCadets = New-Object 'System.Collections.Generic.List[object]'
    foreach ($person in $fullCadets) {
      if (-not $selectedPayrolls.ContainsKey($person.payroll)) {
        $activeCadets.Add([object[]]@($person.first, $person.last, $person.payroll))
      }
    }
    if ($activeCadets.Count -ne $statusByClass[$className]) {
      throw "Active roster count for '$className' does not match Status; no files were changed."
    }

    $classData['started'] = $fullCadets.Count
    $classData['left'] = $selected.Count
    $classData['cadets'] = $activeCadets.ToArray()
    foreach ($person in $selected) { $resignedRecords.Add([object[]]@($person.class, $person.first, $person.last, $person.payroll)) }
  }

  $rosterJson = ConvertTo-Json -InputObject @($roster.ToArray()) -Depth 10 -Compress
  $resignedJson = ConvertTo-Json -InputObject @($resignedRecords.ToArray()) -Depth 10 -Compress
  $datesJson = ConvertTo-Json -InputObject $resignedDates -Depth 5 -Compress
  $output = @(
    "window.ROSTER_DATA = $rosterJson;"
    "window.RESIGNED_DATA = $resignedJson;"
    "window.RESIGNED_DATES = $datesJson;"
  ) -join [Environment]::NewLine
} finally {
  if ($workbook) {
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }
  if ($excel) {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

$activeCount = 0
foreach ($classData in $roster) { $activeCount += $classData['cadets'].Count }

if ($Preview) {
  Write-Output ("Preview only: {0} classes, {1} active cadets, {2} resigned records, {3} verified name aliases. roster-data.js was not changed." -f $roster.Count,$activeCount,$resignedRecords.Count,$usedAliases)
} else {
  $temporaryPath = "$outputPath.tmp"
  [IO.File]::WriteAllText($temporaryPath, $output, (New-Object System.Text.UTF8Encoding($false)))
  Move-Item -LiteralPath $temporaryPath -Destination $outputPath -Force
  Write-Output ("Sync complete: {0} classes, {1} active cadets, {2} resigned records, {3} verified name aliases." -f $roster.Count,$activeCount,$resignedRecords.Count,$usedAliases)
}