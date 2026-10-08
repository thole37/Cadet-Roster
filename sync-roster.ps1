param(
  [switch]$Preview
)

$ErrorActionPreference = 'Stop'

$csvPath = Join-Path $PSScriptRoot "Hedith's Roster.csv"
$outputPath = Join-Path $PSScriptRoot 'roster-data.js'

if (-not (Test-Path -LiteralPath $csvPath)) {
  throw "Roster CSV not found: $csvPath"
}
if (-not (Test-Path -LiteralPath $outputPath)) {
  throw "Existing roster-data.js not found: $outputPath"
}

function Normalize-Text([object]$Value) {
  if ($null -eq $Value) { return '' }
  return ([regex]::Replace([string]$Value, '\s+', ' ')).Trim()
}

$csvRows = @(Import-Csv -LiteralPath $csvPath -Header @('No.', 'EMT Cert', 'LastName', 'FirstName', 'PR#', 'PH#', 'COMMENTS', 'Class', 'Extra1', 'Extra2', 'Extra3') | Select-Object -Skip 1)
if ($csvRows.Count -eq 0) {
  throw 'Roster CSV contains no data rows; no files were changed.'
}

$requiredColumns = @('No.', 'LastName', 'FirstName', 'PR#', 'Class')
foreach ($column in $requiredColumns) {
  if (-not $csvRows[0].PSObject.Properties[$column]) {
    throw "Roster CSV is missing the '$column' column; no files were changed."
  }
}

$roster = New-Object 'System.Collections.Generic.List[object]'
$resignedRecords = New-Object 'System.Collections.Generic.List[object]'
$resignedDates = [ordered]@{}
$peopleByClass = [ordered]@{}
foreach ($row in $csvRows) {
  $firstName = Normalize-Text $row.FirstName
  $first = ($firstName -split ' ')[0]
  $last = Normalize-Text $row.LastName
  if (-not $first -and -not $last) { continue }
  if (-not $first -or -not $last) {
    throw 'A CSV roster row has an incomplete name; no files were changed.'
  }

  $classCode = Normalize-Text $row.Class
  if ($classCode -notmatch '^(?<year>\d{4})\s*(?<letter>[A-Z])$') {
    throw "Could not find a class code for '$first $last'; no files were changed."
  }
  $className = $Matches['year'] + ' ' + $Matches['letter']

  $status = Normalize-Text $row.'No.'
  if ($status -notin @('0', '1')) {
    throw "Unexpected status '$status' for '$first $last'; expected 0 or 1. No files were changed."
  }
  $payroll = Normalize-Text $row.'PR#'
  if (-not $payroll) {
    throw "Missing payroll ID for '$first $last'; no files were changed."
  }

  if (-not $peopleByClass.Contains($className)) {
    $peopleByClass[$className] = New-Object 'System.Collections.Generic.List[object]'
  }
  $peopleByClass[$className].Add([ordered]@{
    class = $className
    first = $first
    last = $last
    payroll = $payroll
    status = $status
  })
}

if ($peopleByClass.Count -eq 0) {
  throw 'No roster rows with class codes were found in the CSV; no files were changed.'
}

foreach ($className in $peopleByClass.Keys) {
  $people = $peopleByClass[$className]
  $activeCadets = New-Object 'System.Collections.Generic.List[object]'
  $left = 0
  foreach ($person in $people) {
    if ($person.status -eq '1') {
      $activeCadets.Add([object[]]@($person.first, $person.last, $person.payroll))
    } else {
      $left++
      $resignedRecords.Add([object[]]@($person.class, $person.first, $person.last, $person.payroll))
    }
  }

  $roster.Add([ordered]@{
    name = $className
    started = $people.Count
    left = $left
    cadets = $activeCadets.ToArray()
  })
}

$rosterJson = ConvertTo-Json -InputObject @($roster.ToArray()) -Depth 10 -Compress
$resignedJson = ConvertTo-Json -InputObject @($resignedRecords.ToArray()) -Depth 10 -Compress
$datesJson = ConvertTo-Json -InputObject $resignedDates -Depth 5 -Compress
$output = @(
  "window.ROSTER_DATA = $rosterJson;"
  "window.RESIGNED_DATA = $resignedJson;"
  "window.RESIGNED_DATES = $datesJson;"
) -join [Environment]::NewLine

$activeCount = 0
foreach ($classData in $roster) { $activeCount += $classData['cadets'].Count }

if ($Preview) {
  Write-Output ("Preview only: {0} classes, {1} active cadets, {2} resigned records from the CSV. roster-data.js was not changed." -f $roster.Count,$activeCount,$resignedRecords.Count)
} else {
  $temporaryPath = "$outputPath.tmp"
  [IO.File]::WriteAllText($temporaryPath, $output, (New-Object System.Text.UTF8Encoding($false)))
  Move-Item -LiteralPath $temporaryPath -Destination $outputPath -Force
  Write-Output ("Sync complete: {0} classes, {1} active cadets, {2} resigned records from the CSV." -f $roster.Count,$activeCount,$resignedRecords.Count)
}