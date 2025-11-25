<#
    Plex Missing Media GUI Tool (Responsive)
    - GUI for scanning Plex TV, Movies, and Anime libraries for missing items on a given drive/root
    - Uses sqlite3.exe at C:\tools\sqlite3.exe
    - Outputs CSVs listing series/movies that need to be reacquired
    - READ ONLY: does not modify Plex
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web   # for HttpUtility.HtmlEncode

#-----------------------------
# GLOBAL CONFIG
#-----------------------------
$Global:PlexDbPath = Join-Path $env:LOCALAPPDATA `
    "Plex Media Server\Plug-in Support\Databases\com.plexapp.plugins.library.db"

$Global:SqlitePath = "C:\tools\sqlite3.exe"

function Write-Log {
    param(
        [System.Windows.Forms.TextBox]$LogBox,
        [string]$Message
    )
    if (-not $LogBox) { return }
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $LogBox.AppendText("[$timestamp] $Message`r`n")
    $LogBox.SelectionStart = $LogBox.Text.Length
    $LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Run-PlexScanGrouped {
    param(
        [int]$SectionType,  # 1 = Movies, 2 = TV/Anime
        [string]$LostRoot,
        [string]$OutputCsv,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.WebBrowser]$Browser
    )

    Write-Log $LogBox "----------------------------------------"
    Write-Log $LogBox "Starting scan (section_type=$SectionType)..."
    Write-Log $LogBox "Plex DB: $Global:PlexDbPath"
    Write-Log $LogBox "sqlite3: $Global:SqlitePath"
    Write-Log $LogBox "Lost root: $LostRoot"
    Write-Log $LogBox "Output CSV: $OutputCsv"

    if (-not (Test-Path $Global:PlexDbPath)) {
        Write-Log $LogBox "ERROR: Plex DB not found at $Global:PlexDbPath"
        [System.Windows.Forms.MessageBox]::Show("Plex DB not found.`r`n$Global:PlexDbPath",
            "Error", 'OK', 'Error') | Out-Null
        return
    }
    if (-not (Test-Path $Global:SqlitePath)) {
        Write-Log $LogBox "ERROR: sqlite3.exe not found at $Global:SqlitePath"
        [System.Windows.Forms.MessageBox]::Show("sqlite3.exe not found.`r`n$Global:SqlitePath",
            "Error", 'OK', 'Error') | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($LostRoot)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a drive/root (e.g. F:\).",
            "Input Required", 'OK', 'Warning') | Out-Null
        return
    }

    # Normalize LostRoot like "F:" -> "F:\"
    if ($LostRoot.Length -eq 2 -and $LostRoot[1] -ne '\') {
        $LostRoot = $LostRoot + "\"
    }

    try {
        $tempDb = Join-Path $env:TEMP ("plex_scan_{0}.db" -f (Get-Date -Format "yyyyMMddHHmmss"))
        Copy-Item $Global:PlexDbPath $tempDb -Force
        Write-Log $LogBox "Copied Plex DB to temp: $tempDb"
    } catch {
        Write-Log $LogBox "ERROR copying DB: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to copy Plex DB: $_",
            "Error", 'OK', 'Error') | Out-Null
        return
    }

    # Choose SQL based on section type
    if ($SectionType -eq 2) {
        # TV / Anime
        $sql = @"
SELECT
  COALESCE(grandparent.title, parent.title, mi.title) AS SeriesTitle,
  CASE WHEN mi.metadata_type = 4 THEN mi.title ELSE '' END AS EpisodeTitle,
  parent.[index] AS SeasonNumber,
  mi.[index] AS EpisodeNumber,
  mp.file AS FilePath
FROM media_parts mp
JOIN media_items mitems ON mp.media_item_id = mitems.id
JOIN metadata_items mi ON mitems.metadata_item_id = mi.id
LEFT JOIN metadata_items parent ON mi.parent_id = parent.id
LEFT JOIN metadata_items grandparent ON parent.parent_id = grandparent.id
JOIN library_sections ls ON mitems.library_section_id = ls.id
WHERE ls.section_type = 2
  AND mp.file IS NOT NULL
ORDER BY SeriesTitle, SeasonNumber, EpisodeNumber;
"@
    } else {
        # Movies (section_type = 1)
        $sql = @"
SELECT
  mi.title AS MovieTitle,
  mp.file AS FilePath
FROM media_parts mp
JOIN media_items mitems ON mp.media_item_id = mitems.id
JOIN metadata_items mi ON mitems.metadata_item_id = mi.id
JOIN library_sections ls ON mitems.library_section_id = ls.id
WHERE ls.section_type = 1
  AND mp.file IS NOT NULL
ORDER BY MovieTitle;
"@
    }

    try {
        Write-Log $LogBox "Running sqlite3 query (can be slow on big libraries)..."
        $rawLines = & $Global:SqlitePath -batch -noheader $tempDb $sql
    } catch {
        Write-Log $LogBox "ERROR running sqlite3: $_"
        [System.Windows.Forms.MessageBox]::Show("sqlite3 query failed: $_",
            "Error", 'OK', 'Error') | Out-Null
        Remove-Item $tempDb -ErrorAction SilentlyContinue
        return
    }

    if (-not $rawLines -or $rawLines.Count -eq 0) {
        Write-Log $LogBox "No records returned from DB for this section type."
        Remove-Item $tempDb -ErrorAction SilentlyContinue
        return
    }

    Write-Log $LogBox "Got $($rawLines.Count) DB rows, processing..."
    [System.Windows.Forms.Application]::DoEvents()

    if ($SectionType -eq 2) {
        # TV / Anime grouping: Series -> Seasons with missing episodes
        $missingBySeries = @{}
        $i = 0
        $total = $rawLines.Count

        foreach ($line in $rawLines) {
            $i++
            if ($i % 1000 -eq 0) {
                Write-Log $LogBox ("Processed {0}/{1} rows..." -f $i, $total)
            }

            $parts = $line -split '\|', 5
            if ($parts.Count -lt 5) { continue }

            $series   = $parts[0]
            $season   = $parts[2]
            $filePath = $parts[4]

            if ($filePath -notlike "$LostRoot*") { continue }

            $exists = Test-Path -LiteralPath $filePath
            if ($exists) { continue }

            if ([string]::IsNullOrWhiteSpace($series)) { $series = "<Unknown Series>" }

            $seasonNum = 0
            [int]::TryParse($season, [ref]$seasonNum) | Out-Null

            if (-not $missingBySeries.ContainsKey($series)) {
                $missingBySeries[$series] = New-Object System.Collections.ArrayList
            }
            $seasonList = $missingBySeries[$series]
            if ($seasonNum -gt 0 -and -not $seasonList.Contains($seasonNum)) {
                [void]$seasonList.Add($seasonNum)
            }
        }

        Remove-Item $tempDb -ErrorAction SilentlyContinue

        if ($missingBySeries.Count -eq 0) {
            Write-Log $LogBox "No missing TV/Anime episodes detected under $LostRoot."
            if ($Browser) {
                $Browser.DocumentText = "<html><body><h2>No missing items</h2><p>No missing TV/Anime detected for <b>$LostRoot</b>.</p></body></html>"
            }
            [System.Windows.Forms.MessageBox]::Show("No missing TV/Anime detected for drive/root $LostRoot.",
                "Scan Complete", 'OK', 'Information') | Out-Null
            return
        }

        Write-Log $LogBox ("Found {0} series with missing seasons." -f $missingBySeries.Count)

        # Build output objects for CSV
        $outputRows = @()
        foreach ($series in ($missingBySeries.Keys | Sort-Object)) {
            $seasonList = $missingBySeries[$series]
            $sortedSeasons = $seasonList | Sort-Object
            $seasonText = ($sortedSeasons -join ", ")
            $outputRows += [pscustomobject]@{
                SeriesTitle    = $series
                SeasonsMissing = $seasonText
                LostRoot       = $LostRoot
            }
        }

        # Ensure output dir exists
        try {
            $outDir = Split-Path $OutputCsv -Parent
            if (-not (Test-Path $outDir)) {
                New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            }

            $outputRows |
                Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

            Write-Log $LogBox "CSV written to $OutputCsv"
        } catch {
            Write-Log $LogBox "ERROR writing CSV: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to write CSV: $_",
                "Error", 'OK', 'Error') | Out-Null
        }

        # Build simple HTML summary for browser
        if ($Browser) {
            $html = @"
<html>
<head>
<meta charset=""utf-8"" />
<title>Plex Missing TV/Anime</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 10px; }
h2 { margin-top: 0; }
.series { margin-bottom: 6px; }
.series-title { font-weight: bold; }
.seasons { color: #444; }
</style>
</head>
<body>
<h2>Missing TV / Anime for $LostRoot</h2>
<p>Series and seasons that have at least one missing episode on this drive.</p>
<ul>
"@
            foreach ($series in ($missingBySeries.Keys | Sort-Object)) {
                $seasonList = $missingBySeries[$series] | Sort-Object
                $seasonText = ($seasonList -join ", ")
                $safeSeries = [System.Web.HttpUtility]::HtmlEncode($series)
                $html += "<li class='series'><span class='series-title'>$safeSeries</span>: <span class='seasons'>Seasons $seasonText</span></li>`n"
            }

            $html += @"
</ul>
<p><em>CSV saved to $OutputCsv</em></p>
</body>
</html>
"@
            $Browser.DocumentText = $html
        }

        [System.Windows.Forms.MessageBox]::Show("Scan complete.`r`nSeries with missing seasons: $($missingBySeries.Count)`r`nCSV: $OutputCsv",
            "Scan Complete", 'OK', 'Information') | Out-Null

    } else {
        # Movies: group by MovieTitle (only list each movie once)
        $missingMovies = New-Object System.Collections.Generic.HashSet[string]
        $i = 0
        $total = $rawLines.Count

        foreach ($line in $rawLines) {
            $i++
            if ($i % 1000 -eq 0) {
                Write-Log $LogBox ("Processed {0}/{1} rows..." -f $i, $total)
            }

            $parts = $line -split '\|', 2
            if ($parts.Count -lt 2) { continue }

            $movie    = $parts[0]
            $filePath = $parts[1]

            if ($filePath -notlike "$LostRoot*") { continue }

            $exists = Test-Path -LiteralPath $filePath
            if ($exists) { continue }

            if ([string]::IsNullOrWhiteSpace($movie)) { $movie = "<Unknown Movie>" }

            [void]$missingMovies.Add($movie)
        }

        Remove-Item $tempDb -ErrorAction SilentlyContinue

        if ($missingMovies.Count -eq 0) {
            Write-Log $LogBox "No missing movies detected under $LostRoot."
            if ($Browser) {
                $Browser.DocumentText = "<html><body><h2>No missing movies</h2><p>No missing movies detected for <b>$LostRoot</b>.</p></body></html>"
            }
            [System.Windows.Forms.MessageBox]::Show("No missing movies detected for drive/root $LostRoot.",
                "Scan Complete", 'OK', 'Information') | Out-Null
            return
        }

        Write-Log $LogBox ("Found {0} missing movies." -f $missingMovies.Count)

        # Build CSV rows
        $movieRows = @()
        foreach ($movie in ($missingMovies | Sort-Object)) {
            $movieRows += [pscustomobject]@{
                MovieTitle = $movie
                LostRoot   = $LostRoot
            }
        }

        try {
            $outDir = Split-Path $OutputCsv -Parent
            if (-not (Test-Path $outDir)) {
                New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            }

            $movieRows |
                Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

            Write-Log $LogBox "CSV written to $OutputCsv"
        } catch {
            Write-Log $LogBox "ERROR writing CSV: $_"
            [System.Windows.Forms.MessageBox]::Show("Failed to write CSV: $_",
                "Error", 'OK', 'Error') | Out-Null
        }

        if ($Browser) {
            $html = @"
<html>
<head>
<meta charset=""utf-8"" />
<title>Plex Missing Movies</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 10px; }
h2 { margin-top: 0; }
.movie { margin-bottom: 4px; }
</style>
</head>
<body>
<h2>Missing Movies for $LostRoot</h2>
<p>Movies that have at least one missing file on this drive.</p>
<ul>
"@
            foreach ($movie in ($missingMovies | Sort-Object)) {
                $safeMovie = [System.Web.HttpUtility]::HtmlEncode($movie)
                $html += "<li class='movie'>$safeMovie</li>`n"
            }
            $html += @"
</ul>
<p><em>CSV saved to $OutputCsv</em></p>
</body>
</html>
"@
            $Browser.DocumentText = $html
        }

        [System.Windows.Forms.MessageBox]::Show("Scan complete.`r`nMissing movies: $($missingMovies.Count)`r`nCSV: $OutputCsv",
            "Scan Complete", 'OK', 'Information') | Out-Null
    }

    Write-Log $LogBox "Scan finished."
}

#-----------------------------
# Build GUI with TabControl
#-----------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Plex Missing Media Scanner"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$form.Controls.Add($tabs)

function New-ScanTab {
    param(
        [string]$Title,
        [string]$DefaultDrive,
        [string]$DefaultOutput
    )

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $Title

    $lblDrive = New-Object System.Windows.Forms.Label
    $lblDrive.Text = "Drive / Root to scan (e.g. F:\):"
    $lblDrive.Location = New-Object System.Drawing.Point(10, 15)
    $lblDrive.AutoSize = $true
    $tab.Controls.Add($lblDrive)

    $txtDrive = New-Object System.Windows.Forms.TextBox
    $txtDrive.Location = New-Object System.Drawing.Point(220, 12)
    $txtDrive.Size = New-Object System.Drawing.Size(150, 20)
    $txtDrive.Text = $DefaultDrive
    $tab.Controls.Add($txtDrive)

    $lblOutput = New-Object System.Windows.Forms.Label
    $lblOutput.Text = "Output CSV path:"
    $lblOutput.Location = New-Object System.Drawing.Point(10, 45)
    $lblOutput.AutoSize = $true
    $tab.Controls.Add($lblOutput)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(220, 42)
    $txtOutput.Size = New-Object System.Drawing.Size(450, 20)
    $txtOutput.Text = $DefaultOutput
    $tab.Controls.Add($txtOutput)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point(680, 40)
    $btnBrowse.Size = New-Object System.Drawing.Size(80, 23)
    $tab.Controls.Add($btnBrowse)

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan for Missing"
    $btnScan.Location = New-Object System.Drawing.Point(10, 75)
    $btnScan.Size = New-Object System.Drawing.Size(150, 30)
    $tab.Controls.Add($btnScan)

    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text = "Log:"
    $lblLog.Location = New-Object System.Drawing.Point(10, 115)
    $lblLog.AutoSize = $true
    $tab.Controls.Add($lblLog)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location = New-Object System.Drawing.Point(10, 135)
    $txtLog.Size = New-Object System.Drawing.Size(300, 400)
    $txtLog.Multiline = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.ReadOnly = $true
    $tab.Controls.Add($txtLog)

    $browser = New-Object System.Windows.Forms.WebBrowser
    $browser.Location = New-Object System.Drawing.Point(320, 115)
    $browser.Size = New-Object System.Drawing.Size(540, 420)
    $tab.Controls.Add($browser)

    $btnBrowse.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $sfd.FileName = $txtOutput.Text
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutput.Text = $sfd.FileName
        }
    })

    return [pscustomobject]@{
        Tab      = $tab
        TxtDrive = $txtDrive
        TxtOutput= $txtOutput
        BtnScan  = $btnScan
        TxtLog   = $txtLog
        Browser  = $browser
    }
}

$tvTab    = New-ScanTab -Title "TV Shows" -DefaultDrive "F:\"       -DefaultOutput "C:\tools\PlexMissing_TV_F_Drive.csv"
$movieTab = New-ScanTab -Title "Movies"   -DefaultDrive "F:\"       -DefaultOutput "C:\tools\PlexMissing_Movies_F_Drive.csv"
$animeTab = New-ScanTab -Title "Anime"    -DefaultDrive "F:\Anime\" -DefaultOutput "C:\tools\PlexMissing_Anime_F_Drive.csv"

$tabs.TabPages.Add($tvTab.Tab)
$tabs.TabPages.Add($movieTab.Tab)
$tabs.TabPages.Add($animeTab.Tab)

$tvTab.BtnScan.Add_Click({
    $drive  = $tvTab.TxtDrive.Text.Trim()
    $output = $tvTab.TxtOutput.Text.Trim()
    $tvTab.BtnScan.Enabled = $false
    try {
        Run-PlexScanGrouped -SectionType 2 -LostRoot $drive -OutputCsv $output -LogBox $tvTab.TxtLog -Browser $tvTab.Browser
    } finally {
        $tvTab.BtnScan.Enabled = $true
    }
})

$movieTab.BtnScan.Add_Click({
    $drive  = $movieTab.TxtDrive.Text.Trim()
    $output = $movieTab.TxtOutput.Text.Trim()
    $movieTab.BtnScan.Enabled = $false
    try {
        Run-PlexScanGrouped -SectionType 1 -LostRoot $drive -OutputCsv $output -LogBox $movieTab.TxtLog -Browser $movieTab.Browser
    } finally {
        $movieTab.BtnScan.Enabled = $true
    }
})

$animeTab.BtnScan.Add_Click({
    $drive  = $animeTab.TxtDrive.Text.Trim()
    $output = $animeTab.TxtOutput.Text.Trim()
    $animeTab.BtnScan.Enabled = $false
    try {
        Run-PlexScanGrouped -SectionType 2 -LostRoot $drive -OutputCsv $output -LogBox $animeTab.TxtLog -Browser $animeTab.Browser
    } finally {
        $animeTab.BtnScan.Enabled = $true
    }
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
