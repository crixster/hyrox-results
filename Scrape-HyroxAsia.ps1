param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Women', 'Men', 'ProMen', 'DoublesMen', 'DoublesMenPro', 'DoublesWomen', 'DoublesMixed', 'TeamRelayMen', 'TeamRelayMixed', 'All')]
    [string[]]$Category,            # one or more names (e.g. Women,DoublesMen) or All

    [string]$OutDir = "",
    [string[]]$Only = @(),
    [int[]]$Seasons = @(8, 9),
    [string[]]$AcceptedSeasonYears = @('7-2025'),
    [int]$DelayMs = 2000,           # base pause after every request (random jitter of up to +50% is added)
    [int]$EventPauseSec = 5,        # extra pause after each event that was scraped
    [switch]$PlainTimes,
    [switch]$SkipAG,                # no longer used (age groups come from the ranking pages); kept so old commands still work
    [switch]$Only40Plus
)

$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$root   = "https://www.hyresult.com"
$hostRx = '(?:https://www\.hyresult\.com)?'

if (-not $OutDir) {
    # Default: the same folder as this script (the repo root), so the PWA finds the CSVs
    # next to its own files. $PSScriptRoot is empty if the code is pasted into a console,
    # so fall back to the current folder in that case.
    $OutDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$allCategories = @('Women', 'Men', 'ProMen', 'DoublesMen', 'DoublesMenPro', 'DoublesWomen', 'DoublesMixed', 'TeamRelayMen', 'TeamRelayMixed')
$categoriesToRun = if ($Category -contains 'All') {
    $allCategories
} else {
    @($allCategories | Where-Object { $Category -contains $_ })
}

$script:abort = $false
$script:blockedStreak = 0

function Get-Html([string]$url) {
    if ($script:abort) { return $null }
    for ($try = 1; $try -le 4; $try++) {
        try {
            $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 `
                 -Headers @{ "User-Agent" = "Mozilla/5.0 (personal research script)" }
            $script:blockedStreak = 0
            # base delay + random jitter so requests are not perfectly regular
            $jitter = Get-Random -Minimum 0 -Maximum ([Math]::Max(1, [int]($DelayMs / 2)))
            Start-Sleep -Milliseconds ($DelayMs + $jitter)
            return [string]$r.Content
        } catch {
            $status = 0
            if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch {} }

            # Page does not exist (e.g. past the last ranking page): retrying will not help
            if ($status -eq 404) { return $null }

            if ($status -in 403, 429, 503) {
                $script:blockedStreak++
                if ($script:blockedStreak -ge 5) {
                    $script:abort = $true
                    Write-Warning "The server keeps answering HTTP $status. This looks like rate limiting or an IP block. Stopping and saving what was collected so far. Wait a few hours before running again."
                    return $null
                }
                $wait = 30 * $try
                Write-Warning "  HTTP $status (rate limited?) for $url - waiting $wait s before retrying"
                Start-Sleep -Seconds $wait
            } else {
                Write-Warning "  attempt $try failed for $url : $($_.Exception.Message)"
                Start-Sleep -Seconds (2 * $try)
            }
        }
    }
    return $null
}

function Strip-Tags([string]$s, [string]$replacement = " ") {
    return [System.Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', $replacement)).Trim()
}

function Export-PlainCsv($objects, [string[]]$columns, [string]$path) {
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add(($columns -join ','))
    foreach ($o in $objects) {
        $vals = foreach ($c in $columns) {
            $v = [string]$o.$c
            if ($v -match '^="[^"]*"$') { $v }
            elseif ($v -match '[,"\r\n]') { '"' + $v.Replace('"', '""') + '"' }
            else { $v }
        }
        $lines.Add(($vals -join ','))
    }
    [System.IO.File]::WriteAllLines($path, $lines, (New-Object System.Text.UTF8Encoding($true)))
}

# ---------------------------------------------------------------
Write-Host "Step 1: discovering Asian locations..."

# NOTE: the Region filter on hyresult runs in the browser - "?con=AS" is ignored, so this page
# always lists every location worldwide. Also, only a few of the location cards are real
# <a href> links, so scraping links alone finds just a handful (it used to find only
# Jakarta, Bangkok and Singapore). So we combine three sources:
#   (a) the cards themselves: <h2>Name</h2> followed by "Country, Asia"
#   (b) any /location/ links that belong to an Asian city
#   (c) the built-in list below (a wrong guess is just a 404 in Step 2 and is skipped)
$locHtml = Get-Html "$root/locations"
if (-not $locHtml) { throw "Could not load the locations page." }

$asiaNames = @('Greater Noida','Mumbai','Osaka','Yokohama','Sanya','Seoul','Sharjah','Singapore','Wuhan',
               'Doha','Nagoya','Dubai','Guangzhou','Shanghai','Jakarta','Shenzhen','Taipei','Gandhinagar',
               'Gujarat','Chengdu','Hangzhou','Hong Kong','Bengaluru','New Delhi','Chiba','Incheon','Bangkok',
               'Kuala Lumpur','Abu Dhabi','Beijing','Hanoi')
$asiaSlugs = $asiaNames | ForEach-Object { ($_.ToLower() -replace '[^a-z0-9]+', '-').Trim('-') }

$locations = [ordered]@{}

# (a) location cards
$fromCards = 0
foreach ($hm in [regex]::Matches($locHtml, '<h2\b[^>]*>(.*?)</h2>', 'Singleline')) {
    $name = Strip-Tags $hm.Groups[1].Value
    if (-not $name) { continue }
    $tailStart = $hm.Index + $hm.Length
    $tail = Strip-Tags $locHtml.Substring($tailStart, [Math]::Min(400, $locHtml.Length - $tailStart))
    if ($tail -match '^[^,]{1,60},\s*Asia\b') {
        $slug = ($name.ToLower() -replace '[^a-z0-9]+', '-').Trim('-')
        if ($slug -and -not $locations.Contains($slug)) { $locations[$slug] = ""; $fromCards++ }
    }
}

# (b) plain links, keeping only Asian cities
$linkHits = [regex]::Matches($locHtml, 'href="' + $hostRx + '/location/([A-Za-z0-9_-]+)"')
if ($linkHits.Count -eq 0) {
    $linkHits = [regex]::Matches($locHtml, '\\?/location\\?/([A-Za-z0-9_-]+)')
}
$fromLinks = 0
foreach ($h in $linkHits) {
    $slug = $h.Groups[1].Value
    if ($locations.Contains($slug)) { continue }
    foreach ($a in $asiaSlugs) {
        if ($slug -eq $a -or $slug.StartsWith($a + "-")) { $locations[$slug] = ""; $fromLinks++; break }
    }
}

# (c) built-in list
$fromList = 0
foreach ($a in $asiaSlugs) {
    if (-not $locations.Contains($a)) { $locations[$a] = ""; $fromList++ }
}
Write-Host "  sources: $fromCards from page cards, $fromLinks from links, $fromList from built-in list"

if ($Only -and $Only.Count -gt 0) {
    $keep = [ordered]@{}
    foreach ($o in $Only) {
        $hit = @($locations.Keys | Where-Object { $_ -ieq $o.Trim() })
        if ($hit.Count -gt 0) { $keep[$hit[0]] = "" }
    }
    $locations = $keep
}
Write-Host "  Found $($locations.Count) location(s): $(@($locations.Keys) -join ', ')"
if ($Only -and $Only.Count -gt 0 -and $locations.Count -eq 0) {
    Write-Warning "None of the -Only names matched a discovered location slug. Run without -Only to see the available slugs."
}

# ---------------------------------------------------------------
Write-Host "Step 2: collecting event links from each location..."
$seenEvents = New-Object 'System.Collections.Generic.HashSet[string]'
$events     = New-Object 'System.Collections.Generic.List[object]'

$locIndex = 0
foreach ($slug in @($locations.Keys)) {
    $locIndex++
    Write-Host ("  [{0}/{1}] {2}" -f $locIndex, $locations.Count, $slug)

    $html = Get-Html "$root/location/$slug"
    if (-not $html) {
        Write-Host "      could not load location page" -ForegroundColor Yellow
        continue
    }

    $linksFound = 0
    $newEvents  = 0
    $ignored    = 0
    foreach ($m in [regex]::Matches($html, 'href="' + $hostRx + '/event/([^"#?/]+)"')) {
        $ev = $m.Groups[1].Value
        $linksFound++

        # Show whether the event passes the season filter used later in Step 3
        $sm0 = [regex]::Match($ev, '^s(\d+)-(\d{4})-')
        $inSeason = $false
        if ($sm0.Success) {
            $pair0 = "{0}-{1}" -f [int]$sm0.Groups[1].Value, $sm0.Groups[2].Value
            $inSeason = ($Seasons -contains [int]$sm0.Groups[1].Value) -or ($AcceptedSeasonYears -contains $pair0)
        }

        # With -Only, location pages may also link to events from other cities.
        # Keep only events whose slug is s<N>-<year>-<location>...
        if ($Only -and $Only.Count -gt 0) {
            if ($ev -notmatch ('^s\d+-\d{4}-' + [regex]::Escape($slug) + '(-|$)')) {
                $ignored++
                continue
            }
        }

        if ($seenEvents.Add($ev)) {
            $newEvents++
            $events.Add([pscustomobject]@{ LocationSlug = $slug; EventSlug = $ev })
            if ($inSeason) {
                Write-Host "      + $ev  (in selected seasons)" -ForegroundColor Green
            } else {
                Write-Host "      - $ev  (outside selected seasons, will be skipped)" -ForegroundColor DarkGray
            }
        }
    }
    Write-Host ("      {0} event link(s) on page, {1} new, {2} ignored (other locations)" -f $linksFound, $newEvents, $ignored)
}
Write-Host "  $($events.Count) unique event(s) in total."

# ---------------------------------------------------------------
$written = New-Object 'System.Collections.Generic.List[string]'

foreach ($cat in $categoriesToRun) {
    Write-Host ""
    Write-Host "=========================================="
    Write-Host "Scraping Category: $cat"
    Write-Host "=========================================="

    # $layout decides which columns the CSV gets: solo (1 athlete), doubles (2 athletes), relay (up to 4)
    switch ($cat) {
        'Women'          { $suffix = '-hyrox-women';              $layout = 'solo';    $fileName = 'Hyrox_Asia_Women.csv' }
        'Men'            { $suffix = '-hyrox-men';                $layout = 'solo';    $fileName = 'Hyrox_Asia_Men.csv' }
        'ProMen'         { $suffix = '-hyrox-pro-men';            $layout = 'solo';    $fileName = 'Hyrox_Asia_ProMen.csv' }
        'DoublesMen'     { $suffix = '-hyrox-doubles-men';        $layout = 'doubles'; $fileName = 'Hyrox_Asia_DoublesMen.csv' }
        'DoublesMenPro'  { $suffix = '-hyrox-pro-doubles-men';    $layout = 'doubles'; $fileName = 'Hyrox_Asia_DoublesMenPro.csv' }
        'DoublesWomen'   { $suffix = '-hyrox-doubles-women';      $layout = 'doubles'; $fileName = 'Hyrox_Asia_DoublesWomen.csv' }
        'DoublesMixed'   { $suffix = '-hyrox-doubles-mixed';      $layout = 'doubles'; $fileName = 'Hyrox_Asia_DoublesMixed.csv' }
        'TeamRelayMen'   { $suffix = '-hyrox-team-relay-men';     $layout = 'relay';   $fileName = 'Hyrox_Asia_TeamRelayMen.csv' }
        'TeamRelayMixed' { $suffix = '-hyrox-team-relay-mixed';   $layout = 'relay';   $fileName = 'Hyrox_Asia_TeamRelayMixed.csv' }
    }

    $OutFile = Join-Path $OutDir $fileName
    $logFile = [System.IO.Path]::ChangeExtension($OutFile, $null).TrimEnd('.') + "_EventLog.csv"

    $seenResults = New-Object 'System.Collections.Generic.HashSet[string]'
    $rows        = New-Object 'System.Collections.Generic.List[object]'
    $log         = New-Object 'System.Collections.Generic.List[object]'
    $textInfo    = (Get-Culture).TextInfo

    foreach ($e in $events) {
        if ($script:abort) { Write-Warning "Stopping early (rate limited). Remaining events skipped."; break }
        $sm = [regex]::Match($e.EventSlug, '^s(\d+)-(\d{4})-')
        $seasonOk = $false
        if ($sm.Success) {
            $pair = "{0}-{1}" -f [int]$sm.Groups[1].Value, $sm.Groups[2].Value
            $seasonOk = ($Seasons -contains [int]$sm.Groups[1].Value) -or ($AcceptedSeasonYears -contains $pair)
        }
        if (-not $seasonOk) { continue }

        $evHtml = Get-Html "$root/event/$($e.EventSlug)"
        if (-not $evHtml) { continue }

        $rm = [regex]::Match($evHtml, "(/ranking/[A-Za-z0-9_-]*$suffix)(?![A-Za-z0-9_-])")
        if (-not $rm.Success) {
            $log.Add([pscustomobject]@{ Event = $e.EventSlug; Outcome = "no $cat ranking link"; Entries = 0 })
            continue
        }
        $rankPath = $rm.Groups[1].Value

        $sn      = [int]$sm.Groups[1].Value
        $season  = "S$sn ({0}/{1})" -f (17 + $sn), (18 + $sn)
        $year    = $sm.Groups[2].Value
        $locName = $textInfo.ToTitleCase(($e.LocationSlug -replace '-', ' '))

        $eventDate = ""
        $h1 = $evHtml.IndexOf("<h1")
        if ($h1 -ge 0) {
            $evText = Strip-Tags $evHtml.Substring($h1) " "
            $dm = [regex]::Match($evText, '((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:\s*[\u2013-]\s*\d{1,2})?(?:,\s*\d{4})?)')
            if ($dm.Success) { $eventDate = $dm.Groups[1].Value }
        }

        $rank = 0
        $entriesThisEvent = 0
        $firstRow = $rows.Count

        for ($page = 1; $page -le 50; $page++) {
            $html = Get-Html "$root$rankPath`?p=$page"
            if (-not $html) { break }
            Write-Host "    [$cat] $($e.EventSlug) page $page ..." -NoNewline
            $newOnPage = 0

            foreach ($tr in [regex]::Matches($html, '<tr\b[^>]*>(.*?)</tr>', 'Singleline')) {
                $rowHtml = $tr.Groups[1].Value
                $idm = [regex]::Match($rowHtml, 'href="' + $hostRx + '/result/([A-Za-z0-9]+)"')
                if (-not $idm.Success) { continue }
                $rid = $idm.Groups[1].Value
                if (-not $seenResults.Add($rid)) { continue }
                $newOnPage++

                $names = New-Object 'System.Collections.Generic.List[string]'
                $ctrys = New-Object 'System.Collections.Generic.List[string]'
                $seenAth = @{}

                foreach ($am in [regex]::Matches($rowHtml, '<a\b[^>]*href="' + $hostRx + '/athlete/([^"#?]+)"[^>]*>(.*?)</a>', 'Singleline')) {
                    $aslug = $am.Groups[1].Value
                    if ($seenAth.ContainsKey($aslug)) { continue }
                    $seenAth[$aslug] = $true

                    $name = Strip-Tags $am.Groups[2].Value
                    $country = ""
                    $imgs = [regex]::Matches($rowHtml.Substring(0, $am.Index), '<img\b[^>]*>')
                    if ($imgs.Count -gt 0) {
                        $img = $imgs[$imgs.Count - 1].Value
                        $cm = [regex]::Match($img, '\b(?:alt|title)="([^"]+)"')
                        if ($cm.Success) { $country = $cm.Groups[1].Value }
                    }
                    $names.Add($name)
                    $ctrys.Add($country)
                }

                $text  = Strip-Tags $rowHtml " "
                $timeM = [regex]::Match($text, '(\d{1,2}(?::\d{2}){1,2})')
                $diffM = [regex]::Match($text, '\+\s*(\d{1,2}(?::\d{2}){1,2})')
                $time  = if ($timeM.Success) { $timeM.Groups[1].Value } else { "" }
                $diff  = if ($diffM.Success) { $diffM.Groups[1].Value } else { "" }

                # Age group badge shown on the row, e.g. 16-24, 25-29, 40-44, 80+ (solo)
                # or U40 / 40+ (team relay)
                $agM = [regex]::Match($text, '(?<![\d:+])(\d{2}\s*-\s*\d{2}|\d{2}\+|\bU(?:nder)?\s*\d{2})(?![\d:])')
                $ag  = if ($agM.Success) { ($agM.Groups[1].Value -replace '\s', '') -replace '^Under', 'U' } else { "" }

                $overall = ""
                if ($time) { $rank++; $overall = $rank }

                $teamCountries = ($ctrys | Where-Object { $_ } | Select-Object -Unique) -join "; "

                $timeOut = $time
                $diffOut = $diff
                if (-not $PlainTimes) {
                    if ($time) { $timeOut = '="' + $time + '"' }
                    if ($diff) { $diffOut = '="' + $diff + '"' }
                }

                if ($layout -eq 'relay') {
                    $rows.Add([pscustomobject]@{
                        Category      = $cat
                        Location      = $locName
                        Season        = $season
                        Year          = $year
                        EventDate     = $eventDate
                        EventSlug     = $e.EventSlug
                        OverallRank   = $overall
                        AG            = $ag
                        Time          = $timeOut
                        GapToWinner   = $diffOut
                        Athlete1      = if ($names.Count -ge 1) { $names[0] } else { "" }
                        Athlete2      = if ($names.Count -ge 2) { $names[1] } else { "" }
                        Athlete3      = if ($names.Count -ge 3) { $names[2] } else { "" }
                        Athlete4      = if ($names.Count -ge 4) { $names[3] } else { "" }
                        TeamCountries = $teamCountries
                        ResultId      = $rid
                    })
                } elseif ($layout -eq 'doubles') {
                    $rows.Add([pscustomobject]@{
                        Category    = $cat
                        Location    = $locName
                        Season      = $season
                        Year        = $year
                        EventDate   = $eventDate
                        EventSlug   = $e.EventSlug
                        OverallRank = $overall
                        AG          = $ag
                        Time        = $timeOut
                        GapToWinner = $diffOut
                        Athlete1    = if ($names.Count -ge 1) { $names[0] } else { "" }
                        Country1    = if ($ctrys.Count -ge 1) { $ctrys[0] } else { "" }
                        Athlete2    = if ($names.Count -ge 2) { $names[1] } else { "" }
                        Country2    = if ($ctrys.Count -ge 2) { $ctrys[1] } else { "" }
                        ResultId    = $rid
                    })
                } else {
                    $rows.Add([pscustomobject]@{
                        Category    = $cat
                        Location    = $locName
                        Season      = $season
                        Year        = $year
                        EventDate   = $eventDate
                        EventSlug   = $e.EventSlug
                        OverallRank = $overall
                        AG          = $ag
                        Time        = $timeOut
                        GapToWinner = $diffOut
                        Athlete     = if ($names.Count -ge 1) { $names[0] } else { "" }
                        Country     = if ($ctrys.Count -ge 1) { $ctrys[0] } else { "" }
                        ResultId    = $rid
                    })
                }
                $entriesThisEvent++
            }

            Write-Host " $newOnPage new entry/entries"
            if ($newOnPage -eq 0) { break }
        }

        if ($entriesThisEvent -gt 0) {
            $withAg = 0
            for ($i = $firstRow; $i -lt $rows.Count; $i++) { if ($rows[$i].AG) { $withAg++ } }
            Write-Host "    age group found for $withAg of $entriesThisEvent rows"
            if ($withAg -lt $entriesThisEvent) {
                Write-Warning "  $($entriesThisEvent - $withAg) of $entriesThisEvent rows have no age group for $($e.EventSlug) (check the badge text in the page HTML)"
            }
        }

        $outcome = if ($entriesThisEvent -gt 0) { "scraped" } else { "no results" }
        Write-Host ("  {0,-8} {1,-28} {2} entries" -f $outcome.Substring(0, [Math]::Min(8, $outcome.Length)), $e.EventSlug, $entriesThisEvent)
        $log.Add([pscustomobject]@{ Event = $e.EventSlug; Outcome = $outcome; Entries = $entriesThisEvent })
        if ($entriesThisEvent -gt 0 -and $EventPauseSec -gt 0 -and -not $script:abort) { Start-Sleep -Seconds $EventPauseSec }
    }

    if ($layout -eq 'relay') {
        $columns = 'Category','Location','Season','Year','EventDate','EventSlug','OverallRank','AG','Time','GapToWinner',
                   'Athlete1','Athlete2','Athlete3','Athlete4','TeamCountries'
    } elseif ($layout -eq 'doubles') {
        $columns = 'Category','Location','Season','Year','EventDate','EventSlug','OverallRank','AG','Time','GapToWinner',
                   'Athlete1','Country1','Athlete2','Country2'
    } else {
        $columns = 'Category','Location','Season','Year','EventDate','EventSlug','OverallRank','AG','Time','GapToWinner',
                   'Athlete','Country'
    }

    Export-PlainCsv $rows $columns $OutFile
    Export-PlainCsv $log @('Event','Outcome','Entries') $logFile
    Write-Host "Completed $cat -> saved to $OutFile"
    $written.Add($OutFile)
    if ($script:abort) { Write-Warning "Stopped early because of rate limiting; remaining categories were skipped."; break }
}

Write-Host ""
if ($script:abort) { Write-Warning "Run stopped early. The CSVs contain only what was collected before the block." } else { Write-Host "All requested categories successfully completed!" }
Write-Host ""
Write-Host "Result files (upload these to the PWA repo):"
foreach ($f in $written) { Write-Host "  $f" }
Write-Host "(The *_EventLog.csv files are for debugging only; no need to upload them.)"
