<#
.SYNOPSIS
  Scrape HYROX TEAM RELAY MEN results for every Asian location on hyresult.com.

.DESCRIPTION
  1. Reads https://www.hyresult.com/locations?con=AS and discovers every /location/<slug> link
     (so the number of locations is never hard-coded).
  2. Opens each location page and collects every /event/<slug> link (past AND upcoming).
  3. Opens each event page and looks for the ranking link ending in "-team-relay-men".
     Events without that link (older seasons, upcoming events) are skipped and logged.
  4. Pages through that ranking (?p=1,2,...) until a page produces no NEW result IDs.
  5. De-duplicates on the /result/<ID> link, which is unique per team result.

  Season filter: events whose slug starts with s8- (25/26) or s9- (26/27) are visited, plus
  s7 events held in 2025 (e.g. s7-2025-bangkok), controlled by -AcceptedSeasonYears.
  All other seasons are skipped without loading their pages. Events in accepted seasons
  whose relay ranking has zero results (i.e. upcoming) are skipped too, and logged.

  Time text: Time and GapToWinner are written as ="h:mm:ss" so Excel keeps them as text instead of
  converting them to time values. Use -PlainTimes to write raw values instead.

  AG column: each team's age group is derived from the ranking's age-group filter. The script
  fetches the 40+ filtered ranking (?ag=40%2B) once per event; teams whose result ID is in that
  set are labelled 40+, all others U40. If the filter returns every team (i.e. the server ignored
  it), that event falls back to reading each team's result page. Use -SkipAG to leave the column
  blank, or -Only40Plus to write only the 40+ teams.

  CSV: values are not wrapped in quotes (only quoted when they contain a comma or quote). The
  ResultId / ResultUrl / RankingUrl columns are not exported.

  The default output path uses the real Desktop folder (works when the Desktop is redirected
  to OneDrive), and the output folder is created if it does not exist.

.EXAMPLE
  .\Get-HyroxAsiaRelayMen.ps1 -Only bangkok,jakarta,singapore   # test with a few locations first
  .\Get-HyroxAsiaRelayMen.ps1                  # all Asian locations, seasons 25/26 + 26/27 (+ s7-2025)
  .\Get-HyroxAsiaRelayMen.ps1 -Seasons 9       # only 26/27 (+ s7-2025)
  .\Get-HyroxAsiaRelayMen.ps1 -Seasons 9 -AcceptedSeasonYears @()   # strictly 26/27
#>
param(
    [string]$OutFile = (Join-Path ([Environment]::GetFolderPath('Desktop')) "Hyrox_Asia_TeamRelayMen_2526_2627.csv"),
    [string[]]$Only  = @(),         # location slug(s) for testing, e.g. -Only bangkok,jakarta,singapore
    [int[]]$Seasons  = @(8, 9),     # s8 = 25/26, s9 = 26/27
    [string[]]$AcceptedSeasonYears = @('7-2025'),   # extra season-year pairs accepted (s7 events held in 2025, e.g. s7-2025-bangkok)
    [int]$DelayMs    = 400,         # pause between requests, be polite to the server
    [switch]$PlainTimes,            # write Time/GapToWinner as plain values (Excel will then turn them into time values)
    [switch]$SkipAG,                # skip the age-group lookup entirely (AG column left blank)
    [switch]$Only40Plus,            # keep only teams in the 40+ age group in the output
    [switch]$AllowPerTeamAG         # allow the slow fallback of 1 request per team to read the age group
)

$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$root    = "https://www.hyresult.com"
$hostRx  = '(?:https://www\.hyresult\.com)?'
$logFile = [System.IO.Path]::ChangeExtension($OutFile, $null).TrimEnd('.') + "_EventLog.csv"

# Make sure the output folder exists (safety net for redirected / missing Desktop folders).
$outDir = Split-Path $OutFile -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

# ---------------------------------------------------------------- helpers
function Get-Html([string]$url) {
    for ($try = 1; $try -le 3; $try++) {
        try {
            $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30 `
                 -Headers @{ "User-Agent" = "Mozilla/5.0 (personal research script)" }
            Start-Sleep -Milliseconds $DelayMs
            return [string]$r.Content
        } catch {
            Write-Warning "  attempt $try failed for $url : $($_.Exception.Message)"
            Start-Sleep -Seconds (2 * $try)
        }
    }
    return $null
}

function Strip-Tags([string]$s, [string]$replacement = " ") {
    return [System.Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', $replacement)).Trim()
}

# Age group of ONE team from its /result/<ID> page ("#1 in AG U40" / "40+"). Slow (1 request per team);
# only used as a fallback when the ?ag= filter turns out not to be applied by the server.
function Get-AgeGroup([string]$resultId) {
    $h = Get-Html "$root/result/$resultId"
    if (-not $h) { return "" }
    $txt = Strip-Tags $h " "
    $m = [regex]::Match($txt, '(?i)\bin\s+AG\s+(U\s?\d{2}|\d{2}\s?\+)')
    if (-not $m.Success) { $m = [regex]::Match($txt, '(?i)\bAG\s*:?\s+(U\s?\d{2}|\d{2}\s?\+)') }
    if ($m.Success) { return ($m.Groups[1].Value -replace '\s', '').ToUpper() }
    return ""
}

# Collect the result IDs of one age-group filter of a ranking, e.g. '40+' => ?ag=40%2B
function Get-AgeGroupIds([string]$rankPath, [string]$ag, [int]$pageSize = 100) {
    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    $agQ = [uri]::EscapeDataString($ag)              # "40+" -> "40%2B"
    for ($page = 1; $page -le 10; $page++) {       # hard cap: a 40+ list never needs more than 10 pages
        $h = Get-Html "$root$rankPath`?ag=$agQ&p=$page"
        if (-not $h) { break }
        if (-not $script:agSampleSaved) {            # keep one filtered page so it can be inspected
            try { Set-Content -Path (Join-Path $env:TEMP "hyresult_ag_sample.html") -Value $h -Encoding UTF8 } catch {}
            $script:agSampleSaved = $true
        }
        # The filtered page's HTML can contain more rows than the filtered table shows (e.g. the normal list
        # is embedded as well), so the rows are split into separate lists in document order: a new list starts
        # when the finish time goes DOWN again (the ranking is sorted by time) or when a result ID repeats.
        # The filtered table is always the smallest of those lists, so that one is used.
        $segments = New-Object 'System.Collections.Generic.List[object]'
        $cur = $null; $lastT = -1
        foreach ($tr in [regex]::Matches($h, '<tr\b[^>]*>(.*?)</tr>', 'Singleline')) {
            $rowHtml = $tr.Groups[1].Value
            $m = [regex]::Match($rowHtml, 'href="' + $hostRx + '/result/([A-Za-z0-9]+)"')
            if (-not $m.Success) { continue }
            $id = $m.Groups[1].Value
            $t = $null
            $tm = [regex]::Match((Strip-Tags $rowHtml " "), '(\d{1,2}(?::\d{2}){1,2})')
            if ($tm.Success) {
                $secs = 0
                foreach ($part in ($tm.Groups[1].Value -split ':')) { $secs = $secs * 60 + [int]$part }
                $t = $secs
            }
            $startNew = ($null -eq $cur) -or $cur.Contains($id) -or (($null -ne $t) -and ($lastT -ge 0) -and ($t -lt $lastT))
            if ($startNew) {
                $cur = New-Object 'System.Collections.Generic.HashSet[string]'
                $segments.Add($cur)
                $lastT = -1
            }
            [void]$cur.Add($id)
            if ($null -ne $t) { $lastT = $t }
        }
        $pageIds = New-Object 'System.Collections.Generic.HashSet[string]'
        if ($segments.Count -gt 0) {
            $best = $segments | Sort-Object { $_.Count } | Select-Object -First 1
            foreach ($id in $best) { [void]$pageIds.Add($id) }
        }
        Write-Host ("      40+ filter page ${page}: row lists found [{0}] -> using the smallest ({1})" -f (($segments | ForEach-Object { $_.Count }) -join ', '), $pageIds.Count)
        $before = $ids.Count
        foreach ($id in $pageIds) { [void]$ids.Add($id) }
        if ($pageIds.Count -lt $pageSize) { break }  # short page -> last page; do NOT request the next one
        if ($ids.Count -eq $before) { break }        # nothing new (site repeating a page) -> stop
    }
    return ,$ids
}

# Write objects to CSV without wrapping every value in quotes. A value is only quoted when it contains
# a comma, a quote or a line break; Excel text formulas such as ="1:05:17" are written as they are.
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

# ---------------------------------------------------------------- 1. discover locations
Write-Host "Step 1: discovering Asian locations..."
$locHtml = Get-Html "$root/locations?con=AS"
if (-not $locHtml) { throw "Could not load the locations page." }

$locations = [ordered]@{}   # slug -> country
$expected  = @{}            # slug -> number of events the card claims

# Keep a copy of the page so it can be inspected if nothing is found.
$dump = Join-Path $env:TEMP "hyresult_locations.html"
try { Set-Content -Path $dump -Value $locHtml -Encoding UTF8 } catch {}

# Real links only: href="/location/<slug>". The page also embeds data for ALL continents, so loose
# matching of "/location/" would pull in the whole world and make the run very slow.
$hits = [regex]::Matches($locHtml, 'href="' + $hostRx + '/location/([A-Za-z0-9_-]+)"')
$usedFallback = $false
if ($hits.Count -eq 0) {
    $hits = [regex]::Matches($locHtml, '\\?/location\\?/([A-Za-z0-9_-]+)')
    $usedFallback = $true
}
Write-Host "  Page length: $($locHtml.Length) chars; /location/ references found: $($hits.Count)"
Write-Host "  (page saved to $dump for inspection)"

$info = [ordered]@{}        # slug -> @{ Country; Continent; Expected }
foreach ($h in $hits) {
    $slug = $h.Groups[1].Value
    if (-not $info.Contains($slug)) { $info[$slug] = @{ Country = ""; Continent = ""; Expected = $null } }
    $rec = $info[$slug]

    # Text that follows this link, up to the next /location/ link: "Country, Continent" and "N events".
    $start = $h.Index + $h.Length
    $next  = $locHtml.IndexOf('/location/', $start)
    $len   = if ($next -gt $start) { $next - $start } else { $locHtml.Length - $start }
    if ($len -gt 1500) { $len = 1500 }
    $raw = $locHtml.Substring($start, $len)
    $gt  = $raw.IndexOf('>')
    if ($gt -ge 0 -and $gt -lt 400) { $raw = $raw.Substring($gt + 1) }      # drop the rest of the <a ...> tag
    $win = Strip-Tags $raw "|"

    if (-not $rec.Country) {
        $gm = [regex]::Match($win, '([^|,]+?)\s*,\s*(Asia|Europe|Africa|North America|South America|Oceania|Antarctica)')
        if ($gm.Success) { $rec.Country = $gm.Groups[1].Value.Trim(); $rec.Continent = $gm.Groups[2].Value }
    }
    if ($null -eq $rec.Expected) {
        $cm = [regex]::Match($win, '(\d+)\s+events?')
        if ($cm.Success) { $rec.Expected = [int]$cm.Groups[1].Value }
    }
}

foreach ($slug in $info.Keys) {
    $rec = $info[$slug]
    # The URL already asks for Asia; only drop a card if the page explicitly says another continent.
    if ($rec.Continent -and $rec.Continent -ne "Asia") { continue }
    $locations[$slug] = $rec.Country
    if ($null -ne $rec.Expected) { $expected[$slug] = $rec.Expected }
}

# Safety net: the Asian locations listed at https://www.hyresult.com/locations?con=AS
# If the page handed us far more than that (i.e. the whole world), keep only these.
$asiaNames = @('Greater Noida','Mumbai','Osaka','Yokohama','Sanya','Seoul','Sharjah','Singapore','Wuhan',
               'Doha','Nagoya','Dubai','Guangzhou','Shanghai','Jakarta','Shenzhen','Taipei','Gandhinagar',
               'Gujarat','Chengdu','Hangzhou','Hong Kong','Bengaluru','New Delhi','Chiba','Incheon','Bangkok',
               'Kuala Lumpur','Abu Dhabi','Beijing')
$asiaSlugs = $asiaNames | ForEach-Object { ($_.ToLower() -replace '[^a-z0-9]+', '-').Trim('-') }
if ($locations.Count -gt 45 -or $usedFallback) {
    Write-Warning "  Page returned $($locations.Count) locations (looks worldwide) - restricting to the known Asian list."
    $filtered = [ordered]@{}
    foreach ($slug in @($locations.Keys)) {
        foreach ($a in $asiaSlugs) {
            if ($slug -eq $a -or $slug.StartsWith($a + "-")) { $filtered[$slug] = $locations[$slug]; break }
        }
    }
    $locations = $filtered
}

if ($Only -and $Only.Count -gt 0) {
    $keep = [ordered]@{}
    foreach ($o in $Only) {
        $hit = @($locations.Keys | Where-Object { $_ -ieq $o.Trim() })
        if ($hit.Count -eq 0) {
            Write-Warning "  '$o' is not a known location slug. Available: $($locations.Keys -join ', ')"
        } else {
            $keep[$hit[0]] = $locations[$hit[0]]
        }
    }
    $locations = $keep
}
Write-Host "  Found $($locations.Count) location(s): $($locations.Keys -join ', ')"
if ($locations.Count -eq 0) { throw "No locations found. Open $dump in a browser/notepad: if it has no /location/ links the page is built by JavaScript or blocked, and a different approach is needed." }

# ---------------------------------------------------------------- 2. discover events
Write-Host "Step 2: collecting event links from each location..."
$seenEvents = New-Object 'System.Collections.Generic.HashSet[string]'
$events     = New-Object 'System.Collections.Generic.List[object]'

foreach ($slug in @($locations.Keys)) {
    $html = Get-Html "$root/location/$slug"
    if (-not $html) { Write-Warning "  could not load location $slug"; continue }

    $found = 0
    foreach ($m in [regex]::Matches($html, 'href="' + $hostRx + '/event/([^"#?/]+)"')) {
        $ev = $m.Groups[1].Value
        if ($seenEvents.Add($ev)) {               # Add() returns $false if already present
            $events.Add([pscustomobject]@{ LocationSlug = $slug; Country = $locations[$slug]; EventSlug = $ev })
            $found++
        }
    }
    $exp = if ($expected.ContainsKey($slug)) { $expected[$slug] } else { "?" }
    Write-Host ("  {0,-16} {1} event link(s) found (locations page says {2})" -f $slug, $found, $exp)
    if ($expected.ContainsKey($slug) -and $found -lt $expected[$slug]) {
        Write-Warning "    fewer event links than expected for $slug - an older event may be hidden or loaded dynamically."
    }
}
Write-Host "  $($events.Count) unique event(s) in total."

# ---------------------------------------------------------------- 3. scrape relay-men rankings
Write-Host "Step 3: scraping TEAM RELAY MEN rankings..."
$seenResults = New-Object 'System.Collections.Generic.HashSet[string]'
$rows        = New-Object 'System.Collections.Generic.List[object]'
$log         = New-Object 'System.Collections.Generic.List[object]'
$textInfo    = (Get-Culture).TextInfo

foreach ($e in $events) {
    # Season filter first, so unwanted events cost no requests at all.
    # Accepted: seasons listed in -Seasons, OR a season-year pair listed in -AcceptedSeasonYears
    # (default '7-2025' => s7-2025-<city>, e.g. s7-2025-bangkok, for every location).
    $sm = [regex]::Match($e.EventSlug, '^s(\d+)-(\d{4})-')
    $seasonOk = $false
    if ($sm.Success) {
        $pair = "{0}-{1}" -f [int]$sm.Groups[1].Value, $sm.Groups[2].Value      # e.g. "7-2025"
        $seasonOk = ($Seasons -contains [int]$sm.Groups[1].Value) -or ($AcceptedSeasonYears -contains $pair)
    }
    if (-not $seasonOk) {
        Write-Host "  skip  $($e.EventSlug) (outside selected seasons)"
        $log.Add([pscustomobject]@{ Event = $e.EventSlug; Country = $e.Country; Outcome = "outside selected seasons"; Teams = 0 })
        continue
    }

    $evHtml = Get-Html "$root/event/$($e.EventSlug)"
    if (-not $evHtml) {
        $log.Add([pscustomobject]@{ Event = $e.EventSlug; Country = $e.Country; Outcome = "event page failed to load"; Teams = 0 })
        continue
    }

    if (-not $e.Country) {
        $h1c = $evHtml.IndexOf("<h1")
        if ($h1c -ge 0) {
            $cm = [regex]::Match((Strip-Tags $evHtml.Substring($h1c) "|"), '([^|,\d]+?)\s*,\s*Asia')
            if ($cm.Success) { $e.Country = $cm.Groups[1].Value.Trim() }
        }
    }

    # "-team-relay-men" at the END of the path. This deliberately does NOT match
    # "-team-relay-women" (which also ends in "men") or "-team-relay-mixed".
    $rm = [regex]::Match($evHtml, '(/ranking/[A-Za-z0-9_-]*-team-relay-men)(?![A-Za-z0-9_-])')
    if (-not $rm.Success) {
        Write-Host "  skip  $($e.EventSlug) (no relay-men ranking)"
        $log.Add([pscustomobject]@{ Event = $e.EventSlug; Country = $e.Country; Outcome = "no relay-men ranking link"; Teams = 0 })
        continue
    }
    $rankPath = $rm.Groups[1].Value

    $sn     = [int]$sm.Groups[1].Value
    $season = "S$sn ({0}/{1})" -f (17 + $sn), (18 + $sn)     # S7 -> 24/25, S8 -> 25/26, S9 -> 26/27
    $year   = $sm.Groups[2].Value
    $locName = $textInfo.ToTitleCase(($e.LocationSlug -replace '-', ' '))

    # Event date: the first "Mon D" / "Mon D, YYYY" / "Mon D-D, YYYY" text after the page's <h1>.
    $eventDate = ""
    $h1 = $evHtml.IndexOf("<h1")
    if ($h1 -ge 0) {
        $evText = Strip-Tags $evHtml.Substring($h1) " "
        $dm = [regex]::Match($evText, '((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:\s*[\u2013-]\s*\d{1,2})?(?:,\s*\d{4})?)')
        if ($dm.Success) { $eventDate = $dm.Groups[1].Value }
    }

    $rank = 0
    $teamsThisEvent = 0
    $firstRow = $rows.Count          # index of this event's first row (used for the AG step below)

    for ($page = 1; $page -le 50; $page++) {
        $html = Get-Html "$root$rankPath`?p=$page"
        if (-not $html) { break }
        Write-Host "    $($e.EventSlug) page $page ..." -NoNewline
        $newOnPage = 0

        foreach ($tr in [regex]::Matches($html, '<tr\b[^>]*>(.*?)</tr>', 'Singleline')) {
            $rowHtml = $tr.Groups[1].Value

            # The result link is the unique key of a team result. Rows without it
            # (header, skeleton placeholders) are ignored.
            $idm = [regex]::Match($rowHtml, 'href="' + $hostRx + '/result/([A-Za-z0-9]+)"')
            if (-not $idm.Success) { continue }
            $rid = $idm.Groups[1].Value
            if (-not $seenResults.Add($rid)) { continue }     # already stored -> skip (page repeats its table)
            $newOnPage++

            # Athletes: each is <img flag> + <a href=/athlete/slug>Name</a>, and the page lists them twice.
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
                    else {
                        $cm = [regex]::Match($img, '/m/([A-Z]{2})\.svg')
                        if ($cm.Success) { $country = $cm.Groups[1].Value }
                    }
                }
                $names.Add($name)
                $ctrys.Add($country)
            }

            # Time and gap to winner, e.g. "1:08:38" and "+3:21"
            $text  = Strip-Tags $rowHtml " "
            $timeM = [regex]::Match($text, '(\d{1,2}(?::\d{2}){1,2})')
            $diffM = [regex]::Match($text, '\+\s*(\d{1,2}(?::\d{2}){1,2})')
            $time  = if ($timeM.Success) { $timeM.Groups[1].Value } else { "" }
            $diff  = if ($diffM.Success) { $diffM.Groups[1].Value } else { "" }

            # Rank = position in the (time-sorted) table; blank if the team has no time (DNF/DQ).
            $overall = ""
            if ($time) { $rank++; $overall = $rank }

            $teamCountries = ($ctrys | Where-Object { $_ } | Select-Object -Unique) -join "; "

            $ag = ""     # filled in after the event's pages are scraped (see the AG step below)

            # Force Excel to treat times as text: ="1:05:17" displays as 1:05:17 but stays text.
            $timeOut = $time; $diffOut = $diff
            if (-not $PlainTimes) {
                if ($time) { $timeOut = '="' + $time + '"' }
                if ($diff) { $diffOut = '="' + $diff + '"' }
            }

            $rows.Add([pscustomobject]@{
                Country       = $e.Country
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
                ResultUrl     = "$root/result/$rid"
                RankingUrl    = "$root$rankPath"
            })
            $teamsThisEvent++
        }

        Write-Host " $newOnPage new team(s)"
        if ($newOnPage -eq 0) { break }     # nothing new -> last page reached
    }

    # Age group. Fetch the 40+ filtered ranking (?ag=40%2B) and keep only IDs that belong to THIS event's
    # unfiltered rows. If the filter returns every team, the server did not apply it, so fall back to the
    # per-team result page for this event instead of labelling everybody 40+.
    if (-not $SkipAG -and $teamsThisEvent -gt 0) {
        $eventIds = @{}
        for ($i = $firstRow; $i -lt $rows.Count; $i++) { $eventIds[$rows[$i].ResultId] = $true }
        $ids40 = Get-AgeGroupIds $rankPath '40+'
        $match40 = @($ids40 | Where-Object { $eventIds.ContainsKey($_) }).Count
        Write-Host "    $($e.EventSlug): 40+ filter returned $($ids40.Count) id(s), $match40 match this event's $teamsThisEvent team(s)"

        if ($match40 -ge $teamsThisEvent -and $teamsThisEvent -gt 1) {
            if ($AllowPerTeamAG) {
                Write-Warning "    40+ filter was NOT applied for $($e.EventSlug) - reading each team's result page instead (slower)."
                for ($i = $firstRow; $i -lt $rows.Count; $i++) {
                    $rows[$i].AG = Get-AgeGroup $rows[$i].ResultId
                    if ((($i - $firstRow) % 10) -eq 9) { Write-Host "      AG lookup $($i - $firstRow + 1)/$teamsThisEvent" }
                }
            } else {
                Write-Warning "    40+ filter was NOT applied for $($e.EventSlug) - AG left blank (use -AllowPerTeamAG for the slow per-team lookup)."
            }
        } else {
            for ($i = $firstRow; $i -lt $rows.Count; $i++) {
                $rows[$i].AG = if ($ids40.Contains($rows[$i].ResultId)) { "40+" } else { "U40" }
            }
        }
    }

    $outcome = if ($teamsThisEvent -gt 0) { "scraped" } else { "ranking exists but has no results (upcoming?)" }
    Write-Host ("  {0,-8} {1,-28} {2} team(s)" -f $outcome.Substring(0, [Math]::Min(8, $outcome.Length)), $e.EventSlug, $teamsThisEvent)
    $log.Add([pscustomobject]@{ Event = $e.EventSlug; Country = $e.Country; Outcome = $outcome; Teams = $teamsThisEvent })
}

# ---------------------------------------------------------------- 4. export
$rows = $rows | Sort-Object Country, Location, Year, EventSlug, { if ($_.OverallRank -eq "") { [int]::MaxValue } else { [int]$_.OverallRank } }

# Overlap check: every ResultId must appear exactly once.
$dupes = $rows | Group-Object ResultId | Where-Object { $_.Count -gt 1 }
if ($dupes) { Write-Warning "$(@($dupes).Count) duplicate ResultId(s) found - please report this." }
else        { Write-Host "Overlap check passed: all $(@($rows).Count) ResultIds are unique." }

if ($Only40Plus) { $rows = @($rows | Where-Object { $_.AG -eq '40+' }) }

$columns = 'Country','Location','Season','Year','EventDate','EventSlug','OverallRank','AG','Time','GapToWinner',
           'Athlete1','Athlete2','Athlete3','Athlete4','TeamCountries'
Export-PlainCsv $rows $columns $OutFile
Export-PlainCsv $log ('Event','Country','Outcome','Teams') $logFile

Write-Host ""
Write-Host "Teams per event:"
$rows | Group-Object Country, Location, EventSlug |
    ForEach-Object { "  {0,-45} {1}" -f $_.Name, $_.Count } | Write-Host

Write-Host ""
Write-Host "Done. $(@($rows).Count) team result(s) from $(@($rows | Select-Object -ExpandProperty EventSlug -Unique).Count) event(s)."
Write-Host "Results : $OutFile"
Write-Host "Event log: $logFile"