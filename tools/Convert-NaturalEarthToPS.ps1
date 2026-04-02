<#
.SYNOPSIS
    Downloads Natural Earth vector data and converts it to simplified continent polygon JSON
    for the PSExchangeClock dashboard map views.

.DESCRIPTION
    This script:
    1. Downloads Natural Earth 110m land polygon GeoJSON from GitHub
    2. Classifies each polygon into named regions (NorthAmerica, Europe, Asia, etc.)
    3. Applies Douglas-Peucker simplification to reduce point count
    4. Outputs Data/continent-coords.json in [lat, lon] format
    5. Optionally downloads ne_50m_admin_0_boundary_lines_land for political borders

    DATA SOURCES & ATTRIBUTION:
    ──────────────────────────────────────────────────────────────────────
    Natural Earth — https://www.naturalearthdata.com/
    License: Public Domain
    Maintained by: Nathaniel Vaughn Kelso & Tom Patterson
    Used by: National Geographic, NASA, Washington Post, and many others
    GitHub: https://github.com/nvkelso/natural-earth-vector

    Thank you to the Natural Earth team for providing the gold standard
    in free, open cartographic data.
    ──────────────────────────────────────────────────────────────────────

.PARAMETER Resolution
    Natural Earth resolution: '110m' (default, ~10K points) or '50m' (~100K points).
    110m is recommended — after classification + simplification it produces
    1,500-2,500 points which is 4-7x more detailed than the original hand-crafted data.

.PARAMETER Tolerance
    Douglas-Peucker simplification tolerance in degrees. Default: 0.5
    Lower = more detail, more points. Higher = less detail, fewer points.
    Recommended range: 0.2 (high detail, ~3,000 pts) to 1.0 (low detail, ~800 pts).

.PARAMETER OutputPath
    Path for the output JSON file. Default: ../PSExchangeClock/Data/continent-coords.json

.PARAMETER IncludePolitical
    Also download and process political boundary lines for the overlay system.

.EXAMPLE
    .\Convert-NaturalEarthToPS.ps1
    # Downloads 110m data, simplifies with tolerance 0.5, outputs continent-coords.json

.EXAMPLE
    .\Convert-NaturalEarthToPS.ps1 -Resolution '50m' -Tolerance 0.3
    # Higher resolution source, more detailed output
#>
[CmdletBinding()]
param(
    [ValidateSet('110m', '50m')]
    [string]$Resolution = '110m',

    [ValidateRange(0.05, 5.0)]
    [double]$Tolerance = 0.5,

    [string]$OutputPath = (Join-Path $PSScriptRoot '..\PSExchangeClock\Data\continent-coords.json'),

    [switch]$IncludePolitical
)

$ErrorActionPreference = 'Stop'

# ── Region Classification Rules ──────────────────────────────────────
# Each polygon is classified by its centroid. Order matters — earlier rules take priority.
# Named regions match the existing Get-ContinentData() structure.

$regionRules = @(
    @{ Name = 'Iceland'; MinLat = 62; MaxLat = 67; MinLon = -25; MaxLon = -12; MaxArea = 200 }
    @{ Name = 'UK'; MinLat = 49; MaxLat = 61; MinLon = -11; MaxLon = 3; MaxArea = 500 }
    @{ Name = 'Ireland'; MinLat = 51; MaxLat = 56; MinLon = -11; MaxLon = -5; MaxArea = 200 }
    @{ Name = 'Japan'; MinLat = 30; MaxLat = 46; MinLon = 128; MaxLon = 150; MaxArea = 600 }
    @{ Name = 'Taiwan'; MinLat = 21; MaxLat = 26; MinLon = 119; MaxLon = 123; MaxArea = 100 }
    @{ Name = 'SriLanka'; MinLat = 5; MaxLat = 10; MinLon = 79; MaxLon = 82; MaxArea = 100 }
    @{ Name = 'NewZealand'; MinLat = -48; MaxLat = -34; MinLon = 165; MaxLon = 179; MaxArea = 500 }
    @{ Name = 'Madagascar'; MinLat = -26; MaxLat = -12; MinLon = 43; MaxLon = 51; MaxArea = 800 }
    @{ Name = 'Cuba'; MinLat = 19; MaxLat = 24; MinLon = -85; MaxLon = -74; MaxArea = 300 }
    @{ Name = 'Hispaniola'; MinLat = 17; MaxLat = 21; MinLon = -75; MaxLon = -68; MaxArea = 200 }
    @{ Name = 'Greenland'; MinLat = 59; MaxLat = 84; MinLon = -75; MaxLon = -10; MaxArea = 5000 }
    @{ Name = 'India'; MinLat = 6; MaxLat = 36; MinLon = 67; MaxLon = 90; MaxArea = 6000 }
    @{ Name = 'ArabianPeninsula'; MinLat = 12; MaxLat = 32; MinLon = 34; MaxLon = 60; MaxArea = 5000 }
    @{ Name = 'Philippines'; MinLat = 4; MaxLat = 20; MinLon = 116; MaxLon = 128; MaxArea = 500 }
    @{ Name = 'Papua'; MinLat = -12; MaxLat = 0; MinLon = 130; MaxLon = 155; MaxArea = 1500 }
    @{ Name = 'Borneo'; MinLat = -5; MaxLat = 8; MinLon = 108; MaxLon = 120; MaxArea = 1200 }
    @{ Name = 'Sumatra'; MinLat = -6; MaxLat = 6; MinLon = 95; MaxLon = 107; MaxArea = 800 }
    @{ Name = 'Java'; MinLat = -9; MaxLat = -5; MinLon = 105; MaxLon = 115; MaxArea = 400 }
    @{ Name = 'Sulawesi'; MinLat = -6; MaxLat = 2; MinLon = 119; MaxLon = 126; MaxArea = 400 }
    @{ Name = 'Australia'; MinLat = -45; MaxLat = -10; MinLon = 112; MaxLon = 155; MaxArea = 99999 }
    @{ Name = 'NorthAmerica'; MinLat = 7; MaxLat = 84; MinLon = -170; MaxLon = -50; MaxArea = 99999 }
    @{ Name = 'CentralAmerica'; MinLat = 7; MaxLat = 19; MinLon = -92; MaxLon = -77; MaxArea = 1000 }
    @{ Name = 'SouthAmerica'; MinLat = -60; MaxLat = 15; MinLon = -82; MaxLon = -34; MaxArea = 99999 }
    @{ Name = 'Europe'; MinLat = 35; MaxLat = 72; MinLon = -12; MaxLon = 40; MaxArea = 99999 }
    @{ Name = 'Africa'; MinLat = -36; MaxLat = 38; MinLon = -18; MaxLon = 52; MaxArea = 99999 }
    @{ Name = 'Asia'; MinLat = -12; MaxLat = 78; MinLon = 25; MaxLon = 180; MaxArea = 99999 }
)

# ── Douglas-Peucker Simplification ──────────────────────────────────
function Simplify-Polyline {
    param(
        [double[][]]$Points,
        [double]$Epsilon
    )
    if ($Points.Count -le 2) { return $Points }

    # Find the point with the maximum distance from the line between first and last
    $first = $Points[0]
    $last = $Points[$Points.Count - 1]
    $maxDist = 0.0
    $maxIdx = 0

    for ($i = 1; $i -lt ($Points.Count - 1); $i++) {
        $d = Get-PerpendicularDistance -Point $Points[$i] -LineStart $first -LineEnd $last
        if ($d -gt $maxDist) {
            $maxDist = $d
            $maxIdx = $i
        }
    }

    if ($maxDist -gt $Epsilon) {
        $left = Simplify-Polyline -Points $Points[0..$maxIdx] -Epsilon $Epsilon
        $right = Simplify-Polyline -Points $Points[$maxIdx..($Points.Count - 1)] -Epsilon $Epsilon
        # Merge, removing duplicate junction point
        $result = [System.Collections.Generic.List[double[]]]::new()
        foreach ($p in $left) { $result.Add($p) }
        for ($i = 1; $i -lt $right.Count; $i++) { $result.Add($right[$i]) }
        return $result.ToArray()
    }
    else {
        return @($first, $last)
    }
}

function Get-PerpendicularDistance {
    param([double[]]$Point, [double[]]$LineStart, [double[]]$LineEnd)
    $dx = $LineEnd[0] - $LineStart[0]
    $dy = $LineEnd[1] - $LineStart[1]
    $lenSq = $dx * $dx + $dy * $dy
    if ($lenSq -eq 0) {
        $dx2 = $Point[0] - $LineStart[0]
        $dy2 = $Point[1] - $LineStart[1]
        return [Math]::Sqrt($dx2 * $dx2 + $dy2 * $dy2)
    }
    $num = [Math]::Abs($dy * $Point[0] - $dx * $Point[1] + $LineEnd[0] * $LineStart[1] - $LineEnd[1] * $LineStart[0])
    return $num / [Math]::Sqrt($lenSq)
}

# ── Polygon Area (for classification — approximate, in degree²) ─────
function Get-PolygonArea {
    param([double[][]]$Coords)
    $area = 0.0
    $n = $Coords.Count
    for ($i = 0; $i -lt $n; $i++) {
        $j = ($i + 1) % $n
        $area += $Coords[$i][1] * $Coords[$j][0]
        $area -= $Coords[$j][1] * $Coords[$i][0]
    }
    return [Math]::Abs($area) / 2.0
}

# ── Centroid Calculation ─────────────────────────────────────────────
function Get-PolygonCentroid {
    param([double[][]]$Coords)
    $latSum = 0.0; $lonSum = 0.0
    foreach ($c in $Coords) { $latSum += $c[0]; $lonSum += $c[1] }
    return @($latSum / $Coords.Count, $lonSum / $Coords.Count)
}

# ── Region Classification ───────────────────────────────────────────
function Get-RegionName {
    param([double[]]$Centroid, [double]$Area)
    foreach ($rule in $regionRules) {
        if ($Centroid[0] -ge $rule.MinLat -and $Centroid[0] -le $rule.MaxLat -and
            $Centroid[1] -ge $rule.MinLon -and $Centroid[1] -le $rule.MaxLon -and
            $Area -le $rule.MaxArea) {
            return $rule.Name
        }
    }
    return $null  # Unclassified (tiny islands, etc.)
}

# ── Extract coordinate rings from GeoJSON geometry ──────────────────
function Get-PolygonRings {
    param($Geometry)
    $rings = [System.Collections.Generic.List[object]]::new()
    if ($Geometry.type -eq 'Polygon') {
        # $Geometry.coordinates is array of rings; first ring is exterior
        $ring = $Geometry.coordinates[0]
        if ($ring.Count -ge 3) {
            # GeoJSON = [lon, lat]; convert to [lat, lon]
            $coords = foreach ($pt in $ring) { , @([double]$pt[1], [double]$pt[0]) }
            $rings.Add([double[][]]$coords)
        }
    }
    elseif ($Geometry.type -eq 'MultiPolygon') {
        foreach ($poly in $Geometry.coordinates) {
            $ring = $poly[0]  # exterior ring only
            if ($ring.Count -ge 3) {
                $coords = foreach ($pt in $ring) { , @([double]$pt[1], [double]$pt[0]) }
                $rings.Add([double[][]]$coords)
            }
        }
    }
    # Use comma operator to prevent PowerShell from unrolling a single-element list
    return , $rings
}

# ══════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Natural Earth → PSExchangeClock Continent Data Converter" -ForegroundColor Cyan
Write-Host "  Source: https://www.naturalearthdata.com/ (Public Domain)" -ForegroundColor DarkGray
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Step 1: Download NE data (admin_0_countries — individual country polygons with CONTINENT field)
$neUrl = "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/master/geojson/ne_${Resolution}_admin_0_countries.geojson"
$cacheDir = Join-Path $PSScriptRoot '.cache'
if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
$cachePath = Join-Path $cacheDir "ne_${Resolution}_admin_0_countries.geojson"

if (Test-Path $cachePath) {
    Write-Host "[CACHED] Using cached $Resolution land data: $cachePath" -ForegroundColor Green
}
else {
    Write-Host "[DOWNLOAD] Fetching Natural Earth $Resolution land data..." -ForegroundColor Yellow
    Write-Host "  URL: $neUrl" -ForegroundColor DarkGray
    try {
        Invoke-WebRequest -Uri $neUrl -OutFile $cachePath -UseBasicParsing -TimeoutSec 120
        Write-Host "[OK] Downloaded successfully ($([Math]::Round((Get-Item $cachePath).Length / 1KB)) KB)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download Natural Earth data: $_"
        Write-Host ""
        Write-Host "Manual download: Visit https://www.naturalearthdata.com/downloads/${Resolution}-physical-vectors/" -ForegroundColor Yellow
        exit 1
    }
}

# Step 2: Parse GeoJSON
Write-Host "[PARSE] Reading GeoJSON..." -ForegroundColor Yellow
$rawJson = Get-Content $cachePath -Raw
$geoJson = $rawJson | ConvertFrom-Json
Write-Host "[OK] Found $($geoJson.features.Count) features" -ForegroundColor Green

# Step 3: Extract all polygon rings and classify using NE CONTINENT/SUBREGION properties
Write-Host "[CLASSIFY] Extracting country polygons and classifying by continent..." -ForegroundColor Yellow
$regionPolygons = @{}
$tinySkipped = 0
$processedCountries = 0

# Map NE continent names to our region names
# Special countries get their own region for visual clarity
$specialCountryMap = @{
    'United Kingdom'       = 'UK'
    'Japan'                = 'Japan'
    'Iceland'              = 'Iceland'
    'Ireland'              = 'Ireland'
    'Greenland'            = 'Greenland'
    'New Zealand'          = 'NewZealand'
    'Australia'            = 'Australia'
    'Madagascar'           = 'Madagascar'
    'Cuba'                 = 'Cuba'
    'Taiwan'               = 'Taiwan'
    'Sri Lanka'            = 'SriLanka'
    'Philippines'          = 'Philippines'
    'Papua New Guinea'     = 'Papua'
    'Indonesia'            = 'Indonesia'
    'India'                = 'India'
    'Saudi Arabia'         = 'ArabianPeninsula'
    'Yemen'                = 'ArabianPeninsula'
    'Oman'                 = 'ArabianPeninsula'
    'United Arab Emirates' = 'ArabianPeninsula'
    'Qatar'                = 'ArabianPeninsula'
    'Kuwait'               = 'ArabianPeninsula'
    'Bahrain'              = 'ArabianPeninsula'
}

$continentMap = @{
    'North America' = 'NorthAmerica'
    'South America' = 'SouthAmerica'
    'Europe'        = 'Europe'
    'Africa'        = 'Africa'
    'Asia'          = 'Asia'
    'Oceania'       = 'Oceania'
    'Antarctica'    = $null  # Skip Antarctica
}

foreach ($feature in $geoJson.features) {
    $props = $feature.properties
    $countryName = if ($props.NAME) { $props.NAME } elseif ($props.name) { $props.name } else { '' }
    $continent = if ($props.CONTINENT) { $props.CONTINENT } elseif ($props.continent) { $props.continent } else { '' }

    # Determine region name
    $regionName = $null
    if ($specialCountryMap.ContainsKey($countryName)) {
        $regionName = $specialCountryMap[$countryName]
    }
    elseif ($continentMap.ContainsKey($continent)) {
        $regionName = $continentMap[$continent]
    }

    if (-not $regionName) {
        Write-Verbose "Skipping: $countryName (continent: $continent)"
        continue
    }

    $rings = Get-PolygonRings -Geometry $feature.geometry
    foreach ($ring in $rings) {
        if ($ring.Count -lt 3) { continue }

        $area = Get-PolygonArea -Coords $ring
        # Skip very tiny island polygons (< 0.3 degree² ≈ < ~3,700 km² at equator)
        if ($area -lt 0.3) { $tinySkipped++; continue }

        if (-not $regionPolygons.ContainsKey($regionName)) {
            $regionPolygons[$regionName] = [System.Collections.Generic.List[object]]::new()
        }
        $regionPolygons[$regionName].Add($ring)
        $processedCountries++
    }
}

Write-Host "[OK] Classified $processedCountries polygons into $($regionPolygons.Count) regions, skipped $tinySkipped tiny fragments" -ForegroundColor Green

# Step 4: Simplify each region
Write-Host "[SIMPLIFY] Applying Douglas-Peucker (tolerance=$Tolerance°)..." -ForegroundColor Yellow
$totalOriginal = 0
$totalSimplified = 0
$output = [System.Collections.Generic.List[object]]::new()

foreach ($regionName in ($regionPolygons.Keys | Sort-Object)) {
    $polygons = $regionPolygons[$regionName]
    foreach ($polygon in $polygons) {
        $originalCount = $polygon.Count
        $totalOriginal += $originalCount

        $simplified = Simplify-Polyline -Points $polygon -Epsilon $Tolerance
        $totalSimplified += $simplified.Count

        # Convert to simple arrays for JSON output
        $coordArrays = foreach ($pt in $simplified) {
            , @([Math]::Round($pt[0], 2), [Math]::Round($pt[1], 2))
        }

        $output.Add(@{
                Name   = $regionName
                Coords = $coordArrays
                Points = $simplified.Count
            })

        Write-Verbose "  $regionName`: $originalCount -> $($simplified.Count) points"
    }
}

Write-Host "[OK] Simplified: $totalOriginal -> $totalSimplified points ($([Math]::Round(($totalSimplified / [Math]::Max($totalOriginal,1)) * 100, 1))% retained)" -ForegroundColor Green

# Step 5: Write output JSON
Write-Host "[WRITE] Writing $OutputPath..." -ForegroundColor Yellow
$outputDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

$jsonOutput = @{
    meta    = @{
        source        = 'Natural Earth'
        sourceUrl     = 'https://www.naturalearthdata.com/'
        license       = 'Public Domain'
        resolution    = $Resolution
        tolerance     = $Tolerance
        dataVersion   = '5.1.2'
        generatedDate = (Get-Date -Format 'yyyy-MM-dd')
        generatedBy   = 'Convert-NaturalEarthToPS.ps1'
        totalPoints   = $totalSimplified
        totalRegions  = $regionPolygons.Count
        totalPolygons = $output.Count
    }
    regions = $output | ForEach-Object {
        @{ Name = $_.Name; Coords = $_.Coords }
    }
}

$jsonOutput | ConvertTo-Json -Depth 10 -Compress:$false | Set-Content $OutputPath -Encoding UTF8
$fileSize = [Math]::Round((Get-Item $OutputPath).Length / 1KB, 1)
Write-Host "[OK] Written: $OutputPath ($fileSize KB)" -ForegroundColor Green

# Step 6: Summary
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Source:      Natural Earth $Resolution (Public Domain)" -ForegroundColor White
Write-Host "  Regions:     $($regionPolygons.Count)" -ForegroundColor White
Write-Host "  Polygons:    $($output.Count)" -ForegroundColor White
Write-Host "  Points:      $totalOriginal -> $totalSimplified (tolerance=$Tolerance°)" -ForegroundColor White
Write-Host "  Output:      $OutputPath ($fileSize KB)" -ForegroundColor White
Write-Host ""

# Per-region breakdown
Write-Host "  Per-region breakdown:" -ForegroundColor DarkGray
$regionSummary = $output | Group-Object { $_.Name } | ForEach-Object {
    $pts = ($_.Group | Measure-Object -Property Points -Sum).Sum
    [PSCustomObject]@{ Region = $_.Name; Polygons = $_.Count; Points = $pts }
} | Sort-Object Points -Descending

foreach ($r in $regionSummary) {
    Write-Host ("    {0,-20} {1,3} polygon(s)  {2,5} points" -f $r.Region, $r.Polygons, $r.Points) -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  Done! The dashboard will automatically load this data." -ForegroundColor Green

# Optional: Political boundaries
if ($IncludePolitical) {
    Write-Host ""
    Write-Host "[POLITICAL] Downloading boundary lines..." -ForegroundColor Yellow
    $polUrl = "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/master/geojson/ne_${Resolution}_admin_0_boundary_lines_land.geojson"
    $polCache = Join-Path $cacheDir "ne_${Resolution}_admin_0_boundary_lines_land.geojson"
    $polOutput = Join-Path (Split-Path $OutputPath -Parent) 'overlays' 'political-boundaries.json'

    if (-not (Test-Path $polCache)) {
        try {
            Invoke-WebRequest -Uri $polUrl -OutFile $polCache -UseBasicParsing -TimeoutSec 120
            Write-Host "[OK] Downloaded political boundaries" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to download political boundaries: $_"
            return
        }
    }

    $polJson = Get-Content $polCache -Raw | ConvertFrom-Json
    $boundaries = [System.Collections.Generic.List[object]]::new()
    $polPointCount = 0

    foreach ($feature in $polJson.features) {
        $geom = $feature.geometry
        $lines = @()
        if ($geom.type -eq 'LineString') {
            $lines = @(, $geom.coordinates)
        }
        elseif ($geom.type -eq 'MultiLineString') {
            $lines = $geom.coordinates
        }
        foreach ($line in $lines) {
            if ($line.Count -lt 2) { continue }
            $coords = foreach ($pt in $line) { , @([Math]::Round([double]$pt[1], 2), [Math]::Round([double]$pt[0], 2)) }
            $simplified = Simplify-Polyline -Points ([double[][]]$coords) -Epsilon ($Tolerance * 0.5)
            $polPointCount += $simplified.Count
            $simplifiedCoords = @(foreach ($pt in $simplified) { , @([Math]::Round($pt[0], 2), [Math]::Round($pt[1], 2)) })
            $boundaries.Add(@{
                    Coords = $simplifiedCoords
                })
        }
    }

    $polOutputDir = Split-Path $polOutput -Parent
    if (-not (Test-Path $polOutputDir)) { New-Item -ItemType Directory -Path $polOutputDir -Force | Out-Null }

    @{
        meta       = @{
            source        = 'Natural Earth'
            sourceUrl     = 'https://www.naturalearthdata.com/'
            license       = 'Public Domain'
            resolution    = $Resolution
            dataVersion   = '5.1.2'
            generatedDate = (Get-Date -Format 'yyyy-MM-dd')
            totalLines    = $boundaries.Count
            totalPoints   = $polPointCount
        }
        boundaries = $boundaries
    } | ConvertTo-Json -Depth 10 -Compress:$false | Set-Content $polOutput -Encoding UTF8

    Write-Host "[OK] Political boundaries: $($boundaries.Count) lines, $polPointCount points -> $polOutput" -ForegroundColor Green
}
