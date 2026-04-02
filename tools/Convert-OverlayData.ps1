<#
.SYNOPSIS
    Downloads and converts overlay data sources to simplified JSON for the PSExchangeClock map.

.DESCRIPTION
    Extends the data preparation pipeline for additional map overlays.
    Supports: SubmarineCables, EEZ, TZBoundaries, PowerPlants, ConflictZones

    DATA SOURCES & ATTRIBUTION:
    ──────────────────────────────────────────────────────────────────────
    Submarine Cables — TeleGeography / tbotnz GeoJSON fork
    Source: https://github.com/tbotnz/submarine-cables-geojson
    License: Open (non-commercial)
    ──────────────────────────────────────────────────────────────────────

.PARAMETER Source
    Which overlay data to process.

.PARAMETER Tolerance
    Douglas-Peucker simplification tolerance in degrees. Default: 0.3

.EXAMPLE
    .\Convert-OverlayData.ps1 -Source SubmarineCables
    .\Convert-OverlayData.ps1 -Source SubmarineCables -Tolerance 0.2
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('SubmarineCables', 'TZBoundaries', 'EEZ', 'PowerPlants', 'ConflictZones')]
    [string]$Source,

    [ValidateRange(0.05, 5.0)]
    [double]$Tolerance = 0.3
)

$ErrorActionPreference = 'Stop'
$cacheDir = Join-Path $PSScriptRoot '.cache'
$outputDir = Join-Path $PSScriptRoot '..\PSExchangeClock\Data\overlays'
if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

# ── Douglas-Peucker Simplification ──────────────────────────────────
function Simplify-Polyline {
    param([double[][]]$Points, [double]$Epsilon)
    if ($Points.Count -le 2) { return $Points }
    $first = $Points[0]; $last = $Points[$Points.Count - 1]
    $maxDist = 0.0; $maxIdx = 0
    for ($i = 1; $i -lt ($Points.Count - 1); $i++) {
        $d = Get-PerpendicularDistance -Point $Points[$i] -LineStart $first -LineEnd $last
        if ($d -gt $maxDist) { $maxDist = $d; $maxIdx = $i }
    }
    if ($maxDist -gt $Epsilon) {
        $left = Simplify-Polyline -Points $Points[0..$maxIdx] -Epsilon $Epsilon
        $right = Simplify-Polyline -Points $Points[$maxIdx..($Points.Count - 1)] -Epsilon $Epsilon
        $result = [System.Collections.Generic.List[double[]]]::new()
        foreach ($p in $left) { $result.Add($p) }
        for ($i = 1; $i -lt $right.Count; $i++) { $result.Add($right[$i]) }
        return $result.ToArray()
    }
    else { return @($first, $last) }
}

function Get-PerpendicularDistance {
    param([double[]]$Point, [double[]]$LineStart, [double[]]$LineEnd)
    $dx = $LineEnd[0] - $LineStart[0]; $dy = $LineEnd[1] - $LineStart[1]
    $lenSq = $dx * $dx + $dy * $dy
    if ($lenSq -eq 0) {
        $dx2 = $Point[0] - $LineStart[0]; $dy2 = $Point[1] - $LineStart[1]
        return [Math]::Sqrt($dx2 * $dx2 + $dy2 * $dy2)
    }
    $num = [Math]::Abs($dy * $Point[0] - $dx * $Point[1] + $LineEnd[0] * $LineStart[0] - $LineEnd[1] * $LineStart[0])
    return $num / [Math]::Sqrt($lenSq)
}

# ── Compiled C# GeoJSON Simplifier (for large datasets) ─────────────
# Uses System.Text.Json for fast parsing and iterative Douglas-Peucker
if (-not ([System.Management.Automation.PSTypeName]'GeoSimplifier').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;

public class GeoSimplifier
{
    public class Zone
    {
        public string TzId;
        public List<List<double[]>> Rings = new List<List<double[]>>();
    }

    public static List<Zone> ParseTZGeoJson(string filePath)
    {
        var zones = new List<Zone>();
        var text = File.ReadAllText(filePath);
        using var doc = JsonDocument.Parse(text);
        var features = doc.RootElement.GetProperty("features");
        foreach (var feat in features.EnumerateArray())
        {
            var props = feat.GetProperty("properties");
            string tzid = props.GetProperty("tzid").GetString();
            var geom = feat.GetProperty("geometry");
            string gtype = geom.GetProperty("type").GetString();
            var coords = geom.GetProperty("coordinates");
            var zone = new Zone { TzId = tzid };

            if (gtype == "Polygon")
            {
                foreach (var ring in coords.EnumerateArray())
                    zone.Rings.Add(ParseRing(ring));
            }
            else if (gtype == "MultiPolygon")
            {
                foreach (var poly in coords.EnumerateArray())
                    foreach (var ring in poly.EnumerateArray())
                        zone.Rings.Add(ParseRing(ring));
            }
            zones.Add(zone);
        }
        return zones;
    }

    static List<double[]> ParseRing(JsonElement ring)
    {
        var pts = new List<double[]>();
        foreach (var coord in ring.EnumerateArray())
        {
            double lon = coord[0].GetDouble();
            double lat = coord[1].GetDouble();
            pts.Add(new double[] { lat, lon });
        }
        return pts;
    }

    public static List<double[]> DecimateAndSimplify(List<double[]> ring, int step, double epsilon)
    {
        // Pre-decimate
        var decimated = new List<double[]>();
        decimated.Add(ring[0]);
        for (int i = step; i < ring.Count - 1; i += step)
            decimated.Add(ring[i]);
        decimated.Add(ring[ring.Count - 1]);

        // Douglas-Peucker (iterative)
        return DouglasPeucker(decimated, epsilon);
    }

    static List<double[]> DouglasPeucker(List<double[]> points, double epsilon)
    {
        if (points.Count <= 2) return points;
        var keep = new bool[points.Count];
        keep[0] = true;
        keep[points.Count - 1] = true;
        var stack = new Stack<(int, int)>();
        stack.Push((0, points.Count - 1));
        while (stack.Count > 0)
        {
            var (start, end) = stack.Pop();
            double maxDist = 0;
            int maxIdx = start;
            for (int i = start + 1; i < end; i++)
            {
                double d = PerpDist(points[i], points[start], points[end]);
                if (d > maxDist) { maxDist = d; maxIdx = i; }
            }
            if (maxDist > epsilon)
            {
                keep[maxIdx] = true;
                if (maxIdx - start > 1) stack.Push((start, maxIdx));
                if (end - maxIdx > 1) stack.Push((maxIdx, end));
            }
        }
        var result = new List<double[]>();
        for (int i = 0; i < points.Count; i++)
            if (keep[i]) result.Add(points[i]);
        return result;
    }

    static double PerpDist(double[] p, double[] a, double[] b)
    {
        double dx = b[0] - a[0], dy = b[1] - a[1];
        double lenSq = dx * dx + dy * dy;
        if (lenSq == 0) return Math.Sqrt((p[0]-a[0])*(p[0]-a[0]) + (p[1]-a[1])*(p[1]-a[1]));
        double num = Math.Abs(dy * p[0] - dx * p[1] + b[0] * a[1] - b[1] * a[0]);
        return num / Math.Sqrt(lenSq);
    }
}
'@
}

# ══════════════════════════════════════════════════════════════════════
# SUBMARINE CABLES
# ══════════════════════════════════════════════════════════════════════
if ($Source -eq 'SubmarineCables') {
    $rawFile = Join-Path $cacheDir 'cables-raw.json'
    $outputFile = Join-Path $outputDir 'submarine-cables.json'

    # Step 1: Download
    if (-not (Test-Path $rawFile)) {
        $url = 'https://raw.githubusercontent.com/tbotnz/submarine-cables-geojson/main/cables.json'
        Write-Host "[DOWNLOAD] Fetching submarine cable GeoJSON..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $url -OutFile $rawFile -UseBasicParsing
        $size = [Math]::Round((Get-Item $rawFile).Length / 1KB, 1)
        Write-Host "[OK] Downloaded: $size KB" -ForegroundColor Green
    }
    else {
        Write-Host "[CACHE] Using cached cables-raw.json" -ForegroundColor DarkGray
    }

    # Step 2: Parse
    Write-Host "[PARSE] Reading GeoJSON..." -ForegroundColor Yellow
    $geojson = Get-Content $rawFile -Raw | ConvertFrom-Json

    # Step 3: Deduplicate — multiple features can share the same cable name
    # Group by cable name and merge all line segments
    $cableMap = [ordered]@{}
    foreach ($feature in $geojson.features) {
        $name = $feature.properties.name
        $color = $feature.properties.color
        if (-not $name) { continue }

        if (-not $cableMap.Contains($name)) {
            $cableMap[$name] = @{
                Name     = $name
                Color    = if ($color) { $color } else { '#00BCD4' }
                Segments = [System.Collections.Generic.List[object]]::new()
            }
        }

        # Each MultiLineString has one or more line segments
        foreach ($line in $feature.geometry.coordinates) {
            # Convert from [lon, lat] to [lat, lon] double[][] for simplification
            $points = [System.Collections.Generic.List[double[]]]::new()
            foreach ($coord in $line) {
                $lon = [double]$coord[0]
                $lat = [double]$coord[1]
                $points.Add(@($lat, $lon))
            }
            if ($points.Count -ge 2) {
                $cableMap[$name].Segments.Add($points.ToArray())
            }
        }
    }

    Write-Host "[OK] Found $($cableMap.Count) unique cables from $($geojson.features.Count) features" -ForegroundColor Green

    # Step 4: Simplify each segment
    Write-Host "[SIMPLIFY] Applying Douglas-Peucker (tolerance=$Tolerance°)..." -ForegroundColor Yellow
    $totalOriginal = 0
    $totalSimplified = 0
    $cables = [System.Collections.Generic.List[object]]::new()

    foreach ($key in $cableMap.Keys) {
        $cable = $cableMap[$key]
        $simplifiedSegments = [System.Collections.Generic.List[object]]::new()

        foreach ($segment in $cable.Segments) {
            $totalOriginal += $segment.Count
            $simplified = Simplify-Polyline -Points $segment -Epsilon $Tolerance
            $totalSimplified += $simplified.Count

            # Only keep segments with 2+ points after simplification
            if ($simplified.Count -ge 2) {
                $coords = foreach ($pt in $simplified) {
                    , @([Math]::Round($pt[0], 2), [Math]::Round($pt[1], 2))
                }
                $simplifiedSegments.Add($coords)
            }
        }

        if ($simplifiedSegments.Count -gt 0) {
            $cables.Add(@{
                    Name     = $cable.Name
                    Color    = $cable.Color
                    Segments = $simplifiedSegments
                })
        }
    }

    Write-Host "[OK] Simplified: $totalOriginal -> $totalSimplified points" -ForegroundColor Green

    # Step 5: Write output
    Write-Host "[WRITE] Writing $outputFile..." -ForegroundColor Yellow
    $jsonOutput = @{
        meta   = @{
            source        = 'TeleGeography / tbotnz'
            sourceUrl     = 'https://github.com/tbotnz/submarine-cables-geojson'
            license       = 'Open (non-commercial)'
            tolerance     = $Tolerance
            generatedDate = (Get-Date -Format 'yyyy-MM-dd')
            generatedBy   = 'Convert-OverlayData.ps1'
            totalCables   = $cables.Count
            totalPoints   = $totalSimplified
        }
        cables = $cables
    }

    $jsonOutput | ConvertTo-Json -Depth 10 -Compress:$false | Set-Content $outputFile -Encoding UTF8
    $fileSize = [Math]::Round((Get-Item $outputFile).Length / 1KB, 1)
    Write-Host "[OK] Written: $outputFile ($fileSize KB)" -ForegroundColor Green

    # Summary
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  SUBMARINE CABLES SUMMARY" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Source:    TeleGeography (Open, non-commercial)" -ForegroundColor White
    Write-Host "  Cables:    $($cables.Count)" -ForegroundColor White
    Write-Host "  Points:    $totalOriginal -> $totalSimplified (tolerance=$Tolerance°)" -ForegroundColor White
    Write-Host "  Output:    $outputFile ($fileSize KB)" -ForegroundColor White
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════
# TIMEZONE BOUNDARIES
# ══════════════════════════════════════════════════════════════════════
if ($Source -eq 'TZBoundaries') {
    $zipFile = Join-Path $cacheDir 'timezones-now.geojson.zip'
    $extractDir = Join-Path $cacheDir 'tz-extract'
    $rawFile = Join-Path $extractDir 'combined-now.json'
    $outputFile = Join-Path $outputDir 'tz-boundaries.json'

    # Step 1: Download
    if (-not (Test-Path $rawFile)) {
        $url = 'https://github.com/evansiroky/timezone-boundary-builder/releases/download/2026a/timezones-now.geojson.zip'
        Write-Host "[DOWNLOAD] Fetching timezone-boundary-builder (timezones-now, 2026a)..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $url -OutFile $zipFile -UseBasicParsing
        $sz = [Math]::Round((Get-Item $zipFile).Length / 1MB, 1)
        Write-Host "[OK] Downloaded: $sz MB" -ForegroundColor Green

        Write-Host "[EXTRACT] Unzipping..." -ForegroundColor Yellow
        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force
        Write-Host "[OK] Extracted to $extractDir" -ForegroundColor Green
    }
    else {
        Write-Host "[CACHE] Using cached combined-now.json" -ForegroundColor DarkGray
    }

    # Step 2+3: Parse and simplify via compiled C# (PowerShell too slow for 3.5M points)
    $tzTolerance = [Math]::Max($Tolerance, 0.5)
    $decimateStep = 20
    Write-Host "[PARSE+SIMPLIFY] C# fast path: parse, decimate (1/$decimateStep), DP (tolerance=$tzTolerance°)..." -ForegroundColor Yellow

    $parsedZones = [GeoSimplifier]::ParseTZGeoJson($rawFile)
    Write-Host "[OK] Parsed $($parsedZones.Count) timezone zones" -ForegroundColor Green

    $totalOriginal = 0
    $totalSimplified = 0
    $zones = [System.Collections.Generic.List[object]]::new()

    foreach ($pz in $parsedZones) {
        $tzid = $pz.TzId
        if (-not $tzid) { continue }

        # Compute UTC offset (Windows TZ lookup from IANA id)
        try {
            $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($tzid)
            $offset = $tz.GetUtcOffset([DateTime]::UtcNow)
            $utcOffset = '{0}{1:00}:{2:00}' -f $(if ($offset.TotalHours -ge 0) { '+' } else { '' }), [int][Math]::Truncate($offset.TotalHours), [Math]::Abs($offset.Minutes)
        }
        catch {
            $utcOffset = '?'
        }

        $simplifiedRings = [System.Collections.Generic.List[object]]::new()
        foreach ($ring in $pz.Rings) {
            $totalOriginal += $ring.Count
            $simplified = [GeoSimplifier]::DecimateAndSimplify($ring, $decimateStep, $tzTolerance)
            $totalSimplified += $simplified.Count
            if ($simplified.Count -ge 3) {
                $coords = foreach ($pt in $simplified) {
                    , @([Math]::Round($pt[0], 2), [Math]::Round($pt[1], 2))
                }
                $simplifiedRings.Add($coords)
            }
        }

        if ($simplifiedRings.Count -gt 0) {
            $zones.Add(@{
                    TzId      = $tzid
                    UtcOffset = $utcOffset
                    Rings     = $simplifiedRings
                })
        }
        Write-Verbose "  $tzid`: $($pz.Rings.Count) rings, offset $utcOffset"
    }

    Write-Host "[OK] Simplified: $totalOriginal -> $totalSimplified points ($([Math]::Round(($totalSimplified / [Math]::Max($totalOriginal,1)) * 100, 1))% retained)" -ForegroundColor Green

    # Step 4: Write output
    Write-Host "[WRITE] Writing $outputFile..." -ForegroundColor Yellow
    $jsonOutput = @{
        meta  = @{
            source        = 'timezone-boundary-builder'
            sourceUrl     = 'https://github.com/evansiroky/timezone-boundary-builder'
            license       = 'ODbL 1.0'
            version       = '2026a'
            variant       = 'combined-now (merged same-timekeeping zones)'
            tolerance     = $tzTolerance
            generatedDate = (Get-Date -Format 'yyyy-MM-dd')
            generatedBy   = 'Convert-OverlayData.ps1'
            totalZones    = $zones.Count
            totalPoints   = $totalSimplified
        }
        zones = $zones
    }

    $jsonOutput | ConvertTo-Json -Depth 10 -Compress | Set-Content $outputFile -Encoding UTF8
    $fileSize = [Math]::Round((Get-Item $outputFile).Length / 1KB, 1)
    Write-Host "[OK] Written: $outputFile ($fileSize KB)" -ForegroundColor Green

    # Summary
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  TIMEZONE BOUNDARIES SUMMARY" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Source:    timezone-boundary-builder 2026a (ODbL 1.0)" -ForegroundColor White
    Write-Host "  Variant:   combined-now (merged same-timekeeping)" -ForegroundColor White
    Write-Host "  Zones:     $($zones.Count)" -ForegroundColor White
    Write-Host "  Points:    $totalOriginal -> $totalSimplified (tolerance=$tzTolerance°)" -ForegroundColor White
    Write-Host "  Output:    $outputFile ($fileSize KB)" -ForegroundColor White
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════
#  EEZ BOUNDARIES — Maritime Boundaries from Marine Regions WFS
# ══════════════════════════════════════════════════════════════════════
if ($Source -eq 'EEZ') {
    $rawFile = Join-Path $cacheDir 'eez-boundaries-raw.json'
    $outputFile = Join-Path $outputDir 'eez-boundaries.json'

    # Step 1: Download from WFS (all ~2349 boundary lines)
    if (-not (Test-Path $rawFile)) {
        $wfsUrl = 'https://geo.vliz.be/geoserver/MarineRegions/ows?service=WFS&version=1.1.0&request=GetFeature&typeName=MarineRegions:eez_boundaries&outputFormat=application/json'
        Write-Host "[DOWNLOAD] Fetching EEZ boundaries from Marine Regions WFS (~2349 features)..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $wfsUrl -OutFile $rawFile -UseBasicParsing -TimeoutSec 120
        $sz = [Math]::Round((Get-Item $rawFile).Length / 1MB, 1)
        Write-Host "[OK] Downloaded: $sz MB" -ForegroundColor Green
    }
    else {
        Write-Host "[CACHE] Using cached eez-boundaries-raw.json" -ForegroundColor DarkGray
    }

    # Step 2+3: Parse and simplify via compiled C#
    $eezTolerance = [Math]::Max($Tolerance, 0.3)
    Write-Host "[PARSE+SIMPLIFY] C# fast path: parse, simplify (tolerance=$eezTolerance°)..." -ForegroundColor Yellow

    # Add EEZ-specific C# parser
    if (-not ([System.Management.Automation.PSTypeName]'EEZParser').Type) {
        Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.IO;

public class EEZParser
{
    public class Boundary
    {
        public string Name;
        public string Territory1;
        public string Territory2;
        public string LineType;
        public double LengthKm;
        public List<List<double[]>> Segments = new List<List<double[]>>();
    }

    public static List<Boundary> ParseEEZGeoJson(string filePath)
    {
        var boundaries = new List<Boundary>();
        var text = File.ReadAllText(filePath);
        using var doc = JsonDocument.Parse(text);
        var features = doc.RootElement.GetProperty("features");
        foreach (var feat in features.EnumerateArray())
        {
            var props = feat.GetProperty("properties");
            string name = GetStr(props, "line_name");
            string ter1 = GetStr(props, "territory1");
            string ter2 = GetStr(props, "territory2");
            string ltype = GetStr(props, "line_type");
            double lenKm = 0;
            if (props.TryGetProperty("length_km", out var lkm) && lkm.ValueKind == JsonValueKind.Number)
                lenKm = lkm.GetDouble();

            var geom = feat.GetProperty("geometry");
            string gtype = geom.GetProperty("type").GetString();
            var coords = geom.GetProperty("coordinates");
            var b = new Boundary { Name = name, Territory1 = ter1, Territory2 = ter2, LineType = ltype, LengthKm = lenKm };

            if (gtype == "LineString")
            {
                b.Segments.Add(ParseLine(coords));
            }
            else if (gtype == "MultiLineString")
            {
                foreach (var line in coords.EnumerateArray())
                    b.Segments.Add(ParseLine(line));
            }
            boundaries.Add(b);
        }
        return boundaries;
    }

    static string GetStr(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.String)
            return v.GetString();
        return "";
    }

    static List<double[]> ParseLine(JsonElement line)
    {
        var pts = new List<double[]>();
        foreach (var coord in line.EnumerateArray())
        {
            // WFS 1.1.0 EPSG:4326 axis order = lat, lon
            double v0 = coord[0].GetDouble();
            double v1 = coord[1].GetDouble();
            double lat, lon;
            if (Math.Abs(v0) > 90) { lon = v0; lat = v1; }
            else if (Math.Abs(v1) > 180) { lat = v0; lon = v1; }
            else { lat = v0; lon = v1; }
            pts.Add(new double[] { lat, lon });
        }
        return pts;
    }
}
'@
    }

    $parsedBoundaries = [EEZParser]::ParseEEZGeoJson($rawFile)
    Write-Host "[OK] Parsed $($parsedBoundaries.Count) boundary lines" -ForegroundColor Green

    $totalOriginal = 0
    $totalSimplified = 0
    $boundaries = [System.Collections.Generic.List[object]]::new()
    $decimateStep = 5

    foreach ($b in $parsedBoundaries) {
        $simplifiedSegments = [System.Collections.Generic.List[object]]::new()
        foreach ($seg in $b.Segments) {
            $totalOriginal += $seg.Count
            if ($seg.Count -gt 20) {
                $simplified = [GeoSimplifier]::DecimateAndSimplify($seg, $decimateStep, $eezTolerance)
            }
            else {
                $simplified = $seg
            }
            $totalSimplified += $simplified.Count
            if ($simplified.Count -ge 2) {
                $coords = foreach ($pt in $simplified) {
                    , @([Math]::Round($pt[0], 2), [Math]::Round($pt[1], 2))
                }
                $simplifiedSegments.Add($coords)
            }
        }

        if ($simplifiedSegments.Count -gt 0) {
            $boundaries.Add(@{
                    Name     = $b.Name
                    Type     = $b.LineType
                    Segments = $simplifiedSegments
                })
        }
    }

    Write-Host "[OK] Simplified: $totalOriginal -> $totalSimplified points ($([Math]::Round(($totalSimplified / [Math]::Max($totalOriginal,1)) * 100, 1))% retained)" -ForegroundColor Green

    # Step 4: Write output
    Write-Host "[WRITE] Writing $outputFile..." -ForegroundColor Yellow
    $jsonOutput = @{
        meta       = @{
            source          = 'Marine Regions (Flanders Marine Institute)'
            sourceUrl       = 'https://www.marineregions.org/'
            license         = 'CC-BY 4.0'
            dataset         = 'eez_boundaries v12 (2023)'
            tolerance       = $eezTolerance
            generatedDate   = (Get-Date -Format 'yyyy-MM-dd')
            generatedBy     = 'Convert-OverlayData.ps1'
            totalBoundaries = $boundaries.Count
            totalPoints     = $totalSimplified
        }
        boundaries = $boundaries
    }

    $jsonOutput | ConvertTo-Json -Depth 10 -Compress | Set-Content $outputFile -Encoding UTF8
    $fileSize = [Math]::Round((Get-Item $outputFile).Length / 1KB, 1)
    Write-Host "[OK] Written: $outputFile ($fileSize KB)" -ForegroundColor Green

    # Summary
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  EEZ BOUNDARIES SUMMARY" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Source:    Marine Regions v12 (CC-BY 4.0)" -ForegroundColor White
    Write-Host "  Boundaries: $($boundaries.Count)" -ForegroundColor White
    Write-Host "  Points:    $totalOriginal -> $totalSimplified (tolerance=$eezTolerance°)" -ForegroundColor White
    Write-Host "  Output:    $outputFile ($fileSize KB)" -ForegroundColor White
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════
#  POWER PLANTS — WRI Global Power Plant Database (CSV → filtered JSON)
# ══════════════════════════════════════════════════════════════════════
if ($Source -eq 'PowerPlants') {
    $rawFile = Join-Path $cacheDir 'global_power_plant_database.csv'
    $outputFile = Join-Path $outputDir 'power-plants.json'

    # Step 1: Download CSV from GitHub
    if (-not (Test-Path $rawFile)) {
        $url = 'https://raw.githubusercontent.com/wri/global-power-plant-database/master/output_database/global_power_plant_database.csv'
        Write-Host "[DOWNLOAD] Fetching WRI Global Power Plant Database CSV..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $url -OutFile $rawFile -UseBasicParsing
        $sz = [Math]::Round((Get-Item $rawFile).Length / 1MB, 1)
        Write-Host "[OK] Downloaded: $sz MB" -ForegroundColor Green
    }
    else {
        Write-Host "[CACHE] Using cached global_power_plant_database.csv" -ForegroundColor DarkGray
    }

    # Step 2: Parse and filter — keep plants >= 500 MW capacity for manageable count
    Write-Host "[PARSE] Reading CSV and filtering (>= 500 MW)..." -ForegroundColor Yellow
    $csv = Import-Csv -Path $rawFile
    Write-Host "[OK] Total plants in DB: $($csv.Count)" -ForegroundColor Green

    $minCapacity = 500  # MW — keeps ~800-1200 plants
    $filtered = $csv | Where-Object {
        [double]$_.capacity_mw -ge $minCapacity -and
        $_.latitude -and $_.longitude
    }
    Write-Host "[OK] Filtered to $($filtered.Count) plants (>= $minCapacity MW)" -ForegroundColor Green

    # Fuel type color mapping
    $fuelColors = @{
        'Coal'       = '#4A4A4A'  # dark gray
        'Gas'        = '#FF8C00'  # orange
        'Oil'        = '#8B4513'  # brown
        'Nuclear'    = '#FFD700'  # gold
        'Hydro'      = '#1E90FF'  # blue
        'Wind'       = '#32CD32'  # green
        'Solar'      = '#FFD93D'  # yellow
        'Biomass'    = '#228B22'  # forest green
        'Geothermal' = '#DC143C'  # crimson
        'Waste'      = '#808080'  # gray
        'Wave and Tidal' = '#00CED1' # teal
        'Petcoke'    = '#2F2F2F'
        'Cogeneration' = '#B8860B'
        'Storage'    = '#9370DB'
        'Other'      = '#AAAAAA'
    }

    # Step 3: Build output
    $plants = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $filtered) {
        $fuel = $row.primary_fuel
        $color = if ($fuelColors.ContainsKey($fuel)) { $fuelColors[$fuel] } else { '#AAAAAA' }
        $plants.Add(@{
                Name     = $row.name
                Country  = $row.country_long
                Fuel     = $fuel
                MW       = [Math]::Round([double]$row.capacity_mw, 0)
                Lat      = [Math]::Round([double]$row.latitude, 3)
                Lon      = [Math]::Round([double]$row.longitude, 3)
                Color    = $color
            })
    }

    # Step 4: Write output
    Write-Host "[WRITE] Writing $outputFile..." -ForegroundColor Yellow
    $jsonOutput = @{
        meta   = @{
            source        = 'WRI Global Power Plant Database v1.3.0'
            sourceUrl     = 'https://datasets.wri.org/dataset/globalpowerplantdatabase'
            license       = 'CC-BY 4.0'
            minCapacityMW = $minCapacity
            generatedDate = (Get-Date -Format 'yyyy-MM-dd')
            generatedBy   = 'Convert-OverlayData.ps1'
            totalPlants   = $plants.Count
            fuelColors    = $fuelColors
        }
        plants = $plants
    }

    $jsonOutput | ConvertTo-Json -Depth 5 -Compress | Set-Content $outputFile -Encoding UTF8
    $fileSize = [Math]::Round((Get-Item $outputFile).Length / 1KB, 1)
    Write-Host "[OK] Written: $outputFile ($fileSize KB)" -ForegroundColor Green

    # Fuel breakdown
    $byFuel = $plants | Group-Object -Property Fuel | Sort-Object -Property Count -Descending
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  POWER PLANTS SUMMARY" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Source:    WRI GPPD v1.3.0 (CC-BY 4.0)" -ForegroundColor White
    Write-Host "  Filter:    >= $minCapacity MW" -ForegroundColor White
    Write-Host "  Plants:    $($plants.Count)" -ForegroundColor White
    Write-Host "  Output:    $outputFile ($fileSize KB)" -ForegroundColor White
    Write-Host "  By fuel:" -ForegroundColor White
    foreach ($g in $byFuel) {
        Write-Host "    $($g.Name): $($g.Count)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════
#  CONFLICT ZONES — UCDP Georeferenced Event Dataset (GED)
# ══════════════════════════════════════════════════════════════════════
if ($Source -eq 'ConflictZones') {
    $zipFile = Join-Path $cacheDir 'ged251-csv.zip'
    $rawFile = Join-Path $cacheDir 'GEDEvent_v25_1.csv'
    $outputFile = Join-Path $outputDir 'conflict-zones.json'

    # Step 1: Download UCDP GED v25.1 static CSV (no API token needed)
    if (-not (Test-Path $rawFile)) {
        $url = 'https://ucdp.uu.se/downloads/ged/ged251-csv.zip'
        Write-Host "[DOWNLOAD] Fetching UCDP GED v25.1 CSV (~150 MB zip)..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $url -OutFile $zipFile -UseBasicParsing
        $sz = [Math]::Round((Get-Item $zipFile).Length / 1MB, 1)
        Write-Host "[OK] Downloaded: $sz MB" -ForegroundColor Green

        Write-Host "[EXTRACT] Unzipping..." -ForegroundColor Yellow
        Expand-Archive -Path $zipFile -DestinationPath $cacheDir -Force
        # Find the actual CSV file
        $csvFiles = Get-ChildItem -Path $cacheDir -Filter 'GEDEvent*.csv'
        if ($csvFiles.Count -eq 0) {
            throw "No GEDEvent CSV found in extracted files"
        }
        $rawFile = $csvFiles[0].FullName
        Write-Host "[OK] Extracted: $($csvFiles[0].Name)" -ForegroundColor Green
    }
    else {
        Write-Host "[CACHE] Using cached UCDP GED CSV" -ForegroundColor DarkGray
    }

    # Step 2: Parse CSV — filter to recent events (last 5 years) for relevance
    Write-Host "[PARSE] Reading CSV and filtering recent events (>= 2020)..." -ForegroundColor Yellow
    $allEvents = Import-Csv -Path $rawFile | Where-Object {
        [int]$_.year -ge 2020 -and
        $_.latitude -and $_.longitude -and
        [double]$_.latitude -ne 0
    }
    Write-Host "[OK] Filtered to $($allEvents.Count) events (2020+)" -ForegroundColor Green

    # Step 3: Aggregate to grid cells (~1° resolution)
    Write-Host "[AGGREGATE] Clustering events to 1° grid cells..." -ForegroundColor Yellow

    $gridSize = 1.0  # degree resolution
    $grid = @{}
    foreach ($evt in $allEvents) {
        $lat = [double]$evt.latitude
        $lon = [double]$evt.longitude
        if ($lat -eq 0 -and $lon -eq 0) { continue }  # skip unknown locations

        $gridLat = [Math]::Floor($lat / $gridSize) * $gridSize + $gridSize / 2
        $gridLon = [Math]::Floor($lon / $gridSize) * $gridSize + $gridSize / 2
        $key = "$gridLat,$gridLon"

        if (-not $grid.ContainsKey($key)) {
            $grid[$key] = @{
                Lat       = [Math]::Round($gridLat, 2)
                Lon       = [Math]::Round($gridLon, 2)
                Events    = 0
                Deaths    = 0
                Country   = $evt.country
                LastYear  = 0
                Types     = [System.Collections.Generic.HashSet[string]]::new()
            }
        }
        $grid[$key].Events++
        $grid[$key].Deaths += [int]$evt.best
        $year = [int]$evt.year
        if ($year -gt $grid[$key].LastYear) { $grid[$key].LastYear = $year }
        if ($evt.type_of_violence) {
            $grid[$key].Types.Add([string]$evt.type_of_violence) | Out-Null
        }
    }

    # Convert to output
    $zones = [System.Collections.Generic.List[object]]::new()
    foreach ($cell in $grid.Values) {
        $zones.Add(@{
                Lat      = $cell.Lat
                Lon      = $cell.Lon
                Events   = $cell.Events
                Deaths   = $cell.Deaths
                Country  = $cell.Country
                LastYear = $cell.LastYear
            })
    }

    Write-Host "[OK] Aggregated to $($zones.Count) grid cells" -ForegroundColor Green

    # Step 3: Write output
    Write-Host "[WRITE] Writing $outputFile..." -ForegroundColor Yellow
    $jsonOutput = @{
        meta  = @{
            source        = 'UCDP Georeferenced Event Dataset (GED) v25.1'
            sourceUrl     = 'https://ucdp.uu.se/downloads/'
            license       = 'CC-BY 4.0'
            gridSizeDeg   = $gridSize
            generatedDate = (Get-Date -Format 'yyyy-MM-dd')
            generatedBy   = 'Convert-OverlayData.ps1'
            totalEvents   = $allEvents.Count
            totalCells    = $zones.Count
        }
        zones = $zones
    }

    $jsonOutput | ConvertTo-Json -Depth 5 -Compress | Set-Content $outputFile -Encoding UTF8
    $fileSize = [Math]::Round((Get-Item $outputFile).Length / 1KB, 1)
    Write-Host "[OK] Written: $outputFile ($fileSize KB)" -ForegroundColor Green

    # Summary
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  CONFLICT ZONES SUMMARY" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Source:    UCDP GED v25.1 (CC-BY 4.0)" -ForegroundColor White
    Write-Host "  Events:    $($allEvents.Count)" -ForegroundColor White
    Write-Host "  Grid cells: $($zones.Count) (at $gridSize° resolution)" -ForegroundColor White
    Write-Host "  Output:    $outputFile ($fileSize KB)" -ForegroundColor White
    Write-Host ""
}
