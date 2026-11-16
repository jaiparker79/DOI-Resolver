#requires -Version 5.1
<##
Resolve-Dois.ps1 (Windows PowerShell 5.1 friendly)
--------------------------------------------------
• Scans the folder for CSVs: all_list_items_*.csv
• Normalizes DOIs (adds https://doi.org/ when needed)
• Resolves each DOI to its final landing URL using Invoke-WebRequest (GET), with cookies/session and manual Location capture
• Outputs: Title, Chapter/Article Title, Item Link, List Link, DOI, ResolvedURL
• Prints progress to the console and pauses at the end

Expected input headers:
  Title (A), Chapter/Article Title (B), Item Link (C), List Link (G), DOI (O)

Usage examples:
  powershell -ExecutionPolicy Bypass -File .\Resolve-Dois.ps1
  powershell -ExecutionPolicy Bypass -File .\Resolve-Dois.ps1 -InputFolder C:\path\to\csvs -Delimiter ';' -TimeoutSeconds 25 -SleepMs 150 -LogFailures
##>

[CmdletBinding()] 
param(
    [string]$InputFolder = $null,       # If not supplied, falls back to script folder or current directory
    [string]$Delimiter   = ',',         # Change to ';' or "`t" if needed
    [int]$TimeoutSeconds = 30,          # Web request timeout per DOI
    [int]$MaxRedirects   = 30,          # Maximum redirects to follow
    [int]$SleepMs        = 150,         # Milliseconds to sleep between DOI requests (be kind to servers)
    [switch]$LogFailures                 # Write a simple log of DOIs that did not resolve
)

# --- Determine working folder ---------------------------------------------
if ([string]::IsNullOrWhiteSpace($InputFolder)) {
    if ($PSScriptRoot) { $InputFolder = $PSScriptRoot }
    else { $InputFolder = (Get-Location).Path }
}

# Ensure TLS 1.2 (often required by doi.org and publisher sites)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Optional failure log
$failureLog = $null
if ($LogFailures) {
    $failureLog = Join-Path $InputFolder ("doi_resolver_failures_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
}

# --- Helpers ---------------------------------------------------------------
function Get-DoiUrl {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }

    $d = $Raw.Trim()
    # Strip wrappers like doi:, dx.doi.org/, quotes, etc.
    $d = ($d -replace '^(?i)doi\s*:\s*', '')
    $d = ($d -replace '^(?i)https?://(dx\.)?doi\.org/', '')
    $d = $d.Trim('"', "'", ' ', '[', ']')
    # Remove typical trailing punctuation
    $d = $d.TrimEnd('.', ';', ',', ')')

    if ([string]::IsNullOrWhiteSpace($d)) { return $null }

    # If it already looks like a URL, pass through; otherwise prepend doi.org
    if ($d -match '^(?i)https?://') { return $d }
    else { return ("https://doi.org/{0}" -f $d) }
}

function Resolve-DoiUrl {
    param(
        [string]$DoiUrl,
        [int]$TimeoutSec = 25,
        [int]$MaxRedirs  = 30,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session = $null
    )

    if ([string]::IsNullOrWhiteSpace($DoiUrl)) { return $null }

    # Persistent session (cookies across hops)
    if (-not $Session) { $Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession }

    # Primary headers (HTML-like)
    $headers1 = @{
        'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) DOI-Resolver/2.0'
        'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        'Accept-Language' = 'en-AU,en;q=0.9'
        'Cache-Control'   = 'no-cache'
    }

    # Alternate headers (some sites behave differently)
    $headers2 = @{
        'User-Agent'      = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.1 Safari/605.1.15'
        'Accept'          = 'text/html,*/*;q=0.8'
        'Accept-Language' = 'en-US,en;q=0.9'
        'Cache-Control'   = 'no-cache'
    }

    # Helper: extract final URL from a 3xx response by disabling auto-redirect
    function Get-LocationHeader {
        param($Uri, $Headers)
        try {
            # max redirects = 0 makes IWR throw on 3xx; the exception has the response
            $null = Invoke-WebRequest -Uri $Uri -Method Get -MaximumRedirection 0 -TimeoutSec $TimeoutSec -Headers $Headers -WebSession $Session -UseBasicParsing -ErrorAction Stop
            return $null
        } catch {
            try {
                $resp = $_.Exception.Response
                if ($resp -and $resp.Headers -and $resp.Headers['Location']) {
                    $loc = $resp.Headers['Location']
                    if ($loc -match '^(?i)https?://') { return $loc }
                    else {
                        $base = $resp.ResponseUri.AbsoluteUri
                        return ([System.Uri]::new([System.Uri]$base, $loc)).AbsoluteUri
                    }
                }
            } catch {}
            return $null
        }
    }

    # Try 1: normal GET with headers1 (auto redirects)
    try {
        $r = Invoke-WebRequest -Uri $DoiUrl -Method Get -MaximumRedirection $MaxRedirs -TimeoutSec $TimeoutSec -Headers $headers1 -WebSession $Session -UseBasicParsing -ErrorAction Stop
        if ($r.BaseResponse -and $r.BaseResponse.ResponseUri) { return $r.BaseResponse.ResponseUri.AbsoluteUri }
    } catch {
        $loc = Get-LocationHeader -Uri $DoiUrl -Headers $headers1
        if ($loc) { return $loc }
    }

    # Try 2: alternate headers (sometimes helps with SAGE/T&F/Wiley)
    try {
        $r = Invoke-WebRequest -Uri $DoiUrl -Method Get -MaximumRedirection $MaxRedirs -TimeoutSec $TimeoutSec -Headers $headers2 -WebSession $Session -UseBasicParsing -ErrorAction Stop
        if ($r.BaseResponse -and $r.BaseResponse.ResponseUri) { return $r.BaseResponse.ResponseUri.AbsoluteUri }
    } catch {
        $loc = Get-LocationHeader -Uri $DoiUrl -Headers $headers2
        if ($loc) { return $loc }
    }

    # Try 3: one more manual Location probe (if both above failed)
    $finalLoc = Get-LocationHeader -Uri $DoiUrl -Headers $headers1
    if ($finalLoc) { return $finalLoc }

    return $null
}

# --- Find inputs -----------------------------------------------------------
$csvFiles = Get-ChildItem -Path $InputFolder -Filter 'all_list_items_*.csv' -File | Sort-Object Name

if (-not $csvFiles -or $csvFiles.Count -eq 0) {
    Write-Host ("No files found matching 'all_list_items_*.csv' in {0}." -f $InputFolder) -ForegroundColor Yellow
    Write-Host "Press Enter to exit"
    [void] (Read-Host)
    exit 1
}

# Output file (timestamped)
$outputFilename = Join-Path $InputFolder ("doi_resolutions_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))

# Collector for all rows
$results = New-Object System.Collections.Generic.List[Object]

# Shared session for cookies across all DOI resolutions
$sharedSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# --- Process each CSV ------------------------------------------------------
foreach ($csv in $csvFiles) {
    Write-Host ("Processing file: {0}" -f $csv.Name) -ForegroundColor Cyan

    # Import CSV with the configured delimiter
    $rows = Import-Csv -Path $csv.FullName -Delimiter $Delimiter

    foreach ($row in $rows) {
        # Safely read expected columns
        $title               = $row.Title
        $chapterArticleTitle = $row.'Chapter/Article Title'
        $itemLink            = $row.'Item Link'
        $listLink            = $row.'List Link'
        $rawDoi              = $row.DOI

        # Normalize DOI and attempt resolution
        $normalizedDoiUrl = Get-DoiUrl -Raw $rawDoi

        $resolvedUrl = $null
        if (-not [string]::IsNullOrWhiteSpace($normalizedDoiUrl)) {
            $resolvedUrl = Resolve-DoiUrl -DoiUrl $normalizedDoiUrl -TimeoutSec $TimeoutSeconds -MaxRedirs $MaxRedirects -Session $sharedSession

            if ([string]::IsNullOrWhiteSpace($resolvedUrl)) {
                Write-Host ("DOI: {0} -> {1}" -f $normalizedDoiUrl, "[no resolution]") -ForegroundColor Green
                if ($failureLog) { Add-Content -Path $failureLog -Value ("{0}`t{1}" -f $normalizedDoiUrl, "no resolution after retries") }
            } else {
                Write-Host ("DOI: {0} -> {1}" -f $normalizedDoiUrl, $resolvedUrl) -ForegroundColor Green
            }
        } else {
            Write-Host ("DOI: {0} -> {1}" -f $rawDoi, "[skipped - empty/invalid]") -ForegroundColor DarkYellow
            if ($failureLog) { Add-Content -Path $failureLog -Value ("{0}`t{1}" -f $rawDoi, "skipped - empty/invalid") }
        }

        # Decide the output DOI value BEFORE creating the hashtable
        $doiOut = if ([string]::IsNullOrWhiteSpace($normalizedDoiUrl)) { $rawDoi } else { $normalizedDoiUrl }

        # Build output record with requested columns + resolved URL
        $results.Add([pscustomobject]@{
            'Title'                 = $title
            'Chapter/Article Title' = $chapterArticleTitle
            'Item Link'             = $itemLink
            'List Link'             = $listLink
            'DOI'                   = $doiOut
            'ResolvedURL'           = $resolvedUrl
        })

        if ($SleepMs -gt 0) { Start-Sleep -Milliseconds $SleepMs }
    }
}

# --- Write output ----------------------------------------------------------
$results | Export-Csv -Path $outputFilename -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("Link checking complete. Please open {0}" -f $outputFilename) -ForegroundColor Magenta
Write-Host "Press Enter to exit"
[void] (Read-Host)