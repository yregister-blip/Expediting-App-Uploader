<#
================================================================================
 Create-ExpeditingRelease.ps1
================================================================================
 PURPOSE
   Loads a Tekla Production-Control export into a new per-release Smartsheet in
   the QMI Expediting workspace. For each run it:
     1. Copies the "Tekla Import Template" SHEET into the Job # folder
        (creating the Job # folder first if it does not exist yet).
     2. Renames the copy  "<Job> Seq <Seq> Rel <Rel>"  (Seq/Rel are read FROM
        the Tekla file, not typed - they live in the export).
     3. Wipes the copy and loads the dropped Tekla file (CSV or .xls/.xlsx;
        Excel files are read via Excel COM).
     4. Stamps the "Job #" sheet-summary field (Sequence(s)/Release(s) are
        formula summary fields that derive themselves from the loaded rows).
     5. VERIFIES every intended row actually landed; if not, the run fails
        LOUDLY and the copied sheet is renamed so it cannot be mistaken for a
        good release.

 HOW THIS DIFFERS FROM THE Create-Release SAMPLE (same engine, three changes)
   1. ONE input file (the Tekla export). There is NO optional bolt-hole list.
   2. "Main Piece" is a Smartsheet CHECKBOX, not a validated picklist. The
      Tekla CSV emits 1/0 and the Excel export emits TRUE/FALSE; both are
      normalized to a real JSON boolean (true/false) - that is exactly how a
      checkbox cell is stored (verified against an existing release sheet).
      (A "Weight" transform also strips the trailing "#" the CSV adds, e.g.
       "14.802083#" -> 14.802083; the .xlsx already exports a clean number.)
   3. Seq and Release are COLUMNS inside the Tekla file, so the GUI collects
      only the Job #. Seq/Rel for the sheet name are derived from the file.
   Architecturally: we copy a single SHEET (the release artifact here is one
   sheet per release inside the Job folder), not a whole folder.

 DESIGN NOTES (for MSP review)
   - No external modules. In-box Windows components only:
       * Windows.Forms / Drawing  (GUI)            - .NET Framework, preinstalled
       * Excel COM                (xlsx -> csv)    - Microsoft Excel, already used
       * Invoke-RestMethod        (Smartsheet API) - PowerShell built-in
       * DPAPI via Export/Import-Clixml (token)    - Windows built-in
   - Outbound network: HTTPS to api.smartsheet.com ONLY.
   - API token is encrypted at rest with DPAPI; only the Windows user that
     saved it on this machine can decrypt it. Token is never written in clear.
   - Targets Windows PowerShell 5.1 so it compiles cleanly with PS2EXE. No
     PowerShell 7 syntax (no ternary, no ??). See Build.ps1.
================================================================================
#>

# ============================== CONFIGURATION =================================
# Get IDs from the Smartsheet UI (right-click an item > Properties) or a GET call.
$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }

# >>> DEFAULT (TEST) ENVIRONMENT VALUES <<<
# The "Tekla Import Template" sheet that gets copied for each release. Unlike a
# COPIED sheet, the template itself keeps a STABLE id, so we can address it
# directly. (Prefilled with the test template id from the provided schema.)
$TemplateSheetId   = 2860137659191172
$TemplateSheetName = "Tekla Import Template"   # for the run log / sanity only

# The PARENT container that holds the Job # folders. This is NOT a Job folder
# itself - the tool finds/creates the "<Job>" folder inside it and copies each
# release sheet there. In TEST this is the "QMI Expediting" FOLDER, which lives
# inside the "Test Files" workspace (id 2497329237387140) and holds the Job
# folders (e.g. "26-10") plus the "Template" folder.
$DestinationId   = 4375630767777668   # "QMI Expediting" folder
$DestinationType = "folder"           # "workspace" | "folder"

# Shown in the window title + run log so a TEST build is never mistaken for PRODUCTION.
$EnvironmentName = "TEST"

# --- Optional external config (config.json next to the tool) -------------------
# Any key present here OVERRIDES the matching default above; missing keys keep the
# default; a missing or invalid file silently uses the defaults. Lets you flip
# test<->production by editing a text file - no rebuild/re-whitelist.
$ConfigPath   = Join-Path $ScriptPath "config.json"
$ConfigSource = "built-in defaults (no config.json)"
if (Test-Path $ConfigPath) {
    try {
        $cfg = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if (("$($cfg.TemplateSheetId)").Trim() -ne "") { $TemplateSheetId = [long]$cfg.TemplateSheetId }
        if (("$($cfg.DestinationId)").Trim()   -ne "") { $DestinationId   = [long]$cfg.DestinationId }
        if ($cfg.DestinationType) { $DestinationType = [string]$cfg.DestinationType }
        if ($cfg.EnvironmentName) { $EnvironmentName = [string]$cfg.EnvironmentName }
        $ConfigSource = "config.json"
    } catch {
        $ConfigSource = "built-in defaults (config.json could not be read: $_)"
    }
}

# Tool version - shown in the window title and run log. Bump on each released change.
$ScriptVersion = "1.0.0"

# Sheet-SUMMARY field on the copied sheet to stamp with the Job #. (Sequence(s)
# and Release(s) are FORMULA summary fields that derive themselves from the
# loaded Seq/Release # columns, so we never write them.) Resolved by title at
# runtime because a copy gets fresh field ids.
$SummaryJobField = "Job #"

# Per-user data folder for the encrypted token and the trace log.
$AppDataDir = Join-Path $env:LOCALAPPDATA "Qualico\ExpeditingUploader"
try { if (-not (Test-Path $AppDataDir)) { New-Item -ItemType Directory -Force -Path $AppDataDir | Out-Null } } catch { }
if (-not (Test-Path $AppDataDir)) { $AppDataDir = $ScriptPath }
$CredPath = Join-Path $AppDataDir "SmartsheetToken.cred"

# Header-row position (1-based) of the Tekla export: header on row 1, data from row 2.
$TeklaHeaderRow = 1

# Plain-text API trace log (in the per-user data folder above). Every request +
# response is appended; the API TOKEN is NEVER written here.
$ApiLogPath = Join-Path $AppDataDir "Expediting-Uploader-API.log"

# Keep the single trace log from growing without bound: trim in place if over cap.
$ApiLogMaxBytes  = 5MB
$ApiLogKeepBytes = 4MB
try {
    if ((Test-Path $ApiLogPath) -and ((Get-Item $ApiLogPath).Length -ge $ApiLogMaxBytes)) {
        $raw = Get-Content -Path $ApiLogPath -Raw
        if ($raw.Length -gt $ApiLogKeepBytes) {
            $tail = $raw.Substring($raw.Length - $ApiLogKeepBytes)
            $nl   = $tail.IndexOf("`n")
            if ($nl -ge 0) { $tail = $tail.Substring($nl + 1) }
            $header = ("==== [log trimmed {0}; older lines dropped to cap size] ====" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
            Set-Content -Path $ApiLogPath -Value ($header + "`r`n" + $tail) -Encoding UTF8 -NoNewline
        }
    }
} catch { }

# Suffixes stamped onto the sheet name when a load does NOT fully complete, so an
# incomplete release can never be mistaken for a good one.
$IncompleteSuffix = "  !! INCOMPLETE - DO NOT USE !!"
$DryRunSuffix     = "  (DRY RUN - no data loaded)"

# ----------------------------- COLUMN MAPPINGS -------------------------------
# Map by HEADER NAME in the Tekla file -> Smartsheet column TITLE. Only columns
# listed here are written; everything else is left for Smartsheet to compute.
# Blank source values are skipped. Titles resolve to live column ids at runtime
# (a copied sheet gets fresh column ids - see Resolve-ColumnMap).
#
# In this export the file headers match the Smartsheet column titles 1:1.
$TeklaColumnMap = @{
    "Seq"          = "Seq"
    "Release #"    = "Release #"
    "Dwg"          = "Dwg"
    "Main Piece"   = "Main Piece"
    "Main Mk"      = "Main Mk"
    "Piece Mk"     = "Piece Mk"
    "Shape"        = "Shape"
    "Dimension"    = "Dimension"
    "Grade"        = "Grade"
    "Qty"          = "Qty"
    "Length"       = "Length"
    "Weight"       = "Weight"
    "Pay Category" = "Pay Category"
    "Remarks"      = "Remarks"
}

# Source headers used to derive the Seq / Release values for the sheet NAME.
$SeqHeader     = "Seq"
$ReleaseHeader = "Release #"

# ----------------------------- VALUE TRANSFORMS ------------------------------
# Per-column value fixers (source header -> scriptblock). A transform may return
# a [bool] (for a CHECKBOX column) or a string. Headers not listed pass through.
$TeklaValueTransforms = @{
    # "Main Piece" is a Smartsheet CHECKBOX. Tekla emits it differently per
    # format: the CSV uses 1/0, the Excel export uses TRUE/FALSE. Normalize both
    # (case-insensitively) to a real boolean - exactly how a checkbox cell is
    # stored. Anything unexpected passes through unchanged so a stray value
    # surfaces loudly rather than being silently rewritten.
    "Main Piece" = { param($v)
        $t = ([string]$v).Trim().ToUpperInvariant()
        if     ($t -eq '1' -or $t -eq 'TRUE'  -or $t -eq 'Y' -or $t -eq 'YES') { $true }
        elseif ($t -eq '0' -or $t -eq 'FALSE' -or $t -eq 'N' -or $t -eq 'NO')  { $false }
        else   { $v }
    }
    # The Tekla CSV appends "#" to Weight (e.g. "14.802083#"); strip it so the
    # value lands as a number. (The .xlsx already exports a clean number, so this
    # is a no-op there.)
    "Weight" = { param($v) (([string]$v) -replace '#', '').Trim() }
}
# =============================================================================


# ============================== TOKEN HANDLING ===============================
function Get-ApiToken {
    $token = $null
    if (Test-Path $CredPath) {
        try {
            $sec  = Import-Clixml $CredPath
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            try   { $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        } catch { $token = $null }
    }
    return $token
}

function Set-ApiToken {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Smartsheet API Token"
    $dlg.ClientSize      = New-Object System.Drawing.Size(440, 150)
    $dlg.StartPosition   = "CenterScreen"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MinimizeBox     = $false
    $dlg.MaximizeBox     = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "First-run setup: paste your Smartsheet API token.`r`n(Saved encrypted with DPAPI; you won't be asked again on this PC.)"
    $lbl.SetBounds(12, 12, 416, 40)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.UseSystemPasswordChar = $true
    $tb.SetBounds(12, 58, 416, 24)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Save"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.SetBounds(262, 100, 80, 30)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.SetBounds(348, 100, 80, 30)

    $dlg.Controls.AddRange(@($lbl, $tb, $btnOk, $btnCancel))
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    $result = $dlg.ShowDialog()
    $entry  = $tb.Text
    $dlg.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK -or [string]::IsNullOrWhiteSpace($entry)) { return $null }
    $sec = ConvertTo-SecureString $entry -AsPlainText -Force
    $sec | Export-Clixml $CredPath -Force
    return $entry
}
# =============================================================================


# ============================== SMARTSHEET API ===============================
$script:Headers = $null

function Write-ApiLog {
    param([string]$Text)
    try {
        $line = ("{0}  {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff"), $Text)
        Add-Content -Path $ApiLogPath -Value $line -Encoding UTF8
    } catch { }
}

# PS 5.1 collapses a single-element array to a JSON object. Smartsheet's bulk row
# endpoints REQUIRE an array, so force array semantics here.
function ConvertTo-SSJson {
    param($Obj)
    if (($Obj -is [System.Collections.IEnumerable]) -and ($Obj -isnot [string]) -and ($Obj -isnot [hashtable])) {
        $tmp = New-Object System.Collections.ArrayList
        foreach ($item in $Obj) { [void]$tmp.Add($item) }
        $arr = $tmp.ToArray()
        if ($arr.Count -eq 1) {
            return '[' + ($arr[0] | ConvertTo-Json -Depth 10 -Compress) + ']'
        }
        return $arr | ConvertTo-Json -Depth 10
    }
    return $Obj | ConvertTo-Json -Depth 10
}

# Prefer Smartsheet's Retry-After header; else bounded exponential backoff. Cap 300s.
function Get-RetryDelaySeconds {
    param([int]$Status, $Response, [int]$Try)
    $retryAfter = $null
    if ($Response) {
        try { $retryAfter = $Response.Headers['Retry-After'] } catch { $retryAfter = $null }
    }
    if ($retryAfter) {
        $seconds = 0
        if ([int]::TryParse([string]$retryAfter, [ref]$seconds)) {
            if ($seconds -gt 0) { return [math]::Min($seconds, 300) }
        } else {
            try {
                $retryAt = [DateTime]::Parse([string]$retryAfter)
                $delta = [int][math]::Ceiling(($retryAt.ToUniversalTime() - (Get-Date).ToUniversalTime()).TotalSeconds)
                if ($delta -gt 0) { return [math]::Min($delta, 300) }
            } catch { }
        }
    }
    $base   = 5 * [math]::Pow(2, [math]::Max(0, $Try - 1))
    $jitter = Get-Random -Minimum 0 -Maximum 6
    $delay  = [int]($base + $jitter)
    if ($Status -eq 429) { $delay = [math]::Max($delay, 30) }
    return [math]::Min($delay, 300)
}

# Single choke point for every API call. Retries on 429 and 5xx; other 4xx surface.
function Invoke-SS {
    param([string]$Method, [string]$Url, $Body = $null, [int]$MaxTries = 4)
    $params = @{ Uri = $Url; Method = $Method; Headers = $script:Headers; TimeoutSec = 300 }
    $bodyJson = $null
    if ($null -ne $Body) { $bodyJson = (ConvertTo-SSJson $Body); $params.Body = $bodyJson }
    $bodyLen = if ($bodyJson) { $bodyJson.Length } else { 0 }

    $try = 0
    while ($true) {
        $try++
        $reqStart = Get-Date
        Write-ApiLog ("--> [try {0}] {1} {2}  (body {3} bytes)" -f $try, $Method, $Url, $bodyLen)
        try {
            $resp = Invoke-RestMethod @params
            $ms = [int]((Get-Date) - $reqStart).TotalMilliseconds
            Write-ApiLog ("<-- 200 OK   {0} {1}  ({2} ms)" -f $Method, $Url, $ms)
            return $resp
        } catch {
            $ms = [int]((Get-Date) - $reqStart).TotalMilliseconds
            $status = 0
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
            }
            $errBody = $null
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $errBody = $_.ErrorDetails.Message }
            elseif ($_.Exception.Response) {
                try { $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $errBody = $sr.ReadToEnd() } catch { }
            }
            if (-not $errBody) { $errBody = $_.Exception.Message }
            $oneLine = ($errBody -replace "\s+", " ").Trim()
            if ($oneLine.Length -gt 500) { $oneLine = $oneLine.Substring(0, 500) + "..." }
            Write-ApiLog ("<-- {0} ERR  {1} {2}  ({3} ms)  body: {4}" -f $status, $Method, $Url, $ms, $oneLine)

            if ($try -ge $MaxTries) { Write-ApiLog ("    giving up after {0} try/tries" -f $try); throw }
            if ($status -eq 429 -or $status -ge 500) {
                $delay = Get-RetryDelaySeconds -Status $status -Response $_.Exception.Response -Try $try
                if ($status -eq 429) { Write-ApiLog ("    429 rate-limited; sleeping {0}s before retry" -f $delay) }
                else                 { Write-ApiLog ("    5xx server error; sleeping {0}s before retry" -f $delay) }
                Start-Sleep -Seconds $delay
            } else {
                Write-ApiLog "    non-retryable; throwing"; throw
            }
        }
    }
}

# Copy the template SHEET into the Job folder under a new name. Returns the new
# sheet id. include=... keeps data/forms/filters etc.; summary fields (incl. the
# Sequence(s)/Release(s) formulas) come along automatically with a sheet copy.
function Copy-TemplateSheet {
    param([long]$DestFolderId, [string]$NewName, [scriptblock]$Log)
    & $Log "Copying template sheet -> '$NewName'..."
    $inc = "data,attachments,discussions,filters,forms,rules,ruleRecipients,cellLinks"
    $url = "https://api.smartsheet.com/2.0/sheets/$TemplateSheetId/copy?include=$inc"
    $body = @{
        destinationType = "folder"
        destinationId   = $DestFolderId
        newName         = $NewName
    }
    $resp = Invoke-SS -Method Post -Url $url -Body $body
    $id = $resp.result.id
    if (-not $id) { throw "Sheet copy did not return a new sheet id (response: $($resp | ConvertTo-Json -Depth 5 -Compress))." }
    & $Log "  Created sheet id $id"
    $script:LastReleaseUrl = [string]$resp.result.permalink
    if (-not $script:LastReleaseUrl) {
        try { $script:LastReleaseUrl = [string](Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($id)?page=1&pageSize=1").permalink } catch { }
    }
    return $id
}

function Rename-Sheet {
    param([long]$SheetId, [string]$NewName)
    Invoke-SS -Method Put -Url "https://api.smartsheet.com/2.0/sheets/$SheetId" -Body @{ name = $NewName } | Out-Null
}

# Try to stamp a bad/partial copy as INCOMPLETE and REPORT whether it worked. A
# failed rename must NOT be reported as success.
function Set-SheetIncomplete {
    param([long]$SheetId, [string]$BaseName, [scriptblock]$Log)
    $badName = New-SuffixedName $BaseName $IncompleteSuffix
    try {
        Rename-Sheet -SheetId $SheetId -NewName $badName
        & $Log "  Sheet renamed '$badName'."
        return "The copied sheet was renamed '$badName' - DELETE it and re-run."
    } catch {
        & $Log "  WARNING: could NOT rename the sheet to mark it incomplete: $_"
        Write-ApiLog ("    Mark-incomplete FAILED for sheet id {0}: {1}" -f $SheetId, $_)
        return "WARNING: the copied sheet (id $SheetId) could NOT be marked incomplete - find it in Smartsheet and DELETE it manually before using any release."
    }
}

# Smartsheet sheet/folder names are capped at 50 chars. Always KEEP the status
# suffix and truncate the base name to fit.
$MaxName = 50
function New-SuffixedName {
    param([string]$Base, [string]$Suffix)
    $full = $Base + $Suffix
    if ($full.Length -le $MaxName) { return $full }
    $keep = $MaxName - $Suffix.Length
    if ($keep -lt 0) { return $Suffix.Substring(0, $MaxName) }
    return $Base.Substring(0, $keep) + $Suffix
}

# --- Job # folders -----------------------------------------------------------
# Each Job keeps its release sheets inside a folder named after the Job # (e.g.
# "26-10"). The configured DESTINATION is the PARENT that holds those Job folders.

function Get-DestinationFolders {
    param([scriptblock]$Log)
    $folders = New-Object System.Collections.Generic.List[object]
    if ($DestinationType -eq 'workspace') {
        $d = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/workspaces/$DestinationId"
    } else {
        $d = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/folders/$DestinationId"
    }
    foreach ($f in $d.folders) { if ($f.name) { [void]$folders.Add($f) } }
    return $folders
}

# Find the Job # folder by name; create it if missing. Returns its id. We do NOT
# swallow a listing failure - creating blindly could split a job across two
# identically-named folders.
function Get-OrCreateJobFolder {
    param([string]$JobName, [scriptblock]$Log)
    $target = $JobName.Trim()
    $folders = @(Get-DestinationFolders -Log $Log)
    foreach ($f in $folders) {
        if (([string]$f.name).Trim().ToLowerInvariant() -eq $target.ToLowerInvariant()) {
            & $Log "Using existing Job folder '$target' (id $($f.id))."
            Write-ApiLog ("    JobFolder: found existing '{0}' id {1}." -f $target, $f.id)
            return [long]$f.id
        }
    }
    & $Log "Job folder '$target' does not exist; creating it..."
    if ($DestinationType -eq 'workspace') {
        $url = "https://api.smartsheet.com/2.0/workspaces/$DestinationId/folders"
    } else {
        $url = "https://api.smartsheet.com/2.0/folders/$DestinationId/folders"
    }
    $resp = Invoke-SS -Method Post -Url $url -Body @{ name = $target }
    $id = $resp.result.id
    if (-not $id) { throw "Failed to create Job folder '$target' (no id returned)." }
    & $Log "  Created Job folder id $id."
    Write-ApiLog ("    JobFolder: created '{0}' id {1}." -f $target, $id)
    return [long]$id
}

# Existing SHEET names already in the Job folder (release sheets live directly in
# it). FAIL CLOSED: if we can't list them we can't dedup, so abort before copy.
function Get-JobSheetNames {
    param([long]$FolderId, [scriptblock]$Log)
    $names = New-Object System.Collections.Generic.List[string]
    try {
        $d = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/folders/$FolderId"
        foreach ($s in $d.sheets) { if ($s.name) { [void]$names.Add([string]$s.name) } }
        Write-ApiLog ("    Dedup: Job folder listing returned {0} sheet name(s)." -f $names.Count)
    } catch {
        Write-ApiLog ("    Dedup: Job folder listing FAILED ($_) - aborting before copy.")
        throw "Could not list existing release sheets in the Job folder (id $FolderId) to check for duplicates: $_  >> Aborting before copy so a duplicate is not created. Re-run; if it persists, check connectivity to Smartsheet."
    }
    return $names
}

function Get-UniqueName {
    param([string]$BaseName, [string[]]$Existing)
    $taken = @{}
    foreach ($n in $Existing) { if ($n) { $taken[$n.ToLowerInvariant()] = $true } }
    if (-not $taken.ContainsKey($BaseName.ToLowerInvariant())) { return $BaseName }
    $i = 1
    while ($true) {
        $candidate = New-SuffixedName $BaseName "-$i"
        if (-not $taken.ContainsKey($candidate.ToLowerInvariant())) { return $candidate }
        $i++
    }
}

function Get-AllRowIds {
    param([long]$SheetId)
    $ids  = New-Object System.Collections.Generic.List[object]
    $page = 1
    $size = 5000
    while ($true) {
        $r = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)?page=$page&pageSize=$size"
        if ($r.rows) { foreach ($row in $r.rows) { [void]$ids.Add($row.id) } }
        $total = $r.totalRowCount
        $got   = if ($r.rows) { @($r.rows).Count } else { 0 }
        if ((-not $total) -or ($ids.Count -ge $total) -or ($got -lt $size)) { break }
        $page++
    }
    return $ids
}

function Get-SheetRowCount {
    param([long]$SheetId)
    $r = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)?page=1&pageSize=1"
    return [int]$r.totalRowCount
}

# Turn a header->TITLE map into the header->columnId map Add-Rows needs, by
# reading the ACTUAL columns of the just-copied sheet (copies get fresh ids).
function Resolve-ColumnMap {
    param([long]$SheetId, [hashtable]$TitleMap, [string]$SheetName, [scriptblock]$Log)
    $r = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)?page=1&pageSize=1"
    $titleToId = @{}
    foreach ($c in $r.columns) { $titleToId[[string]$c.title] = $c.id }

    $idMap   = @{}
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($hdr in $TitleMap.Keys) {
        $title = [string]$TitleMap[$hdr]
        if ($titleToId.ContainsKey($title)) { $idMap[$hdr] = $titleToId[$title] }
        else { [void]$missing.Add($title) }
    }
    if ($missing.Count -gt 0) {
        $have = (($titleToId.Keys | Sort-Object) -join ", ")
        throw "Sheet '$SheetName' (id $SheetId) is missing expected column(s): $($missing -join ', '). Columns present: $have"
    }
    & $Log ("  Resolved {0} columns by title on '{1}'." -f $idMap.Count, $SheetName)
    return $idMap
}

# Update sheet-SUMMARY fields by TITLE (fields get new ids on a copy). Values are
# sent as STRING so an alphanumeric Job # (e.g. "26-10") is preserved verbatim.
function Set-SheetSummary {
    param([long]$SheetId, [hashtable]$ValuesByTitle, [scriptblock]$Log)
    $r = Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)/summary"
    $titleToId = @{}
    foreach ($fld in $r.fields) { $titleToId[[string]$fld.title] = $fld.id }

    $updates = New-Object System.Collections.Generic.List[object]
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($title in $ValuesByTitle.Keys) {
        if ($titleToId.ContainsKey($title)) {
            [void]$updates.Add(@{ id = $titleToId[$title]; objectValue = [string]$ValuesByTitle[$title] })
        } else { [void]$missing.Add($title) }
    }
    if ($missing.Count -gt 0) {
        $have = (($titleToId.Keys | Sort-Object) -join ", ")
        throw "Sheet summary (id $SheetId) is missing expected field(s): $($missing -join ', '). Fields present: $have"
    }
    Invoke-SS -Method Put -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)/summary/fields" -Body $updates | Out-Null
    & $Log ("  Updated {0} sheet-summary field(s) ({1})." -f $updates.Count, (($ValuesByTitle.Keys | Sort-Object) -join ", "))
}

# A sheet copy can be ASYNCHRONOUS: an immediate GET can briefly 404. Poll until
# the sheet answers, or time out. (404 is NOT retried inside Invoke-SS.)
# URL NOTE: wrap the id as $($SheetId) before "?" - PS treats "?" as a legal
# variable-name char, so "$SheetId?page=1" expands to EMPTY -> a permanent 404.
function Wait-SheetReady {
    param([long]$SheetId, [string]$Name, [scriptblock]$Log, [int]$TimeoutSec = 300)
    $start    = Get-Date
    $deadline = $start.AddSeconds($TimeoutSec)
    $delay    = 3
    while ($true) {
        try {
            Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)?page=1&pageSize=1" -MaxTries 1 | Out-Null
            $elapsed = [int]((Get-Date) - $start).TotalSeconds
            if ($elapsed -gt 0) { & $Log "  '$Name' ready after $elapsed s." }
            return
        } catch {
            $status = 0
            if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 } }
            if ($status -ne 404) { throw }
            if ((Get-Date) -ge $deadline) {
                throw "Sheet '$Name' (id $SheetId) was still not available after $TimeoutSec s (post-copy provisioning)."
            }
            $elapsed = [int]((Get-Date) - $start).TotalSeconds
            & $Log "  Waiting for '$Name' to finish provisioning... ($elapsed s)"
            Start-Sleep -Seconds $delay
            if ($delay -lt 10) { $delay += 1 }
        }
    }
}

function Clear-Sheet {
    param([long]$SheetId, [scriptblock]$Log)
    $ids = @(Get-AllRowIds -SheetId $SheetId)
    if ($ids.Count -eq 0) { & $Log "  Sheet already empty."; return }
    & $Log "  Deleting $($ids.Count) existing rows..."
    for ($i = 0; $i -lt $ids.Count; $i += 300) {
        $end   = [math]::Min($i + 299, $ids.Count - 1)
        $idstr = ($ids[$i..$end]) -join ','
        $url   = "https://api.smartsheet.com/2.0/sheets/$SheetId/rows?ids=$idstr&ignoreRowsNotFound=true"
        Invoke-SS -Method Delete -Url $url | Out-Null
    }
    $remaining = Get-SheetRowCount -SheetId $SheetId
    if ($remaining -ne 0) {
        throw "Could not fully clear sheet $SheetId ($remaining rows remain). Aborting before load."
    }
}

# Load records (keyed by header) using a header->columnId map. Writes only mapped,
# non-blank cells; a transform may turn a value into a real boolean (CHECKBOX).
# Posts in batches of <=400 and throws if created != intended (partial-load guard).
function Add-Rows {
    param([long]$SheetId, $Records, [hashtable]$Map, [scriptblock]$Log, [hashtable]$Transforms = @{})

    $payloads = New-Object System.Collections.Generic.List[object]
    $skipped  = 0
    foreach ($rec in $Records) {
        $cells = @()
        foreach ($hdr in $Map.Keys) {
            $raw = $rec.$hdr
            if ($null -eq $raw) { continue }
            $sval = ([string]$raw).Trim()
            if ($sval -eq '') { continue }   # blank source cell -> leave column default
            if ($Transforms.ContainsKey($hdr)) {
                $tv = & $Transforms[$hdr] $sval
                if ($tv -is [bool]) {
                    # Real boolean for a CHECKBOX cell (do NOT stringify or skip).
                    $cells += @{ columnId = $Map[$hdr]; value = $tv }
                    continue
                }
                $sval = [string]$tv
            }
            if ($sval -eq '') { continue }
            # strict=false (PER CELL) accepts free text into picklist-style columns.
            $cells += @{ columnId = $Map[$hdr]; value = $sval; strict = $false }
        }
        if ($cells.Count -gt 0) { [void]$payloads.Add(@{ toBottom = $true; cells = $cells }) }
        else { $skipped++ }
    }
    $intended = $payloads.Count
    if ($skipped -gt 0) { & $Log "  ($skipped blank/empty source rows skipped.)" }
    if ($intended -eq 0) {
        throw "No loadable rows were built for sheet ${SheetId}: the file had rows, but every mapped column was blank/unusable (likely the wrong file or shifted columns)."
    }

    $created = 0
    for ($i = 0; $i -lt $payloads.Count; $i += 400) {
        $end   = [math]::Min($i + 399, $payloads.Count - 1)
        $batch = $payloads[$i..$end]
        $url   = "https://api.smartsheet.com/2.0/sheets/$SheetId/rows"
        $resp  = Invoke-SS -Method Post -Url $url -Body $batch
        $made  = if ($resp.result) { @($resp.result).Count } else { 0 }
        $created += $made
        & $Log ("  ...added {0} rows (total {1}/{2})" -f $made, $created, $intended)
    }

    if ($created -ne $intended) {
        throw "INCOMPLETE LOAD on sheet ${SheetId}: only $created of $intended rows were created."
    }
    try {
        $sheetCount = Get-SheetRowCount -SheetId $SheetId
        if ($sheetCount -eq $created) { & $Log "  Verified: sheet reports $sheetCount rows (matches)." }
        else { & $Log "  NOTE: sheet totalRowCount=$sheetCount vs added=$created (possible API lag; primary count gate passed)." }
    } catch { }

    & $Log "  Uploaded $created rows."
    return @{ Intended = $intended; Created = $created }
}
# =============================================================================


# ============================ FILE READERS ===================================
# Convert an Excel workbook to a temp CSV via Excel COM and return that temp path.
function ConvertTo-TempCsv {
    param([string]$Path)
    $name = [System.IO.Path]::GetFileName($Path)
    $tempCsv = Join-Path $env:TEMP ("ExpeditingUploader_{0}.csv" -f ([guid]::NewGuid().ToString("N")))
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw "Cannot read '$name': Microsoft Excel is not available on this PC (Excel COM is required to read .xls/.xlsx). Either run on a machine with Excel installed, or save the report as .csv first."
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch { }   # force-disable macros
    try { $excel.EnableEvents = $false } catch { }
    $wb = $null
    $ws = $null
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)   # ReadOnly, no link update
        $ws = $wb.Worksheets.Item(1)
        $ws.Activate()
        if (Test-Path $tempCsv) { Remove-Item $tempCsv -Force }
        $wb.SaveAs($tempCsv, 6)   # 6 = xlCSV
    } finally {
        if ($ws) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null } catch { } }
        if ($wb) {
            try { $wb.Close($false) } catch { }
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { }
        }
        try { $excel.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return $tempCsv
}

# Read a CSV or Excel file into objects keyed by the header row. Excel files are
# converted to a temp CSV first so both formats parse through one path.
function Read-DataFile {
    param([string]$Path, [int]$HeaderRow = 1)
    $name = [System.IO.Path]::GetFileName($Path)
    $ext  = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    $tempCsv = $null
    $csvPath = $Path
    if ($ext -eq '.xlsx' -or $ext -eq '.xls') {
        $tempCsv = ConvertTo-TempCsv -Path $Path
        $csvPath = $tempCsv
    } elseif ($ext -ne '.csv') {
        throw "Unsupported file type '$ext' for '$name'. Use .csv, .xls, or .xlsx."
    }

    try {
        if ($HeaderRow -le 1) {
            $records = @(Import-Csv -Path $csvPath)
        } else {
            $lines = Get-Content $csvPath
            $headerIndex = $HeaderRow - 1
            if ($lines.Count -le $headerIndex) {
                throw "File '$name' has fewer than $HeaderRow rows; expected the header on row $HeaderRow."
            }
            $trimmed = $lines[$headerIndex..($lines.Count - 1)]
            $records = @($trimmed | ConvertFrom-Csv)
        }
    } finally {
        if ($tempCsv -and (Test-Path $tempCsv)) { Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue }
    }
    return @(ConvertTo-TrimmedHeaderRecords -Records $records)
}

# Rebuild each record so its property names are trimmed (so header validation and
# row loading agree - property access is case-insensitive but NOT whitespace-insensitive).
function ConvertTo-TrimmedHeaderRecords {
    param($Records)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($rec in @($Records)) {
        $obj = [ordered]@{}
        foreach ($p in $rec.PSObject.Properties) {
            $clean = ([string]$p.Name).Trim()
            if (-not $obj.Contains($clean)) { $obj[$clean] = $p.Value }
        }
        [void]$out.Add([pscustomobject]$obj)
    }
    return $out
}

# Confirm there is >=1 data row and every expected SOURCE header is present.
function Assert-SourceData {
    param($Records, [hashtable]$Map, [string]$FileLabel)
    $rows = @($Records)
    if ($rows.Count -eq 0) {
        throw "$FileLabel contains no data rows (it has only a header, or is empty). Nothing to load."
    }
    $present = @{}
    foreach ($p in $rows[0].PSObject.Properties.Name) { $present[([string]$p).Trim().ToLowerInvariant()] = $true }
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($hdr in $Map.Keys) {
        if (-not $present.ContainsKey(([string]$hdr).Trim().ToLowerInvariant())) { [void]$missing.Add($hdr) }
    }
    if ($missing.Count -gt 0) {
        $have = (($rows[0].PSObject.Properties.Name | Sort-Object) -join ", ")
        throw "$FileLabel is missing expected column header(s): $($missing -join ', '). Headers found in the file: $have"
    }
}

# Every row must carry a Seq and a Release # - they decide which release sheet a
# row belongs to (the upload is SPLIT by Seq+Rel). Fail loudly, listing how many
# rows are missing either, BEFORE we create anything in Smartsheet.
function Assert-SeqRelPresent {
    param($Records, [string]$SeqHeader, [string]$RelHeader)
    $bad = 0
    $n   = 0
    foreach ($rec in @($Records)) {
        $n++
        $seq = ([string]$rec.$SeqHeader).Trim()
        $rel = ([string]$rec.$RelHeader).Trim()
        if ($seq -eq '' -or $rel -eq '') { $bad++ }
    }
    if ($bad -gt 0) {
        throw "$bad of $n data row(s) are missing a '$SeqHeader' and/or '$RelHeader' value. Every row needs both so it can be placed in the correct release sheet. Fix the export and re-run."
    }
}

# Split records into one group per UNIQUE (Seq, Release #) pair, preserving the
# order each pair is first seen. Each group becomes its own release sheet.
function Group-BySeqRel {
    param($Records, [string]$SeqHeader, [string]$RelHeader)
    $order = New-Object System.Collections.Generic.List[string]
    $map   = @{}
    foreach ($rec in @($Records)) {
        $seq = ([string]$rec.$SeqHeader).Trim()
        $rel = ([string]$rec.$RelHeader).Trim()
        $key = ($seq.ToLowerInvariant() + "||" + $rel.ToLowerInvariant())
        if (-not $map.ContainsKey($key)) {
            $map[$key] = @{ Seq = $seq; Rel = $rel; Rows = (New-Object System.Collections.Generic.List[object]) }
            [void]$order.Add($key)
        }
        [void]$map[$key].Rows.Add($rec)
    }
    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($k in $order) { [void]$groups.Add($map[$k]) }
    return $groups
}
# =============================================================================


# =============================== ORCHESTRATION ===============================
# Loads the Tekla file, SPLIT into one release sheet per unique Seq+Release pair
# (e.g. one sequence with two releases -> two sheets). Each sheet is built
# independently and guarded: a failure on one marks THAT sheet incomplete and is
# reported, while the others still load. The run as a whole fails (throws) if ANY
# sheet failed, so a partial result is never silently reported as success.
function Invoke-Release {
    param([string]$Job, [string]$TeklaFile, [bool]$DryRun, [scriptblock]$Log)

    $script:LastReleaseUrl = $null
    Write-ApiLog ("==================== RUN START: Job '{0}'  (v{1}, {2}, DryRun={3}, config={4}) ====================" -f $Job, $ScriptVersion, $EnvironmentName, $DryRun, $ConfigSource)
    & $Log "=== Building releases for Job $Job  [$EnvironmentName  v$ScriptVersion] ==="
    & $Log "Config: $ConfigSource"
    & $Log "API trace log: $ApiLogPath"
    if ($DryRun) { & $Log "*** DRY RUN: will copy + name sheets, but load NO data. ***" }

    # Read AND validate the Tekla file BEFORE touching Smartsheet, so a bad/empty
    # file fails here and leaves no orphan sheet behind.
    & $Log "Reading + checking Tekla file..."
    $teklaRows = Read-DataFile -Path $TeklaFile -HeaderRow $TeklaHeaderRow
    Assert-SourceData -Records $teklaRows -Map $TeklaColumnMap -FileLabel "Tekla file"
    Assert-SeqRelPresent -Records $teklaRows -SeqHeader $SeqHeader -RelHeader $ReleaseHeader
    & $Log ("  Tekla file OK: {0} data row(s)." -f @($teklaRows).Count)

    # Split into one sheet per unique Seq+Release pair (they live IN the file).
    $groups = @(Group-BySeqRel -Records $teklaRows -SeqHeader $SeqHeader -RelHeader $ReleaseHeader)
    & $Log ("File contains {0} unique Sequence/Release combination(s) -> {0} release sheet(s):" -f $groups.Count)
    foreach ($g in $groups) { & $Log ("    - Seq {0} Rel {1}  ({2} row(s))" -f $g.Seq, $g.Rel, $g.Rows.Count) }
    Write-ApiLog ("    Split: {0} Seq/Rel group(s)." -f $groups.Count)

    # Ensure the Job # folder exists (create if missing); releases land inside it.
    $jobFolderId = Get-OrCreateJobFolder -JobName $Job -Log $Log

    # Existing sheet names in the Job folder, so we don't create a duplicate. We
    # ALSO reserve each name we create this run so two groups can't collide.
    $taken = New-Object System.Collections.Generic.List[string]
    foreach ($n in @(Get-JobSheetNames -FolderId $jobFolderId -Log $Log)) { [void]$taken.Add($n) }

    $succeeded = New-Object System.Collections.Generic.List[string]
    $failed    = New-Object System.Collections.Generic.List[string]
    $lastGoodUrl = $null

    foreach ($g in $groups) {
        # MUST reset per group: otherwise a copy failure here would fall into the
        # catch holding the PREVIOUS group's id and mark that good sheet incomplete.
        $sheetId  = $null
        $baseName = New-SuffixedName "$Job Seq $($g.Seq) Rel $($g.Rel)" ""
        $uniqueName = Get-UniqueName -BaseName $baseName -Existing $taken.ToArray()
        if ($uniqueName -ne $baseName) {
            & $Log "A release sheet named '$baseName' already exists; using '$uniqueName' instead."
            Write-ApiLog ("    Dedup: '{0}' taken -> using '{1}'." -f $baseName, $uniqueName)
            $baseName = $uniqueName
        }
        [void]$taken.Add($baseName)   # reserve before creating so the next group can't reuse it

        & $Log ""
        & $Log "--- Release '$baseName'  ($($g.Rows.Count) row(s)) ---"
        try {
            $sheetId = Copy-TemplateSheet -DestFolderId $jobFolderId -NewName $baseName -Log $Log

            if ($DryRun) {
                & $Log ("  DRY RUN: {0} data row(s) would load." -f $g.Rows.Count)
                try { Rename-Sheet -SheetId $sheetId -NewName (New-SuffixedName $baseName $DryRunSuffix) } catch { }
                [void]$succeeded.Add($baseName)
                if ($script:LastReleaseUrl) { $lastGoodUrl = $script:LastReleaseUrl }
                continue
            }

            Wait-SheetReady -SheetId $sheetId -Name $baseName -Log $Log
            $idMap = Resolve-ColumnMap -SheetId $sheetId -TitleMap $TeklaColumnMap -SheetName $baseName -Log $Log
            Clear-Sheet -SheetId $sheetId -Log $Log
            $result = Add-Rows -SheetId $sheetId -Records $g.Rows -Map $idMap -Log $Log -Transforms $TeklaValueTransforms

            # Stamp the Job # summary field. Sequence(s)/Release(s) are formula
            # fields that derive themselves from this sheet's single Seq/Release.
            & $Log "  Updating sheet-summary (Job #)..."
            Set-SheetSummary -SheetId $sheetId -ValuesByTitle @{ $SummaryJobField = $Job } -Log $Log

            & $Log ("  Loaded {0} row(s) into '{1}'." -f $result.Created, $baseName)
            [void]$succeeded.Add($baseName)
            if ($script:LastReleaseUrl) { $lastGoodUrl = $script:LastReleaseUrl }
        } catch {
            $errText = "$_"
            if ($sheetId) {
                $mark = Set-SheetIncomplete -SheetId $sheetId -BaseName $baseName -Log $Log
                $errText = "$errText  >> $mark"
            }
            & $Log "  ERROR on '$baseName': $errText"
            Write-ApiLog ("    Group FAILED '{0}': {1}" -f $baseName, $errText)
            [void]$failed.Add($baseName)
        }
    }

    # Point "Open release" at the last sheet we successfully created (best effort).
    $script:LastReleaseUrl = $lastGoodUrl

    # ---- Final report ----
    & $Log ""
    & $Log ("=== Summary: {0} of {1} release sheet(s) OK. ===" -f $succeeded.Count, $groups.Count)
    if ($succeeded.Count -gt 0) { & $Log ("  OK:     " + ($succeeded -join "  |  ")) }
    if ($failed.Count -gt 0)    { & $Log ("  FAILED: " + ($failed -join "  |  ")) }

    if ($failed.Count -gt 0) {
        throw ("{0} of {1} release sheet(s) FAILED: {2}. The failed sheet(s) were marked INCOMPLETE (or flagged above if even that could not be done) - delete them and re-run. The {3} that succeeded are correct and complete." -f $failed.Count, $groups.Count, ($failed -join ", "), $succeeded.Count)
    }
    if ($DryRun) { & $Log "=== DRY RUN complete. No data loaded. Review/delete the '(DRY RUN)' sheet(s). ===" }
    else         { & $Log "=== DONE: all $($succeeded.Count) release sheet(s) are ready. ===" }
}
# =============================================================================


# ================================== GUI ======================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form
$form.Text = "Create Expediting Release  [$EnvironmentName]  v$ScriptVersion"
$form.Size = New-Object Drawing.Size(560, 540)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

function New-Label($text, $x, $y) {
    $l = New-Object Windows.Forms.Label
    $l.Text = $text; $l.Location = "$x,$y"; $l.AutoSize = $true
    $form.Controls.Add($l); return $l
}
function New-Text($x, $y, $w) {
    $t = New-Object Windows.Forms.TextBox
    $t.Location = "$x,$y"; $t.Width = $w
    $form.Controls.Add($t); return $t
}

New-Label "Job #" 20 22 | Out-Null
$tbJob = New-Text 120 20 160

New-Label "Tekla file (CSV/Excel)" 20 60 | Out-Null
$tbTekla = New-Text 160 58 260
$btnTekla = New-Object Windows.Forms.Button
$btnTekla.Text = "Browse"; $btnTekla.Location = "430,56"; $btnTekla.Width = 90
$form.Controls.Add($btnTekla)

$cbDryRun = New-Object Windows.Forms.CheckBox
$cbDryRun.Text = "Test copy only - creates a DRY RUN sheet in Smartsheet, loads NO data"
$cbDryRun.Location = "20,95"; $cbDryRun.AutoSize = $true
$form.Controls.Add($cbDryRun)

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = "Create Release"; $btnRun.Location = "20,122"; $btnRun.Width = 500; $btnRun.Height = 34
$form.Controls.Add($btnRun)

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Location = "20,165"; $lblStatus.Size = New-Object Drawing.Size(500, 26)
$lblStatus.TextAlign = "MiddleCenter"
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$lblStatus.BackColor = [Drawing.Color]::Gainsboro
$lblStatus.Text = "Idle"
$form.Controls.Add($lblStatus)

$logBox = New-Object Windows.Forms.TextBox
$logBox.Multiline = $true; $logBox.ScrollBars = "Vertical"; $logBox.ReadOnly = $true
$logBox.Location = "20,197"; $logBox.Size = New-Object Drawing.Size(500, 250)
$logBox.Font = New-Object Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = "Open release in Smartsheet"; $btnOpen.Location = "20,457"; $btnOpen.Width = 220; $btnOpen.Height = 28
$btnOpen.Enabled = $false
$btnOpen.Add_Click({
    if ($script:LastReleaseUrl) {
        try { Start-Process $script:LastReleaseUrl }
        catch { [System.Windows.Forms.MessageBox]::Show("Release URL:`r`n$script:LastReleaseUrl", "Open release") | Out-Null }
    }
})
$form.Controls.Add($btnOpen)

$btnLog = New-Object Windows.Forms.Button
$btnLog.Text = "Open Log Folder"; $btnLog.Location = "248,457"; $btnLog.Width = 150; $btnLog.Height = 28
$btnLog.Add_Click({
    try { Start-Process -FilePath explorer.exe -ArgumentList "`"$AppDataDir`"" }
    catch { [System.Windows.Forms.MessageBox]::Show("Log folder:`r`n$AppDataDir", "Open Log Folder") | Out-Null }
})
$form.Controls.Add($btnLog)

$btnCopy = New-Object Windows.Forms.Button
$btnCopy.Text = "Copy Log"; $btnCopy.Location = "404,457"; $btnCopy.Width = 116; $btnCopy.Height = 28
$btnCopy.Add_Click({
    if ($logBox.Text) { try { [System.Windows.Forms.Clipboard]::SetText($logBox.Text) } catch { } }
})
$form.Controls.Add($btnCopy)

# NOTE: this TextBox is '$logBox', NOT '$log' - the functions take a [scriptblock]$Log
# param and PS variable names are case-insensitive, so a '$log' here would shadow it.
$Logger = { param($m) $logBox.AppendText("$m`r`n"); $logBox.Refresh(); [System.Windows.Forms.Application]::DoEvents() }

function Set-Status($text, $color) {
    $lblStatus.Text = $text
    $lblStatus.BackColor = $color
    $lblStatus.Refresh()
}

# Remember the last-browsed folder across launches.
$LastDirPath = Join-Path $AppDataDir "lastdir.txt"
$script:LastDir = $null
if (Test-Path $LastDirPath) { try { $script:LastDir = (Get-Content -Path $LastDirPath -Raw -ErrorAction Stop).Trim() } catch { } }

$pick = {
    param($box, $filter)
    $d = New-Object Windows.Forms.OpenFileDialog
    $d.Filter = $filter
    if ($script:LastDir -and (Test-Path $script:LastDir)) { $d.InitialDirectory = $script:LastDir }
    if ($d.ShowDialog() -eq "OK") {
        $box.Text = $d.FileName
        try {
            $script:LastDir = Split-Path -Parent $d.FileName
            Set-Content -Path $LastDirPath -Value $script:LastDir -Encoding UTF8 -NoNewline
        } catch { }
    }
}
$dataFilter = "Data files (*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx|CSV files (*.csv)|*.csv|Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*"
$btnTekla.Add_Click({ & $pick $tbTekla $dataFilter })

$btnRun.Add_Click({
    $logBox.Clear()
    $btnOpen.Enabled = $false
    Set-Status "Working..." ([Drawing.Color]::Khaki)

    # --- config sanity (fail before any API call if IDs aren't filled in) ---
    if ((-not $TemplateSheetId) -or (-not $DestinationId)) {
        Set-Status "CONFIG ERROR" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: TemplateSheetId / DestinationId are not set. Edit the CONFIGURATION block or config.json (DestinationId must be the QMI Expediting container id)."
        return
    }
    if ($DestinationType -ne 'workspace' -and $DestinationType -ne 'folder') {
        Set-Status "CONFIG ERROR" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: DestinationType must be 'workspace' or 'folder' (currently '$DestinationType')."
        return
    }

    # --- validate inputs BEFORE any API call ---
    if ([string]::IsNullOrWhiteSpace($tbJob.Text)) {
        Set-Status "INPUT ERROR" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: Job # is required."; return
    }
    if (-not (Test-Path $tbTekla.Text)) {
        Set-Status "INPUT ERROR" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: Tekla data file not found."; return
    }

    # --- token ---
    $token = Get-ApiToken
    if (-not $token) { $token = Set-ApiToken }
    if (-not $token) {
        Set-Status "NO TOKEN" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: no API token. Cannot continue."; return
    }
    $script:Headers = @{ "Authorization" = "Bearer $token"; "Content-Type" = "application/json" }

    $btnRun.Enabled = $false
    $script:IsRunning = $true
    try {
        Invoke-Release -Job $tbJob.Text.Trim() -TeklaFile $tbTekla.Text -DryRun $cbDryRun.Checked -Log $Logger
        if ($cbDryRun.Checked) {
            Set-Status "DRY RUN OK - no data loaded" ([Drawing.Color]::LightGreen)
        } else {
            Set-Status "SUCCESS - release(s) ready" ([Drawing.Color]::LightGreen)
        }
        if ($script:LastReleaseUrl) { $btnOpen.Enabled = $true }
    } catch {
        Set-Status "FAILED - see log for details" ([Drawing.Color]::Salmon)
        & $Logger "ERROR: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "$_",
            "Release FAILED - do not use the incomplete sheet(s)",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        $script:IsRunning = $false
        $btnRun.Enabled = $true
    }
})

# Guard window-close mid-run: the DoEvents logger lets the [X] fire during a load,
# which could orphan a half-built sheet that never got marked INCOMPLETE.
$script:IsRunning = $false
$form.Add_FormClosing({
    param($eventSender, $e)
    if ($script:IsRunning) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "A release is still running. Closing now may leave an unfinished sheet in Smartsheet that is NOT marked incomplete - you would have to find and delete it by hand. Close anyway?",
            "Release in progress",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { $e.Cancel = $true }
    }
})

[void]$form.ShowDialog()
# =============================================================================
