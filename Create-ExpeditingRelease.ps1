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
        Excel files are read via Excel COM). Each source row is EXPLODED into
        one row PER PIECE (a row with Qty 4 becomes 4 rows), each stamped with
        an "Instance" number so every physical piece of a Main Mk + Piece Mk
        combination is uniquely identifiable. The per-piece "Weight Each" is
        what lands in the sheet's Weight column; the file's TOTAL Weight column
        is used only to cross-check Qty x Weight Each before anything uploads.
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
      (A "Weight Each" transform also strips the trailing "#" the CSV adds,
       e.g. "14.802083#" -> 14.802083; the .xlsx already exports a clean number.)
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
$ScriptVersion = "1.2.0"

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

# Release-level auto-retry. A Smartsheet sheet COPY is eventually consistent: a
# freshly-copied sheet can briefly return 404 ("Not Found") on a follow-up call
# even though it exists (it answers the next call fine). The per-call layer
# (Invoke-SS + Wait-SheetReady) absorbs most of that; this is the safety net for
# any release that still fails - the whole Seq/Rel group is re-attempted from a
# clean slate up to this many extra times, after a short settle.
$MaxReleaseRetries  = 2
$RetrySettleSeconds = 8

# ----------------------------- COLUMN MAPPINGS -------------------------------
# Map by HEADER NAME in the Tekla file -> Smartsheet column TITLE. Only columns
# listed here are written; everything else is left for Smartsheet to compute.
# Blank source values are skipped. Titles resolve to live column ids at runtime
# (a copied sheet gets fresh column ids - see Resolve-ColumnMap).
#
# PER-PIECE LOADING (v1.2): the file's "Qty" column is NOT uploaded. Instead
# each source row is exploded into Qty rows, and a synthetic "Instance" value
# (1,2,3... per Main Mk + Piece Mk combination) is written to the Smartsheet
# "Instance" column (the renamed Qty column). The per-piece "Weight Each" from
# the file lands in the Smartsheet "Weight" column; the file's TOTAL "Weight"
# column is validation-only and is never uploaded (see Assert-QtyAndWeight).
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
    "Instance"     = "Instance"
    "Length"       = "Length"
    "Weight Each"  = "Weight"
    "Pay Category" = "Pay Category"
    "Remarks"      = "Remarks"
}

# Source headers used to derive the Seq / Release values for the sheet NAME.
$SeqHeader     = "Seq"
$ReleaseHeader = "Release #"

# Headers that drive the per-piece explosion. "Instance" is SYNTHETIC (added by
# Expand-RowsByQty, not present in the file), so source-header validation checks
# every $TeklaColumnMap key EXCEPT it, plus the Qty column that gets exploded.
$QtyHeader         = "Qty"           # exploded into one row per piece; not uploaded
$InstanceField     = "Instance"      # synthetic per-piece counter
$MainMarkHeader    = "Main Mk"       # with Piece Mk: the combination Instance numbers
$PieceMarkHeader   = "Piece Mk"
$WeightEachHeader  = "Weight Each"   # per-piece weight; uploaded as "Weight"
$TotalWeightHeader = "Weight"        # row total (Qty x each); VALIDATION ONLY

$TeklaSourceRequired = @{}
foreach ($h in $TeklaColumnMap.Keys) { if ($h -ne $InstanceField) { $TeklaSourceRequired[$h] = $TeklaColumnMap[$h] } }
$TeklaSourceRequired[$QtyHeader] = $QtyHeader

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
    # The Tekla CSV appends "#" to weight fields (e.g. "14.802083#"); strip it so
    # the value lands as a number. (The .xlsx already exports a clean number, so
    # this is a no-op there.)
    "Weight Each" = { param($v) (([string]$v) -replace '#', '').Trim() }
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

# When TRUE, Invoke-SS treats a 404 ("Not Found") as a TRANSIENT, retryable error
# instead of a fatal one. We flip this on ONLY while operating on a sheet we just
# copied (which is eventually consistent and can flicker 404 for a few seconds);
# everywhere else a 404 stays fatal so a genuine misconfig (wrong/!shared folder
# id) still fails fast.
$script:RetryNotFound = $false

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
    # A 404 we choose to retry is post-copy eventual-consistency flicker, which
    # clears in a second or two - retry FAST (2s, 4s, 6s...) rather than the
    # heavier 5xx/429 backoff below.
    if ($Status -eq 404) { return [math]::Min(2 * $Try, 8) }
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
            $retry404 = ($status -eq 404 -and $script:RetryNotFound)
            if ($status -eq 429 -or $status -ge 500 -or $retry404) {
                $delay = Get-RetryDelaySeconds -Status $status -Response $_.Exception.Response -Try $try
                if     ($status -eq 429) { Write-ApiLog ("    429 rate-limited; sleeping {0}s before retry" -f $delay) }
                elseif ($retry404)       { Write-ApiLog ("    404 on a just-copied sheet (eventual consistency); sleeping {0}s before retry" -f $delay) }
                else                     { Write-ApiLog ("    5xx server error; sleeping {0}s before retry" -f $delay) }
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

# A sheet copy is ASYNCHRONOUS and EVENTUALLY CONSISTENT: not only can the first
# GET 404, but the sheet can answer one call and then 404 the NEXT for a few more
# seconds (different backend node). So we don't just wait for the FIRST success -
# we wait until the sheet answers $NeedConsecutive times IN A ROW, which closes
# the flicker window before the load steps (resolve/clear/add) run against it.
# (Belt-and-suspenders with Invoke-SS's $script:RetryNotFound 404-retry.)
# URL NOTE: wrap the id as $($SheetId) before "?" - PS treats "?" as a legal
# variable-name char, so "$SheetId?page=1" expands to EMPTY -> a permanent 404.
function Wait-SheetReady {
    param([long]$SheetId, [string]$Name, [scriptblock]$Log, [int]$TimeoutSec = 300, [int]$NeedConsecutive = 2)
    $start    = Get-Date
    $deadline = $start.AddSeconds($TimeoutSec)
    $delay    = 2
    $okStreak = 0
    while ($true) {
        try {
            Invoke-SS -Method Get -Url "https://api.smartsheet.com/2.0/sheets/$($SheetId)?page=1&pageSize=1" -MaxTries 1 | Out-Null
            $okStreak++
            if ($okStreak -ge $NeedConsecutive) {
                $elapsed = [int]((Get-Date) - $start).TotalSeconds
                if ($elapsed -gt 0) { & $Log "  '$Name' ready after $elapsed s." }
                return
            }
            # Got one good read; pause briefly and confirm it stays available.
            Start-Sleep -Seconds 1
        } catch {
            $okStreak = 0   # a 404 mid-streak means it's still flickering; start over
            $status = 0
            if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 } }
            if ($status -ne 404) { throw }
            if ((Get-Date) -ge $deadline) {
                throw "Sheet '$Name' (id $SheetId) was still not consistently available after $TimeoutSec s (post-copy provisioning)."
            }
            $elapsed = [int]((Get-Date) - $start).TotalSeconds
            & $Log "  Waiting for '$Name' to finish provisioning... ($elapsed s)"
            Start-Sleep -Seconds $delay
            if ($delay -lt 10) { $delay += 1 }
        }
    }
}

# Best-effort DELETE of a sheet. Used to remove a partial sheet before re-trying
# its release (so retries start clean and no orphan accumulates). Returns $true if
# the sheet is gone. A 404 here means it's already gone -> also success.
function Remove-SheetSafe {
    param([long]$SheetId, [scriptblock]$Log)
    try {
        Invoke-SS -Method Delete -Url "https://api.smartsheet.com/2.0/sheets/$SheetId" | Out-Null
        return $true
    } catch {
        $status = 0
        if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 } }
        if ($status -eq 404) { return $true }
        Write-ApiLog ("    Remove-SheetSafe FAILED for sheet id {0}: {1}" -f $SheetId, $_)
        return $false
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

# Parse a Tekla weight cell to a number. The CSV variant appends "#"; strip it.
# Returns $null if the cell is blank or not numeric.
function ConvertTo-WeightNumber {
    param($Value)
    $s = (([string]$Value) -replace '#', '').Trim()
    if ($s -eq '') { return $null }
    $d = 0.0
    $ok = [double]::TryParse($s, [System.Globalization.NumberStyles]::Float,
                             [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)
    if ($ok) { return $d }
    return $null
}

# Every row must carry a positive WHOLE-NUMBER Qty - it drives the per-piece row
# explosion. If the file also carries the TOTAL weight column, cross-check
# Qty x "Weight Each" against it on every row (small tolerance for Tekla's own
# rounding) so a stale or hand-edited export fails loudly BEFORE anything is
# created in Smartsheet. The total column itself is never uploaded.
function Assert-QtyAndWeight {
    param($Records, [scriptblock]$Log)
    $rows = @($Records)
    $hasTotal = (@($rows[0].PSObject.Properties.Name) -contains $TotalWeightHeader)
    if (-not $hasTotal) {
        & $Log "  NOTE: file has no '$TotalWeightHeader' (total) column - weight cross-check skipped."
    }

    $badQty   = New-Object System.Collections.Generic.List[string]
    $badWt    = New-Object System.Collections.Generic.List[string]
    $nBadQty  = 0
    $nBadWt   = 0
    $sumEach  = 0.0   # file-wide Qty x each
    $sumTotal = 0.0   # file-wide stated totals
    $fileRow  = 1     # header is row 1; data starts on row 2
    foreach ($rec in $rows) {
        $fileRow++
        $qtyRaw = ([string]$rec.$QtyHeader).Trim()
        $qty = 0
        if (-not [int]::TryParse($qtyRaw, [ref]$qty) -or $qty -lt 1) {
            $nBadQty++
            if ($badQty.Count -lt 5) { [void]$badQty.Add("file row ${fileRow}: Qty='$qtyRaw'") }
            continue
        }
        if (-not $hasTotal) { continue }
        $each  = ConvertTo-WeightNumber $rec.$WeightEachHeader
        $total = ConvertTo-WeightNumber $rec.$TotalWeightHeader
        if (($null -eq $each) -or ($null -eq $total)) {
            $nBadWt++
            if ($badWt.Count -lt 5) { [void]$badWt.Add("file row ${fileRow}: '$WeightEachHeader'='$($rec.$WeightEachHeader)' / '$TotalWeightHeader'='$($rec.$TotalWeightHeader)' (not numeric)") }
            continue
        }
        $sumEach  += ($qty * $each)
        $sumTotal += $total
        $tol = [math]::Max(0.05, [math]::Abs($total) * 0.001)
        if ([math]::Abs(($qty * $each) - $total) -gt $tol) {
            $nBadWt++
            if ($badWt.Count -lt 5) { [void]$badWt.Add(("file row {0}: Qty {1} x {2} = {3:0.###} but '{4}' says {5:0.###}" -f $fileRow, $qty, $each, ($qty * $each), $TotalWeightHeader, $total)) }
        }
    }

    if ($nBadQty -gt 0) {
        throw "$nBadQty of $($rows.Count) data row(s) do not have a positive whole-number '$QtyHeader' (e.g. $($badQty -join '; ')). Every row needs one so it can be split into per-piece rows. Fix the export and re-run."
    }
    if ($nBadWt -gt 0) {
        throw "$nBadWt of $($rows.Count) data row(s) FAILED the weight cross-check (Qty x '$WeightEachHeader' should equal '$TotalWeightHeader'). First mismatches: $($badWt -join '; '). The export looks inconsistent - re-export from Tekla and re-run. Nothing was created in Smartsheet."
    }
    if ($hasTotal) {
        & $Log ("  Weight check OK: Qty x each = {0:0.##} matches stated totals = {1:0.##} on all {2} row(s)." -f $sumEach, $sumTotal, $rows.Count)
    }
}

# Explode each source row into one row PER PIECE: a row with Qty 4 becomes 4
# rows, each carrying a synthetic "Instance" number so every physical piece of a
# Main Mk + Piece Mk combination is uniquely identifiable. Numbering runs per
# (Main Mk, Piece Mk) across the WHOLE file, so a combination split over several
# source rows keeps counting (1,2,3,4...) instead of restarting at 1.
function Expand-RowsByQty {
    param($Records)
    $counters = @{}
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($rec in @($Records)) {
        $qty = [int](([string]$rec.$QtyHeader).Trim())   # validated by Assert-QtyAndWeight
        $key = (([string]$rec.$MainMarkHeader).Trim().ToLowerInvariant()) + '||' + (([string]$rec.$PieceMarkHeader).Trim().ToLowerInvariant())
        if (-not $counters.ContainsKey($key)) { $counters[$key] = 0 }
        for ($i = 1; $i -le $qty; $i++) {
            $counters[$key]++
            $obj = [ordered]@{}
            foreach ($p in $rec.PSObject.Properties) { $obj[$p.Name] = $p.Value }
            $obj[$InstanceField] = [string]$counters[$key]
            [void]$out.Add([pscustomobject]$obj)
        }
    }
    return $out
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
# Build ONE release sheet for a single Seq/Rel group: copy the template, load the
# group's rows, stamp the Job # summary. Returns a result hashtable instead of
# throwing, so the caller can decide whether to clean up + retry. The new sheet's
# id is always returned (even on failure) so a partial copy can be removed.
function Invoke-ReleaseGroup {
    param([string]$Job, $Group, [string]$BaseName, [long]$JobFolderId, [bool]$DryRun, [scriptblock]$Log)
    $sheetId = $null
    try {
        $sheetId = Copy-TemplateSheet -DestFolderId $JobFolderId -NewName $BaseName -Log $Log
        $url = $script:LastReleaseUrl

        if ($DryRun) {
            & $Log ("  DRY RUN: {0} data row(s) would load." -f $Group.Rows.Count)
            try { Rename-Sheet -SheetId $sheetId -NewName (New-SuffixedName $BaseName $DryRunSuffix) } catch { }
            return @{ Ok = $true; SheetId = $sheetId; Created = 0; Url = $url }
        }

        Wait-SheetReady -SheetId $sheetId -Name $BaseName -Log $Log
        $idMap = Resolve-ColumnMap -SheetId $sheetId -TitleMap $TeklaColumnMap -SheetName $BaseName -Log $Log
        Clear-Sheet -SheetId $sheetId -Log $Log
        $result = Add-Rows -SheetId $sheetId -Records $Group.Rows -Map $idMap -Log $Log -Transforms $TeklaValueTransforms

        # Stamp the Job # summary field. Sequence(s)/Release(s) are formula fields
        # that derive themselves from this sheet's single Seq/Release.
        & $Log "  Updating sheet-summary (Job #)..."
        Set-SheetSummary -SheetId $sheetId -ValuesByTitle @{ $SummaryJobField = $Job } -Log $Log

        & $Log ("  Loaded {0} row(s) into '{1}'." -f $result.Created, $BaseName)
        return @{ Ok = $true; SheetId = $sheetId; Created = $result.Created; Url = $url }
    } catch {
        return @{ Ok = $false; SheetId = $sheetId; Error = "$_" }
    }
}

# Loads the Tekla file, SPLIT into one release sheet per unique Seq+Release pair
# (e.g. one sequence with two releases -> two sheets). Each sheet is built
# independently; any that fail are AUTO-RETRIED from a clean slate (a just-copied
# sheet is eventually consistent and can transiently 404 - see $MaxReleaseRetries).
# The run as a whole fails (throws) only if a release is STILL failing after all
# retries, so a partial result is never silently reported as success.
function Invoke-Release {
    param([string]$Job, [string]$TeklaFile, [bool]$DryRun, [bool]$SkipExisting, [scriptblock]$Log)

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
    Assert-SourceData -Records $teklaRows -Map $TeklaSourceRequired -FileLabel "Tekla file"
    Assert-SeqRelPresent -Records $teklaRows -SeqHeader $SeqHeader -RelHeader $ReleaseHeader
    Assert-QtyAndWeight -Records $teklaRows -Log $Log
    & $Log ("  Tekla file OK: {0} data row(s)." -f @($teklaRows).Count)

    # One row PER PIECE: explode Qty and stamp each piece's Instance number.
    $pieceRows = @(Expand-RowsByQty -Records $teklaRows)
    & $Log ("  Split {0} source row(s) into {1} per-piece row(s) (one per Qty)." -f @($teklaRows).Count, $pieceRows.Count)
    Write-ApiLog ("    Explode: {0} source rows -> {1} piece rows." -f @($teklaRows).Count, $pieceRows.Count)

    # Split into one sheet per unique Seq+Release pair (they live IN the file).
    $groups = @(Group-BySeqRel -Records $pieceRows -SeqHeader $SeqHeader -RelHeader $ReleaseHeader)
    & $Log ("File contains {0} unique Sequence/Release combination(s) -> {0} release sheet(s):" -f $groups.Count)
    foreach ($g in $groups) { & $Log ("    - Seq {0} Rel {1}  ({2} row(s))" -f $g.Seq, $g.Rel, $g.Rows.Count) }
    Write-ApiLog ("    Split: {0} Seq/Rel group(s)." -f $groups.Count)

    # Ensure the Job # folder exists (create if missing); releases land inside it.
    $jobFolderId = Get-OrCreateJobFolder -JobName $Job -Log $Log

    # Existing sheet names in the Job folder, so we don't create a duplicate. Also
    # reserve each name we assign this run so two groups can't collide.
    $existingNames = @(Get-JobSheetNames -FolderId $jobFolderId -Log $Log)
    $existingLower = @{}
    foreach ($n in $existingNames) { if ($n) { $existingLower[$n.ToLowerInvariant()] = $true } }
    $taken = New-Object System.Collections.Generic.List[string]
    foreach ($n in $existingNames) { [void]$taken.Add($n) }

    if ($SkipExisting) { & $Log "Re-run mode ON: releases that already have a sheet in this Job folder are SKIPPED (not duplicated)." }

    # Assign each group its final sheet name once, up front. In re-run mode, a
    # group whose (clean) name already exists is SKIPPED rather than duplicated;
    # otherwise a name clash is disambiguated with a -1/-2 suffix.
    $work    = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[string]
    foreach ($g in $groups) {
        $baseName = New-SuffixedName "$Job Seq $($g.Seq) Rel $($g.Rel)" ""
        if ($SkipExisting -and $existingLower.ContainsKey($baseName.ToLowerInvariant())) {
            & $Log "  Skipping '$baseName' (already exists)."
            Write-ApiLog ("    Re-run: skip '{0}' (already exists)." -f $baseName)
            [void]$skipped.Add($baseName)
            continue
        }
        $uniqueName = Get-UniqueName -BaseName $baseName -Existing $taken.ToArray()
        if ($uniqueName -ne $baseName) {
            & $Log "A release sheet named '$baseName' already exists; using '$uniqueName' instead."
            Write-ApiLog ("    Dedup: '{0}' taken -> using '{1}'." -f $baseName, $uniqueName)
            $baseName = $uniqueName
        }
        [void]$taken.Add($baseName)
        [void]$work.Add(@{ Group = $g; Name = $baseName })
    }
    if ($skipped.Count -gt 0) { & $Log ("Re-run: {0} existing release(s) skipped; {1} to build." -f $skipped.Count, $work.Count) }
    if ($work.Count -eq 0) {
        & $Log "Nothing to build - every release in the file already exists in the Job folder."
        return
    }

    $succeeded   = New-Object System.Collections.Generic.List[string]
    $lastGoodUrl = $null
    $pending     = $work          # items still to do (all, then only failures)

    # From here on we only touch freshly-copied sheets, which are eventually
    # consistent - tell Invoke-SS to treat their transient 404s as retryable.
    $script:RetryNotFound = $true
    try {
        for ($pass = 0; $pass -le $MaxReleaseRetries; $pass++) {
            if ($pending.Count -eq 0) { break }
            if ($pass -gt 0) {
                & $Log ""
                & $Log ("=== Retry pass {0} of {1}: re-attempting {2} release(s) that failed, after a {3}s settle... ===" -f $pass, $MaxReleaseRetries, $pending.Count, $RetrySettleSeconds)
                Write-ApiLog ("    Retry pass {0}: {1} pending." -f $pass, $pending.Count)
                Start-Sleep -Seconds $RetrySettleSeconds
            }
            $isFinalPass = ($pass -ge $MaxReleaseRetries)
            $stillPending = New-Object System.Collections.Generic.List[object]

            foreach ($item in $pending) {
                $g = $item.Group; $name = $item.Name
                & $Log ""
                $tag = if ($pass -gt 0) { "  [retry $pass]" } else { "" }
                & $Log "--- Release '$name'  ($($g.Rows.Count) row(s))$tag ---"

                $res = Invoke-ReleaseGroup -Job $Job -Group $g -BaseName $name -JobFolderId $jobFolderId -DryRun $DryRun -Log $Log

                if ($res.Ok) {
                    [void]$succeeded.Add($name)
                    if ($res.Url) { $lastGoodUrl = $res.Url }
                    continue
                }

                & $Log "  ERROR on '$name': $($res.Error)"
                Write-ApiLog ("    Group FAILED '{0}': {1}" -f $name, $res.Error)
                if (-not $isFinalPass) {
                    # Will retry: delete the partial sheet so the next attempt is
                    # clean and no orphan piles up. If we can't delete it, fall back
                    # to marking it incomplete so it's never mistaken for good.
                    if ($res.SheetId) {
                        if (Remove-SheetSafe -SheetId $res.SheetId -Log $Log) {
                            & $Log "  Removed the partial sheet; will retry it."
                        } else {
                            $mark = Set-SheetIncomplete -SheetId $res.SheetId -BaseName $name -Log $Log
                            & $Log "  Could not remove the partial sheet. $mark"
                        }
                    }
                } else {
                    # Out of retries: leave a clearly-marked artifact for inspection.
                    if ($res.SheetId) {
                        $mark = Set-SheetIncomplete -SheetId $res.SheetId -BaseName $name -Log $Log
                        & $Log "  $mark"
                    }
                }
                [void]$stillPending.Add($item)
            }
            $pending = $stillPending
        }
    } finally {
        $script:RetryNotFound = $false
    }

    # Point "Open release" at the last sheet we successfully created (best effort).
    $script:LastReleaseUrl = $lastGoodUrl

    # ---- Final report ----
    $failedNames = @($pending | ForEach-Object { $_.Name })
    & $Log ""
    & $Log ("=== Summary: {0} of {1} release sheet(s) OK. ===" -f $succeeded.Count, $work.Count)
    if ($succeeded.Count -gt 0)   { & $Log ("  OK:      " + ($succeeded -join "  |  ")) }
    if ($failedNames.Count -gt 0) { & $Log ("  FAILED:  " + ($failedNames -join "  |  ")) }
    if ($skipped.Count -gt 0)     { & $Log ("  SKIPPED ({0}, already existed): {1}" -f $skipped.Count, ($skipped -join "  |  ")) }

    if ($failedNames.Count -gt 0) {
        throw ("{0} of {1} release sheet(s) still FAILED after {2} retr{3}: {4}. Their partial sheets were removed (or marked INCOMPLETE if removal failed) - just re-run to rebuild only those. The {5} that succeeded are correct and complete." -f $failedNames.Count, $work.Count, $MaxReleaseRetries, $(if ($MaxReleaseRetries -eq 1) { 'y' } else { 'ies' }), ($failedNames -join ", "), $succeeded.Count)
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
$form.Size = New-Object Drawing.Size(560, 580)
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
$cbDryRun.Location = "20,90"; $cbDryRun.AutoSize = $true
$form.Controls.Add($cbDryRun)

# Re-run mode: skip Seq/Rel releases that already have a sheet in the Job folder
# (instead of creating a "-1" duplicate). Lets you re-run the same file to fill in
# only the releases that are missing.
$cbRerun = New-Object Windows.Forms.CheckBox
$cbRerun.Text = "Re-run: skip releases that already exist in this Job folder"
$cbRerun.Location = "20,112"; $cbRerun.AutoSize = $true
$form.Controls.Add($cbRerun)

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = "Create Release"; $btnRun.Location = "20,140"; $btnRun.Width = 380; $btnRun.Height = 34
$form.Controls.Add($btnRun)

# Live elapsed-time readout. Updates as the on-screen log advances (the $Logger's
# DoEvents pump drives the timer ticks during a run).
$lblTimer = New-Object Windows.Forms.Label
$lblTimer.Location = "410,140"; $lblTimer.Size = New-Object Drawing.Size(110, 34)
$lblTimer.TextAlign = "MiddleCenter"
$lblTimer.BorderStyle = "FixedSingle"
$lblTimer.Font = New-Object Drawing.Font("Consolas", 11, [Drawing.FontStyle]::Bold)
$lblTimer.Text = "00:00"
$form.Controls.Add($lblTimer)

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Location = "20,184"; $lblStatus.Size = New-Object Drawing.Size(500, 26)
$lblStatus.TextAlign = "MiddleCenter"
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$lblStatus.BackColor = [Drawing.Color]::Gainsboro
$lblStatus.Text = "Idle"
$form.Controls.Add($lblStatus)

$logBox = New-Object Windows.Forms.TextBox
$logBox.Multiline = $true; $logBox.ScrollBars = "Vertical"; $logBox.ReadOnly = $true
$logBox.Location = "20,216"; $logBox.Size = New-Object Drawing.Size(500, 250)
$logBox.Font = New-Object Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = "Open release in Smartsheet"; $btnOpen.Location = "20,476"; $btnOpen.Width = 220; $btnOpen.Height = 28
$btnOpen.Enabled = $false
$btnOpen.Add_Click({
    if ($script:LastReleaseUrl) {
        try { Start-Process $script:LastReleaseUrl }
        catch { [System.Windows.Forms.MessageBox]::Show("Release URL:`r`n$script:LastReleaseUrl", "Open release") | Out-Null }
    }
})
$form.Controls.Add($btnOpen)

$btnLog = New-Object Windows.Forms.Button
$btnLog.Text = "Open Log Folder"; $btnLog.Location = "248,476"; $btnLog.Width = 150; $btnLog.Height = 28
$btnLog.Add_Click({
    try { Start-Process -FilePath explorer.exe -ArgumentList "`"$AppDataDir`"" }
    catch { [System.Windows.Forms.MessageBox]::Show("Log folder:`r`n$AppDataDir", "Open Log Folder") | Out-Null }
})
$form.Controls.Add($btnLog)

$btnCopy = New-Object Windows.Forms.Button
$btnCopy.Text = "Copy Log"; $btnCopy.Location = "404,476"; $btnCopy.Width = 116; $btnCopy.Height = 28
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

# Elapsed-time clock. A WinForms timer fires its Tick via the message loop, which
# the $Logger's DoEvents() (called on every log line) pumps during a run - so the
# readout advances roughly each log line. $RunStart is set when a run begins.
$script:RunStart = $null
$runTimer = New-Object Windows.Forms.Timer
$runTimer.Interval = 1000
function Update-Elapsed {
    if ($script:RunStart) {
        $el = (Get-Date) - $script:RunStart
        $lblTimer.Text = ("{0:00}:{1:00}" -f [int][math]::Floor($el.TotalMinutes), $el.Seconds)
    }
}
$runTimer.Add_Tick({ Update-Elapsed })

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
    $lblTimer.Text = "00:00"
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
    $script:RunStart = Get-Date
    $runTimer.Start()
    try {
        Invoke-Release -Job $tbJob.Text.Trim() -TeklaFile $tbTekla.Text -DryRun $cbDryRun.Checked -SkipExisting $cbRerun.Checked -Log $Logger
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
        $runTimer.Stop()
        Update-Elapsed          # stamp the final total
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
