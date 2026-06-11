# Expediting App Uploader (Create Expediting Release)

A small Windows desktop tool (PowerShell GUI, compilable to a single `.exe`) that
loads a **Tekla Production-Control export** into a new per-release Smartsheet in the
**QMI Expediting** workspace. For each run it copies the `Tekla Import Template`
sheet into the Job # folder, renames the copy `"<Job> Seq <Seq> Rel <Rel>"`, loads
the dropped Tekla file, stamps the `Job #` summary field, and **verifies** every row
landed — marking the sheet unusable if anything went wrong.

It talks **only** to `https://api.smartsheet.com` over HTTPS, using a Smartsheet API
token the user enters once that is stored **DPAPI-encrypted** on their own machine.
It reads one local file the user picks. It installs nothing and contacts nothing else.

This tool reuses the engine and safety conventions of the team's `Create-Release`
sample (same API choke point, retries, dry-run, fail-loud philosophy). See
[`Smartsheet-Automation-Playbook.md`](Sample%20Script%20from%20another%20project/Smartsheet-Automation-Playbook.md)
in the sample project for the shared patterns and Smartsheet gotchas.

## How it differs from the `Create-Release` sample

1. **One input file** — just the Tekla export. There is **no** optional bolt-hole list.
2. **`Main Piece` is a CHECKBOX.** The Tekla CSV emits `1`/`0` and the Excel export
   emits `TRUE`/`FALSE`; both are normalized to a real JSON **boolean** (`true`/`false`),
   which is exactly how a checkbox cell is stored (verified against an existing release
   sheet). A `Weight Each` transform also strips the trailing `#` and thousands
   comma the export's display format adds (`1,966.93#` → `1966.93`).
3. **Seq and Release are columns in the Tekla file**, so the GUI collects only the
   **Job #**. `Seq`/`Rel` for each sheet name are read from the file, and the
   `Sequence(s)`/`Release(s)` summary fields are formulas that derive themselves from
   the loaded rows.

## One row per piece (v1.2)

The file's `Qty` column is **not uploaded**. Each source row is **exploded into one
row per physical piece** — a row with `Qty` 4 becomes 4 rows — and each piece gets a
synthetic **`Instance`** number written to the Smartsheet `Instance` column (the
renamed `Qty` column). Instance numbers run `1, 2, 3…` per **`Main Mk` + `Piece Mk`
combination across the whole file**, so a combination split over several source rows
keeps counting instead of restarting; `Main Mk` + `Piece Mk` + `Instance` uniquely
identifies every piece in the release.

Weights work per piece too:

- The file's **`Weight Each`** column is what lands in the Smartsheet **`Weight`**
  column (one piece's weight per row).
- The file's **total `Weight`** column (`Qty × Weight Each`) is **validation only** and
  never uploaded: before anything is created in Smartsheet, every row is cross-checked
  (`Qty × Weight Each` must match the stated total within a small rounding tolerance)
  and the run fails loudly if the export is inconsistent. If the total column is absent
  the check is skipped with a note in the log.
- `Qty` itself must be a positive whole number on every row, or the run fails before
  touching Smartsheet.

Architecturally it copies a single **sheet** (the release artifact here is one sheet
per release inside the Job folder), not a whole folder.

## Multiple sequences / releases in one file

The upload is **split by each unique `Seq` + `Release #` combination** — one Smartsheet
sheet per combination. A file with one sequence and two releases produces two sheets
(`26-10 Seq 3 Rel 8` and `26-10 Seq 3 Rel 9`); a normal single-release file produces one.
Every data row must carry a `Seq` and a `Release #` (the tool fails before creating
anything if any row is missing either). Each sheet is built independently and the run
reports `X of N OK`.

### Auto-retry (eventual-consistency handling)

A freshly **copied** Smartsheet sheet is *eventually consistent*: for a few seconds it can
answer one API call and then return `404 Not Found` on the next, even though it exists.
On a big job (e.g. 41 sheets at once) that transient `404` used to fail individual
releases. The tool now defends against it in three layers:

1. `Wait-SheetReady` waits for **two consecutive** successful reads before loading, closing
   the flicker window.
2. `Invoke-SS` treats a `404` on a just-copied sheet as a **transient, retryable** error
   (fast 2–8 s backoff) — while a `404` anywhere else (e.g. a wrong/un-shared folder id)
   stays fatal so real misconfig still fails fast.
3. Any release that *still* fails is **auto-retried from a clean slate** up to
   `$MaxReleaseRetries` (default 2) more times: its partial sheet is deleted and the whole
   Seq/Rel group is rebuilt. Only a release that fails every attempt is reported failed
   (and left marked `!! INCOMPLETE - DO NOT USE !!` for inspection); just re-run to rebuild
   only those.

## Files

| File | Purpose |
|---|---|
| [`Create-ExpeditingRelease.ps1`](Create-ExpeditingRelease.ps1) | The tool. Commented source / review artifact. |
| [`Build.ps1`](Build.ps1) | Compiles the `.ps1` → `.exe` via PS2EXE (build machine only). |
| [`config.example.json`](config.example.json) | Template for an optional `config.json` (environment IDs + label). |

## Configure

The defaults are prefilled and ready to run — the `26-10`-style Job folders live
**directly in the QMI Expediting workspace** (there is no intermediate "QMI
Expediting" folder anymore):

| Setting | Value | Meaning |
|---|---|---|
| `TemplateSheetId` | `2860137659191172` | the `Tekla Import Template` sheet that gets copied |
| `DestinationId` | `2497329237387140` | the **QMI Expediting** workspace (holds the `26-10`-style Job folders) |
| `DestinationType` | `"workspace"` | the destination is a workspace, not a folder |
| `EnvironmentName` | `"TEST"` | shown in the title bar + log |

To move to **production** (or fix an id) without recompiling, copy
`config.example.json` → `config.json` next to the tool and set the production
`TemplateSheetId` / `DestinationId` / `DestinationType` / `EnvironmentName`. Any key you
omit keeps its built-in default; a missing or invalid `config.json` falls back to the
TEST defaults. `config.json` contains only Smartsheet object IDs and a label — **no
secrets**. To get an id, right-click the item in Smartsheet → **Properties**.

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\Create-ExpeditingRelease.ps1
```

Enter the **Job #**, browse to the **Tekla file** (`.csv`, `.xls`, or `.xlsx`), and
click **Create Release**. An elapsed-time readout next to the button counts up while a
job runs (handy for large multi-release files). Two checkboxes:

- **Test copy only** — dry run: copies + names the sheets but loads no data (each sheet
  is suffixed `(DRY RUN - no data loaded)`).
- **Re-run: skip releases that already exist** — when re-running the same file, any
  Seq/Rel whose sheet already exists in the Job folder is **skipped** instead of being
  duplicated as `…-1`. Releases that only exist as `!! INCOMPLETE !!` orphans are *not*
  skipped (their clean name is free), so a re-run rebuilds exactly the missing/failed
  ones. Leave it unchecked for a first load.

Closing the window mid-run asks for confirmation and then **cancels safely**: the run
stops at its next API call, the release being built at that moment is removed (or
marked `!! INCOMPLETE !!` if removal fails), releases already completed are kept, and
the window closes itself when cleanup finishes. Re-run the same file with **Re-run:
skip releases that already exist** to build the rest.

On first run you'll be prompted once for your Smartsheet API token (stored
DPAPI-encrypted under `%LOCALAPPDATA%\Qualico\ExpeditingUploader\`). The same folder
holds the trace log (`Expediting-Uploader-API.log`) — use **Open Log Folder** to find
it, and **Copy Log** to grab the on-screen log for a support note.

## Build (optional)

```powershell
powershell -ExecutionPolicy Bypass -File .\Build.ps1
```

Produces `Create-ExpeditingRelease.exe` (no console window). The `.exe` is **unsigned**,
so SmartScreen/Defender will warn on first run — have the MSP code-sign or whitelist it.

## Safety behavior

- Validates the file (required headers + ≥1 data row) **before** touching Smartsheet.
- Verifies created-row count == intended; a mismatch fails the run.
- On any failure the copied sheet is renamed `!! INCOMPLETE - DO NOT USE !!`, and the
  tool tells you honestly if even that rename failed (so you can delete it by hand).
- Dedups against existing release sheets in the Job folder; fails closed if it can't
  list them.
