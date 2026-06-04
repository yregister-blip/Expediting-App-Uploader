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
   sheet). A `Weight` transform also strips the trailing `#` the CSV adds
   (`14.802083#` → `14.802083`; the `.xlsx` already exports a clean number).
3. **Seq and Release are columns in the Tekla file**, so the GUI collects only the
   **Job #**. `Seq`/`Rel` for each sheet name are read from the file, and the
   `Sequence(s)`/`Release(s)` summary fields are formulas that derive themselves from
   the loaded rows.

Architecturally it copies a single **sheet** (the release artifact here is one sheet
per release inside the Job folder), not a whole folder.

## Multiple sequences / releases in one file

The upload is **split by each unique `Seq` + `Release #` combination** — one Smartsheet
sheet per combination. A file with one sequence and two releases produces two sheets
(`26-10 Seq 3 Rel 8` and `26-10 Seq 3 Rel 9`); a normal single-release file produces one.
Every data row must carry a `Seq` and a `Release #` (the tool fails before creating
anything if any row is missing either). Each sheet is built independently — if one fails
it's marked `!! INCOMPLETE - DO NOT USE !!` and reported, while the others still load, and
the run reports `X of N OK` (failing overall if any sheet failed).

## Files

| File | Purpose |
|---|---|
| [`Create-ExpeditingRelease.ps1`](Create-ExpeditingRelease.ps1) | The tool. Commented source / review artifact. |
| [`Build.ps1`](Build.ps1) | Compiles the `.ps1` → `.exe` via PS2EXE (build machine only). |
| [`config.example.json`](config.example.json) | Template for an optional `config.json` (environment IDs + label). |

## Configure

The **TEST** environment is prefilled and ready to run — everything currently lives in
the **Test Files** workspace (id `2497329237387140`):

| Setting | TEST value | Meaning |
|---|---|---|
| `TemplateSheetId` | `2860137659191172` | the `Tekla Import Template` sheet that gets copied |
| `DestinationId` | `4375630767777668` | the **QMI Expediting** folder (holds the `26-10`-style Job folders) |
| `DestinationType` | `"folder"` | QMI Expediting is a folder inside the workspace |
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
click **Create Release**. Tick **Test copy only** for a dry run (copies + names the
sheet but loads no data; the sheet is suffixed `(DRY RUN - no data loaded)`).

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
