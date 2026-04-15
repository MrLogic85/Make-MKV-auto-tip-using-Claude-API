# Plan: disc script + folder script + shared library

## Overview

Split the single disc-ripping script into three files:

```
common.ps1                    — dot-sourced by both scripts; all shared functions
Rip MKV using Claude.ps1      — disc mode (refactored, much shorter)
Rip MKV from Folder.ps1       — new: process pre-ripped BDMV folder trees
config.example.ps1            — add $defaultSourceRoots
config.ps1                    — user adds their source path(s)
```

The processing pipeline from "ripped MKV on local SSD" onwards is **identical** for both
modes. Only the source phase differs:

| Phase            | Disc script                              | Folder script                              |
|------------------|------------------------------------------|--------------------------------------------|
| Find source      | Poll `makemkvcon disc:0`                 | Scan folder for BDMV subfolders            |
| Read metadata    | Drive letter, bdmt XML, cover art        | Same files, read from folder path          |
| Rip to temp      | `makemkvcon mkv disc:0 N localTemp`      | `makemkvcon mkv file://path N localTemp`   |
| After done       | Eject, loop forever                      | Delete source folder, advance to next      |
| Shared pipeline  | identify · audio · filter · copy         | same                                       |

---

## Task 1 — Create `common.ps1`

Move all functions out of the disc script verbatim. Additionally extract four inline code
blocks into named functions so neither script has to duplicate them.

### Functions moved verbatim

| Function               | Notes                                                      |
|------------------------|------------------------------------------------------------|
| `Write-Log`            | reads `$logFile` from caller scope                         |
| `Invoke-Beep`          | reads `$beepOnManualInput` from caller scope               |
| `Invoke-Claude`        | reads `$claudeApiKey` from caller scope                    |
| `Invoke-Menu`          | pure UI                                                    |
| `Invoke-MultiSelectMenu` | pure UI                                                  |
| `New-Title`            | pure data constructor                                      |
| `Wait-CopyJob`         | **modified**: returns `$true`/`$false` for copy success    |

### New functions extracted from inline blocks

**`Invoke-ParseTitles($infoOutput)`**
- Extracts the title-parsing foreach loop (currently lines 259–308 in disc script)
- Returns `$titles` hashtable

**`Invoke-IdentifyMovie($discInfoLines, $bdmtXml, $titleLines, $discImages)`**
- Extracts the `do { … } while (-not $movieName)` Claude loop
- Also performs movie name sanitization (replace `:`, `"`, `\/*?<>|`)
- Returns `@{ Name = …; Edition = …; TitleNum = … }` (`TitleNum` is `$null` if Claude
  returned UNKNOWN)

**`Invoke-SelectAudioTracks($tempMkv, $movieName)`**
- Extracts Step 2: mkvmerge identify → Claude → manual fallback
- Returns `@{ KeepIds = …; AudioTracks = … }` so the caller can build the audio summary

**`Invoke-FilterAudio($tempMkv, $keepIds, $movieName)`**
- Extracts Step 3: mkvmerge filter + aspect ratio check
- Returns path to `$localFinalMkv`; calls `Wait-CopyJob` + `exit` on fatal errors (same
  as today)

**`Start-DestinationCopy($localFinalMkv, $movieName, $movieEdition, $destRoot)`**
- Extracts Step 4: create folder, write NFO, start background copy job
- Derives `$movieFolder` and `$finalMkv` internally; returns `$finalMkv` for logging

`common.ps1` contains **only function definitions** — no top-level executable code.
All config vars (`$claudeApiKey`, `$beepOnManualInput`, `$preferredAudioLanguages`,
`$mkvmerge`, `$mkvpropedit`, `$localTemp`, `$logFile`) are read from the caller's scope
via normal PowerShell dot-source scoping.

---

## Task 2 — Refactor `Rip MKV using Claude.ps1`

After moving functions to `common.ps1` the disc script becomes:

```
[encoding setup]
[load config.ps1]
[init $logFile]
. (Join-Path $PSScriptRoot "common.ps1")
[verify tools]
[destination selection — unchanged]

$lastDiscName = $null
$copyJob      = $null

while ($true) {
    [wait for disc — unchanged]
    [get $driveLetter, $volumeLabel, $bdmtXml, $discImages — unchanged]

    $titles          = Invoke-ParseTitles $infoOutput
    [build $titleLines + log — unchanged]

    $id              = Invoke-IdentifyMovie $discInfoLines $bdmtXml $titleLines $discImages
    $movieName       = $id.Name
    $movieEdition    = $id.Edition

    [compute $movieFolder/$finalMkv, check if exists — unchanged]

    if ($id.TitleNum -ne $null) {
        $chosenTitle = $id.TitleNum
    } else {
        [manual title menu — unchanged]
    }

    [rip loop with makemkvcon disc:0 — unchanged]

    $audioResult   = Invoke-SelectAudioTracks $tempMkv $movieName
    $localFinalMkv = Invoke-FilterAudio $tempMkv $audioResult.KeepIds $movieName

    Write-Log "Step 4: Starting copy to destination in background..."
    Wait-CopyJob
    $finalMkv = Start-DestinationCopy $localFinalMkv $movieName $movieEdition $destRoot

    [audio summary + DONE log — unchanged]
    $lastDiscName = $discName
    [eject — unchanged]
}
```

No behaviour changes — pure extraction.

---

## Task 3 — Create `Rip MKV from Folder.ps1`

```
[encoding setup]
[load config.ps1]
[init $logFile]
. (Join-Path $PSScriptRoot "common.ps1")
[verify tools]
[source folder selection — mirrors destination, uses $defaultSourceRoots]
[destination selection — same as disc script]

$copyJob            = $null
$lastCopySucceeded  = $false
$lastSourceFolder   = $null

$bdmvFolders = Get-ChildItem -Path $sourceRoot -Directory |
    Where-Object { Test-Path (Join-Path $_.FullName "BDMV") } |
    Sort-Object Name

Write-Log "Found $($bdmvFolders.Count) BDMV folder(s) in $sourceRoot"

foreach ($folder in $bdmvFolders) {

    # Wait for previous copy; delete that source folder if it succeeded
    $lastCopySucceeded = Wait-CopyJob
    if ($lastSourceFolder -and $lastCopySucceeded) {
        Remove-Item -Path $lastSourceFolder -Recurse -Force
        Write-Log "Deleted source folder: $lastSourceFolder"
    }

    $filePath   = $folder.FullName
    $infoOutput = & $makemkvcon -r info "file://$filePath" 2>&1

    [read $bdmtXml from $filePath\BDMV\META\DL\bdmt_eng.xml]
    [read $discImages from same directory]

    $discInfoLines = @("Folder name: $($folder.Name)")

    $titles    = Invoke-ParseTitles $infoOutput
    [build $titleLines + log]
    $id        = Invoke-IdentifyMovie $discInfoLines $bdmtXml $titleLines $discImages
    $movieName = $id.Name;  $movieEdition = $id.Edition

    [compute $movieFolder/$finalMkv, check if exists]

    if ($id.TitleNum -ne $null) {
        $chosenTitle = $id.TitleNum
    } else {
        [manual title menu]
    }

    [rip: makemkvcon mkv "file://$filePath" $chosenTitle $localTemp]
    [rename generated MKV to temp name — same logic as disc script]

    $audioResult   = Invoke-SelectAudioTracks $tempMkv $movieName
    $localFinalMkv = Invoke-FilterAudio $tempMkv $audioResult.KeepIds $movieName
    $finalMkv      = Start-DestinationCopy $localFinalMkv $movieName $movieEdition $destRoot

    [audio summary + DONE log]

    $lastSourceFolder = $filePath
}

# Clean up the last folder once its copy job finishes
$lastCopySucceeded = Wait-CopyJob
if ($lastSourceFolder -and $lastCopySucceeded) {
    Remove-Item -Path $lastSourceFolder -Recurse -Force
    Write-Log "Deleted source folder: $lastSourceFolder"
}
```

Key design decisions:
- **Background copy preserved** — title N+1 rips while title N copies (same overlap
  benefit as disc-swap in the disc script)
- **Delete only on confirmed success** — `Wait-CopyJob` returning `$false` leaves the
  source folder intact
- **No polling loop** — all folders are known upfront; no waiting for media

---

## Task 4 — Update `config.example.ps1`

Add one new variable:

```powershell
# Folder script: root folder(s) containing BDMV subfolders to process
$defaultSourceRoots = @("D:\BDMVRips")
```

---

## Task 5 — Update `README.md`

- Add a section for `Rip MKV from Folder.ps1` with its workflow
- Mention `common.ps1`
- Update Setup to cover `$defaultSourceRoots`

---

## What stays unchanged

- All Claude prompts and response parsing
- All MakeMKV and MKVToolNix command invocations
- All menu UI
- The rip-retry / skip flow on disc read failure
- The aspect ratio check
- The NFO format
- Commit style throughout
