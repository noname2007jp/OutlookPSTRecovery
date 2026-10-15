# Outlook PST Recovery

Windows 11 / classic Outlook / POP accounts.

## Usage
1. Launch `OutlookPSTRecovery.exe`.
2. Click **1. 解析** to inspect common PST locations and display file path, size, modified time, OneDrive status, Outlook process state, and Outlook profile names.
3. Select the intended PST row.
4. Close classic Outlook normally, then click **2. 解析結果から修復**.
5. The program copies the selected PST into a local folder and verifies the copy with SHA-256.
6. Optionally open Outlook Mail setup to inspect the Data Files tab and POP delivery location.

## Safety
- Source PSTs are never deleted or overwritten.
- Existing recovery copies are not overwritten.
- Outlook must be closed before copy/repair.
- The program does not automatically modify Outlook profiles or POP delivery settings. Verify the correct PST and delivery location manually.
- Files marked as offline/cloud placeholders are not copied.
- A verified copy does not guarantee that the PST itself is healthy.

## Local output
- Copies: `%LOCALAPPDATA%\OutlookPSTRecovery\Restored`
- Logs: `%LOCALAPPDATA%\OutlookPSTRecovery\Logs`

## Build
The GitHub Actions workflow **Build Outlook PST Recovery** publishes a self-contained, single-file Windows x64 EXE. Open the repository's Actions tab, select the workflow run, and download the `OutlookPSTRecovery-win-x64` artifact.

## Scope
The current version scans common Outlook/OneDrive folders and does not search every drive or cloud-only online location. It is a safe copy-and-verify recovery assistant, not a full automatic Outlook profile editor.
