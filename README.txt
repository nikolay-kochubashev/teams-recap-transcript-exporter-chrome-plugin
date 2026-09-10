Teams Recap Transcript Exporter v1.6.0

Purpose
-------
Exports a Microsoft Teams / SharePoint Recap transcript that is visible to you but cannot be downloaded by your account.
The extension reads the visible lazy-loaded transcript from the page and exports it as TXT.

Install extension
-----------------
1. Extract this ZIP to a permanent folder.
2. Open chrome://extensions
3. Enable Developer mode.
4. Click Load unpacked.
5. Select the folder teams-recap-transcript-exporter-v1.6.0
6. Reload the SharePoint / Teams Recap tab.

Windows Documents integration (recommended)
-------------------------------------------
Chrome extensions are still browser-sandboxed and cannot silently write to an arbitrary Windows folder by themselves.
This package therefore contains a small local Native Messaging helper.

To enable direct saving to Windows Documents:
1. Double-click Install-Windows-Integration.cmd once.
2. Return to chrome://extensions and click Reload on Teams Recap Transcript Exporter.
3. Reopen the extension side panel.

No administrator rights are required: the helper is installed for the current Windows user under LocalAppData and HKCU.

With Windows integration enabled:
- Save TXT writes directly to the real Windows Documents folder.
- The filename is: Meeting title - YYYYMMDD.txt
- Open folder launches Windows Explorer with the saved file selected.
- Save diagnostics writes a separate diagnostic TXT to Documents and does not replace the transcript shown in the panel.

Without Windows integration:
- Save TXT falls back to Chrome's normal Downloads folder.
- Open folder still shows the downloaded file in Explorer via the Chrome Downloads API.

Usage
-----
1. Open the meeting recording / Recap in Chrome.
2. Open Transcript.
3. Open the extension side panel.
4. Click Collect transcript.
5. Wait for 100%.
6. Click Save TXT.
7. Click Open folder to show the saved file in Windows Explorer.

Changes in v1.6.0
-----------------
- Fixed transcript collection in responsive/narrow layouts where Transcript moves below the video.
- Transcript scroller detection is now based primarily on transcript timestamps and semantic proximity, not right-side geometry.
- Added stronger diagnostics for candidate selection.
- Can re-detect the transcript scroller if SharePoint rebuilds the responsive DOM during collection.

Changes in v1.5.0
-----------------
- TXT filename: Meeting title - YYYYMMDD.txt
- Removed Markdown export.
- Diagnostics no longer overwrite/disappear from the transcript output. The Diagnostics button now saves a diagnostic TXT directly.
- Added optional Windows Documents integration through a local Native Messaging helper.
- Added Open folder, which selects the saved file in Windows Explorer.

Uninstall Windows integration
-----------------------------
Run Uninstall-Windows-Integration.cmd.
