# Teams Admin Check Utility by complicatiion

Menu-driven PowerShell utility for Windows 10/11 that checks Microsoft Teams installation state, versions, cache sizes, possible account identities, running processes, and WebView2 runtime status.

## Included files

- `Teams_Admin_Check_Utility.ps1` - main utility
- `Launch_Teams_Admin_Check_Utility.bat` - simple launcher for double-click use

## Main functions

- Quick Teams health summary
- Installation and version inventory
- Cache size analysis
- Best-effort identity snapshot
- Teams process and WebView2 checks
- Full report export to Desktop `TeamsAdminReports` (Desktop\TeamsAdminReports)
- Teams update trigger
- Teams uninstall
- Teams install or reinstall
- Teams cache cleanup/reset

## Notes

- Run the launcher as administrator for update, uninstall, reinstall, and reset actions.
- If `teamsbootstrapper.exe` is placed next to the script, the utility prefers it for install and update.
- If `MSTeams-x64.msix` is also placed next to the script, the utility uses it as an offline install source.
- Without the bootstrapper, the utility falls back to WinGet when available.
- Account detection is best-effort and uses local Office identity data and device registration information when available.

## License MIT
More Details in LICENSE.md

