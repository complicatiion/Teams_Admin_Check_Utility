[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$Script:UtilityName = 'Teams Admin Check Utility'
$Script:UtilityVersion = '1.0'
$Script:ReportRoot = Join-Path $env:USERPROFILE 'Desktop\TeamsAdminReports'
$Script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$Script:TeamsBootstrapperPath = Join-Path $Script:ScriptRoot 'teamsbootstrapper.exe'
$Script:OfflineMsixPath = Join-Path $Script:ScriptRoot 'MSTeams-x64.msix'

function Test-IsAdmin {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

$Script:IsAdmin = Test-IsAdmin

function Ensure-ReportRoot {
    if (-not (Test-Path -LiteralPath $Script:ReportRoot)) {
        New-Item -Path $Script:ReportRoot -ItemType Directory -Force | Out-Null
    }
}

function Write-Banner {
    Clear-Host
    Write-Host ('=' * 108)
    Write-Host ''
    Write-Host (' ' * 34 + $Script:UtilityName)
    Write-Host ''
    Write-Host '                 Microsoft Teams inventory, cache, account, log, update, uninstall and reinstall checks'
    Write-Host ('=' * 108)
    Write-Host ''
    Write-Host ('Admin status : ' + ($(if ($Script:IsAdmin) { 'YES' } else { 'NO' })))
    Write-Host ('Report folder : ' + $Script:ReportRoot)
    Write-Host ''
}

function Write-Section {
    param([Parameter(Mandatory = $true)][string]$Title)

    Write-Host ('=' * 108)
    Write-Host $Title
    Write-Host ('=' * 108)
    Write-Host ''
}

function Pause-Utility {
    Write-Host ''
    Read-Host 'Press Enter to return to the menu' | Out-Null
}

function Require-Admin {
    if (-not $Script:IsAdmin) {
        Write-Host ''
        Write-Host 'This action requires an elevated PowerShell session.'
        Write-Host 'Close this window, right-click the launcher, and choose "Run as administrator".'
        Pause-Utility
        return $false
    }

    return $true
}

function Convert-ToSizeString {
    param([Parameter(Mandatory = $true)][Int64]$Bytes)

    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return ($Bytes.ToString() + ' B')
}

function Get-FolderSizeBytes {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return [Int64]0
    }

    $measure = Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { -not $_.PSIsContainer } |
        Measure-Object -Property Length -Sum

    if ($null -eq $measure.Sum) {
        return [Int64]0
    }

    return [Int64]$measure.Sum
}

function Get-SafeFileVersion {
    param([Parameter(Mandatory = $true)][string]$Path)

    try {
        return (Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo.FileVersion
    }
    catch {
        return $null
    }
}

function Get-SafeDate {
    param([datetime]$Value)

    if ($null -eq $Value) {
        return $null
    }

    return $Value.ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-UninstallEntries {
    $entries = @()
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($path in $paths) {
        $scope = if ($path -like 'HKCU:*') { 'CurrentUser' } else { 'LocalMachine' }

        $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if ([string]::IsNullOrWhiteSpace($item.DisplayName)) {
                continue
            }

            $entries += [pscustomobject]@{
                Scope               = $scope
                DisplayName         = $item.DisplayName
                DisplayVersion      = $item.DisplayVersion
                Publisher           = $item.Publisher
                InstallLocation     = $item.InstallLocation
                UninstallString     = $item.UninstallString
                QuietUninstallString= $item.QuietUninstallString
                RegistryKey         = $item.PSChildName
            }
        }
    }

    return $entries
}

function Get-TeamsInstallations {
    $results = @()

    try {
        $currentUserPackages = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue
        foreach ($pkg in $currentUserPackages) {
            $results += [pscustomobject]@{
                Component      = 'New Teams (AppX current user)'
                Installed      = $true
                Scope          = 'CurrentUser'
                Version        = $pkg.Version.ToString()
                Source         = 'AppX'
                InstallLocation= $pkg.InstallLocation
                Identifier     = $pkg.PackageFullName
            }
        }
    }
    catch {
    }

    if ($Script:IsAdmin) {
        try {
            $allUserPackages = Get-AppxPackage -AllUsers -Name 'MSTeams' -ErrorAction SilentlyContinue
            foreach ($pkg in $allUserPackages) {
                $results += [pscustomobject]@{
                    Component      = 'New Teams (AppX all users view)'
                    Installed      = $true
                    Scope          = $pkg.NonRemovable
                    Version        = $pkg.Version.ToString()
                    Source         = 'AppXAllUsers'
                    InstallLocation= $pkg.InstallLocation
                    Identifier     = $pkg.PackageFullName
                }
            }
        }
        catch {
        }

        try {
            $provisionedPackages = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -eq 'MSTeams' }
            foreach ($pkg in $provisionedPackages) {
                $results += [pscustomobject]@{
                    Component      = 'New Teams (Provisioned package)'
                    Installed      = $true
                    Scope          = 'Machine'
                    Version        = $pkg.Version
                    Source         = 'ProvisionedAppX'
                    InstallLocation= 'Provisioned'
                    Identifier     = $pkg.PackageName
                }
            }
        }
        catch {
        }
    }

    $classicExe = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\current\Teams.exe'
    if (Test-Path -LiteralPath $classicExe) {
        $results += [pscustomobject]@{
            Component      = 'Classic Teams (per-user executable)'
            Installed      = $true
            Scope          = 'CurrentUser'
            Version        = (Get-SafeFileVersion -Path $classicExe)
            Source         = 'FileSystem'
            InstallLocation= $classicExe
            Identifier     = 'ClassicTeamsExe'
        }
    }

    $uninstallEntries = Get-UninstallEntries |
        Where-Object { $_.DisplayName -match 'Teams' -or $_.DisplayName -match 'Microsoft Teams' }

    foreach ($entry in $uninstallEntries) {
        $results += [pscustomobject]@{
            Component      = $entry.DisplayName
            Installed      = $true
            Scope          = $entry.Scope
            Version        = $entry.DisplayVersion
            Source         = 'UninstallRegistry'
            InstallLocation= $entry.InstallLocation
            Identifier     = $entry.RegistryKey
        }
    }

    if (-not $results) {
        $results += [pscustomobject]@{
            Component      = 'No Teams installation detected'
            Installed      = $false
            Scope          = 'N/A'
            Version        = $null
            Source         = 'Inventory'
            InstallLocation= $null
            Identifier     = $null
        }
    }

    return $results | Sort-Object Component, Version -Unique
}

function Get-TeamsCacheInfo {
    $items = @()

    $cacheTargets = @(
        [pscustomobject]@{ Name = 'New Teams cache'; Path = (Join-Path $env:USERPROFILE 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams') },
        [pscustomobject]@{ Name = 'Classic Teams cache'; Path = (Join-Path $env:APPDATA 'Microsoft\Teams') },
        [pscustomobject]@{ Name = 'New Teams package root'; Path = (Join-Path $env:USERPROFILE 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe') }
    )

    foreach ($target in $cacheTargets) {
        $exists = Test-Path -LiteralPath $target.Path
        $bytes = if ($exists) { Get-FolderSizeBytes -Path $target.Path } else { [Int64]0 }
        $fileCount = if ($exists) {
            (Get-ChildItem -LiteralPath $target.Path -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer }).Count
        }
        else {
            0
        }

        $items += [pscustomobject]@{
            Name      = $target.Name
            Exists    = $exists
            Path      = $target.Path
            SizeBytes = $bytes
            Size      = Convert-ToSizeString -Bytes $bytes
            FileCount = $fileCount
        }
    }

    return $items
}

function Get-TeamsIdentitySnapshot {
    $identities = @()
    $identityPaths = @(
        'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\Identities\*',
        'HKCU:\Software\Microsoft\Office\15.0\Common\Identity\Identities\*'
    )

    foreach ($path in $identityPaths) {
        $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            $email = $null
            $signInName = $null
            $friendlyName = $null

            if ($item.PSObject.Properties.Name -contains 'EmailAddress') { $email = $item.EmailAddress }
            if ($item.PSObject.Properties.Name -contains 'SignInName') { $signInName = $item.SignInName }
            if ($item.PSObject.Properties.Name -contains 'FriendlyName') { $friendlyName = $item.FriendlyName }

            if ([string]::IsNullOrWhiteSpace($email) -and [string]::IsNullOrWhiteSpace($signInName) -and [string]::IsNullOrWhiteSpace($friendlyName)) {
                continue
            }

            $identities += [pscustomobject]@{
                Source       = 'Office Identity Cache'
                FriendlyName = $friendlyName
                SignInName   = $signInName
                EmailAddress = $email
                RegistryPath = $item.PSPath
            }
        }
    }

    $dsregPath = Join-Path $env:SystemRoot 'System32\dsregcmd.exe'
    if (Test-Path -LiteralPath $dsregPath) {
        try {
            $raw = & $dsregPath /status 2>$null
            $tenantLine = ($raw | Where-Object { $_ -match '^\s*TenantName\s*:' } | Select-Object -First 1)
            $userLine = ($raw | Where-Object { $_ -match '^\s*WorkplaceUserEmail\s*:' -or $_ -match '^\s*User Email\s*:' } | Select-Object -First 1)
            if ($tenantLine -or $userLine) {
                $tenant = if ($tenantLine) { (($tenantLine -split ':', 2)[1]).Trim() } else { $null }
                $user = if ($userLine) { (($userLine -split ':', 2)[1]).Trim() } else { $null }
                $identities += [pscustomobject]@{
                    Source       = 'DSREG status'
                    FriendlyName = $tenant
                    SignInName   = $user
                    EmailAddress = $user
                    RegistryPath = 'dsregcmd /status'
                }
            }
        }
        catch {
        }
    }

    if (-not $identities) {
        $identities += [pscustomobject]@{
            Source       = 'Identity snapshot'
            FriendlyName = $null
            SignInName   = $null
            EmailAddress = $null
            RegistryPath = 'No Office or dsreg identities were detected for the current user.'
        }
    }

    return $identities | Sort-Object Source, EmailAddress, SignInName -Unique
}

function Get-TeamsProcessSnapshot {
    $targetNames = @('ms-teams', 'teams', 'msedgewebview2')

    $processes = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $targetNames -contains $_.Name.ToLowerInvariant() }

    if (-not $processes) {
        return @(
            [pscustomobject]@{
                Name         = 'No Teams-related process found'
                Id           = $null
                CPU          = $null
                WorkingSetMB = $null
                StartTime    = $null
                Path         = $null
            }
        )
    }

    $rows = foreach ($process in $processes) {
        $path = $null
        $startTime = $null

        try { $path = $process.Path } catch { }
        try { $startTime = $process.StartTime } catch { }

        [pscustomobject]@{
            Name         = $process.Name
            Id           = $process.Id
            CPU          = $process.CPU
            WorkingSetMB = [math]::Round(($process.WorkingSet64 / 1MB), 2)
            StartTime    = Get-SafeDate -Value $startTime
            Path         = $path
        }
    }

    return $rows | Sort-Object Name, Id
}

function Get-WebView2Info {
    $entries = Get-UninstallEntries |
        Where-Object { $_.DisplayName -like 'Microsoft Edge WebView2 Runtime*' }

    if (-not $entries) {
        return @(
            [pscustomobject]@{
                DisplayName    = 'WebView2 Runtime not detected'
                DisplayVersion = $null
                Scope          = $null
                Publisher      = $null
            }
        )
    }

    return $entries |
        Select-Object DisplayName, DisplayVersion, Scope, Publisher |
        Sort-Object DisplayVersion -Descending -Unique
}

function Get-TeamsLogBundleInfo {
    $downloadsPath = Join-Path $env:USERPROFILE 'Downloads'
    if (-not (Test-Path -LiteralPath $downloadsPath)) {
        return @(
            [pscustomobject]@{
                Name         = 'Downloads folder not found'
                FullName     = $downloadsPath
                LengthMB     = $null
                LastWriteTime= $null
            }
        )
    }

    $patterns = @('MSTeams*', 'PROD-WebLogs*', '*Teams*log*', '*Teams*support*')
    $files = foreach ($pattern in $patterns) {
        Get-ChildItem -LiteralPath $downloadsPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like $pattern }
    }

    $uniqueFiles = $files | Sort-Object FullName -Unique
    if (-not $uniqueFiles) {
        return @(
            [pscustomobject]@{
                Name         = 'No obvious Teams log bundles found in Downloads'
                FullName     = $downloadsPath
                LengthMB     = $null
                LastWriteTime= $null
            }
        )
    }

    return $uniqueFiles |
        Sort-Object LastWriteTime -Descending |
        Select-Object Name, FullName,
            @{ Name = 'LengthMB'; Expression = { [math]::Round(($_.Length / 1MB), 2) } },
            @{ Name = 'LastWriteTime'; Expression = { Get-SafeDate -Value $_.LastWriteTime } }
}

function Get-TeamsHealthSummary {
    $installations = Get-TeamsInstallations
    $caches = Get-TeamsCacheInfo
    $identities = Get-TeamsIdentitySnapshot
    $processes = Get-TeamsProcessSnapshot
    $webview2 = Get-WebView2Info

    $hasNewTeams = $installations.Component -contains 'New Teams (AppX current user)' -or
                   $installations.Component -contains 'New Teams (AppX all users view)' -or
                   $installations.Component -contains 'New Teams (Provisioned package)'
    $hasClassicTeams = ($installations | Where-Object { $_.Component -match 'Classic Teams|Teams Machine-Wide Installer|Microsoft Teams classic|Teams Installer' }).Count -gt 0

    $newTeamsCache = $caches | Where-Object { $_.Name -eq 'New Teams cache' } | Select-Object -First 1
    $classicTeamsCache = $caches | Where-Object { $_.Name -eq 'Classic Teams cache' } | Select-Object -First 1
    $identityCount = ($identities | Where-Object { $_.EmailAddress -or $_.SignInName -or $_.FriendlyName }).Count
    $runningCount = ($processes | Where-Object { $_.Id }).Count

    $notes = New-Object System.Collections.Generic.List[string]
    if (-not $hasNewTeams) { $notes.Add('- New Teams is not clearly detected for the current user.') }
    if ($hasClassicTeams) { $notes.Add('- Classic Teams remnants or installers are still present.') }
    if ($newTeamsCache -and $newTeamsCache.SizeBytes -ge 1GB) { $notes.Add('- New Teams cache is larger than 1 GB. A reset may help.') }
    if ($classicTeamsCache -and $classicTeamsCache.SizeBytes -ge 1GB) { $notes.Add('- Classic Teams cache is larger than 1 GB. Cleanup may be useful.') }
    if ($identityCount -gt 1) { $notes.Add('- Multiple potential Office or Teams identities were detected.') }
    if ($runningCount -eq 0) { $notes.Add('- No running Teams-related process was detected at the moment of the check.') }
    if (($webview2 | Where-Object { $_.DisplayVersion }).Count -eq 0) { $notes.Add('- WebView2 Runtime was not detected in uninstall inventory.') }
    if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) { $notes.Add('- WinGet is not available. Update and reinstall options will be more limited.') }
    if ($notes.Count -eq 0) { $notes.Add('- No obvious Teams-side blocker was detected in the quick summary.') }

    return [pscustomobject]@{
        ComputerName      = $env:COMPUTERNAME
        UserName          = $env:USERNAME
        Admin             = $Script:IsAdmin
        NewTeamsDetected  = $hasNewTeams
        ClassicDetected   = $hasClassicTeams
        IdentityCount     = $identityCount
        RunningProcesses  = $runningCount
        NewTeamsCache     = if ($newTeamsCache) { $newTeamsCache.Size } else { 'N/A' }
        ClassicTeamsCache = if ($classicTeamsCache) { $classicTeamsCache.Size } else { 'N/A' }
        Notes             = $notes
    }
}

function Stop-TeamsProcesses {
    $names = @('ms-teams', 'teams')
    $targets = Get-Process -ErrorAction SilentlyContinue | Where-Object { $names -contains $_.Name.ToLowerInvariant() }

    foreach ($target in $targets) {
        try {
            Stop-Process -Id $target.Id -Force -ErrorAction Stop
        }
        catch {
            Write-Host ('Could not stop process ' + $target.Name + ' (' + $target.Id + '): ' + $_.Exception.Message)
        }
    }
}

function Reset-TeamsCache {
    Write-Section 'Clear or reset Teams cache'
    if (-not (Require-Admin)) { return }

    Stop-TeamsProcesses

    $targets = @(
        (Join-Path $env:USERPROFILE 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams'),
        (Join-Path $env:APPDATA 'Microsoft\Teams')
    )

    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target)) {
            Write-Host ('Not found: ' + $target)
            continue
        }

        Write-Host ('Cleaning: ' + $target)
        Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host ''
    Write-Host 'Teams cache cleanup completed.'
    Pause-Utility
}

function Test-Winget {
    return [bool](Get-Command winget.exe -ErrorAction SilentlyContinue)
}

function Invoke-TeamsUpdate {
    Write-Section 'Trigger Teams update'
    if (-not (Require-Admin)) { return }

    Stop-TeamsProcesses

    if (Test-Path -LiteralPath $Script:TeamsBootstrapperPath) {
        Write-Host ('Using teamsbootstrapper.exe from: ' + $Script:TeamsBootstrapperPath)
        if (Test-Path -LiteralPath $Script:OfflineMsixPath) {
            & $Script:TeamsBootstrapperPath -p -o $Script:OfflineMsixPath
        }
        else {
            & $Script:TeamsBootstrapperPath -p
        }
    }
    elseif (Test-Winget) {
        Write-Host 'Using WinGet to upgrade Microsoft Teams ...'
        & winget upgrade --id Microsoft.Teams --exact --source winget --include-unknown --accept-source-agreements --accept-package-agreements
    }
    else {
        Write-Host 'No supported automatic update method was found.'
        Write-Host 'Place teamsbootstrapper.exe next to this script or install WinGet.'
    }

    Write-Host ''
    Write-Host 'Update action completed.'
    Pause-Utility
}

function Invoke-QuietUninstallEntry {
    param([Parameter(Mandatory = $true)]$Entry)

    $command = if (-not [string]::IsNullOrWhiteSpace($Entry.QuietUninstallString)) { $Entry.QuietUninstallString } else { $Entry.UninstallString }
    if ([string]::IsNullOrWhiteSpace($command)) {
        Write-Host ('No uninstall string available for ' + $Entry.DisplayName)
        return
    }

    if ($command -match '(?i)msiexec(\.exe)?\s+/I\s+(\{[^\}]+\})') {
        $productCode = $Matches[2]
        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow
        return
    }

    if ($command -match '(?i)msiexec(\.exe)?\s+/X\s+(\{[^\}]+\})') {
        $productCode = $Matches[2]
        Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x $productCode /qn /norestart" -Wait -NoNewWindow
        return
    }

    Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', $command -Wait -NoNewWindow
}

function Invoke-TeamsUninstall {
    Write-Section 'Uninstall Microsoft Teams'
    if (-not (Require-Admin)) { return }

    Stop-TeamsProcesses

    Write-Host '[1/3] Removing new Teams AppX package for the current user if present ...'
    try {
        $currentUserPackages = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue
        foreach ($pkg in $currentUserPackages) {
            Write-Host ('Removing package: ' + $pkg.PackageFullName)
            Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host ('AppX uninstall warning: ' + $_.Exception.Message)
    }

    Write-Host ''
    Write-Host '[2/3] Uninstalling Teams-related classic entries and machine-wide installer if present ...'
    $targets = Get-UninstallEntries |
        Where-Object {
            $_.DisplayName -match '^Microsoft Teams classic$' -or
            $_.DisplayName -match '^Teams Machine-Wide Installer$' -or
            $_.DisplayName -match '^Microsoft Teams$' -or
            $_.DisplayName -match '^Teams Installer$'
        }

    foreach ($target in $targets) {
        Write-Host ('Uninstalling: ' + $target.DisplayName + ' [' + $target.Scope + ']')
        Invoke-QuietUninstallEntry -Entry $target
    }

    Write-Host ''
    Write-Host '[3/3] Removing leftover local folders if found ...'
    $cleanupTargets = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Teams'),
        (Join-Path $env:APPDATA 'Microsoft\Teams'),
        (Join-Path $env:USERPROFILE 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams')
    )

    foreach ($cleanupTarget in $cleanupTargets) {
        if (Test-Path -LiteralPath $cleanupTarget) {
            Write-Host ('Cleaning leftover data: ' + $cleanupTarget)
            Remove-Item -LiteralPath $cleanupTarget -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host ''
    Write-Host 'Teams uninstall action completed.'
    Pause-Utility
}

function Invoke-TeamsInstall {
    Write-Section 'Install or reinstall Microsoft Teams'
    if (-not (Require-Admin)) { return }

    Stop-TeamsProcesses

    if (Test-Path -LiteralPath $Script:TeamsBootstrapperPath) {
        Write-Host ('Using teamsbootstrapper.exe from: ' + $Script:TeamsBootstrapperPath)
        if (Test-Path -LiteralPath $Script:OfflineMsixPath) {
            Write-Host ('Using offline MSIX: ' + $Script:OfflineMsixPath)
            & $Script:TeamsBootstrapperPath -p -o $Script:OfflineMsixPath
        }
        else {
            & $Script:TeamsBootstrapperPath -p
        }
    }
    elseif (Test-Winget) {
        Write-Host 'Using WinGet to install Microsoft Teams ...'
        & winget install --id Microsoft.Teams --exact --source winget --accept-source-agreements --accept-package-agreements
    }
    else {
        Write-Host 'No supported installation method was found.'
        Write-Host 'Place teamsbootstrapper.exe next to this script or install WinGet.'
    }

    Write-Host ''
    Write-Host 'Install or reinstall action completed.'
    Pause-Utility
}

function Start-TeamsClient {
    Write-Section 'Start Microsoft Teams'

    try {
        Start-Process -FilePath 'ms-teams.exe' -ErrorAction Stop
        Write-Host 'Started new Teams via ms-teams.exe.'
    }
    catch {
        $classicPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\current\Teams.exe'
        if (Test-Path -LiteralPath $classicPath) {
            Start-Process -FilePath $classicPath
            Write-Host 'Started classic Teams executable.'
        }
        else {
            Write-Host 'Could not start Teams automatically.'
        }
    }

    Pause-Utility
}

function Show-QuickSummary {
    Write-Section 'Quick Teams health summary'

    $summary = Get-TeamsHealthSummary
    [pscustomobject]@{
        ComputerName      = $summary.ComputerName
        UserName          = $summary.UserName
        Admin             = $summary.Admin
        NewTeamsDetected  = $summary.NewTeamsDetected
        ClassicDetected   = $summary.ClassicDetected
        IdentityCount     = $summary.IdentityCount
        RunningProcesses  = $summary.RunningProcesses
        NewTeamsCache     = $summary.NewTeamsCache
        ClassicTeamsCache = $summary.ClassicTeamsCache
    } | Format-List | Out-Host

    Write-Host 'Notes:'
    foreach ($note in $summary.Notes) {
        Write-Host $note
    }

    Pause-Utility
}

function Show-Installations {
    Write-Section 'Detailed Teams installation and version check'
    Get-TeamsInstallations | Format-Table -Wrap -AutoSize | Out-Host
    Pause-Utility
}

function Show-CacheAnalysis {
    Write-Section 'Teams cache size and log analysis'

    Write-Host 'Cache folders:'
    Get-TeamsCacheInfo | Format-Table -AutoSize | Out-Host

    Write-Host ''
    Write-Host 'Potential Teams log bundles in Downloads:'
    Get-TeamsLogBundleInfo | Format-Table -Wrap -AutoSize | Out-Host

    Pause-Utility
}

function Show-IdentitySnapshot {
    Write-Section 'Potential Teams or Office identities (best effort)'
    Get-TeamsIdentitySnapshot | Format-Table -Wrap -AutoSize | Out-Host
    Pause-Utility
}

function Show-ProcessAndRuntimeCheck {
    Write-Section 'Teams processes and WebView2 runtime check'

    Write-Host 'Processes:'
    Get-TeamsProcessSnapshot | Format-Table -Wrap -AutoSize | Out-Host

    Write-Host ''
    Write-Host 'WebView2 runtime:'
    Get-WebView2Info | Format-Table -Wrap -AutoSize | Out-Host

    Pause-Utility
}

function New-TeamsReport {
    Write-Section 'Create full Teams analysis report'
    Ensure-ReportRoot

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportDir = Join-Path $Script:ReportRoot ('Report_' + $stamp)
    $zipPath = $reportDir + '.zip'

    New-Item -Path $reportDir -ItemType Directory -Force | Out-Null

    $installations = Get-TeamsInstallations
    $caches = Get-TeamsCacheInfo
    $identities = Get-TeamsIdentitySnapshot
    $processes = Get-TeamsProcessSnapshot
    $webview2 = Get-WebView2Info
    $logBundles = Get-TeamsLogBundleInfo
    $summary = Get-TeamsHealthSummary

    Write-Host '[1/7] Writing summary text ...'
    $summaryLines = @(
        ($Script:UtilityName + ' report'),
        '',
        ('Generated         : ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
        ('Computer          : ' + $env:COMPUTERNAME),
        ('User              : ' + $env:USERNAME),
        ('Admin             : ' + $Script:IsAdmin),
        ('New Teams         : ' + $summary.NewTeamsDetected),
        ('Classic Teams     : ' + $summary.ClassicDetected),
        ('Identity count    : ' + $summary.IdentityCount),
        ('Running processes : ' + $summary.RunningProcesses),
        ('New Teams cache   : ' + $summary.NewTeamsCache),
        ('Classic cache     : ' + $summary.ClassicTeamsCache),
        '',
        'Notes',
        '-----'
    )
    $summaryLines += $summary.Notes
    Set-Content -LiteralPath (Join-Path $reportDir 'Summary.txt') -Value $summaryLines -Encoding UTF8

    Write-Host '[2/7] Exporting installations ...'
    $installations | Export-Csv -LiteralPath (Join-Path $reportDir 'Installations.csv') -NoTypeInformation -Encoding UTF8
    $installations | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $reportDir 'Installations.json') -Encoding UTF8

    Write-Host '[3/7] Exporting cache analysis ...'
    $caches | Export-Csv -LiteralPath (Join-Path $reportDir 'CacheAnalysis.csv') -NoTypeInformation -Encoding UTF8

    Write-Host '[4/7] Exporting identities ...'
    $identities | Export-Csv -LiteralPath (Join-Path $reportDir 'IdentitySnapshot.csv') -NoTypeInformation -Encoding UTF8

    Write-Host '[5/7] Exporting processes and runtime details ...'
    $processes | Export-Csv -LiteralPath (Join-Path $reportDir 'Processes.csv') -NoTypeInformation -Encoding UTF8
    $webview2 | Export-Csv -LiteralPath (Join-Path $reportDir 'WebView2.csv') -NoTypeInformation -Encoding UTF8

    Write-Host '[6/7] Exporting log bundle inventory ...'
    $logBundles | Export-Csv -LiteralPath (Join-Path $reportDir 'LogBundles.csv') -NoTypeInformation -Encoding UTF8

    Write-Host '[7/7] Writing environment details ...'
    [pscustomobject]@{
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        OSVersion         = [System.Environment]::OSVersion.VersionString
        ScriptVersion     = $Script:UtilityVersion
        WingetAvailable   = (Test-Winget)
        BootstrapperFound = (Test-Path -LiteralPath $Script:TeamsBootstrapperPath)
        OfflineMsixFound  = (Test-Path -LiteralPath $Script:OfflineMsixPath)
        ReportFolder      = $reportDir
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $reportDir 'Environment.json') -Encoding UTF8

    try {
        if (Test-Path -LiteralPath $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        }
        Compress-Archive -Path (Join-Path $reportDir '*') -DestinationPath $zipPath -CompressionLevel Optimal -Force
        Write-Host ''
        Write-Host ('ZIP report created: ' + $zipPath)
    }
    catch {
        Write-Host ('ZIP creation warning: ' + $_.Exception.Message)
    }

    Write-Host ''
    Write-Host ('Report folder: ' + $reportDir)
    Start-Process explorer.exe $reportDir
    Pause-Utility
}

function Open-ReportFolder {
    Ensure-ReportRoot
    Start-Process explorer.exe $Script:ReportRoot
}

Ensure-ReportRoot

while ($true) {
    Write-Banner
    Write-Host '[1] Quick Teams health summary'
    Write-Host '[2] Detailed installation and version check'
    Write-Host '[3] Cache size and log analysis'
    Write-Host '[4] Potential Teams or Office identities (best effort)'
    Write-Host '[5] Process and WebView2 runtime check'
    Write-Host '[6] Create full Teams report'
    Write-Host '[7] Trigger Teams update'
    Write-Host '[8] Uninstall Teams'
    Write-Host '[9] Install or reinstall Teams'
    Write-Host '[A] Clear or reset Teams cache'
    Write-Host '[B] Start Teams'
    Write-Host '[C] Open report folder'
    Write-Host '[0] Exit'
    Write-Host ''

    $choice = Read-Host 'Selection'
    switch ($choice.ToUpperInvariant()) {
        '1' { Show-QuickSummary }
        '2' { Show-Installations }
        '3' { Show-CacheAnalysis }
        '4' { Show-IdentitySnapshot }
        '5' { Show-ProcessAndRuntimeCheck }
        '6' { New-TeamsReport }
        '7' { Invoke-TeamsUpdate }
        '8' { Invoke-TeamsUninstall }
        '9' { Invoke-TeamsInstall }
        'A' { Reset-TeamsCache }
        'B' { Start-TeamsClient }
        'C' { Open-ReportFolder }
        '0' { break }
        default {
            Write-Host ''
            Write-Host 'Invalid selection.'
            Pause-Utility
        }
    }
}
