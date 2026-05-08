[CmdletBinding()]
param(
    [ValidateSet('GA', 'Preview')]
    [string]$Environment = 'GA',

    [ValidateSet('All', 'Word', 'Excel', 'Outlook')]
    [string[]]$AddIns = @('All'),

    [int]$TimeoutSeconds = 300,

    [switch]$SkipClickOnceCacheCleanup,

    [switch]$KeepOfficeAppsOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogPath = Join-Path $env:TEMP 'RedInkAddins-Reset-Reinstall.log'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"
    Write-Host $line
    Add-Content -Path $script:LogPath -Value $line
}

function Write-Section {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    Write-Host ''
    Write-Log "===== $Text ====="
}

function Get-VstoInstallerPath {
    $paths = @(
        (Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'),
        (Join-Path ${env:CommonProgramFiles(x86)} 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe')
    )

    foreach ($path in $paths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) {
            return $path
        }
    }

    throw 'VSTOInstaller.exe was not found. Install Microsoft Visual Studio Tools for Office Runtime first.'
}

function Get-OfficeVersion {
    foreach ($version in @('17.0', '16.0', '15.0', '14.0')) {
        if (Test-Path "HKCU:\Software\Microsoft\Office\$version\Common\General") {
            return $version
        }
    }

    return '16.0'
}

function Get-AddInCatalog {
    return @{
        Word = @{
            OfficeHost = 'Word'
            ProcessName = 'winword'
            ProgId = 'Red Ink for Word'
            FriendlyName = 'Red Ink for Word'
            Urls = @{
                GA = 'https://redink.ai/apps/ga/word/Red%20Ink%20for%20Word.vsto'
                Preview = 'https://redink.ai/apps/preview/word/Red%20Ink%20for%20Word.vsto'
            }
        }
        Excel = @{
            OfficeHost = 'Excel'
            ProcessName = 'excel'
            ProgId = 'Red Ink for Excel'
            FriendlyName = 'Red Ink for Excel'
            Urls = @{
                GA = 'https://redink.ai/apps/ga/excel/Red%20Ink%20for%20Excel.vsto'
                Preview = 'https://redink.ai/apps/preview/excel/Red%20Ink%20for%20Excel.vsto'
            }
        }
        Outlook = @{
            OfficeHost = 'Outlook'
            ProcessName = 'outlook'
            ProgId = 'Red Ink for Outlook'
            FriendlyName = 'Red Ink for Outlook'
            Urls = @{
                GA = 'https://redink.ai/apps/ga/outlook/Red%20Ink%20for%20Outlook.vsto'
                Preview = 'https://redink.ai/apps/preview/outlook/Red%20Ink%20for%20Outlook.vsto'
            }
        }
    }
}

function Get-SelectedAddIns {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Catalog,

        [Parameter(Mandatory = $true)]
        [string[]]$RequestedAddIns
    )

    if ($RequestedAddIns -contains 'All') {
        return @($Catalog.Word, $Catalog.Excel, $Catalog.Outlook)
    }

    $selected = New-Object System.Collections.Generic.List[object]

    foreach ($name in $RequestedAddIns) {
        $selected.Add($Catalog[$name])
    }

    return $selected.ToArray()
}

function Get-AddInRegistryPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    return "HKCU:\Software\Microsoft\Office\$($AddIn.OfficeHost)\Addins\$($AddIn.ProgId)"
}

function Get-RegistryValueSafe {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Properties,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Properties) {
        return $null
    }

    $property = $Properties.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Stop-OfficeApps {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SelectedAddIns
    )

    if ($KeepOfficeAppsOpen) {
        Write-Log 'Skipping Office process shutdown because -KeepOfficeAppsOpen was specified.'
        return
    }

    Write-Section 'Stopping selected Office applications'

    foreach ($addIn in $SelectedAddIns) {
        $processName = [string]$addIn.ProcessName
        Get-Process -Name $processName -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue

        Write-Log "Stopped process if running: $processName"
    }
}

function Remove-RegistryTreeIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Removed registry path: $Path"
    }
}

function Clear-OfficeResiliency {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SelectedAddIns
    )

    Write-Section 'Removing Office resiliency blocks'

    foreach ($version in @('17.0', '16.0', '15.0', '14.0')) {
        foreach ($addIn in $SelectedAddIns) {
            $officeHost = [string]$addIn.OfficeHost
            $base = "HKCU:\Software\Microsoft\Office\$version\$officeHost\Resiliency"

            Remove-RegistryTreeIfExists -Path "$base\DisabledItems"
            Remove-RegistryTreeIfExists -Path "$base\CrashingAddinList"
            Remove-RegistryTreeIfExists -Path "$base\StartupItems"
        }
    }
}

function Invoke-VstoUninstallIfPossible {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $path = Get-AddInRegistryPath -AddIn $AddIn
    if (-not (Test-Path $path)) {
        Write-Log "No add-in registry key found for $($AddIn.ProgId). Skipping VSTO uninstall."
        return
    }

    $properties = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    $manifest = Get-RegistryValueSafe -Properties $properties -Name 'Manifest'

    if ([string]::IsNullOrWhiteSpace([string]$manifest)) {
        Write-Log "No manifest available for $($AddIn.ProgId). Skipping VSTO uninstall."
        return
    }

    $manifestForUninstall = ([string]$manifest).Replace('|vstolocal', '')
    $installer = Get-VstoInstallerPath

    Write-Log "Running VSTO silent uninstall for $($AddIn.ProgId): $manifestForUninstall"

    $process = Start-Process `
        -FilePath $installer `
        -ArgumentList @('/uninstall', $manifestForUninstall, '/silent') `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    Write-Log "VSTO uninstall exit code for $($AddIn.ProgId): $($process.ExitCode)"
}

function Remove-RedInkAddInKeys {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SelectedAddIns
    )

    Write-Section 'Removing Red Ink add-in registry keys'

    foreach ($addIn in $SelectedAddIns) {
        $path = Get-AddInRegistryPath -AddIn $addIn
        Remove-RegistryTreeIfExists -Path $path
    }
}

function Clear-ClickOnceCache {
    if ($SkipClickOnceCacheCleanup) {
        Write-Log 'Skipping ClickOnce cache cleanup because -SkipClickOnceCacheCleanup was specified.'
        return
    }

    Write-Section 'Clearing ClickOnce cache'

    try {
        Start-Process -FilePath 'rundll32.exe' -ArgumentList @('dfshim', 'CleanOnlineAppCache') -Wait -WindowStyle Hidden
        Write-Log 'ClickOnce online application cache cleanup completed.'
    }
    catch {
        Write-Log "ClickOnce cache cleanup warning: $($_.Exception.Message)"
    }
}

function Install-AddInSilent {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    Write-Section "Installing $($AddIn.FriendlyName)"

    $installer = Get-VstoInstallerPath
    $url = [string]$AddIn.Urls[$Environment]

    Write-Log "VSTOInstaller: $installer"
    Write-Log "Manifest URL: $url"

    $process = Start-Process `
        -FilePath $installer `
        -ArgumentList @('/install', $url, '/silent') `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    Write-Log "VSTO install exit code for $($AddIn.ProgId): $($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        throw "Silent VSTO install failed for $($AddIn.ProgId) with exit code $($process.ExitCode)."
    }
}

function Repair-AddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    Write-Section "Repairing registry for $($AddIn.FriendlyName)"

    $path = Get-AddInRegistryPath -AddIn $AddIn
    $manifestUrl = [string]$AddIn.Urls[$Environment]

    New-Item -Path $path -Force | Out-Null

    New-ItemProperty -Path $path -Name 'FriendlyName' -Value ([string]$AddIn.FriendlyName) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name 'Description' -Value ([string]$AddIn.FriendlyName) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name 'Manifest' -Value $manifestUrl -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $path -Name 'RedInkEnvironment' -Value $Environment -PropertyType String -Force | Out-Null

    $officeVersion = Get-OfficeVersion
    $doNotDisablePath = "HKCU:\Software\Microsoft\Office\$officeVersion\$($AddIn.OfficeHost)\Resiliency\DoNotDisableAddinList"
    New-Item -Path $doNotDisablePath -Force | Out-Null
    New-ItemProperty -Path $doNotDisablePath -Name ([string]$AddIn.ProgId) -Value 1 -PropertyType DWord -Force | Out-Null

    Write-Log "Registry repaired: $path"
}

function Validate-AddIn {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    Write-Section "Validating $($AddIn.ProgId)"

    $path = Get-AddInRegistryPath -AddIn $AddIn

    if (-not (Test-Path $path)) {
        throw "Registry key missing after install: $path"
    }

    $properties = Get-ItemProperty -Path $path -ErrorAction Stop
    $friendlyName = Get-RegistryValueSafe -Properties $properties -Name 'FriendlyName'
    $description = Get-RegistryValueSafe -Properties $properties -Name 'Description'
    $manifest = Get-RegistryValueSafe -Properties $properties -Name 'Manifest'
    $loadBehavior = Get-RegistryValueSafe -Properties $properties -Name 'LoadBehavior'
    $marker = Get-RegistryValueSafe -Properties $properties -Name 'RedInkEnvironment'

    Write-Host "FriendlyName      : $friendlyName"
    Write-Host "Description       : $description"
    Write-Host "Manifest          : $manifest"
    Write-Host "LoadBehavior      : $loadBehavior"
    Write-Host "RedInkEnvironment : $marker"

    if ([string]::IsNullOrWhiteSpace([string]$manifest)) {
        throw "Manifest is missing for $($AddIn.ProgId)."
    }

    if ([int]$loadBehavior -ne 3) {
        throw "LoadBehavior is not 3 for $($AddIn.ProgId). Current value: $loadBehavior"
    }
}

Write-Log "Reset and reinstall started. Environment=$Environment AddIns=$($AddIns -join ',')"
Write-Log "Log file: $script:LogPath"

$catalog = Get-AddInCatalog
$selectedAddIns = Get-SelectedAddIns -Catalog $catalog -RequestedAddIns $AddIns

Stop-OfficeApps -SelectedAddIns $selectedAddIns
Clear-OfficeResiliency -SelectedAddIns $selectedAddIns

foreach ($addIn in $selectedAddIns) {
    Invoke-VstoUninstallIfPossible -AddIn $addIn
}

Remove-RedInkAddInKeys -SelectedAddIns $selectedAddIns
Clear-ClickOnceCache

foreach ($addIn in $selectedAddIns) {
    Install-AddInSilent -AddIn $addIn
    Repair-AddInRegistry -AddIn $addIn
    Validate-AddIn -AddIn $addIn
}

Write-Section 'Finished'
Write-Log 'Reset and reinstall finished successfully.'
Write-Host ''
Write-Host 'Now start Word, Excel and Outlook manually and verify File > Options > Add-ins.'
exit 0
