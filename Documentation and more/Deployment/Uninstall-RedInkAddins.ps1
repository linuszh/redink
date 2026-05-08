[CmdletBinding()]
param(
    [ValidateSet('All', 'Word', 'Excel', 'Outlook')]
    [string[]]$AddIns = @('All'),

    [ValidateSet('GA', 'Preview')]
    [string]$Environment,

    [int]$TimeoutSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogPath = Join-Path $env:TEMP 'RedInkAddins-Uninstall.log'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[${timestamp}] $Message"
    Write-Host $line
    Add-Content -Path $script:LogPath -Value $line
}

function Get-OfficeVersion {
    foreach ($version in @('17.0', '16.0', '15.0', '14.0')) {
        if (Test-Path "HKCU:\Software\Microsoft\Office\$version\Common\General") {
            return $version
        }
    }

    return '16.0'
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

    return ''
}

function Get-AddInCatalog {
    return @{
        Word = @{
            Host = 'Word'
            ProgId = 'Red Ink for Word'
            Urls = @{
                GA = 'https://redink.ai/apps/ga/word/Red%20Ink%20for%20Word.vsto'
                Preview = 'https://redink.ai/apps/preview/word/Red%20Ink%20for%20Word.vsto'
            }
        }
        Excel = @{
            Host = 'Excel'
            ProgId = 'Red Ink for Excel'
            Urls = @{
                GA = 'https://redink.ai/apps/ga/excel/Red%20Ink%20for%20Excel.vsto'
                Preview = 'https://redink.ai/apps/preview/excel/Red%20Ink%20for%20Excel.vsto'
            }
        }
        Outlook = @{
            Host = 'Outlook'
            ProgId = 'Red Ink for Outlook'
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

    return "HKCU:\Software\Microsoft\Office\$($AddIn.Host)\Addins\$($AddIn.ProgId)"
}

function Get-AddInProperties {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $path = Get-AddInRegistryPath -AddIn $AddIn
    if (-not (Test-Path $path)) {
        return $null
    }

    return Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]$Properties,

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

function Resolve-EnvironmentFromManifest {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Manifest
    )

    if ([string]::IsNullOrWhiteSpace($Manifest)) {
        return ''
    }

    $manifestLower = $Manifest.ToLowerInvariant()

    if ($manifestLower.Contains('/apps/ga/')) {
        return 'GA'
    }

    if ($manifestLower.Contains('/apps/preview/')) {
        return 'Preview'
    }

    return ''
}

function Get-InstalledEnvironment {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]$Properties
    )

    if ($null -eq $Properties) {
        return ''
    }

    $environmentMarker = Get-PropertyValue -Properties $Properties -Name 'RedInkEnvironment'
    if (-not [string]::IsNullOrWhiteSpace([string]$environmentMarker)) {
        return [string]$environmentMarker
    }

    $manifest = Get-PropertyValue -Properties $Properties -Name 'Manifest'
    if (-not [string]::IsNullOrWhiteSpace([string]$manifest)) {
        return Resolve-EnvironmentFromManifest -Manifest ([string]$manifest)
    }

    return ''
}

function Test-AddInInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $properties = Get-AddInProperties -AddIn $AddIn
    if ($null -eq $properties) {
        return $false
    }

    $manifest = Get-PropertyValue -Properties $properties -Name 'Manifest'
    $loadBehavior = Get-PropertyValue -Properties $properties -Name 'LoadBehavior'

    return (-not [string]::IsNullOrWhiteSpace([string]$manifest)) -or ($null -ne $loadBehavior)
}

function Get-ManifestForUninstall {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn,

        [Parameter(Mandatory = $false)]
        [psobject]$Properties
    )

    $manifest = Get-PropertyValue -Properties $Properties -Name 'Manifest'
    if (-not [string]::IsNullOrWhiteSpace([string]$manifest)) {
        return [string]$manifest
    }

    if (-not [string]::IsNullOrWhiteSpace($Environment)) {
        return [string]$AddIn.Urls[$Environment]
    }

    return ''
}

function Invoke-VstoUninstallSilent {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn,

        [Parameter(Mandatory = $false)]
        [psobject]$Properties
    )

    $vstoInstaller = Get-VstoInstallerPath
    if ([string]::IsNullOrWhiteSpace($vstoInstaller)) {
        Write-Log 'VSTOInstaller.exe was not found. VSTO uninstall command will be skipped.'
        return $false
    }

    $manifest = Get-ManifestForUninstall -AddIn $AddIn -Properties $Properties
    if ([string]::IsNullOrWhiteSpace($manifest)) {
        Write-Log "No manifest available for $($AddIn.ProgId). VSTO uninstall command will be skipped."
        return $false
    }

    Write-Log "Starting silent uninstall for $($AddIn.ProgId) using manifest $manifest"
    Write-Log "Using VSTOInstaller: $vstoInstaller"

    $arguments = @(
        '/uninstall',
        $manifest,
        '/silent'
    )

    $process = Start-Process -FilePath $vstoInstaller -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    Write-Log "VSTOInstaller uninstall exit code for $($AddIn.ProgId): $($process.ExitCode)"

    return ($process.ExitCode -eq 0)
}

function Remove-RegistryProtection {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $addInPath = Get-AddInRegistryPath -AddIn $AddIn
    if (Test-Path $addInPath) {
        Remove-Item -Path $addInPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    $officeVersion = Get-OfficeVersion
    $resiliencyPath = "HKCU:\Software\Microsoft\Office\$officeVersion\$($AddIn.Host)\Resiliency\DoNotDisableAddinList"
    if (Test-Path $resiliencyPath) {
        Remove-ItemProperty -Path $resiliencyPath -Name $AddIn.ProgId -Force -ErrorAction SilentlyContinue
    }
}

$catalog = Get-AddInCatalog
$selectedAddIns = Get-SelectedAddIns -Catalog $catalog -RequestedAddIns $AddIns

Write-Log "Silent uninstall started. AddIns=$($AddIns -join ',') Environment=$Environment"

foreach ($addIn in $selectedAddIns) {
    $properties = Get-AddInProperties -AddIn $addIn
    $installedEnvironment = Get-InstalledEnvironment -Properties $properties

    if (-not [string]::IsNullOrWhiteSpace($Environment) -and -not [string]::IsNullOrWhiteSpace($installedEnvironment) -and $installedEnvironment -ne $Environment) {
        Write-Log "$($addIn.ProgId) is installed in channel $installedEnvironment. Requested uninstall channel is $Environment. Skipping."
        continue
    }

    $vstoUninstallStarted = $false

    if ($null -ne $properties) {
        $vstoUninstallStarted = Invoke-VstoUninstallSilent -AddIn $addIn -Properties $properties
    }
    else {
        Write-Log "$($addIn.ProgId) is not registered for the current user. Removing protection entries if present."
    }

    if ($vstoUninstallStarted) {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep -Seconds 2
            if (-not (Test-AddInInstalled -AddIn $addIn)) {
                break
            }
        }
        while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds)
    }
    else {
        Write-Log "Skipping uninstall wait for $($addIn.ProgId) because no VSTO uninstall command was started."
    }

    Remove-RegistryProtection -AddIn $addIn
    Write-Log "Processed uninstall for $($addIn.ProgId)."
}

Write-Log 'Silent uninstall finished.'
exit 0
