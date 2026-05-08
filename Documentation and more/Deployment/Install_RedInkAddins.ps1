[CmdletBinding()]
param(
    [ValidateSet('GA', 'Preview')]
    [string]$Environment = 'GA',

    [ValidateSet('All', 'Word', 'Excel', 'Outlook')]
    [string[]]$AddIns = @('All'),

    [int]$TimeoutSeconds = 600,

    [switch]$SkipRegistryProtection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogPath = Join-Path $env:TEMP 'RedInkAddins-Install.log'

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

    throw 'VSTOInstaller.exe was not found. Install Microsoft Visual Studio Tools for Office Runtime first.'
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

function Test-AddInInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $properties = Get-AddInProperties -AddIn $AddIn
    if ($null -eq $properties) {
        return $false
    }

    return ($null -ne $properties.Manifest) -or ($null -ne $properties.LoadBehavior)
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
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $properties = Get-AddInProperties -AddIn $AddIn
    if ($null -eq $properties) {
        return ''
    }

    if ($null -ne $properties.RedInkEnvironment -and -not [string]::IsNullOrWhiteSpace([string]$properties.RedInkEnvironment)) {
        return [string]$properties.RedInkEnvironment
    }

    if ($null -ne $properties.Manifest) {
        return Resolve-EnvironmentFromManifest -Manifest ([string]$properties.Manifest)
    }

    return ''
}

function Set-AddInEnvironmentMarker {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $path = Get-AddInRegistryPath -AddIn $AddIn
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name 'RedInkEnvironment' -Value $Environment -PropertyType String -Force | Out-Null
}

function Protect-AddInRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $addInPath = Get-AddInRegistryPath -AddIn $AddIn
    New-Item -Path $addInPath -Force | Out-Null
    New-ItemProperty -Path $addInPath -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

    $officeVersion = Get-OfficeVersion
    $resiliencyPath = "HKCU:\Software\Microsoft\Office\$officeVersion\$($AddIn.Host)\Resiliency\DoNotDisableAddinList"
    New-Item -Path $resiliencyPath -Force | Out-Null
    New-ItemProperty -Path $resiliencyPath -Name $AddIn.ProgId -Value 1 -PropertyType DWord -Force | Out-Null

    Write-Log "Applied registry protection for $($AddIn.ProgId)."
}

function Wait-ForAddInInstallation {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn,

        [Parameter(Mandatory = $true)]
        [int]$SecondsToWait
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    do {
        Start-Sleep -Seconds 3

        if (Test-AddInInstalled -AddIn $AddIn) {
            Write-Log "$($AddIn.ProgId) installation detected."
            return $true
        }
    }
    while ($stopwatch.Elapsed.TotalSeconds -lt $SecondsToWait)

    return $false
}

function Invoke-VstoInstallSilent {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$AddIn
    )

    $url = $AddIn.Urls[$Environment]
    $vstoInstaller = Get-VstoInstallerPath

    Write-Log "Starting silent install for $($AddIn.ProgId) from $url"
    Write-Log "Using VSTOInstaller: $vstoInstaller"

    $arguments = @(
        '/install',
        $url,
        '/silent'
    )

    $process = Start-Process `
        -FilePath $vstoInstaller `
        -ArgumentList $arguments `
        -Wait `
        -PassThru `
        -WindowStyle Hidden

    Write-Log "VSTOInstaller exit code for $($AddIn.ProgId): $($process.ExitCode)"

    if ($process.ExitCode -ne 0) {
        throw "VSTOInstaller failed for $($AddIn.ProgId) with exit code $($process.ExitCode)."
    }

    if (Wait-ForAddInInstallation -AddIn $AddIn -SecondsToWait $TimeoutSeconds) {
        return $true
    }

    Write-Log "Timed out waiting for $($AddIn.ProgId) installation."
    return $false
}

$catalog = Get-AddInCatalog
$selectedAddIns = Get-SelectedAddIns -Catalog $catalog -RequestedAddIns $AddIns

Write-Log "Silent install started. Environment=$Environment AddIns=$($AddIns -join ',')"

$failed = New-Object System.Collections.Generic.List[string]

foreach ($addIn in $selectedAddIns) {
    $alreadyInstalled = Test-AddInInstalled -AddIn $addIn
    $installedEnvironment = Get-InstalledEnvironment -AddIn $addIn

    if ($alreadyInstalled) {
        if ($installedEnvironment -eq $Environment) {
            Write-Log "$($addIn.ProgId) is already installed in channel $Environment."
            Set-AddInEnvironmentMarker -AddIn $addIn

            if (-not $SkipRegistryProtection) {
                Protect-AddInRegistry -AddIn $addIn
            }

            continue
        }

        if ([string]::IsNullOrWhiteSpace($installedEnvironment)) {
            Write-Log "$($addIn.ProgId) is already installed, but the channel could not be determined. Uninstall it first, then install channel $Environment."
        }
        else {
            Write-Log "$($addIn.ProgId) is already installed in channel $installedEnvironment. Uninstall it first, then install channel $Environment."
        }

        $failed.Add($addIn.ProgId)
        continue
    }

    $installed = $false

    try {
        $installed = Invoke-VstoInstallSilent -AddIn $addIn
    }
    catch {
        Write-Log "Installation failed for $($addIn.ProgId): $($_.Exception.Message)"
        $failed.Add($addIn.ProgId)
        continue
    }

    if (-not $installed) {
        $failed.Add($addIn.ProgId)
        continue
    }

    Set-AddInEnvironmentMarker -AddIn $addIn

    if (-not $SkipRegistryProtection) {
        Protect-AddInRegistry -AddIn $addIn
    }

    Write-Log "Configured $($addIn.ProgId) for channel $Environment."
}

if ($failed.Count -gt 0) {
    Write-Log "Silent install finished with failures: $($failed -join ', ')"
    exit 1
}

Write-Log 'Silent install finished successfully.'
exit 0