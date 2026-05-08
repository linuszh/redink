[CmdletBinding()]
param(
    [ValidateSet('All', 'Word', 'Excel', 'Outlook')]
    [string[]]$AddIns = @('All'),

    [ValidateSet('GA', 'Preview')]
    [string]$Environment,

    [switch]$RequireLoadBehavior3,

    [switch]$RequireManifest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Test-AddInInstalled {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]$Properties,

        [Parameter(Mandatory = $false)]
        [bool]$StrictLoadBehavior = $false,

        [Parameter(Mandatory = $false)]
        [bool]$StrictManifest = $false
    )

    if ($null -eq $Properties) {
        return $false
    }

    $manifest = Get-PropertyValue -Properties $Properties -Name 'Manifest'
    $loadBehavior = Get-PropertyValue -Properties $Properties -Name 'LoadBehavior'

    if ($StrictManifest -and [string]::IsNullOrWhiteSpace([string]$manifest)) {
        return $false
    }

    $installed = (-not [string]::IsNullOrWhiteSpace([string]$manifest)) -or ($null -ne $loadBehavior)
    if (-not $installed) {
        return $false
    }

    if ($StrictLoadBehavior) {
        return ($loadBehavior -eq 3)
    }

    return $true
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

$catalog = Get-AddInCatalog
$selectedAddIns = Get-SelectedAddIns -Catalog $catalog -RequestedAddIns $AddIns

$missing = New-Object System.Collections.Generic.List[string]

foreach ($addIn in $selectedAddIns) {
    $properties = Get-AddInProperties -AddIn $addIn
    $installed = Test-AddInInstalled -Properties $properties -StrictLoadBehavior:$RequireLoadBehavior3 -StrictManifest:$RequireManifest

    if (-not $installed) {
        Write-Output "$($addIn.ProgId): Missing"
        $missing.Add($addIn.ProgId)
        continue
    }

    $installedEnvironment = Get-InstalledEnvironment -Properties $properties

    if (-not [string]::IsNullOrWhiteSpace($Environment)) {
        if ([string]::IsNullOrWhiteSpace($installedEnvironment)) {
            Write-Output "$($addIn.ProgId): Installed, but channel marker could not be determined"
            $missing.Add($addIn.ProgId)
            continue
        }

        if ($installedEnvironment -ne $Environment) {
            Write-Output "$($addIn.ProgId): Installed, but channel mismatch. Expected $Environment, found $installedEnvironment"
            $missing.Add($addIn.ProgId)
            continue
        }
    }

    if ([string]::IsNullOrWhiteSpace($installedEnvironment)) {
        Write-Output "$($addIn.ProgId): Installed"
    }
    else {
        Write-Output "$($addIn.ProgId): Installed ($installedEnvironment)"
    }
}

if ($missing.Count -gt 0) {
    exit 1
}

exit 0
