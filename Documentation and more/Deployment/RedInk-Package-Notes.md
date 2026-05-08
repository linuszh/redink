# Red Ink Office Add-ins – Deployment Notes

## Overview

This package deploys the following Red Ink VSTO add-ins silently for the signed-in Windows user:

- Red Ink for Word
- Red Ink for Excel
- Red Ink for Outlook

Supported channels:

- `GA`
- `Preview`

The package uses:

```text
VSTOInstaller.exe /install <vsto-url> /silent
```

Deployment is per-user (`HKCU`) and must run in user context.

---

# Package Contents

- `Install_RedInkAddins.ps1`
- `Detect-RedInkAddins.ps1`
- `Uninstall-RedInkAddins.ps1`
- `Reset-And-Reinstall-RedInk.ps1`
- `Package-Notes.md`

---

# Requirements

Target systems must have:

- Microsoft Word, Excel and/or Outlook desktop applications
- Microsoft Visual Studio Tools for Office Runtime
- Required .NET Framework version
- Access to `https://redink.ai/apps/...`
- Trusted Red Ink code signing certificates
- Office policies allowing VSTO add-ins

---

# Intune Configuration

## Recommended App Type

```text
Windows app (Win32)
```

## Required Settings

### Install behavior

```text
User
```

### Detection method

```text
Custom detection script
```

---

# Install Commands

## GA

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Install_RedInkAddins.ps1 -Environment GA
```

## Preview

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Install_RedInkAddins.ps1 -Environment Preview
```

---

# Detection Commands

## GA

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Detect-RedInkAddins.ps1 -Environment GA
```

## Preview

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Detect-RedInkAddins.ps1 -Environment Preview
```

## Strict detection

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Detect-RedInkAddins.ps1 -Environment GA -RequireLoadBehavior3
```

---

# Uninstall

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Uninstall-RedInkAddins.ps1
```

---

# Cleanup and Reinstall

Use if:

- another version is already installed
- add-ins do not appear in Office
- `LoadBehavior` changes from `3` to `2`
- Office disabled the add-in

Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Reset-And-Reinstall-RedInk.ps1 -Environment GA
```

---

# Registry Keys

```text
HKCU\Software\Microsoft\Office\Word\Addins\Red Ink for Word
HKCU\Software\Microsoft\Office\Excel\Addins\Red Ink for Excel
HKCU\Software\Microsoft\Office\Outlook\Addins\Red Ink for Outlook
```

Expected:

```text
LoadBehavior = 3
```

---

# Logs

```text
%TEMP%\RedInkAddins-Install.log
%TEMP%\RedInkAddins-Uninstall.log
```

---

# Important Limitation

This deployment model is:

- silent
- per-user
- VSTO/ClickOnce based

It is NOT:

- machine-wide
- MSI-based
- HKLM Office add-in deployment

For machine-wide deployment, use an MSI package with HKLM registration.
