<div align="center"><h1> ｅｄｃｔｏｏｌｋｉｔ</h1>
<img src="Scripts/EDCtoolkit/logo-n.png"> <alt="EDCtoolkit logo" width="160">>PoSh toolkit for endpoint triage and troubleshooting.
<br> More troubleshooting assistance coming</div></h4>
<p align="center"></p>
<br><br></c>


## What It Includes

- `Scripts/EDCtoolkit/edctoolkit.ps1`: legacy CLI toolkit (deprecated, retained for compatibility)
- `Scripts/EDCtoolkit/EDCtoolkit.GUI.ps1`: WinForms GUI wrapper
- `Scripts/NetworkMatrix/networkmatrix.ps1`: TUI network survey and camera/NVR deployment mapper
- `EDCtoolkit.cmd`: launcher convenience script
- `EDCtoolkit.GUI.vbs`: hidden-window GUI launcher used by the `.cmd` entrypoint

## Quick Start

### GUI

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\EDCtoolkit\EDCtoolkit.GUI.ps1 -Theme Dark
```

### NetworkMatrix

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\NetworkMatrix\networkmatrix.ps1
```

### AppX Package Build

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Packaging\EDCToolkitApp\build-appx.ps1
```

The current package output is written to:

`Packaging/EDCToolkitApp/output/EDCToolkit_0.1.0.0.appx`

### Legacy CLI

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\EDCtoolkit\edctoolkit.ps1
```

## Reports

Reports are saved under:

`Scripts/EDCtoolkit/EDC_Reports`

If the toolkit is installed from a packaged or otherwise read-only location, it falls back to:

`%LOCALAPPDATA%\EDCtoolkit\EDC_Reports`

The GUI can also export a combined report to a file you choose.

`NetworkMatrix` writes its survey packages under:

`Scripts/NetworkMatrix/NetworkMatrix_Reports`

## Notes

- Run as Administrator for the most complete results and fix actions.
- The GUI is non-interactive for scan checks (no hidden terminal prompts).
- The GUI is the supported primary entrypoint for packaging and distribution. The CLI remains available but is deprecated.
- `EDCtoolkit.cmd` now launches the GUI without a visible console window.
- `Packaging/EDCToolkitApp/build-appx.ps1` builds a test AppX package and currently leaves it unsigned if local test certificate creation fails.

<img src="Scripts/EDCtoolkit/hardy-the-hdd.png">
