# GHCP_setup.ps1

Automated development environment setup script for GitHub Copilot workflows at Intel.
Installs and configures all required tooling in a single, unattended run.

## What It Does

| Component | Action |
|---|---|
| **WinGet** | Bootstraps the Windows Package Manager if missing |
| **Miniforge (conda)** | Installs the conda-forge distribution of conda |
| **Python 3.11 (`py_learn`)** | Creates a dedicated conda environment |
| **Python packages** | Installs data-science and document-processing packages (see below) |
| **Git** | Installs via WinGet |
| **GitHub CLI** | Installs via WinGet and handles authentication |
| **Intel dt** | Downloads and bootstraps Intel's internal `dt` developer tool |
| **VS Code extensions** | Installs `GitHub.copilot`, `GitHub.copilot-chat`, `ms-python.python`, `ms-toolsai.jupyter` |
| **VS Code proxy** | Writes proxy settings to `%APPDATA%\Code\User\settings.json` |
| **Intel certificates** | (Optional) Imports Intel SHA-2/SHA-384 cert bundles into the Windows trust store |

> **Note:** VS Code itself must be installed manually from the Company Portal before running this script.

### Python Packages Installed in `py_learn`

- **Core data science:** `numpy`, `pandas`, `scipy`, `matplotlib`, `seaborn`, `scikit-learn`
- **Notebooks:** `jupyter`, `ipykernel`
- **Excel / spreadsheets:** `openpyxl`, `xlrd`, `xlsxwriter`
- **Word / PowerPoint:** `python-docx`, `python-pptx`
- **PDF processing:** `pypdf`, `reportlab`, `pymupdf`
- **Images:** `pillow`
- **Utilities:** `requests`, `tqdm`, `rich`

---

## Requirements

- Windows 10 or later (x64)
- PowerShell 5.1 or later (`#Requires -Version 5`)
- VS Code installed (from Company Portal)
- Internet access (direct or via Intel proxy)
- Administrator privileges required **only** when using `-InstallIntelCerts`

---

## Quick Start

Open a PowerShell window (does **not** need to be Admin for the default run) and execute:

```powershell
.\GHCP_setup.ps1
```

The script will run all steps automatically and print a live progress bar and per-step timing table.

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Proxy` | `string` | `http://proxy-chain.intel.com:912` | Proxy URL written to VS Code settings and shown in manual instructions |
| `-NoProxy` | `string` | `localhost,127.0.0.0/8,...` | NO_PROXY list (informational) |
| `-InstallIntelCerts` | `switch` | off | Import Intel SHA-2/SHA-384 certificate bundles (**requires Admin**) |
| `-InstallBuildTools` | `switch` | off | Install Visual Studio Build Tools for compiling native Python packages |
| `-UseExternalPyPI` | `switch` | off | Prefer external PyPI during package installs (informational) |
| `-Interactive` | `switch` | off | Pause at each step and prompt Y / N / Skip |
| `-DryRun` | `switch` | off | Print all planned actions without executing any of them |
| `-SkipVSCode` | `switch` | off | Skip VS Code installation step |
| `-SkipPython` | `switch` | off | Skip Miniforge installation |
| `-SkipGit` | `switch` | off | Skip Git installation |
| `-SkipGitHubCLI` | `switch` | off | Skip GitHub CLI installation |
| `-SkipPythonPackages` | `switch` | off | Skip Python package installation |
| `-SkipExtensions` | `switch` | off | Skip VS Code extension installation |
| `-SkipGitHubAuth` | `switch` | off | Skip GitHub CLI authentication step |
| `-SkipDt` | `switch` | off | Skip Intel dt installation and setup |

---

## Usage Examples

```powershell
# Standard run (all defaults)
.\GHCP_setup.ps1

# Install Intel certificates (requires elevated PowerShell)
.\GHCP_setup.ps1 -InstallIntelCerts

# Interactive mode — confirm each step before it runs
.\GHCP_setup.ps1 -Interactive

# Dry run — show everything that would happen without making changes
.\GHCP_setup.ps1 -DryRun

# Custom proxy (e.g. for a different Intel site)
.\GHCP_setup.ps1 -Proxy "http://proxy.iind.intel.com:911"

# Skip components already installed
.\GHCP_setup.ps1 -SkipGit -SkipGitHubCLI
```

---

## After the Script Completes

1. **Git proxy** (if needed):
   ```powershell
   git config --global http.proxy http://proxy-chain.intel.com:912
   git config --global https.proxy http://proxy-chain.intel.com:912
   ```

2. **GitHub Copilot** — ensure you have a Copilot entitlement and complete any required Intel onboarding.

3. **Open VS Code** and sign in to GitHub when prompted by the Copilot extensions.

4. **Select `py_learn` as your Python interpreter:**
   `Ctrl+Shift+P` → *Python: Select Interpreter* → choose the `py_learn` conda environment.

5. **Proxy in command sessions** (for pip / conda installs run manually):
   ```cmd
   set http_proxy=http://proxy-chain.intel.com:912
   set https_proxy=http://proxy-chain.intel.com:912
   ```

---

## Logs

A full transcript is written to `%TEMP%\GHCPSetup_<timestamp>.log` and the path is printed at the end of each run.

---

## Changelog

| Version | Changes |
|---|---|
| **v6.0** | Added Intel `dt` bootstrap flow; automated `py_learn` environment creation and package installation directly from PowerShell; updated banners and summary |
| **v2** | Bug fixes (B-01..B-08), pre-flight checks, interactive mode, per-package pip install, version detection, visual progress, dry-run support, per-component skip flags, PS 5.1 compatibility |

---

## License

This script is intended for internal Intel use. Python packages are installed from open-source distributions; refer to each package's individual license for terms.
