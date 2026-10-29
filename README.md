# Automating patch management for a hybrid Windows/SaaS SME

Fictional case: Coastal Managed IT (~60 staff modelled). Free/local tools only for the laboratory: PowerShell inventory, winget, and an Essential Eight patch reporting workbook. No paid Intune required for this pilot.

## Research questions

1. How can an SME automate OS and third-party patching affordably?
2. What patch cycle balances speed with stability?
3. How is coverage and timeliness best measured for Essential Eight?

## Scripts

1. `01_Inventory_Installed_Software.ps1` — local installed apps (registry + winget list).
2. `02_Find_Outdated_Packages.ps1` — winget upgrade candidates (what can be automated).
3. `03_Apply_Safe_Upgrades.ps1` — optional apply with `-WhatIf` default; use `-Apply` to install.
4. `04_Build_Patch_Report.ps1` — writes JSON evidence for coverage and age.

Run on a lab Windows machine only. Do not point at production endpoints without change control.
