# .NETCHECKER

## Description

This project contains a PowerShell script (**dotnetchecker.ps1**) that performs the following tasks:

- Scans the computer for installed .NET versions (both .NET Framework and .NET Core/5+).
- Logs the installed versions to a CSV file.
- Determines which installed versions are End-Of-Support (EOL) based on predefined dates.
- Creates an Azure DevOps user story listing only the EOL versions.
- Downloads and installs the latest .NET version.

## Usage

Run the script using PowerShell:

```Powershell

.\dotnetchecker.ps1