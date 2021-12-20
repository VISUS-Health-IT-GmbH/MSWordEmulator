# MSWordEmulator

![example workflow](https://github.com/VISUS-Health-IT-GmbH/MSWordEmulator/actions/workflows/msbuild.yml/badge.svg)

A simple Windows driver to emulate a subset of COM objects provided by a Microsoft Word installation
to be used in PowerShell using black magic.

Allows to test interaction between any PowerShell script and the provided COM objects without the
need of an actual Microsoft Word installation.


## Scope of emulation

The only emulated COM objects are so far are

- *word.application*
- *word.document*

which can be used to convert a given DOC / DOCX file to PDF using a few functions / subroutines.


## Installation

There is only one way to install this driver yet! It must be installed using *MSBuild* provided by
Microsoft standalone.

The following script should work:
```powershell
#Requires -RunAsAdministrator

<# download vs_buildtools.exe and install MSBuild / .NET 4.8 #>
$client = [System.Net.WebClient]::New()
$client.DownloadFile("https://aka.ms/vs/17/release/vs_BuildTools.exe", "$(Get-Location)\vs_buildtools.exe")
.\vs_buildtools.exe --quiet --add Microsoft.Components.MSBuild --add Microsoft.Net.Component.4.8.TargetingPack

<# download Git and install it #>
$client.DownloadFile("https://github.com/git-for-windows/git/releases/download/v2.34.1.windows.1/Git-2.34.1-64-bit.exe", "$(Get-Location)\git.exe")
.\git.exe /VERYSILENT

<# wait for installation to be done and reload path #>
Start-Sleep -Seconds 30
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 

<# clone repository and save path for MSBuild #>
git clone "https://github.com/VISUS-Health-IT-GmbH/MSWordEmulator.git" repo
$path = "$(Get-Location)/repo/Word.sln"

<# install Word.dll using MSBuild #>
cd "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin"
.\MSBuild.exe $path /p:Platform=x64 /p:Configuration=Release
```


# Usage

After installation you can run the following PowerShell commands to veryfy everything works
correctly:
```powershell
New-Object -ComObject word.application
New-Object -ComObject word.document
```
