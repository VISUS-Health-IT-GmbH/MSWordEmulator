# MSWordEmulator

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
Microsoft standalone without Visual Studio Community 2022.

The following script should work:
```powershell
#Requires -RunAsAdministrator

Invoke-WebRequest "https://aka.ms/vs/17/vs_buildtools.exe" -OutFile "vs_buildtools.exe"
.\vs_buildtools.exe --quiet --add Microsoft.Components.MSBuild --add Microsoft.Net.Component.4.7.2.TargetingPack

git clone "https://github.com/VISUS-Health-IT-GmbH/MSWordEmulator.git" Word
$path = "$(Get-Location)/Word/Word/Word/Word.vbproj"

cd "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin"
.\MSBuild.exe $path
```


# Usage

After installation you can run the following PowerShell commands to veryfy everything works
correctly:
```powershell
$word = New-Object -ComObject word.application
$word.visible = $False

$doc = $word.documents.open("test.doc")
$doc.saveas("test.pdf", 17)
$doc.close()
```
