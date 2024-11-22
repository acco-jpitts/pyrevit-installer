If ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        & "$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -NonInteractive -NoProfile -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
    Exit
}

####### pyRevit installers
$revit4Versions = @(2020, 2021, 2022, 2023, 2024)
$revit5Versions = @(2025)
$pyrevitinstallerpath = $PSScriptRoot

# Use Get-ChildItem to find installers with wildcard patterns
$pyRevit4Installer = Get-ChildItem -Path $pyrevitinstallerpath -Filter "pyRevit_CLI_4.*_admin_signed.exe" | Select-Object -First 1
$pyRevit5Installer = Get-ChildItem -Path $pyrevitinstallerpath -Filter "pyRevit_5.*_admin_signed.exe" | Select-Object -First 1

# Create a collection of installers found
$pyrevitInstallers = @()
if ($pyRevit4Installer) {
    $pyrevitInstallers += @{
        Version = 4
        Path = $pyRevit4Installer.FullName
    }
} else {
    Write-Output "No pyRevit 4 installer found matching the pattern."
}

if ($pyRevit5Installer) {
    $pyrevitInstallers += @{
        Version = 5
        Path = $pyRevit5Installer.FullName
    }
} else {
    Write-Output "No pyRevit 5 installer found matching the pattern."
}

####### pyRevit variables
$pyrevit4 = "pyRevit-4"
$pyrevit5 = "pyRevit-5"
$pyrevitdeployment = "basepublic"

###### Setup path variables for installers
$pyrevitroot = "C:\pyRevit-Master"
$programFilesPath = ${env:ProgramFiles}
$pyrevit4cli = Join-Path $programFilesPath "pyRevit CLI"
$pyrevit5path = Join-Path $pyrevitroot $pyrevit5
$env:Path = "$env:Path;$pyrevit4cli\bin"

###### Begin Confirm and Create Directories
function Confirm-Path ([string] $targetpath) {
    Write-Output "Confirming $($targetpath)"
    If (Test-Path $targetpath) {
        try {
            Remove-Item -Path $targetpath -Recurse -Force
        } catch {
            Write-Output "Failed to remove directory $targetpath. Error: $_"
        }
    }

    # Create the directory after removing it
    try {
        New-Item -Path $targetpath -ItemType Directory -Force > $null
    } catch {
        Write-Output "Failed to create directory $targetpath. Error: $_"
    }
}

function Create-Directories {
    Confirm-Path $pyrevitroot
    # Confirm path for pyRevit CLI directory for version 4
    Confirm-Path $pyrevit4cli
    # Confirm path for pyRevit 5 directory
    Confirm-Path $pyrevit5path
}

# Create directories
Create-Directories
###### End Confirm and Create Directories

# Install .NET Framework 4.8 if needed
function Install-DotNet48 {
    $dotNetKey = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    $installerPath = "$($PSScriptRoot)\ndp48-devpack-enu.exe"
    
    if ((Get-ItemProperty -Path $dotNetKey -Name Release -ErrorAction SilentlyContinue).Release -lt 528040) {
        if (Test-Path -Path $installerPath) {
            Start-Process -FilePath $installerPath `
                          -ArgumentList "/q /norestart" `
                          -Wait `
                          -Verb RunAs > $null
        } else {
            Write-Output "Installer file ndp48-devpack-enu.exe is missing. Please ensure the installer is in the script directory."
        }
    } else {
        Write-Output ".NET Framework 4.8 is already installed."
    }
}

Install-DotNet48


# Function to Install Executable
function Install-Executable ($installerPath, $installArgs) {
    Write-Output "Installing $installerPath"
    Start-Process -FilePath $installerPath `
                  -ArgumentList $installArgs `
                  -Wait `
                  -Verb RunAs > $null
}

# Install pyRevit 4 CLI and pyRevit 5 WIP
foreach ($installer in $pyrevitInstallers) {
    # Determine the correct installation path for each version
    if ($installer.Version -eq 4) {
        $installLocation = $pyrevit4cli  # Install pyRevit 4 to C:\Program Files\pyRevit CLI
    } elseif ($installer.Version -eq 5) {
        $installLocation = $pyrevit5path  # Install pyRevit 5 to C:\pyRevit-Master\pyRevit-5
    } else {
        Write-Output "Unknown pyRevit version: $($installer.Version). Skipping installation."
        continue
    }

    $installArgs = "/VERYSILENT /NORESTART /DIR=`"$installLocation`""

    Install-Executable "$($installer.Path)" $installArgs
}

function Test-CommandExists {
    param ($command)
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
    try {
        if(Get-Command $command){
            return $true
        }
    }
    catch {
        return $false
    }
    finally {
        $ErrorActionPreference=$oldPreference
    }
}

function Safe-Execute($command, $errorMessage) {
    try {
        Invoke-Expression $command
    } catch {
        Write-Output "$errorMessage. Error Details: $_.Exception.Message"
    }
}

function Attach-PyRevit4($version, $years) {
	foreach ($year in $years) {
		if (Test-Path "C:\ProgramData\Autodesk\Revit\Addins\$year") {
			$command = "pyrevit attach $version DEFAULT $year --allusers"
			Safe-Execute $command "Failed to attach pyRevit $version to Revit $year"
		} else {
			Write-Output "Revit version $year not installed."
		}
	}
}

function Attach-PyRevit5($version, $years) {
	foreach ($year in $years) {
		if (Test-Path "C:\ProgramData\Autodesk\Revit\Addins\$year") {
			$command = "$pyrevit5path\bin\pyrevit.exe attach $version DEFAULT $year --allusers"
			Safe-Execute $command "Failed to attach pyRevit $version to Revit $year"
		} else {
			Write-Output "Revit version $year not installed."
		}
	}
}

####### Clone and Configure pyRevit
if (Test-CommandExists "pyrevit") {
    Safe-Execute "pyrevit revits killall" "Failed to close all Revit processes"
    Safe-Execute "pyrevit clones forget --all" "Failed to forget existing pyRevit clones"
    Safe-Execute "pyrevit clone $pyrevit4 $pyrevitdeployment --dest=$pyrevitroot" `
                 "Failed to clone pyRevit4"
	Attach-PyRevit4 $pyrevit4 $revit4Versions
    Safe-Execute "$pyrevit5path\bin\pyrevit.exe clones add this $pyrevit5" `
                 "Failed to add $pyrevit5"
	Attach-PyRevit5 $pyrevit5 $revit5Versions
}

Read-Host -Prompt "Installation complete. Press Enter to exit"