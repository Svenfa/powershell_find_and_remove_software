# Written by Sven Falk - 23~25 of May 2018
# Description: This script will search for installed Software and is able to remove it if a correct uninstall string is found in the registry
# ------ ToDo:
# - Return Exitcodes of MSI to Matrix-Package for better handling?
# - Return Errormessages to Matrix-Package for better log-entries in Empirum
# - Search for software after uninstallation to make sure uninstallation was successful
# - No correct way to return 11000 - need info if NO software was found
# ------------------------------------------------------- Define environment -------------------------------------------------------
# Param has to be the first line!
# Defines the parameters which are given by calling this script:
# e.g.: .\check_previous_Versions.ps1 -software "7-Zip" -version "16.01" -uninstall "yes" -debug 1
param (
    [string]$software = "",
    [string]$version = "0.0",
    [string]$uninstall = "no",
    [int]$debug = 0,
    [string]$OutputFileLocation = "$env:Temp\check_previous_Versions_debug_output_$(get-date -f yyyy.MM.dd-H.m).log"
)

# ---- Exit Codes ----
# Setup-routines will exit with their own exit-codes.
# Define some custom exit-codes for this script.
#   11000 | Software not found - exited without any action
#   11001 | Undefined error
#   11002 | removed found software-version
#   11003 | found previous version but could not identify uninstall-routine
#	11004 | found software, but deinstallation was disabled
#	11005 | found software, but in recent version
#   11010 | Found Inno-setup but uninstall-exe was not found

# ---- Debugging ----
# Enable debugging (1) or disable (0)
# Powershelldebugging:
Set-PSDebug -Trace 0
# Enable Debug-Write-Host-Messages:
$DebugMessages = $debug
#
# Send all Write-Host messages to console and to the file defined in $OutputFileLocation
if ($DebugMessages -eq "1") {
    # Stop transcript - just in case it's running in another PS-Script:
    $ErrorActionPreference="SilentlyContinue"
    Stop-Transcript | out-null
    # Start transcript of all output into a file:
    $ErrorActionPreference = "Continue"
    Start-Transcript -path $OutputFileLocation -append
}

# ------------------------------------------------------- End definition of environment ---------------------------------------------------


# ------------------------------------------------------- Define functions ----------------------------------------------------------------

function endscript{
    # Debug info:
    if ($DebugMessages -eq "1") {Write-Host "Result is exitcode:" $exitcode }
    if ($DebugMessages -eq "1") {Write-Host "End of script"}
    if ($DebugMessages -eq "1") {Stop-Transcript}
    exit $exitcode
    }

# Define what to do with found entries. Put it into a function so it can be called for 32 and 64-bit searches
function do_compare{

    # Debug info:
    if ($DebugMessages -eq 1) {Write-Host "--- Beginning of loop ---`nFound" $_.DisplayName "in version" $_.DisplayVersion}

    # Check if found software is matrix-package:
    if ("$_." -like "*Setup.inf*") { 
        # Debug info:
        if ($DebugMessages -eq "1") {Write-Host "This is a matrix package entry and not the installed software `nSkipping it."}

        # Debug info:
        if ($DebugMessages -eq "1") {Write-Host "--- End of loop ---`n"}
        # Skip this found software and exit the function (but not the whole script)
        return;
    }

    # Compare installed version
    if ([System.Version]$_.DisplayVersion -lt [System.Version]"$version") {

		# Debug info:
		if ($DebugMessages -eq "1") {Write-Host "Found older version of" $_.DisplayName "(Given parameter version:" $version "> Installed version:" $_.DisplayVersion ")"}

		if($uninstall -eq "yes") {
			
			# Checking uninstall-string type:
			if ($_.UninstallString -match "MsiExec.exe") {

				##------------------------------------------------------- MSI Uninstallation -------------------------------------------------------
				# Debug info:
				if ($DebugMessages -eq "1") {Write-Host "Found MSI-Uninstallstring."}

				# Modify '/I' argument to '/X'
				if ($_.UninstallString -match "MsiExec.exe /I") {
					$_.UninstallString = $_.UninstallString -replace "/I","/X"
					# Debug info:
					if ($DebugMessages -eq "1") {Write-Host "Found MSI-UninstallString and replaced /I with /X - New UninstallString:" $_.UninstallString}
				}

				# Debug Info:
				if ($DebugMessages -eq "1") {Write-Host "Found UninstallString and extracted the parameters:" $_.UninstallString}

				if ($_.UninstallString -match "MsiExec.exe /X") {

					# Extract arguments to uninstall (We have to define arguments separated to use start-process)
					$_.UninstallString = $_.UninstallString -replace "MsiExec.exe ",""
					
					Start-Process MsiExec.exe -wait -ArgumentList "$($_.UninstallString)","/qn","/log $env:Windir\Temp\Uninstalled_$_.Displayname.log" -PassThru
                    $exitcode = $LASTEXITCODE
					if ($DebugMessages -eq "1") {Write-Host "Ran MsiExec.exe and got errorlevel:" $exitcode}
					
					# Check returncode of msiexec. If it's not 0, exit this script.
					if ($exitcode -ne "0") {
                        if ( $exitcode -eq "3010" ) {
                                return;
                        }
                        if ( $exitcode -eq "XXXX" ) {
                                return;
                        }
                        if ( $exitcode -eq "YYYY" ) {
                                return;
                        }
                        exit $exitcode 
                    } else {
                        return;
                    }
				}
				##------------------------------------------------------- End MSI Uninstallation ----------------------------------------------------

			} 	elseif ($_.UninstallString -match "unins000.exe") {

					##------------------------------------------------------- Inno Uninstallation -------------------------------------------------------
					# Debug info:
					if ($DebugMessages -eq "1") {Write-Host "Found Inno-Setup uninstallation"}
					
					# Check if uninstallation file still exists:
					# Test-Path does not like quotes, so we will trim them and save the path into a new variable:
					$InnoFilePath=$_.UninstallString.Replace("`"","")

					# Debug info:
					if ($DebugMessages -eq "1") {Write-Host "Uninstallation activated via arguments"}
					if (Test-Path -Path $InnoFilePath) {
						# Debug info:
						if ($DebugMessages -eq "1") {Write-Host "Inno-Setup-File exists in" $_.UninstallString }

						Start-Process "$InnoFilePath" -ArgumentList "/VERYSILENT","/SUPPRESSMSGBOXES","/LOG=$env:Windir\Temp\Uninstalled_$_.Displayname.log"
						# Check returncode of uninstallation-file. If it's not 0, exit this script.
						if ($LASTEXITCODE -ne "0") { exit $LASTEXITCODE } else {return;}
					} else {
						# Debug info:
						if ($DebugMessages -eq "1") {Write-Host "Could not find Inno-Setup-File in" $_.UninstallString }
						exit 11010
					}
					
				##------------------------------------------------------- End Inno Uninstallation -------------------------------------------------------
			} 	else {
					# Debug info:
					if ($DebugMessages -eq "1") {Write-Host "Found old version of software but could not identify uninstallation-routine" }
					exit 11003
				}

		} else {
			# Debug info:
			if ($DebugMessages -eq "1") {Write-Host "Found software - Uninstallation disabled via arguments." }
			exit 11004
		}
			
			
    } else {
        # Debug info:
        if ($DebugMessages -eq "1") {Write-Host "No older version of" $_.DisplayName "found. (Given parameter version:" $version " - Installed version:" $_.DisplayVersion ")"}
        $exitcode=11005
        endscript
    }

    # Debug info:
    if ($DebugMessages -eq "1") {Write-Host "--- End of loop ---`n"}

}
# ------------------------------------------------------- End definition of functions ---------------------------------------------------
#
#
# -------------------------------------------------------------------- Tasks  -------------------------------------------------------------
# Initiate $counter-Variable
$counter=0
# Search 32-Bit Software
Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*$software*"} | ForEach-Object -process { 
    do_compare
    $counter++
}

# 64-Bit Software
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*$software*"} | ForEach-Object -process { 
    do_compare
    $counter++
}

if ( $counter -ne 0 ) {
<<<<<<< HEAD
    if ($DebugMessages -eq "0") {Write-Host "No version of" $_.DisplayName "found." | Out-File -FilePath C:\temp\MYFILE.log -Append}
=======
    if ($DebugMessages -eq "1") {Write-Host "No version of" $_.DisplayName "found." | Out-File -FilePath C:\temp\MYFILE.log -Append}
>>>>>>> f12b560cfb77e036b0844ce8f14e42531a0d8c17
    exit 11000
}

endscript