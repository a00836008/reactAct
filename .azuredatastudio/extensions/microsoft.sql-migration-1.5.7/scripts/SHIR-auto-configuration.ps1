###############################################################################################
# $Description: This PowerShell script automatically downloads self-hosted integration runtime software, install it on windows local machine, and register it with your Azure Database Migration Service.
# $Id: SHIR-auto-configuration.ps1
# $Author: Team - Azure Database Migration Service
###############################################################################################

<#
.SYNOPSIS
Automatically configures SHIR on a Windows machine including downloading, installing Integration runtime,
and configuring IR with DMS.

.DESCRIPTION
The script downloads, installs and configures SHIR on a Windows machine.

.PARAMETER AdminPriv
(Optional) [switch] True if script is running with admin privileges else false.

.PARAMETER AuthKey1
(Optional) [string] First authentication key to register node to DMS.

.PARAMETER AuthKey2
(Optional) [string] Second authentication key to register node to DMS.
#>

param (
	[switch] $AdminPriv = $false,
	[string] $AuthKey1 = $null,
	[string] $AuthKey2 = $null
)

$timeStamp = [System.DateTime]::Now.ToString("yyyyMMddHHmmss")
# Unique script id for every run - Telemetry purpose
$Global:ScriptId = "Script-$timestamp"

# TODO: Auto populate
$Global:LatestIRVersion = [Version]"5.34.8675.1"

# Minimum NuGet version required to install Az.DataMigration module
$Global:MinRequiredNuGetVersion = [Version]"2.8.5.201"

# Log file path
$Global:Logfile = Join-Path $env:USERPROFILE "shir/shir-$timeStamp.log"

# Guid used to create a script lock (to disable concurrent script instances)
$Global:InstallationLockGUID = $MyInvocation.MyCommand.Name

################################### Main Script Block ##################################
$main = {
	$isFirstInstance = $false
	# Create a new Stopwatch object
	$stopwatch = New-Object System.Diagnostics.Stopwatch
	# Start the stopwatch
	$stopwatch.Start()

	# Create a log file
	try {
		New-Item -ItemType File -Path $Global:Logfile -Force | Out-Null
	}
	catch {
		Write-Host "Failed to create log file at $Global:Logfile.`nError: $($_.Exception.Message)" -ForegroundColor Red
		$Global:Logfile = $null
	}

	try {
		Write-OutputAndLog "Script ID: $($Global:ScriptId)`nDescription: This script automatically configures SHIR on a Windows machine including downloading, installing Integration runtime, and configuring IR with DMS."

		# Create a lock file so that only one instance of script can run at a time
		$isFirstInstance = Lock-ScriptInstance
		if ($isFirstInstance) {
			Install-IR
		}
	}
	catch {
		Write-ErrorAndLog "The script could not configure self-hosted integration runtime due to an internal exception. Please try again. `nError: $($_.Exception.Message)"
	}
	finally {
		# Only allow the first running instance to unlock the script
		if ($isFirstInstance) {
			Unlock-ScriptInstance
		}
		# Stop the stopwatch
		$stopwatch.Stop()

		# Write the total time taken
		Write-OutputAndLog ("Total time taken: {0}" -f $stopwatch.Elapsed.ToString("hh\h\:mm\m\:ss\s"))
		if ($null -ne $Global:Logfile) {
			Write-OutputAndLog ("Log file generated: " + $Global:Logfile)
		}
		Write-Host "Press Enter to continue..."
		$null = Read-Host
	}
}

################################### Download, Install and Configure IR ##################################
Function Install-IR {
	Write-OutputAndLog "Checking if Integration Runtime is already installed on this machine..."
	$installedIRVersion = Get-InstalledShirVersion

	if ($null -ne $installedIRVersion) {
		Write-OutputAndLog "Integration Runtime found: Version $installedIRVersion is installed on the machine."

		# Perform validation checks for installed IR
		# if ($installedIRVersion -ne $Global:LatestIRVersion) {
		# 	Write-ErrorAndLog "Installed Integration Runtime's version $($installedIRVersion) does not meet the version $($Global:LatestIRVersion) or above required for configuring."
		# 	return
		# }
		# Write-OutputAndLog "The installed Integration Runtime satisfies all requirements for successful configuration."

		if (-not (Test-InternetConnectivity)) {
			return
		}

		# Configure SHIR - Register IR with the auth key
		Register-IntegrationRuntime -AuthKey1 $AuthKey1 -AuthKey2 $AuthKey2
	}
	elseif (Test-CanInstallIR) {
		$downloadFolder = Join-Path (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path "SHIR-Installer"
		$downloadUrl = "aka.ms/downloadLatestIR"

		# Download the installer
		$packageDownloadPath = Get-Package -packageName "IntegrationRuntime_$Global:LatestIRVersion.msi" `
			-url $downloadUrl `
			-downloadFolder $downloadFolder `
			-Force

		# Installs the IR msi
		$installationWasSuccessful = Install-MsiPackage -packageName "IntegrationRuntime" `
			-packagePath $packageDownloadPath `
			-installationVerificationCallback { return $null -ne (Get-InstalledShirVersion) }

		if ($installationWasSuccessful) {
			# Configure SHIR - Register IR with the auth key
			Register-IntegrationRuntime -AuthKey1 $AuthKey1 -AuthKey2 $AuthKey2
		}
	}
}

################################### Pre-installation checks ##################################
Function Test-CanInstallIR {
	<#
    .DESCRIPTION
    Performs pre-installation checks to ensure that SHIR can be installed, if any check fails a terminating error is thrown.

    .OUTPUTS
    [bool]: True if pre-installation validation is passed else false.
    #>
	try {
		if (-not [Environment]::Is64BitOperatingSystem) {
			throw "Prerequisites are not met. 64-bit operating system is required."
		}
		if (-not (Test-DotNetVersion)) {
			throw "Prerequisites are not met. Dotnet version 4.7.2 or later is required. Please download the supported version from: https://dotnet.microsoft.com/en-us/download/dotnet-framework."
		}
		if ([string]::IsNullOrWhitespace($AuthKey1) -and [string]::IsNullOrWhitespace($AuthKey2)) {
			throw "Minimum one authentication key is required."
		}
	}
	catch {
		Write-ErrorAndLog -exception $_.Exception
		return $false
	}
	return $true
}

Function Test-DotNetVersion {
	<#
    .DESCRIPTION
    Checks if installed .NET Framework version >= 4.7.2.
    https://learn.microsoft.com/en-us/purview/manage-integration-runtimes
    https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed

    .OUTPUTS
    [bool]: true if intalled .NET Framework version >= 4.7.2 else false
    #>
	return (Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release) -ge 461808
}

Function Get-InstalledShirVersion {
	<#
    .DESCRIPTION
    Gets the version of SHIR that is installed on the machine.

    .OUTPUTS
    SHIR version if SHIR is installed, else $null.
    #>
	$InstalledSoftware = Get-ChildItem "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*"
	foreach ($obj in $InstalledSoftware) {
		if ($obj.GetValue('DisplayName') -like "*Microsoft Integration Runtime*") {
			return $obj.GetValue('DisplayVersion')
		}
	}
	return $null
}

################################### Lock/Unlock Script Instance ##################################
Function Lock-ScriptInstance {
	<#
    .DESCRIPTION
    The function creates a lock file so that only one instance of the script can run at a time.
    Return false if already another powershell console has the lock using the script id, else true.
	Using file based locking since it is simpler, and does not have PowerShell version compatibility issues.

    .OUTPUTS
    [bool] True if locking was successful else false.
    #>
	$lockFile = Join-Path ([System.IO.Path]::GetTempPath()) "$($Global:InstallationLockGUID).lk"

	# Lock exists
	if (Test-Path $lockFile) {
		# If file exists then check if the process is open in powershell and still running
		$existingScriptProcessId = [int]::Parse((Get-Content -Path $lockFile))
		$existingScriptProcess = Get-Process -Id $existingScriptProcessId -ErrorAction SilentlyContinue

		# If it is null and not running in powershell then its safe to take a lock.
		# Else show error to user and ask them to close the existing powershell process.
		if ($null -eq $existingScriptProcess -or -not $existingScriptProcess.ProcessName -like “pwsh”) {
			Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue | Out-Null
		}
		else {
			$errorMessage = "Invalid Operation: Only one instance of the script is allowed to run at a time." + `
				"Another powershell console with processid $existingScriptProcessId has a lock for installation." + `
				"If the last installation was aborted in middle, then please close the console or " + `
				"delete the file '$lockFile' to release the lock."

			Write-ErrorAndLog $errorMessage
			return $false
		}
	}

	# Lock acquired
	Write-OutputAndLog "Installation lock acquired."
	Set-Content -Path $lockFile -Value $PID -NoNewline
	return $true
}

Function Unlock-ScriptInstance {
	<#
    .DESCRIPTION
    The function releases the file lock taken by this instance of the powershell console using the script id,
    so that other instances of this script can be run.

    .OUTPUTS
    None
    #>
	$lockFile = Join-Path ([System.IO.Path]::GetTempPath()) "$($Global:InstallationLockGUID).lk"

	if (Test-Path $lockFile) {
		Write-OutputAndLog "Releasing installation lock..."
		Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue | Out-Null
	}
}

############################### Write Output on the console and Log ##################################
Function Write-OutputAndLog {
	<#
    .DESCRIPTION
    Display the message on the console host and also log the message to the script log file.

    .OUTPUTS
    None.
    #>
	Param (
		# Progress message which needs to be logged
		[Parameter()]
		[string]$message
	)

	$logLine = ("{0}: INFO: " -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")) + $message
	$color = if ($message.EndsWith("...")) { "Yellow" } else { "Green" }
	Write-Host "$logLine" -ForegroundColor $color
	if ($null -ne $Global:Logfile) {
		Add-content $Global:Logfile -value $logLine
	}
}

############################## Write Error on the console and Log ##################################
Function Write-ErrorAndLog {
	<#
    .DESCRIPTION
    Display the error message on the console host and also log the error to the script log file.

    .OUTPUTS
    None.
    #>
	Param (
		# Error message which needs to be logged
		[Parameter()]
		[System.Exception] $message,
		# Exception instance which needs to be logged.
		[Parameter()]
		[System.Exception] $exception
	)

	if ($null -ne $message) {
		$logLine = ("{0}: ERROR: {1} `n{2}" -f ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")), $message, "For more information, refer documentation (aka.ms/single-click-SHIR).")
		Write-Host $logLine -ForegroundColor Red
		Write-LogFile -message $logLine
	}

	if ($null -ne $exception) {
		$logLine = "{0}: ERROR: {1} `n{2}" -f ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff"), ($exception.ToString()), "For more information, refer documentation (aka.ms/single-click-SHIR).")
		Write-Host $logLine -ForegroundColor Red
		Write-LogFile -message $logLine
	}
}

################################# Append to log file #####################################
Function Write-LogFile {
	<#
    .DESCRIPTION
    Append to script log file.

    .OUTPUTS
    None.
    #>
	Param (
		# Log message
		[Parameter()]
		[string]$message
	)

	if ($null -ne $Global:Logfile) {
		Add-content $Global:Logfile -value $message
	}
}

############################################ Utils ########################################
Function Get-Package {
	<#
    .DESCRIPTION
    Downloads the installer packages from the url to the local folder. If the installer is already downloaded, then the
    download is aborted unless the force flag is set to true.

    .OUTPUTS
    [string]: Path to the downloaded install package.
    #>
	Param(
		# Full name of the installer package file which will be saved
		[Parameter(Mandatory)]
		[string] $packageName,
		# Url from where to download the installer package
		[Parameter(Mandatory)]
		[string] $url,
		# Download folder where the installer package will be downloaded
		[Parameter(Mandatory)]
		[string] $downloadFolder,
		# Flag which defines if the installer should be downloaded even if it already exists
		[Parameter()]
		[switch] $force = $false
	)

	$packagePath = Join-Path $downloadFolder $packageName
	$shouldDownload = $true

	# If the package already exists then the download flag respects the force flag
	if (Test-Path -path $packagePath) {
		$shouldDownload = $force
	}

	if ($shouldDownload) {
		# Check for internet connection before going forward
		if (-not (Test-InternetConnectivity)) {
			return
		}

		if (Test-Path -Path $packagePath) {
			# Package already exists, force download
			Remove-Item -Path $packagePath | Out-Null
			Write-OutputAndLog "Successfully removed the cached package from the directory: $downloadFolder."
		}

		if (-not (Test-Path -Path $downloadFolder)) {
			# If folder does not exists then create
			New-Item $downloadFolder -ItemType Directory | Out-Null
			Write-OutputAndLog "Successfully created the directory: $downloadFolder."
		}

		Write-OutputAndLog "Downloading from url $url..."
		Write-OutputAndLog "Downloading $packageName package..."

		$stopwatch = New-Object System.Diagnostics.Stopwatch
		$stopwatch.Start()
		Invoke-WebRequest -Uri $url -OutFile $packagePath
		$stopwatch.Stop()

		Write-OutputAndLog "Package $packageName downloaded at $downloadFolder."
		Write-OutputAndLog ("Downloaded {0} bytes in {1}." -f (Get-Item $packagePath).length, $stopwatch.Elapsed.ToString("hh\h\:mm\m\:ss\s"))
	}
	else {
		# Display message that file already exists
		Write-OutputAndLog "Package $packageName found at $downloadFolder."
	}

	return $packagePath
}

Function Test-InternetConnectivity {
	<#
    .DESCRIPTION
    Checks the internet connection by pinging known sites and stops execution if internet connection is
    not available.

    .OUTPUTS
    $true if the connection is available, otherwise $false
    #>
	Param (
		[Parameter(Mandatory = $false)]
		# Microsoft NCSI url for testing internet connectivity
		[string]$Url = "http://www.msftncsi.com/ncsi.txt"
	)

	try {
		$response = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction SilentlyContinue
		if ($null -ne $response) {
			Write-OutputAndLog "Checking for internet connection... Completed."
			return $true
		}
		Write-ErrorAndLog "Prerequisites are not met. Internet connectivity is required on the machine."
		return $false
	}
	catch {
		Write-ErrorAndLog "Prerequisites are not met. Internet connectivity is required on the machine."
		return $false
	}
}

###################################### Check admin access #####################################
Function Test-AdminAccess {
	<#
    .DESCRIPTION
    Checks if the user executing this script has admin access.

    .OUTPUTS
    [bool]: True if user executing the script has admin access else false.
    #>
	$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
	return $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

#################################### Install an MSI package ###################################
Function Install-MsiPackage {
	<#
    .DESCRIPTION
    Installs the msi package and then verifies the installation.

    .OUTPUTS
    [bool]: True if installation was successful, else false
    #>
	Param(
		# Name of the package to be installed
		[Parameter(Mandatory)]
		[string] $packageName,
		# Full path of the package along with executable name
		[Parameter(Mandatory)]
		[string] $packagePath,
		# Function that checks if package is installed
		[Parameter(Mandatory)]
		[scriptblock] $installationVerificationCallback
	)

	try {
		# Check if the installation is already done
		Write-OutputAndLog "Checking existing installation of $packageName package..."
		try {
			$isPackageInstalled = & $installationVerificationCallback -ErrorAction SilentlyContinue
			if ($null -ne $isPackageInstalled -and $isPackageInstalled -eq $true) {
				Write-OutputAndLog "$packageName package is already installed."
				return $true
			}
			else {
				Write-OutputAndLog "$packageName package is not installed."
			}
		}
		catch {
			Write-ErrorAndLog -exception $_.Exception
			return $false
		}

		# Add necessary arguments to install quietly with no UI
		# For no UI /qn, basic UI /qb, reduced UI /qr, full UI /qf
		$argsx = @('/i', $packagePath, '/quiet', '/qb')

		# Execute installation process
		Write-OutputAndLog "Installing $packageName package..."

		$stopwatch = New-Object System.Diagnostics.Stopwatch
		$stopwatch.Start()
		$process = Start-Process msiexec -Wait -NoNewWindow -ErrorAction SilentlyContinue -ArgumentList $argsx
		$stopwatch.Stop()

		Write-OutputAndLog "MSI Installation process exited with code: $($process.ExitCode)"
		Write-OutputAndLog "Installation completed in $($stopwatch.Elapsed.ToString("hh\h\:mm\m\:ss\s"))."

		# No installation verification callback provided
		if ([string]::IsNullOrWhiteSpace($installationVerificationCallback.ToString())) {
			return $true
		}

		# Verifying the installation
		Write-OutputAndLog "Verifying $packageName installation..."
		try {
			$installationResult = & $installationVerificationCallback
			if ($installationResult -eq $true) {
				Write-OutputAndLog "$packageName installation verification successful."
				return $true
			}
		}
		catch {
			Write-ErrorAndLog "Failed to verify $packageName installation. Check for errors. Error: $($_.Exception.Message)"
			return $false
		}
	}
	catch {
		Write-ErrorAndLog -exception $_.Exception
		return $false
	}
	return $true
}

#################################### Register Node to DMS ###################################
Function Register-IntegrationRuntime {
	<#
    .DESCRIPTION
    Configures the SHIR with the auth key. It registers the IR to the DMS using the authentication keys.

    .OUTPUTS
    None.
    #>
	param (
		[Parameter()]
		[string]$AuthKey1,

		[Parameter()]
		[string]$AuthKey2
	)

	if ([string]::IsNullOrWhitespace($AuthKey1) -and [string]::IsNullOrWhitespace($AuthKey2)) {
		Write-ErrorAndLog "Minimum one authentication key is required."
		return
	}

	$azDataMigrationModule = Get-Module -ListAvailable -Name Az.DataMigration

	try {
		if (-not $azDataMigrationModule) {
			# Install the NugetProvider for Az.DataMigration
			$nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue

			if ($null -eq $nuget -or $nuget.Version -lt $Global:MinRequiredNuGetVersion) {
				Write-OutputAndLog "Installing NuGet Package Provider..."
				$stopwatch = New-Object System.Diagnostics.Stopwatch
				$stopwatch.Start()
				Install-PackageProvider -Name NuGet -MinimumVersion $Global:MinRequiredNuGetVersion -Force
				$stopwatch.Stop()
				Write-OutputAndLog "Installation completed in $($stopwatch.Elapsed.ToString("hh\h\:mm\m\:ss\s"))."
			}

			# Install the Az.DataMigration module
			Write-OutputAndLog "Installing Az.DataMigration Module..."
			$stopwatch = New-Object System.Diagnostics.Stopwatch
			$stopwatch.Start()
			Install-Module -Name Az.DataMigration -Force -AllowClobber
			$stopwatch.Stop()
			Write-OutputAndLog "Installation completed in $($stopwatch.Elapsed.ToString("hh\h\:mm\m\:ss\s"))."
		}
	}
	catch {
		Write-ErrorAndLog -exception $_.Exception
		return
	}

	# Try the first authentication key
	try {
		if (-not [string]::IsNullOrWhitespace($AuthKey1)) {
			Write-OutputAndLog "Trying the first authentication key..."
			Register-AzDataMigrationIntegrationRuntime -AuthKey $AuthKey1 *> $null
			Write-OutputAndLog "Registration successful with the first authentication key."
			return
		}
	}
	catch {
		Write-ErrorAndLog "Failed to register with the first authentication key. Error: $($_.Exception.Message)"
	}

	# Try the second authentication key
	try {
		if (-not [string]::IsNullOrWhitespace($AuthKey2)) {
			Write-OutputAndLog "Trying the second authentication key..."
			Register-AzDataMigrationIntegrationRuntime -AuthKey $AuthKey2 *> $null
			Write-OutputAndLog "Registration successful with the second authentication key."
			return
		}
	}
	catch {
		Write-ErrorAndLog "Failed to register with the second authentication key. Error: $($_.Exception.Message)"
	}
}

###################################### Main Execution #####################################
if (-not (Test-AdminAccess)) {
	if ($AdminPriv) {
		Write-ErrorAndLog -exception "Failed to gain admin privileges."
	}
	# Gain admin access by prompting the user, if the script is executed without admin privileges.
	else {
		$commandLineArgs = "-noprofile -noexit -executionpolicy unrestricted -file $($MyInvocation.MyCommand.Definition) -admin-priv"

		if (-not [string]::IsNullOrWhitespace($AuthKey1)) {
			$commandLineArgs += " -authKey1 $AuthKey1"
		}
		if (-not [string]::IsNullOrWhitespace($AuthKey2)) {
			$commandLineArgs += " -authKey2 $AuthKey2"
		}

		Start-Process powershell.exe -Verb RunAs -ArgumentList $commandLineArgs
	}
	exit
}

& $main

# SIG # Begin signature block
# MIIoRgYJKoZIhvcNAQcCoIIoNzCCKDMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8TNQIyr2MhAN0
# fqRedSAqnIM07EcMZOTfbMq3p8FmDqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGiYwghoiAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIC2QWv5KiEnr3o+5dgOG/OGf
# XgbqanhDtji2WHnzNIjYMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAPmZ73Gouvi34tynkmVf75vAbFdscytzup5GrWFiU2lbNjvaKZrTUuMd9
# mCgBkc4syFxLvO0C1CxHBkC1Q8ccoenNviXTGZhqBE0ixx1swk3OmCefd0VJtPoW
# 4ngEsF6vnxH1Pn3CBuZd6bYCaUMlHe0XIm18aSlhWX5wIegiC8yBt6UoU+av0v8u
# d2t2lyDlkrrTfHsl9XmsJKa4Akt2JVR1bzPdadhnpdK6IqVGrwp8mhSXVtkczueg
# af4u7JxyS8yr+pKBdEsyyxAjBJs5uWobZ8NjB4yccdWVVVgFeFNSLCtbCVGKQz+z
# 1sphunPdMUYuRNmM/+XUHxRgUmRiK6GCF7AwghesBgorBgEEAYI3AwMBMYIXnDCC
# F5gGCSqGSIb3DQEHAqCCF4kwgheFAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFaBgsq
# hkiG9w0BCRABBKCCAUkEggFFMIIBQQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCBULHZ3ySvnJsv/JAlY/Z/SulGGp81n8+hqi+2ygllFOQIGZzvBNTok
# GBMyMDI0MTEyNjEyMjgyNi41MjZaMASAAgH0oIHZpIHWMIHTMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVT
# TjoyRDFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# U2VydmljZaCCEf4wggcoMIIFEKADAgECAhMzAAAB/XP5aFrNDGHtAAEAAAH9MA0G
# CSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI0
# MDcyNTE4MzExNloXDTI1MTAyMjE4MzExNlowgdMxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9w
# ZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjJEMUEt
# MDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAoWWs+D+Ou4JjYnRHRedu
# 0MTFYzNJEVPnILzc02R3qbnujvhZgkhp+p/lymYLzkQyG2zpxYceTjIF7HiQWbt6
# FW3ARkBrthJUz05ZnKpcF31lpUEb8gUXiD2xIpo8YM+SD0S+hTP1TCA/we38yZ3B
# EtmZtcVnaLRp/Avsqg+5KI0Kw6TDJpKwTLl0VW0/23sKikeWDSnHQeTprO0zIm/b
# tagSYm3V/8zXlfxy7s/EVFdSglHGsUq8EZupUO8XbHzz7tURyiD3kOxNnw5ox1eZ
# X/c/XmW4H6b4yNmZF0wTZuw37yA1PJKOySSrXrWEh+H6++Wb6+1ltMCPoMJHUtPP
# 3Cn0CNcNvrPyJtDacqjnITrLzrsHdOLqjsH229Zkvndk0IqxBDZgMoY+Ef7ffFRP
# 2pPkrF1F9IcBkYz8hL+QjX+u4y4Uqq4UtT7VRnsqvR/x/+QLE0pcSEh/XE1w1fcp
# 6Jmq8RnHEXikycMLN/a/KYxpSP3FfFbLZuf+qIryFL0gEDytapGn1ONjVkiKpVP2
# uqVIYj4ViCjy5pLUceMeqiKgYqhpmUHCE2WssLLhdQBHdpl28+k+ZY6m4dPFnEoG
# cJHuMcIZnw4cOwixojROr+Nq71cJj7Q4L0XwPvuTHQt0oH7RKMQgmsy7CVD7v55d
# OhdHXdYsyO69dAdK+nWlyYcCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBTpDMXA4ZW8
# +yL2+3vA6RmU7oEKpDAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBf
# BgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmww
# bAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0El
# MjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAY9hYX+T5AmCr
# YGaH96TdR5T52/PNOG7ySYeopv4flnDWQLhBlravAg+pjlNv5XSXZrKGv8e4s5dJ
# 5WdhfC9ywFQq4TmXnUevPXtlubZk+02BXK6/23hM0TSKs2KlhYiqzbRe8QbMfKXE
# DtvMoHSZT7r+wI2IgjYQwka+3P9VXgERwu46/czz8IR/Zq+vO5523Jld6ssVuzs9
# uwIrJhfcYBj50mXWRBcMhzajLjWDgcih0DuykPcBpoTLlOL8LpXooqnr+QLYE4Bp
# Uep3JySMYfPz2hfOL3g02WEfsOxp8ANbcdiqM31dm3vSheEkmjHA2zuM+Tgn4j5n
# +Any7IODYQkIrNVhLdML09eu1dIPhp24lFtnWTYNaFTOfMqFa3Ab8KDKicmp0Ath
# RNZVg0BPAL58+B0UcoBGKzS9jscwOTu1JmNlisOKkVUVkSJ5Fo/ctfDSPdCTVaIX
# XF7l40k1cM/X2O0JdAS97T78lYjtw/PybuzX5shxBh/RqTPvCyAhIxBVKfN/hfs4
# CIoFaqWJ0r/8SB1CGsyyIcPfEgMo8ceq1w5Zo0JfnyFi6Guo+z3LPFl/exQaRubE
# rsAUTfyBY5/5liyvjAgyDYnEB8vHO7c7Fg2tGd5hGgYs+AOoWx24+XcyxpUkAajD
# hky9Dl+8JZTjts6BcT9sYTmOodk/SgIwggdxMIIFWaADAgECAhMzAAAAFcXna54C
# m0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
# Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMy
# MjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51
# yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY
# 6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9
# cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN
# 7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDua
# Rr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
# kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2
# K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5
# TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZk
# i1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9Q
# BXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3Pmri
# Lq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUC
# BBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9y
# eS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUA
# YgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
# 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2Ny
# bC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIw
# MTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/yp
# b+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulm
# ZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM
# 9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECW
# OKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4
# FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3Uw
# xTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
# fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVX
# VAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGC
# onsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU
# 5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEG
# ahC0HVUzWLOhcGbyoYIDWTCCAkECAQEwggEBoYHZpIHWMIHTMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVT
# TjoyRDFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# U2VydmljZaIjCgEBMAcGBSsOAwIaAxUAoj0WtVVQUNSKoqtrjinRAsBUdoOggYMw
# gYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsF
# AAIFAOrwInswIhgPMjAyNDExMjYxMDMzMzFaGA8yMDI0MTEyNzEwMzMzMVowdzA9
# BgorBgEEAYRZCgQBMS8wLTAKAgUA6vAiewIBADAKAgEAAgImzAIB/zAHAgEAAgIT
# oDAKAgUA6vFz+wIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAow
# CAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBCwUAA4IBAQB0vy9oR0z/
# 1ZJn3j1P8JhwTVZpeCBDKH8bA0Acx461kQ62GjQYpVBFd69HrixfxOAY/BgpRrek
# OZPsalVE5JNlIhKoEZwNYrKtrHu6Dw945qFE5/EXOwJ6PaMdJkaGzC3WnOqh8LNA
# c6FuS3VPtoDUCN6Nb69r0tQ1OxiZ0yYxs4QU64BfJYMqcyYf/sn2ak+TtZ2VNi7J
# 1JyYWMQUBNbCq4MqOgM7mA0QDZ7jJCeh1DQSHaIi9qmRY5nPigDO7fSVXYS3bxex
# GjfYZ5BerM0mDziyIXVqEWMSXiHvxKLZez7IAOZTxIxHS+rBWCtd39lXZRH+VPvW
# Fr8ZM9INhNDyMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAH9c/loWs0MYe0AAQAAAf0wDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQg4pKFgiqT
# 79vP4tA9RokUdCFL0Ng7aALCziq+PzBNanIwgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCCAKEgNyUowvIfx/eDfYSupHkeF1p6GFwjKBs8lRB4NRzCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB/XP5aFrNDGHtAAEA
# AAH9MCIEINnztOp5U1v/NdEbTTyb1TAmENaLjsSUbeh3Rp02B5bzMA0GCSqGSIb3
# DQEBCwUABIICADZrJIrNQktnuWnYxLS7091zZWlb0zHvzTJjNDpV6GOGMWTbaApx
# 5XkveSSSj7pZfEFWhELhhAb5Y24ufcDx0AtJzhl/hF2W490VuxkOEUbOuV6D9lO7
# Ik/oydtWXo6KBaiuZGqiMWBIPLmAqYKlpjWZxif5Y+ZEVlSH4o7m67C1mLWzcmhR
# 6HfWecVFnRnOhi6wiWV5fw/BfLT6uurKeySNqvrtqXzn/+DUmEL/qkTss0cOYb/z
# PObugjaSkjvPan93nE5Be2QJGakRoJpPGmqqJ3XoVNEMpZTNkNMbuNOaAXJTduPV
# pCGSO/Xd/+uke3GfH/MyIswm/WuQT0Fvko+5xJOGxGMVV0SvIZC6XZgAIeR8pRtj
# ueQC0vGN/tYtfX/1I5GB0LiPJG0GPOwX3OoUg8ddTHLQk+vu3To8oh76OMaFofib
# iGp5ib8ZKPcvMeZ0yAnhFG3KQb7GUOa8MOgXfTSGWQC7Ju6BEg1UjWgeofEErJ5h
# HLvNnlAEMsGpyLAI6F5KRVJ+GhH8jjGHLGQuxAnMPf1cfClIyejPh5btdj/bPULV
# 3CvtTuhNIEzkujhWGmxZ9qxX13RIMlnE9Dk5dGekzMg5K+NzcFFVfjMmzkMUgX+M
# v659TfbMHzKj/AoLUmZ/1UMwCSahTF51z29l1xqQat+M544nX3ifnEXx
# SIG # End signature block
