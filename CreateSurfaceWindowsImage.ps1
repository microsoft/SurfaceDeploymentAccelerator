<#
.SYNOPSIS
    This script creates a Surface Windows image.

.DESCRIPTION
    This script creates a Surface Windows image, including Office 365 and requisite Visual C runtime libraries as required.
    
.EXAMPLE
    .\CreateSurfaceWindowsImage.ps1 -ISO <ISO path> -OSSKU Pro -DestinationFolder "C:\Temp" -Device SurfacePro7

.NOTES
    Author:       Microsoft
    Last Update:  20th October 2020
    Version:      1.2.5.4

    Version 1.2.5.4
    - Added support for Surface Laptop Go
    - Added support for Windows 10 20H2.
    - Added language support:
      - Chinese (Simplified and Traditional)
      - French
      - Russian

    Version 1.2.5.3
    - Fixed default PowerShell execution policy to match OOB defaults (Restricted)

    Version 1.2.5.2
    - Split prereq installation check into two checks for ADK installation to avoid WinPE not installed bug after ADK install check succeeds
    - Office 365 does not install if not passed full path parameter, added explicit check and handler
    
    Version 1.2.5.1
    - Fixed typos in code causing first-run errors
    
    Version 1.2.5
    - Prevent usage of spaces in file paths for DestinationFolder and LocalDriverPath
    - Prevent script from executing if prior execution failed at a specific point

    Version 1.2.4
    - Support for ESD file format added

    Version 1.2.3
    - LocalDriverPath can now point to a flat (extracted) driver path, or a Surface platform MSI file
    - Logging functionality added
    - Changed all Get-WmiObject calls with Get-CimInstance calls to be more compatible with PowerShell Core
    - Added registry tattoo
    - Handles install.esd files properly
    - Performance improvements

    Version 1.2.2
    - Added USB drive picker

    Version 1.2.1
    - Fixed sysprep audit bugs

    Version 1.2.0
    - Added support for Surface Laptop 3 AMD SKUs (please note the "Device" name change from 1.0 and 1.1 versions for SurfaceLaptop3* variants)
    - Added support for including Office 365 into images
    - Bugfixes / performance improvements

    Version 1.1.0
    - Added support for local driver paths
    - Added support for Surface Go 2 and Surface Book 3

    Version 1.0.0
    - Initial release
#>



# Parse Params:
[CmdletBinding()]
Param(
    [Parameter(
        Position=1,
        Mandatory=$True,
        HelpMessage="Location of ISO containing Windows image (ex. D:\18362.1.190318-1202.19h1_release_CLIENT_BUSINESS_VOL_x64FRE_en-us.iso) to use as template"
        )]
        [string]$ISO,

    [Parameter(
        Position=2,
        Mandatory=$False,
        HelpMessage="What SKU should be used inside ISO (valid parameters are 'Pro' or 'Enterprise'), default is Pro - note checking is disabled currently as language support is added"
        )]
        #[ValidateSet('Pro', 'Enterprise')]
        [string]$OSSKU = 'Pro',

    [Parameter(
        Position=3,
        Mandatory=$True,
        HelpMessage="Destination folder to where resulting WIM image(s) should be placed"
        )]
        [string]$DestinationFolder,

    [Parameter(
        Position=4,
        Mandatory=$False,
        HelpMessage="Architecture of image being used (valid options are x64 and ARM64), default is x64"
        )]
        [ValidateSet('x64', 'ARM64')]
        [string]$Architecture = 'x64',

    [Parameter(
        Position=5,
        Mandatory=$False,
        HelpMessage="Install .NET 3.5 (bool true/false, default is false)"
        )]
        [bool]$DotNet35 = $True,

    [Parameter(
        Position=6,
        Mandatory=$False,
        HelpMessage="Add latest servicing stack update (bool true/false, default is true)"
        )]
        [bool]$ServicingStack = $True,

    [Parameter(
        Position=7,
        Mandatory=$False,
        HelpMessage="Add latest cumulative update (bool true/false, default is true)"
        )]
        [bool]$CumulativeUpdate = $True,

    [Parameter(
        Position=8,
        Mandatory=$False,
        HelpMessage="Add latest cumulative .NET update (bool true/false, default is true)"
        )]
        [bool]$CumulativeDotNetUpdate = $True,

    [Parameter(
        Position=9,
        Mandatory=$False,
        HelpMessage="Add latest Adobe Flash Player Security update (bool true/false, default is true)"
        )]
        [bool]$AdobeFlashUpdate = $True,

    [Parameter(
        Position=10,
        Mandatory=$False,
        HelpMessage="Add Office 365 C2R (bool true/false, default is true)"
        )]
        [bool]$Office365 = $True,

    [Parameter(
        Position=11,
        Mandatory=$False,
        HelpMessage="Surface device type to add drivers to image for, if not specified no drivers injected - Custom can be used if using with a non-Surface device"
        )]
        [ValidateSet('SurfacePro4', 'SurfacePro5', 'SurfacePro6', 'SurfacePro7', 'SurfaceLaptop', 'SurfaceLaptop2', 'SurfaceLaptop3Intel', 'SurfaceLaptop3AMD', 'SurfaceLaptopGo', 'SurfaceBook', 'SurfaceBook2', 'SurfaceBook3', 'SurfaceStudio', 'SurfaceStudio2', 'SurfaceGo', 'SurfaceGoLTE', 'SurfaceGo2', 'SurfaceHub2', 'Custom')]
        [string]$Device = "SurfacePro7",

    [Parameter(
        Position=12,
        Mandatory=$False,
        HelpMessage="Create USB key when finished (bool true/false, default is true)"
        )]
        [bool]$CreateUSB = $True,

    [Parameter(
        Position=13,
        Mandatory=$False,
        HelpMessage="Create bootable ISO file (useful for testing) when finished (bool true/false, default is true)"
        )]
        [bool]$CreateISO = $True,

    [Parameter(
        Position=14,
        Mandatory=$False,
        HelpMessage="Location of Windows ADK installation"
        )]
        [string]$WindowsKitsInstall = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit",

    [Parameter(
        Position=15,
        Mandatory=$False,
        HelpMessage="Use BITS for downloads"
        )]
        [bool]$BITSTransfer = $True,

    [Parameter(
        Position=16,
        Mandatory=$False,
        HelpMessage="Edit Install.wim"
        )]
        [bool]$InstallWIM = $True,

    [Parameter(
        Position=17,
        Mandatory=$False,
        HelpMessage="Edit boot.wim"
        )]
        [bool]$BootWIM = $True,

    [Parameter(
        Position=18,
        Mandatory=$False,
        HelpMessage="Keep original unsplit WIM even if resulting image size >4GB (bool true false, default is true)"
        )]
        [bool]$KeepOriginalWIM = $True,

    [Parameter(
        Position=19,
        Mandatory=$False,
        HelpMessage="Use a local driver path instead of downloading an MSI (bool true false, default is false)"
        )]
        [bool]$UseLocalDriverPath = $False,

    [Parameter(
        Position=20,
        Mandatory=$False,
        HelpMessage="Path to an MSI or extracted driver folder - required if you set UseLocalDriverPath variable to true or script will not find any drivers to inject"
        )]
        [string]$LocalDriverPath
    )



$SDAVersion = "1.2.5.4"
$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = New-Object System.Text.UTF8Encoding



Function Start-Log
{
    Param (
        [Parameter(Mandatory = $True)]
	    [String]$FilePath,

        [Parameter(Mandatory = $True)]
        [String]$FileName
    )
	
    Try
    {
        If (!(Test-Path $FilePath))
	    {
	        ## Create the log file
	        New-Item -Path "$FilePath" -ItemType "directory" | Out-Null
            New-Item -Path "$FilePath\$FileName" -ItemType "file"
	    }
        Else
        {
            New-Item -Path "$FilePath\$FileName" -ItemType "file"
        }
		
	    ## Set the global variable to be used as the FilePath for all subsequent Write-Log calls in this session
	    $global:ScriptLogFilePath = "$FilePath\$FileName"
    }
    Catch
    {
        Write-Error $_.Exception.Message
        Exit
    }
}



Function Write-Log
{
    Param (
        [Parameter(Mandatory = $True)]
        [String]$Message,
		
        [Parameter(Mandatory = $False)]
        # 1 == "Informational"
        # 2 == "Warning'
        # 3 == "Error"
        [ValidateSet(1, 2, 3)]
        [Int]$LogLevel = 1,

        [Parameter(Mandatory = $False)]
	    [String]$LogFilePath = $ScriptLogFilePath,

        [Parameter(Mandatory = $False)]
        [String]$ScriptLineNumber
    )

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$ScriptLineNumber", $LogLevel
    $Line = $Line -f $LineFormat

    #Add-Content -Path $LogFilePath -Value $Line
    Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $ScriptLogFilePath
}



Function Receive-Output
{
    Param(
        $Color,
        $BGColor,
        $LogLevel,
        $LogFile,
        $LineNumber
    )

    Process
    {
        If ($BGColor)
        {
            Write-Host $_ -ForegroundColor $Color -BackgroundColor $BGColor
        }
        Else
        {
        Write-Host $_ -ForegroundColor $Color
        }

        If (($LogLevel) -or ($LogFile))
        {
            Write-Log -Message $_ -LogLevel $LogLevel -LogFilePath $ScriptLogFilePath -ScriptLineNumber $LineNumber
        }
    }
}



Function AddHeaderSpace
{
    Write-Output "This space intentionally left blank..." | Receive-Output -Color Gray
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
}



Function CheckIfRunAsAdmin
{
    If (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] “Administrator”))
    {
        Write-Output “You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator to continue.” | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Break
    }
}



Function Check-Internet
{
    While (([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]‘{DCB00C01-570F-4A9B-8D69-199FDBA5723B}’)).IsConnectedToInternet) -eq $False)
    {
        Write-Output "No internet connection detected. Retrying in 60 seconds..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep -Seconds 60
    }
}



Function Get-RedirectedUrl
{
    Param(
        $URL
    )

    $Request = [System.Net.WebRequest]::Create($URL)
    $Request.AllowAutoRedirect=$false
    $Request.Timeout = 3000
    $Response = $Request.GetResponse()

    If ($Response.ResponseUri)
    {        
        $Response.GetResponseHeader("Location")
    }
    $Response.Close()
}



Function DownloadFile
{
    Param(
        [System.Uri]$URL,
        [System.String]$Path
    )

    # Get file name
    Start-Sleep 1

    If ($URL.Host -like "*aka.ms*")
    {
        $ActualURL = Get-RedirectedUrl -URL "$URL" -ErrorAction Continue -WarningAction Continue
        $FileName = $ActualURL.Substring($ActualURL.LastIndexOf("/") + 1)
        Write-Output "aka.ms link: $URL" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output "Actual URL:  $ActualURL" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output "File name:   $FileName" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
    }
    Else
    {
        $ActualURL = $URL
        $FileName = $URL.AbsoluteUri.Substring($URL.AbsoluteUri.LastIndexOf("/") +1)
        Write-Output "Actual URL:  $URL" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output "File name:   $FileName" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
    }

    $global:Output = "$Path\$Filename"

    # If file does not exist, download file
    If (!(Test-Path -Path "$global:Output"))
    {
        Write-Output "Using BITS to download files" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output "Downloading $FileName to $Path..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $ActualURL -Destination "$global:Output" -Priority Foreground -RetryTimeout 60 -RetryInterval 120
    }
    Else
    {
        Write-Output "File $global:Output exists, skipping file download." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
    }

    Return $global:Output
}



# Using this to avoid reinstalling and breaking installed Win32 MSI apps via WMI calls to Win32_Product!
Function GetInstalledAppStatus
{
    Param(
        $AppName,
        $AppVersion
    )

    $OSArch = Get-CimInstance -ClassName Win32_OperatingSystem

    If ($OSArch.OSArchitecture -eq "64-bit")
    {
        $InstalledPrograms32 = Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse
        $InstalledPrograms64 = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse

        ForEach ($Item in $InstalledPrograms32)
        {
            If ($Item.GetValue("DisplayName") -like "*$AppName*" -and ($Item.GetValue("DisplayVersion")) -like "*$AppVersion*")
            {
                $global:IsInstalled = $true
                Break
            }
        }

        ForEach ($Item in $InstalledPrograms64)
        {
            If ($Item.GetValue("DisplayName") -like "*$AppName*" -and ($Item.GetValue("DisplayVersion")) -like "*$AppVersion*")
            {
                $global:IsInstalled = $true
                Break
            }
        }
    }
    Else
    {
        $InstalledPrograms32 = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse

        ForEach ($Item in $InstalledPrograms32)
        {
            If ($Item.GetValue("DisplayName") -like "*$AppName*" -and ($Item.GetValue("DisplayVersion")) -like "*$AppVersion*")
            {
                $global:IsInstalled = $true
                Break
            }
        }
    }
}



Function PrereqCheck
{
    # Check variables for spaces and not fully-defined paths
    If ($DestinationFolder.Contains(" "))
    {
        Write-Output "`$DestinationFolder cannot contain spaces: $DestinationFolder" | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
    $IsThisAFullLocalPath = $DestinationFolder.Substring(1,1)
    If ($IsThisAFullLocalPath -ne ":")
    {
        Write-Output "$DestinationFolder was not passed as a full path to a local folder.  Please pass the full path to the DestinationFolder parameter." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
    If ($UseLocalDriverPath -and $LocalDriverPath.Contains(" "))
    {
        Write-Output "`$LocalDriverPath cannot contain spaces: $LocalDriverPath" | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
    
    # Check for admin rights
    CheckIfRunAsAdmin

    # Windows Version Check
    $OSCaption = (Get-CimInstance -ClassName win32_operatingsystem).caption
    If ($OSCaption -like "Microsoft Windows 10*" -or $OSCaption -like "Microsoft Windows Server 2019*")
    {
        # All OK
    }
    Else
    {
        Write-Output "$Env:Computername You must use Windows 10 1809 or newer, or Windows Server 2019 or newer when servicing Windows 10 offline, with the latest ADK installed." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output "$Env:Computername Aborting script..." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
    
    # Validating that the ADK Deployment Tools are installed
    If (!(Test-Path $DISMFile))
    {
        Write-Output "DISM in Windows ADK not found, attempting installation..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
        $global:Output = $null
        $global:IsInstalled = $null

        $ScriptFolder = $DestinationFolder
        $ADKSourceFile = "$ScriptFolder\adksetup.exe"
        $ADKArguments = " /features OptionId.DeploymentTools /quiet"

        GetInstalledAppStatus -AppName "Windows Assessment and Deployment Kit - Windows 10" -AppVersion "10.1.19041"

        If ($global:IsInstalled -eq $null)
        {
            # ADK cannot do an "in place" upgrade.  Do we need to uninstall the old version?
            $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Windows Assessment and Deployment Kit - Windows 10*" } | select UninstallString
            $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Windows Assessment and Deployment Kit - Windows 10*" } | select UninstallString

            If ($uninstall64) 
            {
                ForEach ($u in $uninstall64)
                {
                    $u = $u.UninstallString -Replace "/uninstall","" 
                    $u = $u.Trim()
                    Write-Output "Removing old 64bit ADK components.  Command is $u and args are /uninstall /quiet" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ($uninstall32)
            {
                ForEach ($u in $uninstall32)
                {
                    $u = $u.UninstallString -Replace "/uninstall",""
                    $u = $u.Trim()
                    Write-Output "Removing old 32bit ADK components.  Command is $u and args are /uninstall /quiet" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ((Test-Path -Path $ADKSourceFile) -eq $true)
            {
                $SourceFilePath = $(Get-Item $SourceFile).FullName
                Write-Output "Found Installation files for ADK at $SourceFilePath" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            }	
            Else
            {
                Check-Internet
                $URL = "https://aka.ms/sdaadk/2004"
                $Path = "$env:TEMP"
                DownloadFile $URL $Path
                $SourceFilePath = $global:Output
            }

            Try
            {
                Write-Output "Installing Windows Assessment and Deployment Kit Deployment Tools" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Start-Process -File  $SourceFilePath -Arg $ADKArguments -passthru | Wait-Process

                Write-Output "$AppName - ADK INSTALLATION SUCCESSFULLY COMPLETED" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Write-Output  ""

            }
            Catch
            {
                Write-Output "$AppName - INSTALLATION ERROR - check logs in $env:TEMP\adk for more info." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Write-Output  ""
                Exit
            }
        }
    }

    # Validating that the ADK WinPE Add-ons are installed
    If (!(Test-Path $ADKWinPEFile))
    {
        Write-Output "Windows ADK Windows PE components not found, attempting installation..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
        $global:Output = $null
        $global:IsInstalled = $null

        $ScriptFolder = $DestinationFolder
        $WinPESourceFile = "$ScriptFolder\adkwinpesetup.exe"
        $WinPEArguments = " /features OptionId.WindowsPreinstallationEnvironment /quiet"

        GetInstalledAppStatus -AppName "Windows Assessment and Deployment Kit Windows Preinstallation Environment Add-ons - Windows 10" -AppVersion "10.1.19041"

        If ($global:IsInstalled -eq $null)
        {
            # ADK cannot do an "in place" upgrade.  Do we need to uninstall the old version?
            $uninstall32 = Get-ChildItem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Windows Assessment and Deployment Kit Windows Preinstallation Environment Add-ons - Windows 10*" } | select UninstallString
            $uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Windows Assessment and Deployment Kit Windows Preinstallation Environment Add-ons - Windows 10*" } | select UninstallString

            If ($uninstall64) 
            {
                ForEach ($u in $uninstall64)
                {
                    $u = $u.UninstallString -Replace "/uninstall","" 
                    $u = $u.Trim()
                    Write-Output "Removing old 64bit ADK WinPE components.  Command is $u and args are /uninstall /quiet" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ($uninstall32)
            {
                ForEach ($u in $uninstall32)
                {
                    $u = $u.UninstallString -Replace "/uninstall",""
                    $u = $u.Trim()
                    Write-Output "Removing old 32bit ADK WinPE components.  Command is $u and args are /uninstall /quiet" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ((Test-Path -Path $WinPESourceFile) -eq $true)
            {
                $SourceFilePath = $(Get-Item $SourceFile).FullName
                Write-Output "Found Installation files for ADK WinPE at $SourceFilePath" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            }	
            Else
            {
                Check-Internet
                $URL = "https://aka.ms/sdaadkpe/2004"
                $Path = "$env:TEMP"
                DownloadFile $URL $Path
                $SourceFilePath = $global:Output
            }

            Try
            {
                Write-Output "Installing Windows Assessment and Deployment Kit Windows Preinstallation Environment Add-Ons" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Start-Process -File  $SourceFilePath -Arg $WinPEArguments -passthru | Wait-Process
  
                Write-Output  "$AppName - ADK WinPE Add-Ons INSTALLATION SUCCESSFULLY COMPLETED" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Write-Output  ""
            }

            Catch
            {
                Write-Output  "$AppName - INSTALLATION ERROR - check logs in $env:TEMP\adkwinpeaddons for more info." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Write-Output  ""
                Exit
            }
        }
    }
}



Function Download-LatestUpdates
{
    Param(
        $uri,
        $Path,
        $Date,
        $Servicing,
        $Cumulative,
        $CumulativeDotNet,
        $Adobe,
        $OSBuild
    )

    $kbObj = Invoke-WebRequest -Uri $uri -UseBasicParsing

    # Parse the Response
    $global:KBGUID = $null
    $kbObjectLinks = ($kbObj.Links | Where-Object {$_.id -match "_link"})
    $array = @()

    ForEach ($link in $kbObjectLinks)
    {
        $xmlNode = [XML]($link.outerHTML)

        If ($xmlNode.HasChildNodes)
        {
            $kbId = $link.id -replace "_link", ""
            $description = $xmlNode.FirstChild.InnerText.Trim()
            $array += [PSCustomObject]@{
                kbId = $kbId
                description = $description
            }
        }
    }

    If ($array.count -gt 0)
    {
        If ($Servicing)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*$Date*") -and ($_.description -like "*Servicing Stack Update for Windows 10*") -and ($_.description -like "*$OSBuild*") -and ($_.description -like "*$Architecture*")}
            If ($global:KBGUID.Count -gt 1)
            {
                $largest = ($global:KBGUID | Measure-Object -Property description -Maximum)
                $global:KBGUID = $global:KBGUID | Where-Object {$_.description -eq $largest.Maximum}
            }
        }
        If ($Cumulative)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*$Date*") -and ($_.description -like "*Cumulative Update for Windows 10*") -and ($_.description -like "*$OSBuild*") -and ($_.description -like "*$Architecture*")}
            If ($global:KBGUID.Count -gt 1)
            {
                $largest = ($global:KBGUID | Measure-Object -Property description -Maximum)
                $global:KBGUID = $global:KBGUID | Where-Object {$_.description -eq $largest.Maximum}
            }
        }
        If ($CumulativeDotNet)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*$Date*") -and ($_.description -like "*Cumulative Update for .NET Framework*") -and ($_.description -like "*Windows 10*") -and ($_.description -like "*$OSBuild*")}
        }
        If ($Adobe)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*$Date*") -and ($_.description -like "*Security Update for Adobe Flash Player for Windows 10*") -and ($_.description -like "*$OSBuild*")}
        }

        $updatesFound = $false

        ForEach ($Object in $global:KBGUID)
        {
            $kb = $Object.kbId
            $curTxt = $Object.description
            
            ##Create Post Request to get the Download URL of the Update
            $Post = @{ size = 0; updateID = $kb; uidInfo = $kb } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" }
            
            ## Fetch and parse the download URL
            $PostRes = (Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody).content
            ## Seeing two potentially different hosts serving download links:
            $DLWUDOTCOM = ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
            $DLDELDOTCOM = ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://dl\.delivery\.mp\.microsoft\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
            If ($DLWUDOTCOM)
            {
                $DownloadLinks = $DLWUDOTCOM
            }
            ElseIf ($DLDELDOTCOM)
            {
                $DownloadLinks = $DLDELDOTCOM
            }
            If ($DownloadLinks)
            {
                $updatesFound = $true
                If ($DownloadLinks.Count -gt 1)
                {
                    ForEach ($URL in $DownloadLinks)
                    {
                        Write-Output "Download found:" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Write-Output $curTxt | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Write-Output ""
                        Write-Output ""
                        DownloadFile -URL $URL -Path "$Path"
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                    }
                }
                Else
                {
                    Write-Output "Download found:" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output $curTxt | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output ""
                    Write-Output ""
                    DownloadFile -URL $DownloadLinks -Path "$Path"
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                }
            }
        }

        If (!($updatesFound))
        {
            $global:KBGUID = $null
        }
    }
}



Function Get-LatestUpdates
{
    Param(
        $Servicing = $False,
        $Cumulative = $False,
        $CumulativeDotNet = $False,
        $Adobe = $False,
        $Path,
        $Date,
        $OSBuild,
        $Architecture
    )

    If (!($Path))
    {
        $Path = $WorkingDirPath
    }

    If (!(Test-Path -Path $Path))
    {
        New-Item -path "$Path" -ItemType "directory" | Out-Null
    }
    
    If (!($Date))
    {
        $Date = Get-Date -Format "yyyy-MM"
    }
    
    $ServicingURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + " Servicing Stack " + $Architecture + " windows 10 " + $OSBuild
    $CumulativeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for Windows 10" ' + $Architecture + " " + $OSBuild
    $CumulativeDotNetURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for .NET Framework" ' + $Architecture + " windows 10 " + $OSBuild
    $AdobeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "Security Update for Adobe Flash Player for Windows 10" ' + $Architecture + " " + $OSBuild

    If ($Servicing)
    {
        Write-Output "Attempting to find and download Servicing Stack updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $uri = $ServicingURI
        Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $True -Cumulative $False -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
        If (!($global:KBGUID))
        {
            While (!($global:KBGUID))
            {
                If ($LoopBreak -le 5)
                {
                    $LoopBreak++
                    Start-Sleep 1
                    $NewDate = (Get-Date).AddMonths(-$LoopBreak)
                    $NewDate = $NewDate.ToString("yyyy-MM")
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

                    $Date = $NewDate
                    $ServicingURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + " Servicing Stack " + $Architecture + " windows 10 " + $OSBuild

                    $uri = $ServicingURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $True -Cumulative $False -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Break
                }
            }
        }
        $LoopBreak = $null
        $Date = Get-Date -Format "yyyy-MM"
    }
    If ($Cumulative)
    {
        Write-Output "Attempting to find and download Cumulative Update updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $uri = $CumulativeURI
        Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $True -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
        If (!($global:KBGUID))
        {
            While (!($global:KBGUID))
            {
                If ($LoopBreak -le 5)
                {
                    $LoopBreak++
                    Start-Sleep 1
                    $NewDate = (Get-Date).AddMonths(-$LoopBreak)
                    $NewDate = $NewDate.ToString("yyyy-MM")
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

                    $Date = $NewDate
                    $CumulativeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for Windows 10" ' + $Architecture + " " + $OSBuild

                    $uri = $CumulativeURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $True -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Break
                }
            }
        }
        $Date = Get-Date -Format "yyyy-MM"
        $LoopBreak = $null
    }
    If ($CumulativeDotNet)
    {
        Write-Output "Attempting to find and download Cumulative .NET Framework Update updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $uri = $CumulativeDotNetURI
        Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $True -Adobe $False -OSBuild $OSBuild
        If (!($global:KBGUID))
        {
            While (!($global:KBGUID))
            {
                If ($LoopBreak -le 5)
                {
                    $LoopBreak++
                    Start-Sleep 1
                    $NewDate = (Get-Date).AddMonths(-$LoopBreak)
                    $NewDate = $NewDate.ToString("yyyy-MM")
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

                    $Date = $NewDate
                    $CumulativeDotNetURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for .NET Framework" ' + $Architecture + " windows 10 " + $OSBuild

                    $uri = $CumulativeDotNetURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $True -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Break
                }
            }
        }
        $Date = Get-Date -Format "yyyy-MM"
        $LoopBreak = $null
    }
    If ($Adobe)
    {
        Write-Output "Attempting to find and download Adobe Flash Player updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $uri = $AdobeURI
        Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $False -Adobe $True -OSBuild $OSBuild
        If (!($global:KBGUID))
        {
            While (!($global:KBGUID))
            {
                If ($LoopBreak -le 11)
                {
                    $LoopBreak++
                    Start-Sleep 1
                    $NewDate = (Get-Date).AddMonths(-$LoopBreak)
                    $NewDate = $NewDate.ToString("yyyy-MM")
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

                    $Date = $NewDate
                    $AdobeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "Security Update for Adobe Flash Player for Windows 10" ' + $Architecture + " " + $OSBuild

                    $uri = $AdobeURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $False -Adobe $True -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak month's of searches.  Continuing..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Break
                }
            }
        }
        $Date = Get-Date -Format "yyyy-MM"
        $LoopBreak = $null
        
    }
}



Function ExtractMSIFile
{
    Param
    (
        $MsiFile,
        $Path
    )

    If (Test-Path "$Path\Extract")
    {
        Write-Output "Deleting $Path\Extract\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$Path\Extract\" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$Path\Extract" -Force
    }
    If (!(Test-Path "$Path\Extract"))
    {
        New-Item -Path "$Path\Extract" -ItemType "directory" | Out-Null
    }

    Write-Output "Extracting file $MsiFile to $Path\Extract..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Start-Process "msiexec" -ArgumentList "/a $MsiFile /qn TARGETDIR=$Path\Extract" -Wait -NoNewWindow
}



Function Get-LatestSurfaceEthernetDrivers
{
    Param(
        $Device,
        $TempFolder
    )
    Write-Output ""
    Write-Output ""

    $DeviceDriverPath = "$TempFolder\$Device"

    If (!($Device))
    {
        # Nothing yet
    }
    ElseIf ($Device -eq "SurfaceHub2")
    {
        # Nothing yet
    }
    Else
    {
        $URI = "http://www.catalog.update.microsoft.com/Search.aspx?q=Surface net Windows 10"
        $kbObj = Invoke-WebRequest -Uri $URI -UseBasicParsing

        $global:KBGUID = $null
        $kbObjectLinks = ($kbObj.Links | Where-Object {$_.id -match "_link"})
        $array = @()


        $kbObj = Invoke-WebRequest -Uri $uri -UseBasicParsing

        # Parse the Response
        $global:KBGUID = $null
        $kbObjectLinks = ($kbObj.Links | Where-Object {$_.id -match "_link"})
        $array = @()

        ForEach ($link in $kbObjectLinks)
        {
            $xmlNode = [XML]($link.outerHTML)

            If ($xmlNode.HasChildNodes)
            {
                $kbId = $link.id -replace "_link", ""
                $description = $xmlNode.FirstChild.InnerText.Trim()
                $array += [PSCustomObject]@{
                    kbId = $kbId
                    description = $description
                }
            }
        }

        If ($array.count -gt 0)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*Surface - Net - 10.*")}

            If ($global:KBGUID.Count -gt 1)
            {
                $largest = ($global:KBGUID | Measure-Object -Property description -Maximum)
                $global:KBGUID = $global:KBGUID | Where-Object {$_.description -eq $largest.Maximum}
            }
        }

        ForEach ($Object in $global:KBGUID)
        {
            $kb = $Object.kbId
            $curTxt = $Object.description
    
            ##Create Post Request to get the Download URL of the Update
            $Post = @{ size = 0; updateID = $kb; uidInfo = $kb } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" }
    
            ## Fetch and parse the download URL
            $PostRes = (Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody).content
            $DownloadLinks = ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
            If ($DownloadLinks)
            {
                If ($DownloadLinks.Count -gt 1)
                {
                    ForEach ($URL in $DownloadLinks)
                    {
                        Write-Output "Download found:" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Write-Output $curTxt | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Write-Output ""
                        Write-Output ""
                        $TempCAB = DownloadFile -URL $URL -Path "$DeviceDriverPath"
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        $expand = "$env:WINDIR\System32\expand.exe"
                        $args = "-f:* $TempCAB $DeviceDriverPath"
                        Start-Process -FilePath $expand -ArgumentList $args -Wait -NoNewWindow
                        Write-Output ""
                        Write-Output ""
                    }
                }
                Else
                {
                    Write-Output "Download found:" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output $curTxt | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output ""
                    Write-Output ""
                    $TempCAB = DownloadFile -URL $DownloadLinks -Path "$DeviceDriverPath"
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    $expand = "$env:WINDIR\System32\expand.exe"
                    $args = "-f:* $TempCAB $DeviceDriverPath"
                    Start-Process -FilePath $expand -ArgumentList $args -Wait -NoNewWindow
                    Write-Output ""
                    Write-Output ""
                }
            }
        }
    }
}



Function Get-LatestWinUSBDrivers
{
    Param(
        $Device,
        $TempFolder
    )
    Write-Output ""
    Write-Output ""

    $DeviceDriverPath = "$TempFolder\$Device"

    If (!($Device -eq "SurfaceHub2"))
    {
        # Nothing yet
    }
    Else
    {
        $URI = "http://www.catalog.update.microsoft.com/Search.aspx?q=SMSC-Microchip WinUSB USB2534 Device Windows 10"
        $kbObj = Invoke-WebRequest -Uri $URI -UseBasicParsing

        $global:KBGUID = $null
        $kbObjectLinks = ($kbObj.Links | Where-Object {$_.id -match "_link"})
        $array = @()

        $kbObj = Invoke-WebRequest -Uri $uri -UseBasicParsing

        # Parse the Response
        $global:KBGUID = $null
        $kbObjectLinks = ($kbObj.Links | Where-Object {$_.id -match "_link"})
        $array = @()

        ForEach ($link in $kbObjectLinks)
        {
            $xmlNode = [XML]($link.outerHTML)
            If ($xmlNode.HasChildNodes)
            {
                $kbId = $link.id -replace "_link", ""
                $description = $xmlNode.FirstChild.InnerText.Trim()
                $array += [PSCustomObject]@{
                    kbId = $kbId
                    description = $description
                }
            }
        }

        If ($array.count -gt 0)
        {
            $global:KBGUID = $array | Where-Object {($_.description -like "*SMSC-Microchip WinUSB USB2534 Device*")}
            If ($global:KBGUID.Count -gt 1)
            {
                #32 and 64bit driver packages have both, so just grab the first in the array which should be amd64
                $global:KBGUID = $global:KBGUID[0]
            }
        }

        ForEach ($Object in $global:KBGUID)
        {
            $kb = $Object.kbId
            $curTxt = $Object.description
    
            ##Create Post Request to get the Download URL of the Update
            $Post = @{ size = 0; updateID = $kb; uidInfo = $kb } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" }
    
            ## Fetch and parse the download URL
            $PostRes = (Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody).content
            $DownloadLinks = ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
            If ($DownloadLinks)
            {
                If ($DownloadLinks.Count -gt 1)
                {
                    ForEach ($URL in $DownloadLinks)
                    {
                        Write-Output "Download found:"
                        Write-Output $curTxt
                        Write-Output ""
                        Write-Output ""
                        $TempCAB = DownloadFile -URL $URL -Path "$DeviceDriverPath"
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        $expand = "$env:WINDIR\System32\expand.exe"
                        $args = "-f:* $TempCAB $DeviceDriverPath"
                        Start-Process -FilePath $expand -ArgumentList $args -Wait -NoNewWindow
                        Write-Output ""
                        Write-Output ""
                    }
                }
                Else
                {
                    Write-Output "Download found:"
                    Write-Output $curTxt
                    Write-Output ""
                    Write-Output ""
                    $TempCAB = DownloadFile -URL $DownloadLinks -Path "$DeviceDriverPath"
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
                    $expand = "$env:WINDIR\System32\expand.exe"
                    $args = "-f:* $TempCAB $DeviceDriverPath"
                    Start-Process -FilePath $expand -ArgumentList $args -Wait -NoNewWindow
                    Write-Output ""
                    Write-Output ""
                }
            }
        }
    }
}



Function Get-LatestDrivers
{
    Param(
        $Device,
        $TempFolder
    )
    Write-Output ""
    Write-Output ""

    $DeviceDriverPath = "$TempFolder\$Device"

    If (Test-Path "$DeviceDriverPath")
    {
        Write-Output "Deleting $DeviceDriverPath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$DeviceDriverPath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$DeviceDriverPath" -Force
    }
    If (!(Test-Path "$DeviceDriverPath"))
    {
        New-Item -path "$DeviceDriverPath" -ItemType "directory" | Out-Null
    }

    If ($UseLocalDriverPath -eq $True)
    {
        If (!(Test-Path "$LocalDriverPath"))
        {
            Write-Output "$LocalDriverPath not found, continuing without drivers..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            $Device = $null
        }
        Else
        {
            $TempLocalDriverPath = (Get-Item $LocalDriverPath) -is [System.IO.DirectoryInfo]
            If ($TempLocalDriverPath -eq $False)
            {
                ExtractMSIFile -MsiFile $LocalDriverPath -Path "$DeviceDriverPath"
            }
            ElseIf ($TempLocalDriverPath -eq $True)
            {
                $TempDeviceDriverPath = "$DeviceDriverPath\Extract"
                If (Test-Path "$TempDeviceDriverPath")
                {
                    Write-Output "Deleting $TempDeviceDriverPath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Get-ChildItem -Path "$TempDeviceDriverPath" -Recurse | Remove-Item -Force -Recurse
                    Remove-Item -Path "$TempDeviceDriverPath" -Force
                }
                If (!(Test-Path "$TempDeviceDriverPath"))
                {
                    New-Item -path "$TempDeviceDriverPath" -ItemType "directory" | Out-Null
                }
                # Use local drivers
                Write-Output "Copying drivers from $LocalDriverPath to $TempDeviceDriverPath..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                & xcopy.exe /herky "$LocalDriverPath" "$TempDeviceDriverPath"
            }
            Write-Output ""
        }
    }
    Else
    {
        Write-Output "Downloading latest drivers for $Device, Windows 10 version $global:OSVersion..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $OSBuild = New-Object string (,@($global:OSVersion.ToCharArray() | Select-Object -Last 5))

        If ($Device -eq "SurfaceLaptop3Intel")
        {
            $TempDevice = "SurfaceLaptop3"
            $TempDeviceType = "Intel"
            $URL = "https://aka.ms/" + $TempDevice + "/" + $TempDeviceType + "/" + $OSBuild
        }
        ElseIf ($Device -eq "SurfaceLaptop3AMD")
        {
            $TempDevice = "SurfaceLaptop3"
            $TempDeviceType = "AMD"
            $URL = "https://aka.ms/" + $TempDevice + "/" + $TempDeviceType + "/" + $OSBuild
        }
        Else
        {
            $URL = "https://aka.ms/" + $Device + "/" + $OSBuild
        }
        
        $DownloadedFile = DownloadFile -URL $URL -Path "$DeviceDriverPath"
        Write-Output "Downloaded File: $DownloadedFile" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

        $FileToExtract = $DownloadedFile
        ExtractMSIFile -MsiFile $FileToExtract -Path $DeviceDriverPath
        Write-Output ""
    }

    If ($Device -eq "SurfaceHub2")
    {
        Write-Output "Downloading latest WinUSB drivers for $Device..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-LatestWinUSBDrivers -Device $Device -TempFolder $TempFolder
        Write-Output ""
    }
    Else
    {
        Write-Output "Downloading latest Surface Ethernet drivers for $Device..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-LatestSurfaceEthernetDrivers -Device $Device -TempFolder $TempFolder
        Write-Output ""
    }
}



Function Get-LatestVCRuntimes
{
    Param(
        $TempFolder
    )
    Write-Output ""
    Write-Output ""

    $VisualCRuntimePath = "$TempFolder\VCRuntimes"

    If (Test-Path "$VisualCRuntimePath")
    {
        Write-Output "Deleting $VisualCRuntimePath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$VisualCRuntimePath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$VisualCRuntimePath" -Force
    }
    If (!(Test-Path "$VisualCRuntimePath"))
    {
        New-Item -path "$VisualCRuntimePath" -ItemType "directory" | Out-Null
    }
    If (!(Test-Path "$VisualCRuntimePath\2013"))
    {
        New-Item -path "$VisualCRuntimePath\2013" -ItemType "directory" | Out-Null
    }
    If (!(Test-Path "$VisualCRuntimePath\2019"))
    {
        New-Item -path "$VisualCRuntimePath\2019" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest VisualC++ Runtimes..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

    $VC2013x86URL = "https://aka.ms/vcpp2013x86"
    $VC2013x64URL = "https://aka.ms/vcpp2013x64"
    $VC2019X86URL = "https://aka.ms/vcpp2019x86"
    $VC2019X64URL = "https://aka.ms/vcpp2019x64"

    # 2013
    $VC2013x86 = DownloadFile -URL $VC2013x86URL -Path "$VisualCRuntimePath\2013"
    Write-Output "Downloaded File: $VC2013x86" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
    $VC2013x64 = DownloadFile -URL $VC2013x64URL -Path "$VisualCRuntimePath\2013"
    Write-Output "Downloaded File: $VC2013x64" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
    # 2019
    $VC2019x86 = DownloadFile -URL $VC2019x86URL -Path "$VisualCRuntimePath\2019"
    Write-Output "Downloaded File: $VC2019x86" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
    $VC2019x64 = DownloadFile -URL $VC2019x64URL -Path "$VisualCRuntimePath\2019"
    Write-Output "Downloaded File: $VC2019x64" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
}



Function Get-Office365
{
    Param(
        [string]$TempFolder
    )

    Write-Output ""
    Write-Output ""
    Write-Output ""

    $Office365Path = "$TempFolder\Office365"

    If (Test-Path "$Office365Path")
    {
        Write-Output "Deleting $Office365Path\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$Office365Path" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$Office365Path" -Force    
    }
    If (!(Test-Path "$Office365Path"))
    {
        New-Item -Path "$Office365Path" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading Office 365 $Office365SKU..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    
    $Office365OfflineURL = "https://aka.ms/sdao365"

    $Office365TempFile = DownloadFile -URL $Office365OfflineURL -Path "$Office365Path"
    Write-Output "Downloaded File: $Office365TempFile" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""

    Write-Output "Extracting Office 365 offline installer..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Start-Process -FilePath "$Office365TempFile" -ArgumentList "/extract:$Office365Path /quiet" -Wait
    Write-Output ""

    If (!(Test-Path "$Office365Path\setup.exe"))
    {
        #File not downloaded, bail
        Write-Output "Office offline setup file download appears to have failed.  Exiting..." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
    Else
    {
        $O365ProPlusDownloadXMLPath = "$WorkingDirPath\O365_Download.xml"
        $O365ProPlusConfigurationXMLPath = "$WorkingDirPath\O365_Configuration.xml"
        Copy-Item -Path $O365ProPlusDownloadXMLPath -Destination $Office365Path
        Copy-Item -Path $O365ProPlusConfigurationXMLPath -Destination $Office365Path
        
        $TempFile = "$Office365Path\O365_Download.xml"
        If (Test-Path $TempFile)
        {
            $O365ProPlusDownloadXMLPath = "$Office365Path\O365_Download.xml"
        }
        Write-Output "Downloading Office 365 offline package..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $Argumentlist = "/download $O365ProPlusDownloadXMLPath"
        Set-Location -Path "$Office365Path"
        Start-Process -FilePath "$Office365Path\setup.exe" -ArgumentList $Argumentlist -Wait
        Start-Sleep 2
        Set-Location $WorkingDirPath
    }
    Write-Output ""
    Write-Output ""
    Write-Output ""
}



Function Get-AdobeFlashUpdates
{
    Param(
        [string]$TempFolder
    )

    $adobeUpdatePath = "$TempFolder\Adobe"

    If (Test-Path "$adobeUpdatePath")
    {
        Write-Output "Deleting $adobeUpdatePath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$adobeUpdatePath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$adobeUpdatePath" -Force
    }
    If (!(Test-Path "$adobeUpdatePath"))
    {
        New-Item -path "$adobeUpdatePath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Adobe Flash update for $global:OSVersion..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-LatestUpdates -Adobe $True -Path $adobeUpdatePath -OSBuild $global:ReleaseId -Architecture $Architecture
}



Function Get-CumulativeUpdates
{
    Param(
        [string]$TempFolder
    )

    $CumulativeUpdatePath = "$TempFolder\Cumulative"

    If (Test-Path "$CumulativeUpdatePath")
    {
        Write-Output "Deleting $CumulativeUpdatePath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$CumulativeUpdatePath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$CumulativeUpdatePath" -Force
    }
    If (!(Test-Path "$CumulativeUpdatePath"))
    {
        New-Item -path "$CumulativeUpdatePath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Cumulative Update for $global:OSVersion..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-LatestUpdates -Cumulative $True -Path $CumulativeUpdatePath -OSBuild $global:ReleaseId -Architecture $Architecture
}



Function Get-ServicingStackUpdates
{
    Param(
        [string]$TempFolder
    )

    $ServicingStackPath = "$TempFolder\Servicing"

    If (Test-Path "$ServicingStackPath")
    {
        Write-Output "Deleting $ServicingStackPath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$ServicingStackPath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$ServicingStackPath" -Force
    }
    If (!(Test-Path "$ServicingStackPath"))
    {
        New-Item -Path "$ServicingStackPath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Servicing Stack update for $global:OSVersion..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-LatestUpdates -Servicing $True -Path $ServicingStackPath -OSBuild $global:ReleaseId -Architecture $Architecture
}



Function Get-CumulativeDotNetUpdates
{
    Param(
        [string]$TempFolder
    )

    $CumulativeDotNetPath = "$TempFolder\DotNet"

    If (Test-Path "$CumulativeDotNetPath")
    {
        Write-Output "Deleting $CumulativeDotNetPath\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$CumulativeDotNetPath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$CumulativeDotNetPath" -Force
    }
    If (!(Test-Path "$CumulativeDotNetPath"))
    {
        New-Item -Path "$CumulativeDotNetPath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Dot Net Cumulative updates for $global:OSVersion..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-LatestUpdates -CumulativeDotNet $True -Path $CumulativeDotNetPath -OSBuild $global:ReleaseId -Architecture $Architecture
}



Function Get-OSWIMFromISO
{
    Param(
        $ISO,
        $OSSKU,
        $DestinationFolder,
        $Architecture,
        $global:OSVersion,
        $WindowsKitsInstall,
        $ScratchMountFolder
    )

    If (Test-Path "$ScratchMountFolder")
    {
        Write-Output "Deleting $ScratchMountFolder\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$ScratchMountFolder" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$ScratchMountFolder" -Force
    }
    If (!(Test-Path -path $ScratchMountFolder))
    {
        New-Item -path $ScratchMountFolder -ItemType Directory | Out-Null
    }

    Write-Output "Mounting ISO $ISO..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    $ISOPath = (Mount-DiskImage -ImagePath $ISO -StorageType ISO -PassThru | Get-Volume).DriveLetter
    $Drive = $ISOPath + ":"

    If ($ISOPath)
    {
        Write-Output "ISO successfully mounted at $Drive" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""   
    }
    Else
    {
        Write-Output "Failed to mount the ISO. Please verify the ISO path and try again" | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }

    Write-Output "Parsing install.wim/install.esd file(s) in $Drive for images..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    $IsESD = $False
    $WIMs = Get-ChildItem -Path "$Drive" -Filter install.wim -Recurse
    If (!($WIMs))
    {
        $WIMs = Get-ChildItem -Path "$Drive" -Filter install.esd -Recurse
        $IsESD = $True
    }
    If (!($WIMs))
    {
        Write-Output "No WIM or ESD files found in $Drive, aborting." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Dismount-DiskImage -ImagePath $ISO | Out-Null
        Exit
    }
    
    ForEach ($WIM in $WIMs)
    {
        $TempWIM = $WIM.FullName
        #$OSWIM = Get-WindowsImage -ImagePath $TempWIM | Where-Object {($_.ImageName -like "*$($OSSKU)") -or ($_.ImageName -like "*$($OSSKU) Evaluation") -or ($_.ImageName -like "*$OSSKU) LTSC")}
        
        # Handle different language support as per issue #1 (https://github.com/microsoft/SurfaceDeploymentAccelerator/issues/1)
        $OSImages = Get-WindowsImage -ImagePath $TempWIM

        # Read WinPEXML file
        [string]$XmlPath = "$WorkingDirPath\Languages.xml"
        [Xml]$LanguagesXML = Get-Content $XmlPath
        $Editions = $LanguagesXML.Windows10.Editions.$OSSKU.Variants.Variant

        Write-Output "Checking $TempWIM for valid images..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        $OSImageFound = $False
        ForEach ($Edition in $Editions)
        {
            ForEach ($OSImage in $OSImages)
            {
                If ($OSImage.ImageName -eq $Edition.name)
                {
                    $ImagePath = $OSImage.ImagePath
                    $ImageIndex = $OSImage.ImageIndex
                    $OSImage = Get-WindowsImage -ImagePath $ImagePath -Index $ImageIndex
                    $ImageName = $OSImage.ImageName
                    $ImageVersion = $OSImage.Version
                    $ImageArch = $OSImage.Architecture
                    $OSImageFound = $True
                }
                Else
                {
                    # Do Nothing
                }
            }
        }

        If ($OSImageFound -eq $False)
        {
            # $OSImage not found
            Write-Output "No OS Image found in $TempWIM matching $OSSKU, exiting." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Images inside $TempWIM :"
            $OSImages
            Write-Output ""
            Write-Output ""
            Start-Sleep 2
            Dismount-DiskImage -ImagePath $ISO | Out-Null
            Exit
        }
        Else
        {
            If ($ImageArch -eq "0")
            {
                $ImageArch = "x86"
            }
            If ($ImageArch -eq "9")
            {
                $ImageArch = "x64"
            }
            ElseIf ($ImageArch -eq "Unknown")
            {
                $ImageArch = "ARM64"
            }
            
            Write-Output "Found image matching $OSSKU :" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Path:          $ImagePath" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Index:         $ImageIndex" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Name:          $ImageName" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Version:       $ImageVersion" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "Architecture:  $ImageArch" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output ""
            Start-Sleep 3
            If ($IsESD -eq $True)
            {
                $TmpESDConvertWIM = "$env:TEMP\install.wim"
                $TempWIM = $TmpESDConvertWIM
                If (Test-Path "$TmpESDConvertWIM")
                {
                    Write-Output "Deleting $TmpESDConvertWIM..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Remove-Item -Path "$TmpESDConvertWIM" -Force | Out-Null
                    Write-Output ""
                }
                Write-Output "Exporting $ImagePath to $TmpESDConvertWIM..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                $Process = $DISMFile
                $ArgumentList = "/Export-Image /SourceImageFile:$ImagePath /SourceIndex:$ImageIndex /DestinationImageFile:$TmpESDConvertWIM /CheckIntegrity /Compress:fast"
                Start-Process -FilePath $Process -ArgumentList $Argumentlist -Wait -NoNewWindow
                Write-Output ""
                $ImagePath = $TmpESDConvertWIM
                $ImageIndex = "1"
                $global:OriginalOSIndex = "1"
            }
            Else
            {
                $global:OriginalOSIndex = $ImageIndex
            }

            $global:OSVersionFull = (Get-WindowsImage -ImagePath "$ImagePath" -Index "$ImageIndex").Version
            If ($global:OSVersionFull)
            {
                $global:OSVersion = $global:OSVersionFull.Substring(0, $global:OSVersionFull.LastIndexOf('.'))
                If (($global:OSVersion -like "10.0.18362*") -or ($global:OSVersion -like "10.0.19041*"))
                {
                    Write-Output "$ImagePath contains image version $global:OSVersion, validating if H1 or H2 build of $global:OSVersion - checking..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output ""
                    Write-Output "Mounting $ImagePath in $ScratchMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $ScratchMountFolder -ReadOnly | Out-Null
                    Start-Sleep 5
                    Write-Output "Querying image registry for ReleaseId..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    & reg.exe load "HKLM\Mount" "$ScratchMountFolder\Windows\system32\config\SOFTWARE"
                    $Key = "HKLM:\Mount\Microsoft\Windows NT\CurrentVersion"
                    $global:ReleaseId = (Get-ItemProperty -Path $Key -Name ReleaseId).ReleaseId
                    Start-Sleep 5
                    Write-Output "Unloading image registry..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    & reg.exe unload "HKLM\Mount"
                    Start-Sleep 5
                    Write-Output "Dismounting $ScratchMountFolder..." |Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Dismount-WindowsImage -Path $ScratchMountFolder  -Discard | Out-Null
                    Write-Output ""
                    # Specific 1909 check as it will report as 10.0.18362 still when offline
                    If ($global:ReleaseId -eq "1909")
                    {
                        $global:OSVersion = "10.0.18363"
                    }
                    # Specific 20H2 check as it will report as 10.0.19041 still when offline
                    If ($global:ReleaseId -eq "2009")
                    {
                        $global:OSVersion = "10.0.19042"
                    }
                }
                Else
                {
                    $global:ReleaseId = Switch ($global:OSVersion)
                    {
                        10.0.17763 {"1809"}
                        10.0.19041 {"2004"}
                    }
                }

                If (!($global:ReleaseID))
                {
                    Write-Output "Unknown Windows release found ( $global:OSVersion ), aborting." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Outupt ""
                    Exit
                }
            }
            Else
            {
                Write-Output "OS Version not pulled from $ImagePath, aborting." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Exit
            }
        }
    }

    If ($OSImageFound -eq $False)
    {
        Dismount-DiskImage -ImagePath $ISO | Out-Null
        Write-Output "$OSSKU not found in $WIMs on $ISO.  Please make sure to use an ISO file that contains $OSSKU, and try again." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }

    If (!(Test-Path "$Mount"))
    {
        New-Item -path "$Mount" -ItemType "directory" | Out-Null
    }

    If (!(Test-Path "$DestinationFolder"))
    {
        New-Item -path "$DestinationFolder" -ItemType "directory" | Out-Null
    }

    If (!(Test-Path "$DestinationFolder\$OSSKU"))
    {
        New-Item -path "$DestinationFolder\$OSSKU" -ItemType "directory" | Out-Null
    }

    If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion"))
    {
        New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion" -ItemType "directory" | Out-Null
    }

    If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture"))
    {
        New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture" -ItemType "directory" | Out-Null
    }

    If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp")
    {
        Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp" -Recurse -Filter *.wim | Remove-Item -Force -Recurse
        Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp" -Force -Recurse
    }
    If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp"))
    {
        New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp" -ItemType "directory" | Out-Null
    }

    If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs"))
    {
        New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -ItemType "directory" | Out-Null
    }

    If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim")
    {
        Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim" -Force
        Start-Sleep 5
    }
    
    If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim")
    {
        Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim" -Force
        Start-Sleep 5
    }

    If ($Architecture -eq "x64")
    {
        $Arch = "amd64"
    }
    ElseIf ($Architecture -eq "ARM64")
    {
        $Arch = "arm64"
    }

    Write-Output "Copying $WindowsKitsInstall\Windows Preinstallation Environment\$Arch\en-us\winpe.wim to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Copy-Item -Path "$WindowsKitsInstall\Windows Preinstallation Environment\$Arch\en-us\winpe.wim" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim"
    $SourceBootWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -filter boot.wim -Recurse
    ForEach ($SourceBootWIM in $SourceBootWIMs)
    {
        $TempBootWIM = $SourceBootWIM.FullName
        $PEWIM = Get-WindowsImage -ImagePath $TempBootWIM | Where-Object {$_.ImageName -like "*Windows PE*"}

        $ImagePath = $PEWIM.ImagePath
        $ImageIndex = $PEWIM.ImageIndex
        $ImageName = $PEWIM.ImageName
        $global:WinPEVersion = (& $DISMFile /Get-WimInfo /WimFile:$ImagePath /index:$ImageIndex | Select-String "Version ").ToString().Split(":")[1].Trim()

    }

    If ($DotNet35 -eq $true)
    {
        If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs")
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -Recurse | Remove-Item -Force -Recurse
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -Force
        }
        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs"))
        {
            New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -ItemType "directory" | Out-Null
        }
        Write-Output "Copying $Drive\Sources\sxs\* to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs\..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Copy-Item -Path "$Drive\Sources\sxs\*" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -PassThru | Set-ItemProperty -Name IsReadOnly -Value $false
    }

    If ($IsESD -eq $True)
    {
        Write-Output "Copying $TempWIM to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Copy-Item -Path $TempWIM -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -PassThru | Set-ItemProperty -Name IsReadOnly -Value $false
        Start-Sleep 2
    }
    Else
    {
        ForEach ($WIM in $WIMs)
        {
            $TempWIM = $WIM.FullName
            Write-Output "Copying $TempWIM to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Copy-Item -Path $TempWIM -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -PassThru | Set-ItemProperty -Name IsReadOnly -Value $false
            Start-Sleep 2
        }
    }
    
    If ($TmpESDConvertWIM -eq $Null)
    {
        # Do Nothing
    }
    ElseIf (Test-Path $TmpESDConvertWIM)
    {
        Write-Output "Deleting $TmpESDConvertWIM..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Remove-Item -Path $TmpESDConvertWIM -Force
        Write-Output ""
    }
    Else
    {
        # This shouldn't be possible
    }
    Dismount-DiskImage -ImagePath $ISO | Out-Null
    Write-Output ""
}



Function Add-PackageIntoWindowsImage
{
    Param(
        [string]$ImageMountFolder,
        [string]$PackagePath,
        [string]$TempImagePath,
        [bool]$DismountImageOnCompletion = $true
    )

    Add-WindowsPackage -Path $ImageMountFolder -PackagePath $PackagePath
    Write-Output ""
    Write-Output ""

    If ($DismountImageOnCompletion -eq $True)
    {
        # Dismount the image to avoid PSFX/non-PSFX update compression issues in RS5+
        Write-Output "Saving $TempImagePath..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
        Write-Output ""
        Write-Output ""
        Start-Sleep 2

        # Re-mount the image
        Write-Output "Mounting $TempImagePath in $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Mount-WindowsImage -ImagePath $TempImagePath -Index 1 -Path $ImageMountFolder -CheckIntegrity
        Write-Output ""
        Write-Output ""
    }
}



Function UpdateMenu
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]
        $MenuTitle,

        [Parameter(Mandatory=$true)]
        [string[]]
        $MenuItems,

        [Parameter(Mandatory=$false)]
        [switch]
        $ShowOnlyLastFolder,

        [Parameter(Mandatory=$true)]
        [int]
        $selection
    )

    Clear-Host

    $UpA = [char]0x2191
    $DownA = [char]0x2193
    $Count = $MenuItems.Length
    $HelperText = " $UpA, $DownA, Num (1-$Count), Enter to select:"

    Write-Host -ForegroundColor White "`n $MenuTitle "
    Write-Host -ForegroundColor White ("-"*($MenuTitle.Length + 4))

    $itemCount = 0
    foreach($item in $MenuItems){

        If ($ShowOnlyLastFolder -eq $true){
            $line = [string]$(Split-Path -Path $item -Leaf)
        } Else {
            $line = $item
        }

        If ($selection -eq $itemCount) {
            $itemCount++
            Write-Host -BackgroundColor White -ForegroundColor Black "$itemCount ] $line"
        } Else {
            $itemCount++
            Write-Host -ForegroundColor White "$itemCount ] $line"
        }
    }
    
    $viewSelection = $selection+1
    Write-Host -ForegroundColor White ("-"*($MenuTitle.Length + 4))
    If ($HelperText) {
        Write-Host -ForegroundColor Yellow $HelperText
    }
    Write-Host -ForegroundColor White ">>: $viewSelection" -NoNewline
}



Function Select-MenuItem
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]
        $MenuTitle,

        [Parameter(Mandatory=$true)]
        [string[]]
        $MenuItems,

        [Parameter(Mandatory=$false)]
        [switch]
        $ShowOnlyLastFolder
    )

    Clear-Host
    
    #Menu input type defines
    $ENTER = 13
    $UPARROW = 38
    $DOWNARROW = 40
    $LEFTARROW = 37
    $RIGHTARROW = 39
    $BACKSPACE = 8
    $DELETE = 46

    #init selection variables
    $selection = 0
    $ExitEvent = $false
    $UserInput = $null

    Do {
        #   clear key input before getting new up/down/enter etc
        $host.UI.RawUI.FlushInputBuffer()

        UpdateMenu -MenuTitle $MenuTitle -MenuItems $MenuItems -selection $selection -ShowOnlyLastFolder

        $key = ($host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")).VirtualKeyCode
        #Below lines useful for debugging
        #Write-Host -ForegroundColor Magenta $key
        #[Threading.Thread]::Sleep( 800 )

        switch ($key) {
            $UPARROW {
                if($selection -gt 0){
                    $selection--
                    $UserInput = $null
                }
            }

            $DOWNARROW {
                if($selection -lt ($MenuItems.Count - 1) ) {
                    $selection++
                    $UserInput = $null
                }
            }

            {(48..57) -contains $_ } {
                #Number 0-9 key hit
                $num = $key - 48

                #$MenuItems.Count
                $tempUI = $UserInput
                If ( ($UserInput -eq $null) -or (($UserInput.Length -gt 0) -and ($MenuItems.Count -lt 10)) ) {
                    $tempUI = [string]"$num"
                } Else {
                    $tempUI = $UserInput + [string]"$num"
                }
                #Below lines useful for debugging
                #Write-Host -ForegroundColor Magenta "tempUI: $tempUI"
                #[Threading.Thread]::Sleep( 1000 )

                $UINum = [int]$tempUI
                If( ($UINum -le 0) -or ($UINum -gt $MenuItems.Count) ) {
                    #out of range, reset
                    $UserInput = [string]($selection + 1)
                } Else {
                    $UserInput = $tempUI
                    $selection = $UINum - 1
                }

            }

            $ENTER {
                $ExitEvent = $true
            }

            $BACKSPACE {
                #Do back space stuff
                If ( $UserInput.Length -eq 1 ) {
                    $UserInput = $null
                    $selection = 0
                }

                If ( $UserInput.Length -gt 1 ) {
                    $UserInput = $UserInput.Substring(0, $UserInput.Length-1)
                    $selection = ([int]$UserInput) - 1
                }                
            }
        }

    } While ( -not $ExitEvent )

    Return $selection
}



Function Select-USBDrive
{
    $usbDisks = Get-Disk | Where-Object BusType -eq USB | Where-Object isOffline -ne True | Sort-Object Size
    $DriveNumArray = @($usbDisks | Select-Object -ExpandProperty Number)
    $MenuArray = @()
    $usbDisks | 
    Select-Object -Property Number, FriendlyName, Size | 
        ForEach-Object {
            $VolumeLabel = (Get-Disk -Number $_.Number | Get-Partition | Get-Volume).FileSystemLabel
            $MenuArray += "DISK:$("{0:D3}" -f $_.Number) ($("{0:G5} GB" -f ($_.Size /1GB))) [$VolumeLabel] $($_.FriendlyName) "
        }

    If ($DriveNumArray.Count -lt 1)
    {
        Write-output " -- No USB key Found." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Return $null
    }
    
    $SelectIndex = Select-MenuItem -MenuTitle "Select USB Drive to format" -MenuItems $MenuArray
    $diskNumToFlash = $DriveNumArray[$SelectIndex]
    $diskName = $MenuArray[$SelectIndex]

    Write-Output   $diskNumToFlash
}



Function New-RegKey
{
    Param($key)
  
    $key = $key -replace ':',''
    $parts = $key -split '\\'
  
    $tempkey = ''
    $parts | ForEach-Object {
        $tempkey += ($_ + "\")
        If ( (Test-Path "Registry::$tempkey") -eq $false)
        {
            New-Item "Registry::$tempkey" | Out-Null
        }
    }
}



Function TattooRegistry
{
    Param(
        [string]$ImageMountFolder,
        [string]$RefImage,
        [string]$SplitImage
    )

    $TempPath = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp"

    & reg.exe load "HKLM\Mount" "$ImageMountFolder\Windows\system32\config\SOFTWARE"
    Start-Sleep 2
    $SDARegKey = "HKLM:\Mount\Microsoft\Surface\SDA"
    New-RegKey $SDARegKey
    Start-Sleep 2
    
    $ISORegValue = Get-ItemProperty $SDARegKey ISO -ErrorAction SilentlyContinue
    $OSSKURegValue = Get-ItemProperty $SDARegKey OSSKU -ErrorAction SilentlyContinue
    $DotNet35RegValue = Get-ItemProperty $SDARegKey DotNet35 -ErrorAction SilentlyContinue
    $ServicingStackRegValue = Get-ItemProperty $SDARegKey ServicingStackUpdate -ErrorAction SilentlyContinue
    $CumulativeUpdateRegValue = Get-ItemProperty $SDARegKey CumulativeUpdate -ErrorAction SilentlyContinue
    $CumulativeDotNetUpdateRegValue = Get-ItemProperty $SDARegKey CumulativeDotNetUpdate -ErrorAction SilentlyContinue
    $AdobeFlashUpdateRegValue = Get-ItemProperty $SDARegKey AdobeFlashUpdate -ErrorAction SilentlyContinue
    $Office365RegValue = Get-ItemProperty $SDARegKey Office365 -ErrorAction SilentlyContinue
    $DeviceRegValue = Get-ItemProperty $SDARegKey Device -ErrorAction SilentlyContinue
    $DriverRegValue = Get-ItemProperty $SDARegKey Drivers -ErrorAction SilentlyContinue
    $ImageRegValue = Get-ItemProperty $SDARegKey Image -ErrorAction SilentlyContinue
    $OSVersionRegValue = Get-ItemProperty $SDARegKey OSVersion -ErrorAction SilentlyContinue
    $ReleaseIDRegValue = Get-ItemProperty $SDARegKey ReleaseID -ErrorAction SilentlyContinue
    $SDAVersionRegValue = Get-ItemProperty $SDARegKey SDAVersion -ErrorAction SilentlyContinue

    $ISOFileName = $ISO.Substring($ISO.LastIndexOf("\") + 1)
    If ($ISORegValue -eq $null)
    {
        New-ItemProperty -Path $SDARegKey -Name ISO -PropertyType STRING -Value $ISOFilename | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $SDARegKey -Name ISO -Value $ISOFileName
    }

    If ($OSSKURegValue -eq $null)
    {
        New-ItemProperty -Path $SDARegKey -Name OSSKU -PropertyType STRING -Value $OSSKU | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $SDARegKey -Name OSSKU -Value $OSSKU
    }

    If ($DotNet35 -eq $true)
    {
        If ($DotNet35RegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name DotNet35 -PropertyType STRING -Value $DotNet35 | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name DotNet35 -Value $DotNet35
        }
    }
    
    If ($ServicingStack -eq $true)
    {
        $PathToScan = "$TempPath\Servicing"
        $FileName = (Get-ChildItem -Path $PathToScan).Name
        If ($ServicingStackRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name ServicingStackUpdate -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name ServicingStackUpdate -Value $FileName
        }
    }

    If ($CumulativeUpdate -eq $true)
    {
        $PathToScan = "$TempPath\Cumulative"
        $FileName = (Get-ChildItem -Path $PathToScan).Name
        If ($CumulativeUpdateRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name CumulativeUpdate -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name CumulativeUpdate -Value $FileName
        }
    }

    If ($CumulativeDotNetUpdate -eq $true)
    {
        $PathToScan = "$TempPath\DotNet"
        $FileName = (Get-ChildItem -Path $PathToScan).Name
        If ($CumulativeDotNetUpdateRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name CumulativeDotNetUpdate -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name CumulativeDotNetUpdate -Value $FileName
        }
    }

    If ($AdobeFlashUpdate -eq $true)
    {
        $PathToScan = "$TempPath\Adobe"
        $FileName = (Get-ChildItem -Path $PathToScan).Name
        If ($AdobeFlashUpdateRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name AdobeFlashUpdate -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name AdobeFlashUpdate -Value $FileName
        }
    }

    If ($Office365 -eq $true)
    {
        $PathToScan = "$TempPath\Office365"
        $FileName = (Get-ChildItem -Path $PathToScan -Recurse | Where-Object { ($_.PSIsContainer) -and ($_.Name -like "16.*") }).Name
        If ($Office365RegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name Office365 -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name Office365 -Value $FileName
        }
    }

    If ($Device)
    {
        If ($DeviceRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name Device -PropertyType STRING -Value $Device | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name Device -Value $Device
        }
    }

    If (($Device) -or ($LocalDriverPath))
    {
        If ($LocalDriverPath)
        {
            $TempLocalDriverPath = (Get-Item $LocalDriverPath) -is [System.IO.DirectoryInfo]
            If ($TempLocalDriverPath -eq $False)
            {
                $FileName = (Get-ChildItem -Path $LocalDriverPath).Name
            }
            Else
            {
                $FileName = (Get-ChildItem -Path $LocalDriverPath -Recurse | Where-Object { $_.Name -like "*.msi" }).Name
            }
        }
        Else
        {
            If (Test-Path "$TempPath\$Device")
            {
                $TempLocalDriverPath = (Get-ChildItem -Path "$TempPath\$Device")
                $FileName = (Get-ChildItem -Path "$Temppath\$Device" -Recurse | Where-Object { $_.Name -like "*.msi" }).Name
            }
        }

        If ($DriverRegValue -eq $null)
        {
            New-ItemProperty -Path $SDARegKey -Name Drivers -PropertyType STRING -Value $FileName | Out-Null
        }
        Else
        {
            Set-ItemProperty -Path $SDARegKey -Name Drivers -Value $FileName
        }
    }

    If (($RefImage) -or ($SplitImage))
    {
        If ($ImageRegValue -eq $null)
        {
            If (Test-path $RefImage)
            {
                If (Test-Path $SplitImage)
                {
                    $SplitImageName = (Get-Item -Path $SplitImage).Name
                    If ($ImageRegValue -eq $null)
                    {
                        New-ItemProperty -Path $SDARegKey -Name Image -PropertyType STRING -Value $SplitImageName | Out-Null
                    }
                    Else
                    {
                        Set-ItemProperty -Path $SDARegKey -Name Image -Value $SplitImageName
                    }
                }
                Else
                {
                    $RefImageName = (Get-Item -Path $RefImage).Name
                    If ($ImageRegValue -eq $null)
                    {
                        New-ItemProperty -Path $SDARegKey -Name Image -PropertyType STRING -Value $RefImageName | Out-Null
                    }
                    Else
                    {
                        Set-ItemProperty -Path $SDARegKey -Name Image -Value $RefImageName
                    }
                }
            }
        }
    }

    If ($OSVersionRegValue -eq $null)
    {
        New-ItemProperty -Path $SDARegKey -Name OSVersion -PropertyType STRING -Value $Build | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $SDARegKey -Name OSVersion -Value $Build
    }

    If ($ReleaseIDRegValue -eq $null)
    {
        New-ItemProperty -Path $SDARegKey -Name ReleaseID -PropertyType STRING -Value $global:ReleaseID | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $SDARegKey -Name ReleaseID -Value $global:ReleaseID
    }

    If ($SDAVersionRegValue -eq $null)
    {
        New-ItemProperty -Path $SDARegKey -Name SDAVersion -PropertyType STRING -Value $SDAVersion | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $SDARegKey -Name SDAVersion -Value $SDAVersion
    }


    & reg.exe unload "HKLM\Mount"
    Start-Sleep 2
}



Function Update-Win10WIM
{
    Param(
        [string]$SourcePath,
        [string]$SourceName,
        [bool]$ServicingStack,
        [bool]$CumulativeUpdate,
        [bool]$DotNet35,
        [bool]$CumulativeDotNetUpdate,
        [bool]$AdobeFlashUpdate,
        [bool]$UpdateBootWIM,
        [string]$ImageMountFolder,
        [string]$BootImageMountFolder,
        [string]$WinREImageMountFolder,
        [string]$TempFolder,
        [string]$WindowsKitsInstall,
        [bool]$MakeUSBMedia,
        [bool]$MakeISOMedia
    )

    # Variables
    $TmpImage = "$TempFolder\tmp_install.wim"
    $TmpWinREImage = "$TempFolder\tmp_winre.wim"
    $TmpBootImage = "$TempFolder\tmp_boot.wim"
    $ServicingStackPath = "$TempFolder\Servicing"
    $CumulativeUpdatePath = "$TempFolder\Cumulative"
    $CumulativeDotNetPath = "$TempFolder\DotNet"
    $AdobeFlashUpdatePath = "$TempFolder\Adobe"
    $Office365Path = "$TempFolder\Office365"
    $DeviceDriverPath = "$TempFolder\$Device"
    $VC2013x86Path = "$TempFolder\VCRuntimes\2013\vcredist_x86.exe"
    $VC2013x64Path = "$TempFolder\VCRuntimes\2013\vcredist_x64.exe"
    $VC2019x86Path = "$TempFolder\VCRuntimes\2019\vc_redist.x86.exe"
    $VC2019x64Path = "$TempFolder\VCRuntimes\2019\vc_redist.x64.exe"
    $ProUnattendXMLPath = "$WorkingDirPath\Win10Pro_Unattend.xml"
    $EntUnattendXMLPath = "$WorkingDirPath\Win10Ent_Unattend.xml"
    $HubUnattendXMLPath = "$WorkingDirPath\Win10Hub_Unattend.xml"
    $OfficeAuditXMLPath = "$WorkingDirPath\Win10_Audit_Office.xml"
    $NoOfficeAuditXMLPath = "$WorkingDirPath\Win10_Audit_NoOffice.xml"
    $InstallOfficeScriptPath = "$WorkingDirPath\InstallOffice.ps1"
    $SetTaskBarPinsScriptPath = "$WorkingDirPath\SetTaskBarPins.ps1"
    $SysprepToOOBEScriptPath = "$WorkingDirPath\SysprepToOOBE.ps1"
    
    <#
    $SourceName = Switch ($SourceName)
    {
        Pro {"Windows 10 Pro"}
        Enterprise {"Windows 10 Enterprise"}
    }
    #>
    
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *           Updating install.wim            *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Start-Sleep 2

    If ($InstallWIM)
    {
        # Export the reference image to a new (temporary) WIM - this will leave the original "install.wim" untouched when finished
        If (Test-Path "$SourcePath\install.wim")
        {
            Write-Output "Exporting $SourcePath\install.wim to $TmpImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Export-WindowsImage -SourceImagePath "$SourcePath\install.wim" -SourceIndex $global:OriginalOSIndex -DestinationImagePath $TmpImage -CheckIntegrity
            #Export-WindowsImage -SourceImagePath "$SourcePath\install.wim" -SourceName "$SourceName" -DestinationImagePath $TmpImage -CheckIntegrity
            Write-Output ""
            Write-Output ""
        }
        Else
        {
            Write-Output "No WIM file found in $SourcePath, aborting." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Exit
        }

        # Mount the image
        Write-Output "Mounting $TmpImage in $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Mount-WindowsImage -ImagePath $TmpImage -Index 1 -Path $ImageMountFolder -CheckIntegrity
        Write-Output ""
        Write-Output ""

        If ($DotNet35 -eq $True)
        {
            # Cleanup the image BEFORE installing .NET to prevent errors
            Write-Output "Running image cleanup on $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            & $DISMFile /Image:$ImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase
            Write-Output ""
            Write-Output ""

            # Dismount the image
            Write-Output "Saving $TmpImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
            Write-Output ""
            Write-Output ""
            Start-Sleep 10

            # Re-mount the image
            Write-Output "Mounting $TmpImage in $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Mount-WindowsImage -ImagePath $TmpImage -Index 1 -Path $ImageMountFolder -CheckIntegrity
            Write-Output ""
            Write-Output ""

            # Add .NET Framework 3.5 to the image
            Write-Output "Adding .NET Framework 3.5 to $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Enable-WindowsOptionalFeature -Path $ImageMountFolder -FeatureName NetFx3 -All -Source "$TempFolder\sxs" -LimitAccess
            Write-Output ""
            Write-Output ""
            Start-Sleep 2
        }

        If ($ServicingStack -eq $true)
        {
            $SSU = Get-ChildItem -Path $ServicingStackPath
            If (!($SSU.Exists))
            {
                $ServicingStack = $False
            }
            Else
            {
                # Add required Servicing Stack updates
                Write-Output "Adding Servicing Stack updates to $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpImage -DismountImageOnCompletion $True
                Start-Sleep 2
            }
        }

        If ($CumulativeUpdate -eq $true)
        {
            $CU = Get-ChildItem -Path $CumulativeUpdatePath
            If (!($CU.Exists))
            {
                $CumulativeUpdate = $False
            }
            Else
            {
                # Add monthly Cumulative update
                Write-Output "Adding Cumulative updates to $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }

        If ($CumulativeDotNetUpdate -eq $true)
        {
            $CUDN = Get-ChildItem -Path $CumulativeDotNetPath
            If (!($CUDN.Exists))
            {
                $CumulativeDotNetUpdate = $False
            }
            Else
            {
                # Add monthly Cumulative update
                Write-Output "Adding Cumulative .NET updates to $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $CumulativeDotNetPath -TempImagePath $TmpImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }
        
        If ($AdobeFlashUpdate -eq $true)
        {
            $AFU = Get-ChildItem -Path $AdobeFlashUpdatePath
            If (!($AFU.Exists))
            {
                $AdobeFlashUpdate = $False
            }
            Else
            {
                # Add Adobe Flash updates
                Write-Output "Adding Adobe Flash updates to $ImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $AdobeFlashUpdatePath -TempImagePath $TmpImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }

        If ($Office365 -eq $True)
        {
            # Copy Office 365 bits to device
            If (Test-Path "$ImageMountFolder\Windows\Temp\Office365")
            {
                Write-Output "$ImageMountFolder\Windows\Temp\Office365..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\Office365" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\Office365" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\Office365"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\Office365" -ItemType Directory | Out-Null
            }

            If (!($Architecture -eq "ARM64"))
            {
                Write-Output "Copying Office365 files to $ImageMountFolder\Windows\Temp..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Copy-Item -Path $InstallOfficeScriptPath -Destination "$Office365Path\InstallOffice.ps1" -Force -ErrorAction Continue
                & xcopy.exe /herky "$Office365Path" "$ImageMountFolder\Windows\Temp\Office365"
                Write-Output ""
            }
        }

        If ($Device)
        {
            $MSIFiles = Get-ChildItem -Path $DeviceDriverPath -Recurse
            # Add drivers/firmware to WIM
            Write-Output "Adding Driver updates for $Device to $ImageMountFolder from $DeviceDriverPath..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Add-WindowsDriver -Path $ImageMountFolder -Driver "$DeviceDriverPath" -Recurse
            Write-Output ""
            Write-Output ""

            # Copy VC++ Runtimes
            If (Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013")
            {
                Write-Output "Deleting $ImageMountFolder\Windows\Temp\VCRuntimes\2013..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -ItemType Directory | Out-Null
            }

            If (Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2019")
            {
                Write-Output "Deleting $ImageMountFolder\Windows\Temp\VCRuntimes\2019..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2019" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2019" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2019"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2019" -ItemType Directory | Out-Null
            }
        }

        If (!($Architecture -eq "ARM64"))
        {
            Write-Output "Copying VC++ Runtime binaries to $ImageMountFolder\Windows\Temp..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Copy-Item -Path $VC2013x86Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"
            Copy-Item -Path $VC2013x64Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"
            Copy-Item -Path $VC2019x86Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2019"
            Copy-Item -Path $VC2019x64Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2019"
            Write-Output ""
        }

        Write-Output "Copying files to disk for unattended installation..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Copy-Item -Path $SysprepToOOBEScriptPath -Destination "$ImageMountFolder\Windows\Temp\SysprepToOOBE.ps1" -Force -ErrorAction Continue
        If ($Office365 -eq $true)
        {
            Copy-Item -Path $OfficeAuditXMLPath -Destination "$ImageMountFolder\Windows\System32\sysprep\unattend.xml" -Force -ErrorAction Continue
            Copy-Item -Path $SetTaskBarPinsScriptPath -Destination "$ImageMountFolder\Windows\Temp\SetTaskBarPins.ps1" -Force -ErrorAction Continue
        }
        Else
        {
            Copy-Item -Path $NoOfficeAuditXMLPath -Destination "$ImageMountFolder\Windows\System32\sysprep\unattend.xml" -Force -ErrorAction Continue
        }

        If ($Device -eq "SurfaceHub2")
        {
            Copy-Item -Path $HubUnattendXMLPath -Destination "$ImageMountFolder\Windows\Temp\Reseal.xml" -Force -ErrorAction Continue
        }
        Else
        {
            If ($OSSKU -like "*Pro*")
            {
                Copy-Item -Path $ProUnattendXMLPath -Destination "$ImageMountFolder\Windows\Temp\Reseal.xml" -Force -ErrorAction Continue
            }
            ElseIf ($OSSKU -like "*Enterprise*")
            {
                Copy-Item -Path $EntUnattendXMLPath -Destination "$ImageMountFolder\Windows\Temp\Reseal.xml" -Force -ErrorAction Continue
            }
            Write-Output ""
        }


        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *           Updating winre.wim              *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep 2


        # Copy WinRE Image to temp location
        Write-Output "Copying WinRE image to $TmpWinREImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Move-Item -Path "$ImageMountFolder\Windows\System32\Recovery\winre.wim" -Destination $TmpWinREImage
        Write-Output ""
        Write-Output ""

        # Mount the temp WinRE Image
        Write-Output "Mounting $TmpWinREImage to $WinREImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Mount-WindowsImage -ImagePath $TmpWinREImage -Index 1 -Path $WinREImageMountFolder -CheckIntegrity
        Write-Output ""
        Write-Output ""

        If ($ServicingStack)
        {
            $SSU = Get-ChildItem -Path $ServicingStackPath
            If (!($SSU.Exists))
            {
                $ServicingStack = $False
            }
            Else
            {
                # Add Servicing Stack updates to the WinRE image 
                Write-Output "Adding Servicing Stack updates to $WinREImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $WinREImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpWinREImage -DismountImageOnCompletion $True
                Start-Sleep 2
            }
        }

        If ($CumulativeUpdate)
        {
            $CU = Get-ChildItem -Path $CumulativeUpdatePath
            If (!($CU.Exists))
            {
                $CumulativeUpdate = $False
            }
            Else
            {
                # Add monthly Cumulative updates to the WinRE image
                Write-Output "Adding Cumulative updates to $WinREImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $WinREImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpWinREImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }

        If ($Device)
        {
            $MSIFiles = Get-ChildItem -Path $DeviceDriverPath -Recurse
            If ($SurfaceDevices.$Device)
            {
                # Add system-level drivers to WIM
                Write-Output "Adding Driver updates for $Device to $WinREImageMountFolder from $DeviceDriverPath..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                $Drivers = $SurfaceDevices.$Device.Drivers.Driver
                ForEach ($Driver in $Drivers)
                {
                    $TempDriverName = $Driver.name
                    ForEach ($MSIFile in $MSIFiles)
                    {
                        If ($MSIFile.Name -eq $TempDriverName)
                        {
                            Add-WindowsDriver -Path $WinREImageMountFolder -Driver $MSIFile.FullName
                        }
                    }
                }
            }
            Write-Output ""
            Write-Output ""
        }

        # Cleanup the WinRE image
        Write-Output "Running image cleanup on $TmpWinREImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        & $DISMFile /Image:$WinREImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase
        Write-Output ""
        Write-Output ""
        Start-Sleep 2

        # Dismount the WinRE image
        Write-Output "Saving $TmpWinREImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        DisMount-WindowsImage -Path $WinREImageMountFolder -Save -CheckIntegrity
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *            Saving winre.wim               *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep 2


        # Export the new WinRE image back to original location
        Write-Output "Exporting $TmpWinREImage to $ImageMountFolder\Windows\System32\Recovery\winre.wim..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Export-WindowsImage -SourceImagePath $TmpWinREImage -SourceName "Microsoft Windows Recovery Environment (x64)" -DestinationImagePath "$ImageMountFolder\Windows\System32\Recovery\winre.wim" -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""


        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *            Saving install.wim             *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep 2


        # Validate Windows WIM build number
        $Build = (Get-Item $ImageMountFolder\Windows\System32\ntoskrnl.exe).VersionInfo.ProductVersion
        If (($global:ReleaseId -eq "1909") -and ($Build -match "18362"))
        {
            $Build = $Build -replace "18362", "18363"
        }
        If (($global:ReleaseId -eq "2009") -and ($Build -match "19041"))
        {
            $Build = $Build -replace "19041", "19042"
        }
        If ($Device)
        {
            $RefImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Install-$Build-$OSSKU-$Now.wim"
            $SplitImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Install-$Build-$OSSKU-$Now--Split.swm"
        }
        Else
        {
            $RefImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Generic-Install-$Build-$OSSKU-$Now.wim"
            $SplitImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Generic-Install-$Build-$OSSKU-$Now--Split.swm"
        }

        Write-Output "Adding registry tattoo..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output ""
        TattooRegistry -ImageMountFolder $ImageMountFolder -RefImage $RefImage -SplitImage $SplitImage
        Start-Sleep 2

        # Dismount the reference image
        Write-Output "Saving $TmpImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""

        # Export the image to a new WIM
        Write-Output "Exporting $TmpImage to $RefImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Export-WindowsImage -SourceImagePath $TmpImage -SourceIndex 1 -DestinationImagePath $RefImage -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""

        $TempRefImageSize = Get-Item $RefImage
        $RefImageSize = ($TempRefImageSize.Length /1GB)
        If ($RefImageSize -ge "4")
        {
            $SplitWIM = $true
            # Split the WIM to fit on FAT32-formatted media (splitting at ~3GB for simplicity)
            Write-Output "Splitting $RefImage into 3GB files as $SplitImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Split-WindowsImage -ImagePath $RefImage -SplitImagePath $SplitImage -FileSize 3096 -CheckIntegrity
            Start-Sleep 2
            Write-Output ""
            Write-Output ""
        }

        # Remove temporary WIMs
        If (Test-Path -path $TmpImage)
        {
            Remove-Item -Path $TmpImage -Force
        }
        If (Test-Path -path $TmpWinREImage)
        {
            Remove-Item -Path $TmpWinREImage -Force
        }
        If ($SplitWIM -eq $True)
        {
            If ($KeepOriginalWIM -eq $True)
            {
                #Don't delete original .wim file
            }
            ElseIf (Test-Path -path $RefImage)
            {
                Remove-Item -Path $RefImage
            }
        }
    }

    If ($UpdateBootWIM -eq $True)
    {

        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *           Updating boot.wim               *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep 2


        # Copy boot.wim for editing
        Write-Output "Copying $SourcePath\boot.wim to $TmpBootImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Copy-Item "$SourcePath\boot.wim" $TempFolder
        Attrib -r "$TempFolder\boot.wim"
        Rename-Item -Path "$TempFolder\boot.wim" -NewName "$TmpBootImage"
        Write-Output ""
        Write-Output ""


        # Mount index 1 of the boot image (WinPE)
        Write-Output "Mounting $TmpBootImage to $BootImageMountFolder using Index 1..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Mount-WindowsImage -ImagePath $TmpBootImage -Index 1 -Path $BootImageMountFolder -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""


        <#If ($ServicingStack)
        {
            $SSU = Get-ChildItem -Path $ServicingStackPath
            If (!($SSU.Exists))
            {
                $ServicingStack = $False
            }
            Else
            {
                # Add required Servicing Stack updates
                Write-Output "Adding Servicing Stack updates to $BootImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpBootImage -DismountImageOnCompletion $True
                Start-Sleep 2
            }
        }

        If ($CumulativeUpdate)
        {
            $CU = Get-ChildItem -Path $CumulativeUpdatePath
            If (!($CU.Exists))
            {
                $CumulativeUpdate = $False
            }
            Else
            {
                # Add monthly Cumulative update
                Write-Output "Adding Cumulative updates to $BootImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpBootImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }

        If ($CumulativeDotNetUpdate)
        {
            $CUDN = Get-ChildItem -Path $CumulativeDotNetPath
            If (!($CUDN.Exists))
            {
                $CumulativeDotNetUpdate = $False
            }
            Else
            {
                # Add monthly Cumulative update
                Write-Output "Adding Cumulative .NET updates to $BootImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $CumulativeDotNetPath -TempImagePath $TmpBootImage -DismountImageOnCompletion $False
                Start-Sleep 2
            }
        }#>

        If ($Device)
        {
            $MSIFiles = Get-ChildItem -Path $DeviceDriverPath -Recurse
            If ($SurfaceDevices.$Device)
            {
                # Add system-level drivers to WIM
                Write-Output "Adding Driver updates for $Device to $BootImageMountFolder from $DeviceDriverPath..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                $Drivers = $SurfaceDevices.$Device.Drivers.Driver
                ForEach ($Driver in $Drivers)
                {
                    $TempDriverName = $Driver.name
                    ForEach ($MSIFile in $MSIFiles)
                    {
                        If ($MSIFile.Name -eq $TempDriverName)
                        {
                            Add-WindowsDriver -Path $BootImageMountFolder -Driver $MSIFile.FullName
                        }
                    }
                }
            }
            Write-Output ""
            Write-Output ""
        }

        # Add support for deployment components
        If ($Architecture -eq "x64")
        {
            $WinPEOCPath = "$WindowsKitsInstall\Windows Preinstallation Environment\amd64\WinPE_OCs"
        }
        ElseIf ($Architecture -eq "ARM64")
        {
            $WinPEOCPath = "$WindowsKitsInstall\Windows Preinstallation Environment\arm64\WinPE_OCs"
        }

        Write-Output "Adding WMI..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-WMI.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-WMI_en-us.cab" | Out-Null

        Write-Output "Adding PE Scripting..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-Scripting.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-Scripting_en-us.cab" | Out-Null

        Write-Output "Adding Enhanced Storage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-EnhancedStorage.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-EnhancedStorage_en-us.cab" | Out-Null

        Write-Output "Adding Bitlocker support..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-SecureStartup.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-SecureStartup_en-us.cab" | Out-Null

        Write-Output "Adding .NET..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-NetFx.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-NetFx_en-us.cab" | Out-Null

        Write-Output "Adding PowerShell..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-PowerShell.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-PowerShell_en-us.cab" | Out-Null

        Write-Output "Adding Storage WMI..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-StorageWMI.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-StorageWMI_en-us.cab" | Out-Null

        Write-Output "Adding DISM support..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-DismCmdlets.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-DismCmdlets_en-us.cab" | Out-Null

        Write-Output "Adding Secure Boot support..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-SecureBootCmdlets.cab" | Out-Null

        Write-Output "Adding Secure Startup support..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-DismCmdlets.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-DismCmdlets_en-us.cab" | Out-Null

        Write-Output "Adding WinRE support..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-WinReCfg.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-WinReCfg_en-us.cab" | Out-Null


        If (($MakeUSBMedia) -or ($MakeISOMedia))
        {
            Write-Output "Copying scripts to $BootImageMountFolder..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI.txt" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI_Source.txt" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\Imaging.ps1" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\Install.cmd" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\startnet.cmd" -Destination "$BootImageMountFolder\Windows\System32" -Force
        }

        Write-Output ""
        Write-Output ""


        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *            Saving boot.wim                *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Start-Sleep 2


        # Variable
        $WinPEBuild = (Get-Item $BootImageMountFolder\Windows\System32\ntoskrnl.exe).VersionInfo.ProductVersion
        If ($Device)
        {
            $RefBootImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Boot-$WinPEBuild-$Now.wim"
        }
        Else
        {
            $RefBootImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Generic-Boot-$WinPEBuild-$Now.wim"
        }


        # Dismount the boot image
        Write-Output "Saving $TmpBootImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        DisMount-WindowsImage -Path $BootImageMountFolder -Save -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""

        # Export the image to a new WIM
        Write-Output "Exporting $TmpBootImage to $RefBootImage..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Export-WindowsImage -SourceImagePath $TmpBootImage -SourceIndex 1 -DestinationImagePath $RefBootImage -CheckIntegrity
        Start-Sleep 2
        Write-Output ""
        Write-Output ""

        # Remove the temporary WIM
        If (Test-Path -path $TmpBootImage)
        {
            Remove-Item -Path $TmpBootImage -Force
            Write-Output ""
            Write-Output ""
        }
    }


    # Make a USB key or ISO
    If (($MakeUSBMedia) -or ($MakeISOMedia))
    {
        If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media")
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -Recurse | Remove-Item -Force -Recurse
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -Force
        }
        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"))
        {
            New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -ItemType "directory" | Out-Null
        }

        If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles")
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles" -Recurse | Remove-Item -Force -Recurse
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles" -Force
        }
        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles"))
        {
            New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles" -ItemType "directory" | Out-Null
        }

        If ($Architecture -eq "x64")
        {
            $Arch = "amd64"
        }
        ElseIf ($Architecture -eq "ARM64")
        {
            $Arch = "arm64"
        }

        Write-Output "Creating WinPE media in $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        & xcopy.exe /herky "$WindowsKitsInstall\Windows Preinstallation Environment\$Arch\Media" "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        Copy-Item -Path "$WindowsKitsInstall\Deployment Tools\$Arch\Oscdimg\efisys.bin" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles"
        Copy-Item -Path "$WindowsKitsInstall\Deployment Tools\$Arch\Oscdimg\etfsboot.com" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles"

        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media\Sources"))
        {
            New-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media\sources" -ItemType Directory | Out-Null
        }
        Copy-Item -Path $RefBootImage -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media\sources\boot.wim"
        Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI.txt" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI_Source.txt" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        Copy-Item -Path "$WorkingDirPath\UsbImage\Imaging.ps1" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        Copy-Item -Path "$WorkingDirPath\UsbImage\Install.cmd" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        Copy-Item -Path "$WorkingDirPath\UsbImage\startnet.cmd" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
        
        If ($MakeUSBMedia)
        {
            Write-Output "Insert USB drive 16GB+ in size, and press ENTER to view the drive selection menu" | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output "  Note that choosing a USB drive on the next screen WILL FORMAT THE DRIVE  " | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output ""
            PAUSE
            Start-Sleep 5

            # Find USB Drive that the image will be copied to.
            $TempUSB = Select-USBDrive
            Write-Output ""

            If (!($TempUSB))
            {
                Write-Output "No USB key found, skipping..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            }
            Else
            {
                $USB = Get-Disk | Where-Object {$_.Number -eq $TempUSB} | Get-Partition | Get-Volume
                If ($USB)
                {
                    $USBVolumeLabel = @($USB.FileSystemLabel)
                }
                $USBDiskName = Get-Disk |
                               Where-Object Number -eq $TempUSB |
                               ForEach-Object { "DISK:$("{0:D3}" -f $_.Number) ($("{0:G5} GB" -f ($_.Size /1GB))) $($_.FriendlyName)"}
                $UserInput = Read-Host -Prompt "`n`nAre you sure you want to format: [$USBVolumeLabel] on ($USBDiskName) (Y/N)?"

                If ( $UserInput -ne "y" )
                {
                    Write-Output " -- Aborting Operation" | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                }
                Else
                {
                    $USBSize = $USB.Size /1GB

                    Get-Disk -Number $TempUSB | Clear-Disk -RemoveData -Confirm:$false
                    Initialize-Disk -Number $TempUSB -PartitionStyle MBR -ErrorAction SilentlyContinue
                    Write-Output ""
                    Write-Output ""
                    Write-Output "DEBUG:   USB disk size:" | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output "DEBUG:   $USBSize" | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Write-Output ""
                    Write-Output ""
    
                    If ($USBSize -ge "30")
                    {
                        $NewUSBDriveLetter = New-Partition -DiskNumber $TempUSB -Size 32GB -AssignDriveLetter | Format-Volume -FileSystem FAT32 -NewFileSystemLabel $Device
                    }
                    ElseIf ($USBSize -lt "14")
                    {
                        Write-Output "USB drive appears to be smaller than 16GB, skipping..." | Receive-Output -Color Yellow -LogLevel 2 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Write-Output ""
                    }
                    Else
                    {
                        $NewUSBDriveLetter = New-Partition -DiskNumber $TempUSB -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem FAT32 -NewFileSystemLabel $Device
                    }

                    $NewUSBDriveLetter = $NewUSBDriveLetter.DriveLetter + ":"

                    Write-Output "Copying WinPE Media contents to $NewUSBDriveLetter..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    & bootsect.exe /nt60 $NewUSBDriveLetter /force /mbr
                    & xcopy /herky "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" $NewUSBDriveLetter
    
                    If ($SplitWIM -eq $True)
                    {
                        $SplitWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture" -Filter *install*$Now*.swm -Recurse
                        ForEach ($TempWIM in $SplitWIMs)
                        {
                            $TempSplitWIM = $TempWIM.FullName
                            Write-Output "Copying $TempSplitWIM to $NewUSBDriveLetter..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                            Copy-Item -Path "$TempSplitWIM" -Destination "$NewUSBDriveLetter\Sources" -Force
                        }
                    }
                    Else
                    {
                        Write-Output "Copying $RefImage to $NewUSBDriveLetter..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                        Copy-Item -Path "$RefImage" -Destination "$NewUSBDriveLetter\Sources" -Recurse
                    }
                }
            }
        }

        If ($MakeISOMedia)
        {
            $oscdimg = "$WindowsKitsInstall\Deployment Tools\$Arch\Oscdimg\oscdimg.exe"
            $efisys = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\efisys.bin"
            $etfsboot = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\etfsboot.com"
            $MediaSource = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
            $args = "-l$Device -bootdata:2#p0,e,b$etfsboot#pEF,e,b$efisys -m -u1 -udfver102 $MediaSource $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-$Build-$Now.iso"
            
            If ($SplitWIM -eq $True)
            {
                $SplitWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture" -Filter *install*$Now*.swm -Recurse
                ForEach ($TempWIM in $SplitWIMs)
                {
                    $TempSplitWIM = $TempWIM.FullName
                    Write-Output "Copying $TempSplitWIM to $MediaSource..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                    Copy-Item -Path "$TempSplitWIM" -Destination "$MediaSource\Sources" -Force
                }
            }
            Else
            {
                Write-Output "Copying $RefImage to $MediaSource..." | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
                Copy-Item -Path "$RefImage" -Destination "$MediaSource\Sources" -Recurse
            }

            Start-Process -FilePath $oscdimg -ArgumentList $args -NoNewWindow -Wait
        }
    }


    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *       Image modifications complete!       *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Start-Sleep 2

    Set-Location -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture"
    Write-Output "Finalized image files can be found here:" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
    If ($CreateISO)
    {
        If (Test-Path("$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-$Build-$Now.iso"))
        {
            Write-Output "ISO:      $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-$Build-$Now.iso" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Write-Output ""
        }
    }
    If ($SplitWIM -eq $True)
    {
        Write-Output "Install:  $SplitImage" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    }
    Else
    {
        Write-Output "Install:  $RefImage" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    }
    Write-Output "Boot:     $RefBootImage" | Receive-Output -Color Green -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Write-Output ""
}



###########################
# Begin script processing #
###########################
Clear-Host


# Get current working directory
$Invocation = (Get-Variable MyInvocation).Value
$WorkingDirPath = Split-Path $Invocation.MyCommand.Path
If (!($DestinationFolder))
{
    $DestinationFolder = $WorkingDirPath
}


# Get script start time (will be used to determine how long execution takes)
$Script_Start_Time = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
$Now = Get-Date -Format yyyy-MM-dd_HH-mm-ss

# Start logging
$SourcePath = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs"
$TempFolder = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp"
$LogFilePath = "$DestinationFolder\Logs"
$LogFileName = "Log--$OSSKU-$Architecture--$Now.log"
Start-Log -FilePath $LogFilePath -FileName $LogFileName
Write-Output "Script start: $Script_Start_Time" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"


If ($Device)
{
    # Read WinPEXML file
    [string]$XmlPath = "$WorkingDirPath\WinPE_Drivers.xml"
    [Xml]$WinPEXML = Get-Content $XmlPath
    
    $SurfaceDevices = $WinPEXML.Surface.Devices
}


# Necessary variables not passed into script directly
$DISMFile = "$WindowsKitsInstall\Deployment Tools\amd64\DISM\dism.exe"
$ADKWinPEFile = "$WindowsKitsInstall\Windows Preinstallation Environment\amd64\en-us\winpe.wim"
$Mount = "$env:TEMP\Mount"
$ScratchMountFolder = "$Mount\Scratch"

if(Test-Path -Path $Mount)
{
    try
    {
        If ((Get-ChildItem -Path $Mount -File -Force -Recurse | Select-Object -First 1 | Measure-Object).Count -gt 0)
        {
            Write-Output "Previous interrupted execution detected. $Mount must be empty for execution to continue." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Exit
        }
    }
    catch
    {
        Write-Output "Previous interrupted execution detected. $Mount must be empty for execution to continue." | Receive-Output -Color Red -BGColor Black -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        Exit
    }
}


# Leave blank space at top of window to not block output by progress bars
AddHeaderSpace


# Check for admin rights and ADK install
PrereqCheck


Write-Output "SDA version:  $SDAVersion" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " *       Parameters passed to script:        *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " *                                           *" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " *********************************************" | Receive-Output -Color Cyan -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "ISO path:                     $ISO" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "OS SKU:                       $OSSKU" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "Architecture:                 $Architecture" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "Output:                       $DestinationFolder" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  .NET 3.5:                   $DotNet35" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Servicing Stack:            $ServicingStack" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Cumulative Update:          $CumulativeUpdate" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Cumulative DotNet Updates:  $CumulativeDotNetUpdate" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Adobe Flash Player Updates: $AdobeFlashUpdate" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Office 365 install:         $Office365" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
If ($Device)
{
    Write-Output "  Device drivers:             $Device" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}
If ($UseLocalDriverPath -eq $True)
{
    Write-Output "  Use Local driver path:      $LocalDriverPath" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}
Write-Output "  Create USB key:             $CreateUSB" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "  Create ISO:                 $CreateISO" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output " " | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Start-Sleep 2


# Pull Windows 10 version and SKU from ISO provided by script param, returns OSVersion and WinPEVersion variable as well
Get-OSWIMFromISO -ISO $ISO -OSSKU $OSSKU -DestinationFolder $DestinationFolder -Architecture $Architecture -WindowsKitsInstall $WindowsKitsInstall -ScratchMountFolder $ScratchMountFolder
Start-Sleep 2
Write-Output "OSVersion:  $global:OSVersion" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "ReleaseId:  $global:ReleaseId" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output ""
Start-Sleep 5


# Variables needed after Get-OSWIMFromISO finishes, passed to Update-Win10WIM
$SourcePath = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs"
$TempFolder = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp"
$ImageMountFolder = "$Mount\OSImage"
$BootImageMountFolder = "$Mount\BootImage"
$WinREImageMountFolder = "$Mount\WinREImage"


If (Test-Path "$ImageMountFolder")
{
    Write-Output "Deleting $ImageMountFolder\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-ChildItem -Path "$ImageMountFolder" -Recurse | Remove-Item -Force -Recurse
    Remove-Item -Path "$ImageMountFolder" -Force
}
If (!(Test-Path -path $ImageMountFolder))
{
    New-Item -path $ImageMountFolder -ItemType Directory | Out-Null
}

If (Test-Path "$BootImageMountFolder")
{
    Write-Output "Deleting $BootImageMountFolder\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-ChildItem -Path "$BootImageMountFolder" -Recurse | Remove-Item -Force -Recurse
    Remove-Item -Path "$BootImageMountFolder" -Force
}
If (!(Test-Path -path $BootImageMountFolder))
{
    New-Item -path $BootImageMountFolder -ItemType Directory | Out-Null
}

If (Test-Path "$WinREImageMountFolder")
{
    Write-Output "Deleting $WinREImageMountFolder\..." | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Get-ChildItem -Path "$WinREImageMountFolder" -Recurse | Remove-Item -Force -Recurse
    Remove-Item -Path "$WinREImageMountFolder" -Force
}
If (!(Test-Path -path $WinREImageMountFolder))
{
    New-Item -path $WinREImageMountFolder -ItemType Directory | Out-Null
}


If ($BootWIM -eq $true)
{
    $UpdateBootWIM = $True
}
Else
{
    $UpdateBootWIM = $False
}

# If installing DotNet 3.5, the latest updates are also required - override any user parameters
If ($DotNet35 -eq $True)
{
    $ServicingStack = $True
    $CumulativeUpdate = $True
    $CumulativeDotNetUpdate = $True
}

# Latest Servicing Stack is likely needed (if it exists) for the latest Cumulative Update to install successfully
If ($CumulativeUpdate -eq $True)
{
    $ServicingStack = $True
}



# Download any components requested
If ($Device)
{
    Get-LatestDrivers -TempFolder $TempFolder -Device $Device
}

If ($Office365 -eq $True)
{
    Get-Office365 -TempFolder $TempFolder
}

# We always need the VC Runtimes for our devices
Get-LatestVCRuntimes -TempFolder $TempFolder

If ($ServicingStack -eq $True)
{
    Get-ServicingStackUpdates -TempFolder $TempFolder
}

If ($CumulativeUpdate -eq $True)
{
    Get-CumulativeUpdates -TempFolder $TempFolder
}

If ($DotNet35 -eq $True)
{
    Get-CumulativeDotNetUpdates -TempFolder $TempFolder
}

If ($AdobeFlashUpdate -eq $True)
{
	Get-AdobeFlashUpdates -TempFolder $TempFolder
}


# Add Servicing Stack / Cumulative updates and necessary drivers to install.wim, winre.wim, and boot.wim
Update-Win10WIM -SourcePath $SourcePath -SourceName $OSSKU -ServicingStack $ServicingStack -CumulativeUpdate $CumulativeUpdate -DotNet35 $DotNet35 -CumulativeDotNetUpdate $CumulativeDotNetUpdate -AdobeFlashUpdate $AdobeFlashUpdate -ImageMountFolder $ImageMountFolder -BootImageMountFolder $BootImageMountFolder -WinREImageMountFolder $WinREImageMountFolder -TempFolder $TempFolder -WindowsKitsInstall $WindowsKitsInstall -UpdateBootWIM $UpdateBootWIM -MakeUSBMedia $CreateUSB -MakeISOMedia $CreateISO


# Determine ending time
$Script_End_Time = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

# How long did this take?
Write-Output "Script start: $Script_Start_Time" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "Script end:   $Script_End_Time" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output ""
Write-Output "Execution time: $Script_Time_Taken seconds" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
