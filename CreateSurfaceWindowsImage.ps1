<#
.SYNOPSIS
    This script downloads the ADK and WinPE addon.

.DESCRIPTION
    This script downloads the ADK and WinPE addon, uninstalls any previous versions and installs the version referenced by the aka.ms link in the script.

    
    // *************
    // *  CAUTION  *
    // *************

    Please review this script THOROUGHLY before applying, and disable changes below as necessary to suit your current environment.

    This script is provided AS-IS - usage of this source assumes that you are at the very least familiar with PowerShell, and the
    tools used to create and debug this script.

    In other words, if you break it, you get to keep the pieces.
    
.EXAMPLE
    .\CreateSurfaceWindowsImage.ps1 -ISO <ISO path> -OSSKU Pro -Device SurfacePro7

.NOTES
    Author:       Microsoft
    Last Update:  6th May 2020
    Version:      1.1.0

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
        HelpMessage="What SKU should be used inside ISO (valid parameters are 'Pro' or 'Enterprise'), default is Pro"
        )]
        [ValidateSet('Pro', 'Enterprise')]
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
        [bool]$DotNet35 = $False,

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
        HelpMessage="Add latest Adobe Flash Player Security update (bool true/false, default is true)"
        )]
        [bool]$AdobeFlashUpdate = $True,

    [Parameter(
        Position=9,
        Mandatory=$False,
        HelpMessage="Surface device type to add drivers to image for, if not specified no drivers injected - Custom can be used if using with a non-Surface device"
        )]
        [ValidateSet('SurfacePro4', 'SurfacePro5', 'SurfacePro6', 'SurfacePro7', 'SurfaceLaptop', 'SurfaceLaptop2', 'SurfaceLaptop3', 'SurfaceBook', 'SurfaceBook2', 'SurfaceBook3', 'SurfaceStudio', 'SurfaceStudio2', 'SurfaceGo', 'SurfaceGoLTE', 'SurfaceGo2', 'Custom')]
        [string]$Device = "SurfacePro7",

    [Parameter(
        Position=10,
        Mandatory=$False,
        HelpMessage="Create USB key when finished (bool true/false, default is false)"
        )]
        [bool]$CreateUSB = $False,

    [Parameter(
        Position=11,
        Mandatory=$False,
        HelpMessage="Create bootable ISO file (useful for testing) when finished (bool true/false, default is false)"
        )]
        [bool]$CreateISO = $False,

    [Parameter(
        Position=12,
        Mandatory=$False,
        HelpMessage="Location of Windows ADK installation"
        )]
        [string]$WindowsKitsInstall = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit",

    [Parameter(
        Position=13,
        Mandatory=$False,
        HelpMessage="Use BITS for downloads"
        )]
        [bool]$BITSTransfer = $True,

    [Parameter(
        Position=14,
        Mandatory=$False,
        HelpMessage="Edit Install.wim"
        )]
        [bool]$InstallWIM = $True,

    [Parameter(
        Position=15,
        Mandatory=$False,
        HelpMessage="Edit boot.wim"
        )]
        [bool]$BootWIM = $True,

    [Parameter(
        Position=16,
        Mandatory=$False,
        HelpMessage="Keep original unsplit WIM even if resulting image size >4GB (bool true false, default is true)"
        )]
        [bool]$KeepOriginalWIM = $True,

    [Parameter(
        Position=17,
        Mandatory=$False,
        HelpMessage="Use a local driver path instead of downloading an MSI (bool true false, default is false)"
        )]
        [bool]$UseLocalDriverPath = $False,

    [Parameter(
        Position=18,
        Mandatory=$False,
        HelpMessage="Path to an extracted driver folder - required if you set UseLocalDriverPath variable to true or script will not find any drivers to inject"
        )]
        [string]$LocalDriverPath
    )



Function Receive-Output
{
    Param(
        $Color
    )

    Process { Write-Host $_ -ForegroundColor $Color }
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
        Write-Warning “You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator to continue.”
        Break
    }
}



Function Check-Internet
{
    While (([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]‘{DCB00C01-570F-4A9B-8D69-199FDBA5723B}’)).IsConnectedToInternet) -eq $False)
    {
        Write-Output "No internet connection detected. Retrying in 60 seconds..." | Receive-Output -Color Yellow
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
        Write-Output "aka.ms link: $URL" | Receive-Output -Color Gray
        Write-Output "Actual URL:  $ActualURL" | Receive-Output -Color Gray
        Write-Output "File name:   $FileName" | Receive-Output -Color White
        Write-Output ""
    }
    Else
    {
        $ActualURL = $URL
        $FileName = $URL.AbsoluteUri.Substring($URL.AbsoluteUri.LastIndexOf("/") +1)
        Write-Output "Actual URL:  $URL" | Receive-Output -Color Gray
        Write-Output "File name:   $FileName" | Receive-Output -Color White
        Write-Output ""
    }

    $global:Output = "$Path\$Filename"

    # If file does not exist, download file
    If (!(Test-Path -Path "$global:Output"))
    {
        Write-Output "Using BITS to download files" | Receive-Output -Color White
        Write-Output "Downloading $FileName to $Path..." | Receive-Output -Color White
        Write-Output ""
        Import-Module BitsTransfer
        Start-BitsTransfer -Source $ActualURL -Destination "$global:Output" -Priority Foreground -RetryTimeout 60 -RetryInterval 120
    }
    Else
    {
        Write-Output "File $global:Output exists, skipping file download." | Receive-Output -Color Gray
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


    $OSArch = Get-WmiObject -Class Win32_OperatingSystem

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
    # Check for admin rights
    CheckIfRunAsAdmin

    # Windows Version Check
    $OSCaption = (Get-WmiObject win32_operatingsystem).caption
    If ($OSCaption -like "Microsoft Windows 10*" -or $OSCaption -like "Microsoft Windows Server 2016*" -or $OSCaption -like "Microsoft Windows Server 2019*")
    {
        # All OK
    }
    Else
    {
        Write-Warning "$Env:Computername You must use Windows 10 or Windows Server 2016/2019 when servicing Windows 10 offline, with the latest ADK installed."
        Write-Warning "$Env:Computername Aborting script..."
        Exit
    }
    
    
    # Validating that the ADK is installed
    If (!(Test-Path $DISMFile))
    {
        Write-Warning "DISM in Windows ADK not found, attempting installation..." | Receive-Output -Color Yellow
        Write-Output ""
        $global:Output = $null
        $global:IsInstalled = $null

        $ScriptFolder = $DestinationFolder
        $ADKSourceFile = "$ScriptFolder\adksetup.exe"
        $WinPESourceFile = "$ScriptFolder\adkwinpesetup.exe"
        $ADKArguments = " /features OptionId.DeploymentTools /quiet"
        $WinPEArguments = " /features OptionId.WindowsPreinstallationEnvironment /quiet"

        GetInstalledAppStatus -AppName "Windows Assessment and Deployment Kit - Windows 10" -AppVersion "10.1.18362"

        If ($global:IsInstalled -eq $null)
        {
            # ADK cannot do an "in place" upgrade.  Do we need to uninstall the old version?
            $uninstall32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Assessment and Deployment*" } | select UninstallString
            $uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach { gp $_.PSPath } | ? { $_ -like "*Assessment and Deployment*" } | select UninstallString

            If ($uninstall64) 
            {
                ForEach ($u in $uninstall64)
                {
                    $u = $u.UninstallString -Replace "/uninstall","" 
                    $u = $u.Trim()
                    Write-Output "Command is $u Args are /uninstall /quiet" | Receive-Output -Color Gray
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ($uninstall32)
            {
                ForEach ($u in $uninstall32)
                {
                    $u = $u.UninstallString -Replace "/uninstall",""
                    $u = $u.Trim()
                    Write-Output "Command is $u Args are /uninstall /quiet" | Receive-Output -Color Gray
                    Start-Process -filepath $u -argumentlist "/uninstall /quiet" -wait
                }
            }

            If ((Test-Path -Path $ADKSourceFile) -eq $true)
            {
                $SourceFilePath = $(Get-Item $SourceFile).FullName
                Write-Output "Found Installation files for ADK at $SourceFilePath" | Receive-Output -Color Gray
            }	
            Else
            {
                Check-Internet
                $URL = "https://aka.ms/sdaadk/1903"
                $Path = "$env:TEMP"
                DownloadFile $URL $Path
                $SourceFilePath = $global:Output
            }

            Try
            {
                Write-Output "Installing Windows Assessment and Deployment Kit" | Receive-Output -Color White
                Start-Process -File  $SourceFilePath -Arg $ADKArguments -passthru | wait-process

                Write-Output  "$AppName - ADK INSTALLATION SUCCESSFULLY COMPLETED" | Receive-Output -Color Green
                Write-Output  ""

            }
            Catch
            {
                Write-Output  "$AppName - INSTALLATION ERROR - check logs in $env:TEMP\adk for more info." | Receive-Output -Color Yellow
                Write-Output  ""
            }
        }

        If ((Test-Path -Path $WinPESourceFile) -eq $true)
        {
            $SourceFilePath = $(Get-Item $SourceFile).FullName
            Write-Output "Found Installation files for ADK WinPE at $SourceFilePath" | Receive-Output -Color Gray
        }	
        Else
        {
            Check-Internet
            $URL = "https://aka.ms/sdaadkpe/1903"
            $Path = "$env:TEMP"
            DownloadFile $URL $Path
            $SourceFilePath = $global:Output
        }

        Try
        {
            Write-Output "Installing Windows Assessment and Deployment Kit Windows Preinstallation Environment Add-Ons" | Receive-Output -Color White
            Start-Process -File  $SourceFilePath -Arg $WinPEArguments -passthru | wait-process

            Write-Output  "$AppName - ADK WinPE Add-Ons INSTALLATION SUCCESSFULLY COMPLETED" | Receive-Output -Color Green
            Write-Output  ""
        }

        Catch
        {
            Write-Output  "$AppName - INSTALLATION ERROR - check logs in $env:TEMP\adkwinpeaddons for more info." | Receive-Output -Color Yellow
            Write-Output  ""
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
            $DownloadLinks = ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
            If ($DownloadLinks)
            {
                $updatesFound = $true
                If ($DownloadLinks.Count -gt 1)
                {
                    ForEach ($URL in $DownloadLinks)
                    {
                        Write-Output "Download found:" | Receive-Output -Color Green
                        Write-Output $curTxt | Receive-Output -Color White
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
                    Write-Output "Download found:" | Receive-Output -Color Green
                    Write-Output $curTxt | Receive-Output -Color White
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
        
        if(!($updatesFound))
        {
            $global:KBGUID = $null
            Write-Output "No update found." | Receive-Output -Color Yellow
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
        Write-Output "Attempting to find and download Servicing Stack updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray
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
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow

                    $Date = $NewDate
                    $ServicingURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + " Servicing Stack " + $Architecture + " windows 10 " + $OSBuild

                    $uri = $ServicingURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $True -Cumulative $False -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow
                    Break
                }
            }
        }
        $LoopBreak = $null
        $Date = Get-Date -Format "yyyy-MM"
    }
    If ($Cumulative)
    {
        Write-Output "Attempting to find and download Cumulative Update updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray
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
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow

                    $Date = $NewDate
                    $CumulativeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for Windows 10" ' + $Architecture + " " + $OSBuild

                    $uri = $CumulativeURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $True -CumulativeDotNet $False -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow
                    Break
                }
            }
        }
        $Date = Get-Date -Format "yyyy-MM"
        $LoopBreak = $null
    }
    If ($CumulativeDotNet)
    {
        Write-Output "Attempting to find and download Cumulative .NET Framework Update updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray
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
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow

                    $Date = $NewDate
                    $CumulativeDotNetURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "cumulative update for .NET Framework" ' + $Architecture + " windows 10 " + $OSBuild

                    $uri = $CumulativeDotNetURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $True -Adobe $False -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak months of searches.  Continuing..." | Receive-Output -Color Yellow
                    Break
                }
            }
        }
        $Date = Get-Date -Format "yyyy-MM"
        $LoopBreak = $null
    }
    If ($Adobe)
    {
        Write-Output "Attempting to find and download Adobe Flash Player updates for $Architecture Windows 10 version $OSBuild for month $Date..." | Receive-Output -Color Gray
        $uri = $AdobeURI
        Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $False -Adobe $True -OSBuild $OSBuild
        If (!($global:KBGUID))
        {
            While (!($global:KBGUID))
            {
                If ($LoopBreak -le 10)
                {
                    $LoopBreak++
                    Start-Sleep 1
                    $NewDate = (Get-Date).AddMonths(-$LoopBreak)
                    $NewDate = $NewDate.ToString("yyyy-MM")
                    Write-Output "No update found for month ($Date) - attempting previous month ($NewDate)..." | Receive-Output -Color Yellow

                    $Date = $NewDate
                    $AdobeURI = "http://www.catalog.update.microsoft.com/Search.aspx?q=" + $Date + ' "Security Update for Adobe Flash Player for Windows 10" ' + $Architecture + " " + $OSBuild

                    $uri = $AdobeURI
                    Download-LatestUpdates -uri $uri -Path $Path -Date $Date -Servicing $False -Cumulative $False -CumulativeDotNet $False -Adobe $True -OSBuild $OSBuild
                }
                Else
                {
                    Write-Output "Unable to find update for past $LoopBreak month's of searches.  Continuing..." | Receive-Output -Color Yellow
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
        Write-Output "Deleting $Path\Extract\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$Path\Extract\" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$Path\Extract" -Force
    }
    If (!(Test-Path "$Path\Extract"))
    {
        New-Item -Path "$Path\Extract" -ItemType "directory" | Out-Null
    }

    Write-Output "Extracting file $MsiFile to $Path\Extract..." | Receive-Output -Color White
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

    If ($Device -eq "SurfaceHub2S")
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
                        Write-Output "Download found:" | Receive-Output -Color Green
                        Write-Output $curTxt | Receive-Output -Color White
                        Write-Output ""
                        Write-Output ""
                        DownloadFile -URL $URL -Path "$DeviceDriverPath"
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                        Write-Output ""
                    }
                }
                Else
                {
                    Write-Output "Download found:" | Receive-Output -Color Green
                    Write-Output $curTxt | Receive-Output -Color White
                    Write-Output ""
                    Write-Output ""
                    DownloadFile -URL $DownloadLinks -Path "$DeviceDriverPath"
                    Write-Output ""
                    Write-Output ""
                    Write-Output ""
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
        Write-Output "Deleting $DeviceDriverPath\..." | Receive-Output -Color Gray
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
            Write-Output "$LocalDriverPath not found, continuing without drivers..." | Receive-Output -Color Yellow
            $Device = $null
        }
        Else
        {
            # Use local drivers
            Write-Output "Using $LocalDriverPath..." | Receive-Output -Color White
            $TempDeviceDriverPath = "$DeviceDriverPath\Extract"
            If (Test-Path "$TempDeviceDriverPath")
            {
                Write-Output "Deleting $TempDeviceDriverPath\..." | Receive-Output -Color Gray
                Get-ChildItem -Path "$TempDeviceDriverPath" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$TempDeviceDriverPath" -Force
            }
            If (!(Test-Path "$TempDeviceDriverPath"))
            {
                New-Item -path "$TempDeviceDriverPath" -ItemType "directory" | Out-Null
            }

            Write-Output "Copying drivers from $LocalDriverPath to $TempDeviceDriverPath..." | Receive-Output -Color White
            & xcopy.exe /herky "$LocalDriverPath" "$TempDeviceDriverPath"
            Write-Output ""
        }
    }
    Else
    {
        Write-Output "Downloading latest drivers for $Device, Windows 10 version $global:OSVersion..." | Receive-Output -Color White
        $OSBuild = New-Object string (,@($global:OSVersion.ToCharArray() | Select-Object -Last 5))
        $URL = "https://aka.ms/" + $Device + "/" + $OSBuild

        $DownloadedFile = DownloadFile -URL $URL -Path "$DeviceDriverPath"
        Write-Output "Downloaded File: $DownloadedFile"

        $FileToExtract = $DownloadedFile
        ExtractMSIFile -MsiFile $FileToExtract -Path $DeviceDriverPath
        Write-Output ""
    }

    Write-Output "Downloading latest Surface Ethernet drivers for $Device..." | Receive-Output -Color White
    Get-LatestSurfaceEthernetDrivers -Device $Device -TempFolder $TempFolder
    Write-Output ""
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
        Write-Output "Deleting $VisualCRuntimePath\..." | Receive-Output -Color Gray
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
    If (!(Test-Path "$VisualCRuntimePath\2015"))
    {
        New-Item -path "$VisualCRuntimePath\2015" -ItemType "directory" | Out-Null
    }
    If (!(Test-Path "$VisualCRuntimePath\2017"))
    {
        New-Item -path "$VisualCRuntimePath\2017" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest VisualC++ Runtimes..." | Receive-Output -Color White

    $VC2013x86URL = "https://aka.ms/vcpp2013x86"
    $VC2013x64URL = "https://aka.ms/vcpp2013x64"
    $VC2015x86URL = "https://aka.ms/vcpp2015x86"
    $VC2015x64URL = "https://aka.ms/vcpp2015x64"
    $VC2017X86URL = "https://aka.ms/vcpp2017x86"
    $VC2017X64URL = "https://aka.ms/vcpp2017x64"


    # 2013
    $VC2013x86 = DownloadFile -URL $VC2013x86URL -Path "$VisualCRuntimePath\2013"
    Write-Output "Downloaded File: $VC2013x86"
    Write-Output ""
    $VC2013x64 = DownloadFile -URL $VC2013x64URL -Path "$VisualCRuntimePath\2013"
    Write-Output "Downloaded File: $VC2013x64"
    Write-Output ""

    # 2015
    $VC2015x86 = DownloadFile -URL $VC2015x86URL -Path "$VisualCRuntimePath\2015"
    Write-Output "Downloaded File: $VC2015x86"
    Write-Output ""
    $VC2015x64 = DownloadFile -URL $VC2015x64URL -Path "$VisualCRuntimePath\2015"
    Write-Output "Downloaded File: $VC2015x64"
    Write-Output ""
    
    # 2017
    $VC2017x86 = DownloadFile -URL $VC2017x86URL -Path "$VisualCRuntimePath\2017"
    Write-Output "Downloaded File: $VC2017x86"
    Write-Output ""
    $VC2017x64 = DownloadFile -URL $VC2017x64URL -Path "$VisualCRuntimePath\2017"
    Write-Output "Downloaded File: $VC2017x64"
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
        Write-Output "Deleting $adobeUpdatePath\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$adobeUpdatePath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$adobeUpdatePath" -Force
    }
    If (!(Test-Path "$adobeUpdatePath"))
    {
        New-Item -path "$adobeUpdatePath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Adobe Flash update for $global:OSVersion..." | Receive-Output -Color White
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
        Write-Output "Deleting $CumulativeUpdatePath\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$CumulativeUpdatePath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$CumulativeUpdatePath" -Force
    }
    If (!(Test-Path "$CumulativeUpdatePath"))
    {
        New-Item -path "$CumulativeUpdatePath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Cumulative Update for $global:OSVersion..." | Receive-Output -Color White
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
        Write-Output "Deleting $ServicingStackPath\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$ServicingStackPath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$ServicingStackPath" -Force
    }
    If (!(Test-Path "$ServicingStackPath"))
    {
        New-Item -Path "$ServicingStackPath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Servicing Stack update for $global:OSVersion..." | Receive-Output -Color White
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
        Write-Output "Deleting $CumulativeDotNetPath\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$CumulativeDotNetPath" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$CumulativeDotNetPath" -Force
    }
    If (!(Test-Path "$CumulativeDotNetPath"))
    {
        New-Item -Path "$CumulativeDotNetPath" -ItemType "directory" | Out-Null
    }

    Write-Output "Downloading latest Dot Net Cumulative updates for $global:OSVersion..." | Receive-Output -Color White
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
        Write-Output "Deleting $ScratchMountFolder\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$ScratchMountFolder" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$ScratchMountFolder" -Force
    }
    If (!(Test-Path -path $ScratchMountFolder))
    {
        New-Item -path $ScratchMountFolder -ItemType Directory | Out-Null
    }

    Write-Output "Mounting ISO $ISO..." | Receive-Output -Color White
    $ISOPath = (Mount-DiskImage -ImagePath $ISO -StorageType ISO -PassThru | Get-Volume).DriveLetter
    $Drive = $ISOPath + ":"

    If ($ISOPath)
    {
        Write-Output "ISO successfully mounted at $Drive" | Receive-Output -Color White
        Write-Output ""   
    }
    Else
    {
        Write-Output "Failed to mount the ISO. Please verify the ISO path and try again" | Receive-Output -Color Red
        Exit
    }

    Write-Output "Parsing install.wim file(s) in $Drive for images..." | Receive-Output -Color White
    $WIMs = Get-ChildItem -Path "$Drive" -Filter install.wim -Recurse
    $OSWIMFound = $False

    # Required to get ReleaseId value, which is needed for 1909
    ForEach ($WIM in $WIMs)
    {
        $TempWIM = $WIM.FullName
        $OSWIM = Get-WindowsImage -ImagePath $TempWIM | Where-Object {($_.ImageName -like "*$($OSSKU)") -or ($_.ImageName -like "*$($OSSKU) Evaluation") -or ($_.ImageName -like "*$OSSKU) LTSC")}
        If (!($OSWIM))
        {
            # $OSSKU not found
        }
        Else
        {
            $ImagePath = $OSWIM.ImagePath
            $ImageIndex = $OSWIM.ImageIndex
            $ImageName = $OSWIM.ImageName

            Write-Output "Found image matching $OSSKU :" | Receive-Output -Color Gray
            Write-Output "Image Path:  $ImagePath" | Receive-Output -Color White
            Write-Output "Image Index: $ImageIndex" | Receive-Output -Color White
            Write-Output "Image Name:  $ImageName" | Receive-Output -Color White

            If (($ImageName -like "*$($OSSKU)") -or ($ImageName -like "*$($OSSKU) Evaluation") -or ($ImageName -like "*$OSSKU) LTSC"))
            {
                $global:OSVersion = (Get-WindowsImage -ImagePath "$ImagePath" -Index "$ImageIndex").Version
                $OSWIMFound = $True
            }
            If ($global:OSVersion)
            {
                $global:OSVersion = $global:OSVersion.Substring(0, $global:OSVersion.LastIndexOf('.'))
                Write-Output "Mounting $ImagePath in $ScratchMountFolder..." | Receive-Output -Color White
                Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $ScratchMountFolder -ReadOnly | Out-Null
                Start-Sleep 5
                Write-Output "Querying image registry for ReleaseId..." | Receive-Output -Color White
                & reg.exe load "HKLM\Mount" "$ScratchMountFolder\Windows\system32\config\SOFTWARE"
                $Key = "HKLM:\Mount\Microsoft\Windows NT\CurrentVersion"
                $global:ReleaseId = (Get-ItemProperty -Path $Key -Name ReleaseId).ReleaseId
                Start-Sleep 5
                Write-Output "Unloading image registry..." | Receive-Output -Color White
                & reg.exe unload "HKLM\Mount"
                Start-Sleep 5
                Write-Output "Dismounting $ScratchMountFolder..."
                Dismount-WindowsImage -Path $ScratchMountFolder  -Discard | Out-Null
                Write-Output ""
            }
            # Specific 1909 check as it will report as 10.0.18362 still when offline
            If ($global:ReleaseId -eq "1909")
            {
                $global:OSVersion = "10.0.18363"
            }
        }
    }

    If ($OSWIMFound -eq $False)
    {
        Dismount-DiskImage -ImagePath $ISO | Out-Null
        Write-Output "$OSSKU not found in $WIMs on $ISO.  Please make sure to use an ISO file that contains $OSSKU, and try again." | Receive-Output -Color Red
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
        Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp..." | Receive-Output -Color Gray
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
        $ExistingInstallWIM = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim"
        $TempExistingInstallWIM = Get-WindowsImage -ImagePath $ExistingInstallWIM | Where-Object {$_.ImageName -like "*$($OSSKU)"}
        $TempExistingInstallWIMPath = $TempExistingInstallWIM.ImagePath
        $TempExistingInstallWIMIndex = $TempExistingInstallWIM.ImageIndex
        $TempExistingInstallWIMName = $TempExistingInstallWIM.ImageName

        If ($TempExistingInstallWIMName -like "*$($OSSKU)")
        {
            $TempExistingInstallWIMOSVersion = (& $DISMFile /Get-WimInfo /WimFile:$TempExistingInstallWIMPath /index:$TempExistingInstallWIMIndex | Select-String "Version ").ToString().Split(":")[1].Trim()
            If ($TempExistingInstallWIMOSVersion -eq $global:OSVersion)
            {
                $LeaveInstallWIM = $True
                Write-Output "Leaving existing install.wim in $ExistingInstallWIM" | Receive-Output -Color Gray
            }
        }
        Else
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim..." | Receive-Output -Color Gray
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim" -Force
            Start-Sleep 5
        }
    }
    
    If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim")
    {
        Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim..." | Receive-Output -Color Gray
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

    Write-Output "Copying $WindowsKitsInstall\Windows Preinstallation Environment\$Arch\en-us\winpe.wim to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim..." | Receive-Output -Color White
    Copy-Item -Path "$WindowsKitsInstall\Windows Preinstallation Environment\$Arch\en-us\winpe.wim" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\boot.wim"
    $SourceBootWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -filter boot.wim -Recurse
    ForEach ($SourceBootWIM in $SourceBootWIMs)
    {
        $TempBootWIM = $SourceBootWIM.FullName

        $PEWIM = Get-WindowsImage -ImagePath $TempBootWIM | Where-Object {$_.ImageName -like "*Windows PE*"}

        $ImagePath = $PEWIM.ImagePath
        $ImageIndex = $PEWIM.ImageIndex
        $ImageName = $PEWIM.ImageName
    }

    If ($DotNet35 -eq $true)
    {
        If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs")
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs..." | Receive-Output -Color Gray
            Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -Recurse | Remove-Item -Force -Recurse
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -Force
        }
        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs"))
        {
            New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -ItemType "directory" | Out-Null
        }
        Write-Output "Copying $Drive\Sources\sxs\* to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs\..." | Receive-Output -Color White
        Copy-Item -Path "$Drive\Sources\sxs\*" -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\sxs" -PassThru | Set-ItemProperty -Name IsReadOnly -Value $false
    }

    If (!($LeaveInstallWIM))
    {
        ForEach ($WIM in $WIMs)
        {
            $TempWIM = $WIM.FullName
            Write-Output "Copying $TempWIM to $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs\install.wim..." | Receive-Output -Color White
            Copy-Item -Path $TempWIM -Destination "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs" -PassThru | Set-ItemProperty -Name IsReadOnly -Value $false
        }
    }

    Dismount-DiskImage -ImagePath $ISO | Out-Null

    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
}



Function Add-PackageIntoWindowsImage
{
    Param(
        [string]$ImageMountFolder,
        [string]$PackagePath,
        [string]$TempImagePath
    )

    Add-WindowsPackage -Path $ImageMountFolder -PackagePath $PackagePath
    Write-Output ""
    Write-Output ""

    # Dismount the image to avoid PSFX/non-PSFX update compression issues in RS5+
    Write-Output "Saving $TempImagePath..." | Receive-Output -Color White
    DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
    Write-Output ""
    Write-Output ""
    Start-Sleep 2

    # Re-mount the image
    Write-Output "Mounting $TempImagePath in $ImageMountFolder..." | Receive-Output -Color White
    Mount-WindowsImage -ImagePath $TempImagePath -Index 1 -Path $ImageMountFolder -CheckIntegrity
    Write-Output ""
    Write-Output ""
}



Function Update-Win10WIM
{
    Param(
        [string]$SourcePath,
        [string]$SourceName,
        [bool]$ServicingStack,
        [bool]$CumulativeUpdate,
        [bool]$DotNet35,
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
    

    $SourceName = Switch ($SourceName)
    {
        Pro {"Windows 10 Pro"}
        Enterprise {"Windows 10 Enterprise"}
    }

    If (Test-Path "$ImageMountFolder")
    {
        Write-Output "Deleting $ImageMountFolder\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$ImageMountFolder" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$ImageMountFolder" -Force
    }
    If (!(Test-Path -path $ImageMountFolder))
    {
        New-Item -path $ImageMountFolder -ItemType Directory | Out-Null
    }

    If (Test-Path "$BootImageMountFolder")
    {
        Write-Output "Deleting $BootImageMountFolder\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$BootImageMountFolder" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$BootImageMountFolder" -Force
    }
    If (!(Test-Path -path $BootImageMountFolder))
    {
        New-Item -path $BootImageMountFolder -ItemType Directory | Out-Null
    }

    If (Test-Path "$WinREImageMountFolder")
    {
        Write-Output "Deleting $WinREImageMountFolder\..." | Receive-Output -Color Gray
        Get-ChildItem -Path "$WinREImageMountFolder" -Recurse | Remove-Item -Force -Recurse
        Remove-Item -Path "$WinREImageMountFolder" -Force
    }
    If (!(Test-Path -path $WinREImageMountFolder))
    {
        New-Item -path $WinREImageMountFolder -ItemType Directory | Out-Null
    }


    # Variables
    $Now = Get-Date -Format yyyy-MM-dd_HH-mm-ss
    $TmpImage = "$TempFolder\tmp_install.wim"
    $TmpWinREImage = "$TempFolder\tmp_winre.wim"
    $TmpBootImage = "$TempFolder\tmp_boot.wim"
    $ServicingStackPath = "$TempFolder\Servicing"
    $CumulativeUpdatePath = "$TempFolder\Cumulative"
    $DotNetPath = "$TempFolder\DotNet"
    $AdobeFlashUpdatePath = "$TempFolder\Adobe"
    $DeviceDriverPath = "$TempFolder\$Device"
    $VC2013x86Path = "$TempFolder\VCRuntimes\2013\vcredist_x86.exe"
    $VC2013x64Path = "$TempFolder\VCRuntimes\2013\vcredist_x64.exe"
    $VC2015x86Path = "$TempFolder\VCRuntimes\2015\vc_redist.x86.exe"
    $VC2015x64Path = "$TempFolder\VCRuntimes\2015\vc_redist.x64.exe"
    $VC2017x86Path = "$TempFolder\VCRuntimes\2017\vc_redist.x86.exe"
    $VC2017x64Path = "$TempFolder\VCRuntimes\2017\vc_redist.x64.exe"
    $ProUnattendXMLPath = "$WorkingDirPath\Win10Pro_Unattend.xml"
    $EntUnattendXMLPath = "$WorkingDirPath\Win10Ent_Unattend.xml"



    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    
    Write-Output ""
    Write-Output ""
    Write-Output " *********************************************" | Receive-Output -Color Cyan
    Write-Output " *                                           *" | Receive-Output -Color Cyan
    Write-Output " *           Updating install.wim            *" | Receive-Output -Color Cyan
    Write-Output " *                                           *" | Receive-Output -Color Cyan
    Write-Output " *********************************************" | Receive-Output -Color Cyan
    Write-Output ""
    Write-Output ""
    Start-Sleep 2

    If ($InstallWIM)
    {
        # Export the reference image to a new (temporary) WIM - this will leave the original "install.wim" untouched when finished
        Write-Output "Exporting $SourcePath\install.wim to $TmpImage..." | Receive-Output -Color White
        Export-WindowsImage -SourceImagePath "$SourcePath\install.wim" -SourceName "$SourceName" -DestinationImagePath $TmpImage -CheckIntegrity
        Write-Output ""
        Write-Output ""

        # Mount the image
        Write-Output "Mounting $TmpImage in $ImageMountFolder..." | Receive-Output -Color White
        Mount-WindowsImage -ImagePath $TmpImage -Index 1 -Path $ImageMountFolder -CheckIntegrity
        Write-Output ""
        Write-Output ""

        If ($DotNet35 -eq $True)
        {
            # Cleanup the image BEFORE installing .NET to prevent errors
            Write-Output "Running image cleanup on $ImageMountFolder..." | Receive-Output -Color White
            & $DISMFile /Image:$ImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase
            Write-Output ""
            Write-Output ""

            # Dismount the image
            Write-Output "Saving $TmpImage..." | Receive-Output -Color White
            DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
            Write-Output ""
            Write-Output ""
            Start-Sleep 2

            # Re-mount the image
            Write-Output "Mounting $TmpImage in $ImageMountFolder..." | Receive-Output -Color White
            Mount-WindowsImage -ImagePath $TmpImage -Index 1 -Path $ImageMountFolder -CheckIntegrity
            Write-Output ""
            Write-Output ""

            # Add .NET Framework 3.5 to the image
            Write-Output "Adding .NET Framework 3.5 to $ImageMountFolder..." | Receive-Output -Color White
            Enable-WindowsOptionalFeature -Path $ImageMountFolder -FeatureName NetFx3 -All -Source "$TempFolder\sxs" -LimitAccess
            Write-Output ""
            Write-Output ""
        }

        If ($ServicingStack)
        {
            $SSU = Get-ChildItem -Path $ServicingStackPath
            If (!($SSU.Exists))
            {
                $ServicingStack = $False
            }
            Else
            {
                # Add required Servicing Stack updates
                Write-Output "Adding Servicing Stack updates to $ImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpImage
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
                Write-Output "Adding Cumulative updates to $ImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpImage
            }
        }

        If ($DotNet35)
        {
            $DNU = Get-ChildItem -Path $DotNetPath
            If (!($DNU.Exists))
            {
                $DotNet35 = $False
            }
            Else
            {
                # Add .NET Framework updates
                Write-Output "Adding .NET Framework updates to $ImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $DotNetPath -TempImagePath $TmpImage
            }
        }
        
        if ($AdobeFlashUpdate)
        {
            $AFU = Get-ChildItem -Path $AdobeFlashUpdatePath
            if (!($AFU.Exists))
            {
                $AdobeFlashUpdate = $False
            }
            Else
            {
                # Add Adobe Flash updates
                Write-Output "Adding Adobe Flash updates to $ImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $ImageMountFolder -PackagePath $AdobeFlashUpdatePath -TempImagePath $TmpImage
            }
        }

        If ($Device)
        {
            $MSITempPath = "$DeviceDriverPath\Extract"
            $MSIFiles = Get-ChildItem -Path $MSITempPath -Recurse
            # Add drivers/firmware to WIM
            Write-Output "Adding Driver updates for $Device to $ImageMountFolder from $MSITempPath..." | Receive-Output -Color White
            Add-WindowsDriver -Path $ImageMountFolder -Driver "$MSITempPath" -Recurse
            Write-Output ""
            Write-Output ""

            # Copy VC++ Runtimes
            If (Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013")
            {
                Write-Output "Deleting $ImageMountFolder\Windows\Temp\VCRuntimes\2013..." | Receive-Output -Color Gray
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2013" -ItemType Directory | Out-Null
            }

            If (Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2015")
            {
                Write-Output "Deleting $ImageMountFolder\Windows\Temp\VCRuntimes\2015..." | Receive-Output -Color Gray
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2015" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2015" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2015"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2015" -ItemType Directory | Out-Null
            }

            If (Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2017")
            {
                Write-Output "Deleting $ImageMountFolder\Windows\Temp\VCRuntimes\2017..." | Receive-Output -Color Gray
                Get-ChildItem -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2017" -Recurse | Remove-Item -Force -Recurse
                Remove-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2017" -Force
            }
            If (!(Test-Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2017"))
            {
                New-Item -Path "$ImageMountFolder\Windows\Temp\VCRuntimes\2017" -ItemType Directory | Out-Null
            }

            If (!($Architecture -eq "ARM64"))
            {
                Write-Output "Copying VC++ Runtime binaries to $ImageMountFolder\Windows\Temp..."
                Copy-Item -Path $VC2013x86Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"
                Copy-Item -Path $VC2013x64Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2013"
                Copy-Item -Path $VC2015x86Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2015"
                Copy-Item -Path $VC2015x64Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2015"
                Copy-Item -Path $VC2017x86Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2017"
                Copy-Item -Path $VC2017x64Path -Destination "$ImageMountFolder\Windows\Temp\VCRuntimes\2017"
                Write-Output ""
            }
        }

        Write-Output "Copying unattend.xml to $ImageMountFolder\Windows\System32\sysprep..."
        If ($OSSKU -like "*Pro*")
        {
            Copy-Item -Path $ProUnattendXMLPath -Destination "$ImageMountFolder\Windows\System32\sysprep\unattend.xml"
        }
        If ($OSSKU -like "*Enterprise*")
        {
            Copy-Item -Path $EntUnattendXMLPath -Destination "$ImageMountFolder\Windows\System32\sysprep\unattend.xml"
        }
        Write-Output ""


        Write-Output ""
        Write-Output ""
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *           Updating winre.wim              *" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        # Copy WinRE Image to temp location
        Write-Output "Copying WinRE image to $TmpWinREImage..." | Receive-Output -Color White
        Move-Item -Path "$ImageMountFolder\Windows\System32\Recovery\winre.wim" -Destination $TmpWinREImage
        Write-Output ""
        Write-Output ""

        # Mount the temp WinRE Image
        Write-Output "Mounting $TmpWinREImage to $WinREImageMountFolder..." | Receive-Output -Color White
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
                Write-Output "Adding Servicing Stack updates to $WinREImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $WinREImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpWinREImage
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
                Write-Output "Adding Cumulative updates to $WinREImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $WinREImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpWinREImage
            }
        }
        
        If ($DotNet35)
        {
            $DNU = Get-ChildItem -Path $DotNetPath
            If (!($DNU.Exists))
            {
                $DotNet35 = $False
            }
            Else
            {
                # Add .NET Framework updates
                Write-Output "Adding .NET Framework updates to $WinREImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $WinREImageMountFolder -PackagePath $DotNetPath -TempImagePath $TmpWinREImage
            }
        }

        If ($Device)
        {
            $MSITempPath = "$DeviceDriverPath\Extract"
            $MSIFiles = Get-ChildItem -Path $MSITempPath -Recurse
            If ($SurfaceDevices.$Device)
            {
                # Add system-level drivers to WIM
                Write-Output "Adding Driver updates for $Device to $WinREImageMountFolder from $MSITempPath..." | Receive-Output -Color White
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
        Write-Output "Running image cleanup on $TmpWinREImage..." | Receive-Output -Color White
        & $DISMFile /Image:$WinREImageMountFolder /Cleanup-Image /StartComponentCleanup /ResetBase
        Write-Output ""
        Write-Output ""

        # Dismount the WinRE image
        Write-Output "Saving $TmpWinREImage..." | Receive-Output -Color White
        DisMount-WindowsImage -Path $WinREImageMountFolder -Save -CheckIntegrity
        Write-Output ""
        Write-Output ""


        Write-Output ""
        Write-Output ""
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *            Saving winre.wim               *" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        # Export the new WinRE image back to original location
        Write-Output "Exporting $TmpWinREImage to $ImageMountFolder\Windows\System32\Recovery\winre.wim..." | Receive-Output -Color White
        Export-WindowsImage -SourceImagePath $TmpWinREImage -SourceName "Microsoft Windows Recovery Environment (x64)" -DestinationImagePath "$ImageMountFolder\Windows\System32\Recovery\winre.wim" -CheckIntegrity
        Write-Output ""
        Write-Output ""


        Write-Output ""
        Write-Output ""
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *            Saving install.wim             *" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        # Validate Windows WIM build number
        $Build = (Get-Item $ImageMountFolder\Windows\System32\ntoskrnl.exe).VersionInfo.ProductVersion
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

        # Dismount the reference image
        Write-Output "Saving $TmpImage..." | Receive-Output -Color White
        DisMount-WindowsImage -Path $ImageMountFolder -Save -CheckIntegrity
        Write-Output ""
        Write-Output ""

        # Export the image to a new WIM
        Write-Output "Exporting $TmpImage to $RefImage..." | Receive-Output -Color White
        Export-WindowsImage -SourceImagePath $TmpImage -SourceName "$SourceName" -DestinationImagePath $RefImage -CheckIntegrity
        Write-Output ""
        Write-Output ""

        $TempRefImageSize = Get-Item $RefImage
        $RefImageSize = ($TempRefImageSize.Length /1GB)
        If ($RefImageSize -ge "4")
        {
            $SplitWIM = $true
            # Split the WIM to fit on FAT32-formatted media (splitting at ~3GB for simplicity)
            Write-Output "Splitting $RefImage into 3GB files as $SplitImage..." | Receive-Output -Color White
            Split-WindowsImage -ImagePath $RefImage -SplitImagePath $SplitImage -FileSize 3096 -CheckIntegrity
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

        Write-Output ""
        Write-Output ""
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *           Updating boot.wim               *" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        # Copy boot.wim for editing
        Write-Output "Copying $SourcePath\boot.wim to $TmpBootImage..." | Receive-Output -Color White
        Copy-Item "$SourcePath\boot.wim" $TempFolder
        Attrib -r "$TempFolder\boot.wim"
        Rename-Item -Path "$TempFolder\boot.wim" -NewName "$TmpBootImage"
        Write-Output ""
        Write-Output ""


        # Mount index 1 of the boot image (WinPE)
        Write-Output "Mounting $TmpBootImage to $BootImageMountFolder using Index 1..." | Receive-Output -Color White
        Mount-WindowsImage -ImagePath $TmpBootImage -Index 1 -Path $BootImageMountFolder -CheckIntegrity
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
                # Add required Servicing Stack updates
                Write-Output "Adding Servicing Stack updates to $BootImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $ServicingStackPath -TempImagePath $TmpBootImage
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
                Write-Output "Adding Cumulative updates to $BootImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $CumulativeUpdatePath -TempImagePath $TmpBootImage
            }
        }

        If ($DotNet35)
        {
            $DNU = Get-ChildItem -Path $DotNetPath
            If (!($DNU.Exists))
            {
                $DotNet35 = $False
            }
            Else
            {
                # Add .NET Framework updates
                Write-Output "Adding .NET Framework updates to $BootImageMountFolder..." | Receive-Output -Color White
                Add-PackageIntoWindowsImage -ImageMountFolder $BootImageMountFolder -PackagePath $DotNetPath -TempImagePath $TmpBootImage
            }
        }

        If ($Device)
        {
            $MSITempPath = "$DeviceDriverPath\Extract"
            $MSIFiles = Get-ChildItem -Path $MSITempPath -Recurse
            If ($SurfaceDevices.$Device)
            {
                # Add system-level drivers to WIM
                Write-Output "Adding Driver updates for $Device to $BootImageMountFolder from $MSITempPath..." | Receive-Output -Color White
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

        Write-Output "Adding WMI..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-WMI.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-WMI_en-us.cab" | Out-Null

        Write-Output "Adding PE Scripting..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-Scripting.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-Scripting_en-us.cab" | Out-Null

        Write-Output "Adding Enhanced Storage..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-EnhancedStorage.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-EnhancedStorage_en-us.cab" | Out-Null

        Write-Output "Adding Bitlocker support..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-SecureStartup.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-SecureStartup_en-us.cab" | Out-Null

        Write-Output "Adding .NET..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-NetFx.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-NetFx_en-us.cab" | Out-Null

        Write-Output "Adding PowerShell..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-PowerShell.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-PowerShell_en-us.cab" | Out-Null

        Write-Output "Adding Storage WMI..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-StorageWMI.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-StorageWMI_en-us.cab" | Out-Null

        Write-Output "Adding DISM support..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-DismCmdlets.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-DismCmdlets_en-us.cab" | Out-Null

        Write-Output "Adding Secure Boot support..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-SecureBootCmdlets.cab" | Out-Null

        Write-Output "Adding Secure Startup support..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-DismCmdlets.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-DismCmdlets_en-us.cab" | Out-Null

        Write-Output "Adding WinRE support..." | Receive-Output -Color White
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\WinPE-WinReCfg.cab" | Out-Null
        Add-WindowsPackage -Path $BootImageMountFolder -PackagePath "$WinPEOCPath\en-us\WinPE-WinReCfg_en-us.cab" | Out-Null


        If (($MakeUSBMedia) -or ($MakeISOMedia))
        {
            Write-Host "Copying scripts to $BootImageMountFolder..."
            Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI.txt" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\CreatePartitions-UEFI_Source.txt" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\Imaging.ps1" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\Install.cmd" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\surface_devices.xml" -Destination $BootImageMountFolder
            Copy-Item -Path "$WorkingDirPath\UsbImage\startnet.cmd" -Destination "$BootImageMountFolder\Windows\System32" -Force
        }

        Write-Output ""
        Write-Output ""


        Write-Output ""
        Write-Output ""
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *            Saving boot.wim                *" | Receive-Output -Color Cyan
        Write-Output " *                                           *" | Receive-Output -Color Cyan
        Write-Output " *********************************************" | Receive-Output -Color Cyan
        Write-Output ""
        Write-Output ""
        Start-Sleep 2


        # Variable
        $Build = (Get-Item $BootImageMountFolder\Windows\System32\ntoskrnl.exe).VersionInfo.ProductVersion
        If ($Device)
        {
            $RefBootImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Boot-$Build-$Now.wim"
        }
        Else
        {
            $RefBootImage = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Generic-Boot-$Build-$Now.wim"
        }


        # Dismount the boot image
        Write-Output "Saving $TmpBootImage..." | Receive-Output -Color White
        DisMount-WindowsImage -Path $BootImageMountFolder -Save -CheckIntegrity
        Write-Output ""
        Write-Output ""

        # Export the image to a new WIM
        Write-Output "Exporting $TmpBootImage to $RefBootImage..." | Receive-Output -Color White
        Export-WindowsImage -SourceImagePath $TmpBootImage -SourceIndex 1 -DestinationImagePath $RefBootImage -CheckIntegrity
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
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media\..." | Receive-Output -Color Gray
            Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -Recurse | Remove-Item -Force -Recurse
            Remove-Item -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -Force
        }
        If (!(Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"))
        {
            New-Item -path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" -ItemType "directory" | Out-Null
        }

        If (Test-Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles")
        {
            Write-Output "Deleting $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\..." | Receive-Output -Color Gray
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

        Write-Output "Creating WinPE media in $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media..." | Receive-Output -Color White
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
            Write-Output "Insert USB drive 16GB+ in size, and press ENTER" | Receive-Output -Color Yellow
            Write-Output "      !!!THIS WILL FORMAT THE DRIVE!!!" | Receive-Output -Color Yellow
            Write-Output ""
            PAUSE
            Start-Sleep 5

            $TempUSB = (Get-PhysicalDisk | Where-Object {$_.BusType -eq "USB" -and $_.MediaType -ne "SSD"}).FriendlyName

            If (!($TempUSB))
            {
                Write-Warning "No USB key found, skipping..."
            }
            Else
            {
                Write-Output "Getting USB drive ready..." | Receive-Output -Color White
                $TempUSB = (Get-PhysicalDisk | Where-Object {$_.BusType -eq "USB" -and $_.MediaType -ne "SSD"}).FriendlyName
                $USB = Get-Disk | Where-Object {$_.FriendlyName -like $TempUSB}
                $USBSize = $USB.Size /1GB

                Get-Disk -FriendlyName $TempUSB | Clear-Disk -RemoveData -Confirm:$false
                Initialize-Disk -FriendlyName $TempUSB -PartitionStyle MBR -ErrorAction SilentlyContinue

                If ($USBSize -ge "32")
                {
                    $NewUSBDriveLetter = New-Partition -DiskNumber $USB.DiskNumber -Size 32GB -AssignDriveLetter | Format-Volume -FileSystem FAT32 -NewFileSystemLabel $Device
                }
                ElseIf ($USBSize -lt "14")
                {
                    Write-Warning "USB drive not 16GB or larger, skipping..."
                }
                Else
                {
                    $NewUSBDriveLetter = New-Partition -DiskNumber $USB.DiskNumber -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem FAT32 -NewFileSystemLabel $Device
                }

                $NewUSBDriveLetter = $NewUSBDriveLetter.DriveLetter + ":"

                Write-Output "Copying WinPE Media contents to $NewUSBDriveLetter..." | Receive-Output -Color White
                & bootsect.exe /nt60 $NewUSBDriveLetter /force /mbr
                & xcopy /herky "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media" $NewUSBDriveLetter

                If ($SplitWIM -eq $True)
                {
                    $SplitWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture" -Filter *install*$Now*.swm -Recurse
                    ForEach ($TempWIM in $SplitWIMs)
                    {
                        $TempSplitWIM = $TempWIM.FullName
                        Write-Output "Copying $TempSplitWIM to $NewUSBDriveLetter..." | Receive-Output -Color White
                        Copy-Item -Path "$TempSplitWIM" -Destination "$NewUSBDriveLetter\Sources" -Force
                    }
                }
                Else
                {
                    Write-Output "Copying $RefImage to $NewUSBDriveLetter..." | Receive-Output -Color White
                    Copy-Item -Path "$RefImage" -Destination "$NewUSBDriveLetter\Sources" -Recurse
                }
            }
        }

        If ($MakeISOMedia)
        {
            $oscdimg = "$WindowsKitsInstall\Deployment Tools\$Arch\Oscdimg\oscdimg.exe"
            $efisys = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\efisys.bin"
            $etfsboot = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\fwfiles\etfsboot.com"
            $MediaSource = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp\Media"
            $args = "-l$Device -bootdata:2#p0,e,b$etfsboot#pEF,e,b$efisys -m -u1 -udfver102 $MediaSource $DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Boot-$Build-$Now.iso"
            
            If ($SplitWIM -eq $True)
            {
                $SplitWIMs = Get-ChildItem -Path "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture" -Filter *install*$Now*.swm -Recurse
                ForEach ($TempWIM in $SplitWIMs)
                {
                    $TempSplitWIM = $TempWIM.FullName
                    Write-Output "Copying $TempSplitWIM to $MediaSource..." | Receive-Output -Color White
                    Copy-Item -Path "$TempSplitWIM" -Destination "$MediaSource\Sources" -Force
                }
            }
            Else
            {
                Write-Output "Copying $RefImage to $MediaSource..." | Receive-Output -Color White
                Copy-Item -Path "$RefImage" -Destination "$MediaSource\Sources" -Recurse
            }

            Start-Process -FilePath $oscdimg -ArgumentList $args -NoNewWindow -Wait
        }
    }


    Write-Output ""
    Write-Output ""
    Write-Output " *********************************************" | Receive-Output -Color Cyan
    Write-Output " *                                           *" | Receive-Output -Color Cyan
    Write-Output " *       Image modifications complete!       *" | Receive-Output -Color Cyan
    Write-Output " *                                           *" | Receive-Output -Color Cyan
    Write-Output " *********************************************" | Receive-Output -Color Cyan
    Write-Output ""
    Write-Output ""
    Start-Sleep 2

    Write-Output "Finalized image files can be found here:" | Receive-Output -Color White
    Write-Output ""
    If ($CreateISO)
    {
        If (Test-Path("$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Boot-$Build-$Now.iso"))
        {
            Write-Output "ISO:      $$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\$Device-Boot-$Build-$Now.iso" | Receive-Output -Color Green
        }
    }
    If ($SplitWIM -eq $True)
    {
        Write-Output "Install:  $SplitImage" | Receive-Output -Color Green
    }
    Else
    {
        Write-Output "Install:  $RefImage" | Receive-Output -Color Green
    }
    Write-Output "Boot:     $RefBootImage" | Receive-Output -Color Green
}



###########################
# Begin script processing #
###########################
cls


# Get current working directory
$Invocation = (Get-Variable MyInvocation).Value
$WorkingDirPath = Split-Path $Invocation.MyCommand.Path
If (!($DestinationFolder))
{
    $DestinationFolder = $WorkingDirPath
}


If ($Device)
{
    # Read WinPEXML file
    [string]$XmlPath = "$WorkingDirPath\WinPE_Drivers.xml"
    [Xml]$WinPEXML = Get-Content $XmlPath
    [System.Xml.XmlElement] $root = $WinPEXML.get_DocumentElement()
    
    $SurfaceDevices = $WinPEXML.Surface.Devices
}


# Get script start time (will be used to determine how long execution takes)
$Script_Start_Time = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
Write-Output "Script start: $Script_Start_Time" | Receive-Output -Color Gray


# Necessary variables not passed into script directly
$DISMFile = "$WindowsKitsInstall\Deployment Tools\amd64\DISM\dism.exe"
$Mount = "$env:TEMP\Mount"
$ScratchMountFolder = "$Mount\Scratch"


# Leave blank space at top of window to not block output by progress bars
AddHeaderSpace


# Check for admin rights and ADK install
PrereqCheck


Write-Output ""
Write-Output ""
Write-Output " *********************************************" | Receive-Output -Color Cyan
Write-Output " *                                           *" | Receive-Output -Color Cyan
Write-Output " *       Parameters passed to script:        *" | Receive-Output -Color Cyan
Write-Output " *                                           *" | Receive-Output -Color Cyan
Write-Output " *********************************************" | Receive-Output -Color Cyan
Write-Output ""
Write-Output "ISO path:                     $ISO" | Receive-Output -Color White
Write-Output "OS SKU:                       $OSSKU" | Receive-Output -Color White
Write-Output "Output:                       $DestinationFolder" | Receive-Output -Color White
Write-Output "Architecture:                 $Architecture" | Receive-Output -Color White
Write-Output "  .NET 3.5:                   $DotNet35" | Receive-Output -Color White
Write-Output "  Servicing Stack:            $ServicingStack" | Receive-Output -Color White
Write-Output "  Cumulative Update:          $CumulativeUpdate" | Receive-Output -Color White
Write-Output "  Cumulative DotNet Updates:  $CumulativeUpdate" | Receive-Output -Color White
Write-Output "  Adobe Flash Player Updates: $AdobeFlashUpdate" | Receive-Output -Color White
Write-Output "  Device drivers:             $Device" | Receive-Output -Color White
If ($UseLocalDriverPath -eq $True)
{
    Write-Output "  Use Local driver path:      $LocalDriverPath" | Receive-Output -Color White
}
Write-Output "  Create USB key:             $CreateUSB" | Receive-Output -Color White
Write-Output "  Create ISO:                 $CreateISO" | Receive-Output -Color White
Write-Output ""
Write-Output ""
Start-Sleep 2


# Pull Windows 10 version and SKU from ISO provided by script param, returns OSVersion and WinPEVersion variable as well
Get-OSWIMFromISO -ISO $ISO -OSSKU $OSSKU -DestinationFolder $DestinationFolder -Architecture $Architecture -WindowsKitsInstall $WindowsKitsInstall -ScratchMountFolder $ScratchMountFolder
Start-Sleep 2
Write-Output "OSVersion:  " $global:OSVersion | Receive-Output -Color White
Write-Output "ReleaseId:  " $global:ReleaseId | Receive-Output -Color White
Write-Output ""
Write-Output ""
Write-Output ""
Write-Output ""
Write-Output ""
Start-Sleep 5


# Variables needed after Get-OSWIMFromISO finishes, passed to Update-Win10WIM
$SourcePath = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\SourceWIMs"
$TempFolder = "$DestinationFolder\$OSSKU\$global:OSVersion\$Architecture\Temp"
$ImageMountFolder = "$Mount\OSImage"
$BootImageMountFolder = "$Mount\BootImage"
$WinREImageMountFolder = "$Mount\WinREImage"
If ($BootWIM)
{
    $UpdateBootWIM = $True
}


# Download any components requested
If ($Device)
{
    Get-LatestDrivers -TempFolder $TempFolder -Device $Device
}

# We always need the VC Runtimes for our devices
Get-LatestVCRuntimes -TempFolder $TempFolder

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
Update-Win10WIM -SourcePath $SourcePath -SourceName $OSSKU -ServicingStack $ServicingStack -CumulativeUpdate $CumulativeUpdate -DotNet35 $DotNet35 -AdobeFlashUpdate $AdobeFlashUpdate -ImageMountFolder $ImageMountFolder -BootImageMountFolder $BootImageMountFolder -WinREImageMountFolder $WinREImageMountFolder -TempFolder $TempFolder -WindowsKitsInstall $WindowsKitsInstall -UpdateBootWIM $UpdateBootWIM -MakeUSBMedia $CreateUSB -MakeISOMedia $CreateISO


# Determine ending time
$Script_End_Time = (Get-Date).ToShortDateString()+", "+(Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

# How long did this take?
Write-Output "Script start: $Script_Start_Time" | Receive-Output -Color Gray
Write-Output "Script end:   $Script_End_Time" | Receive-Output -Color Gray
Write-Output ""
Write-Output "Execution time: $Script_Time_Taken seconds" | Receive-Output -Color White
