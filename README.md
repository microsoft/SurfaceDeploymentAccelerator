# SurfaceDeploymentAccelerator
Surface Deployment Accelerator (SDA) is a script-driven tool to create Windows images (WIM) for test or deployment that closely match the configuration of Bare Metal Recovery (BMR) images, minus certain preinstalled applications like Microsoft Office and the Surface UWP application.


# Need help?
Please use this GitHub Repos issue tracking capability to raise issues or feature requests.


# Prerequisites
 - Windows 10 2004 or newer environment to run the script.
 - Internet Explorer "first run experience" must have been completed in environment where script is run (you must run Internet Explorer once, otherwise driver/firmware and other downloads *will* fail and the script will throw errors indicating you have not run IE once in that environment before downloads will work properly).
 - Internet access from machine/environment where script is run (to download drivers and Office 365 - internet access is not required to use the image to (re)image devices, but is *strongly* recommended during that process as well).
 - The script must be run with administrative privileges to succeed.
 - Windows 10 2004 ADK/PE installed (script requires ADK/PE components to be installed in the environment used to create an image, and it will download/install them as necessary if they are not found).
 - USB drive, 16GB+ to hold image created by the script.  This drive will be formatted for use as necessary.
 - At least 20GB of disk space free on the drive used to run the script, to store images and downloaded files.


# How to use (general)
 - Acquire a Windows 10 1809 or newer ISO that includes Windows 10 Professional or Windows 10 Enterprise images
 - Open an elevated PowerShell prompt (not PSCore - currently this tool has not been fully tested with PSCore 6 or 7), change directory to where SDA was placed, and execute one of the following command to create an image - replace paths as appropriate.
 - To create a Windows 10 Professional image:
    X:\SurfaceDeploymentAccelerator\CreateSurfaceWindowsImage.ps1 -ISO "X:\en_windows_10_business_editions_version_2004_x64_dvd_d06ef8c5.iso" -OSSKU Pro -DestinationFolder C:\Output -Device SurfaceHub2

 - To create a Windows 10 Enterprise image for use on a Surface Hub 2:
    X:\SurfaceDeploymentAccelerator\CreateSurfaceWindowsImage.ps1 -ISO "X:\en_windows_10_business_editions_version_2004_x64_dvd_d06ef8c5.iso" -OSSKU Enterprise -DestinationFolder C:\Output -Device SurfaceHub2

- To create an image WITHOUT any additional patches downloaded/injected for a Surface Laptop 3:
    X:\SurfaceDeploymentAccelerator\CreateSurfaceWindowsImage.ps1 -ISO "X:\en_windows_10_business_editions_version_2004_x64_dvd_d06ef8c5.iso" -OSSKU Pro -DestinationFolder C:\Output -Device SurfaceLaptop3 -ServicingStack $false -CumulativeUpdate $false -CumulativeDotNetUpdate $false -AdobeFlashUpdate $false

- To create an image WITHOUT DotNet 3.5 or Office 365 C2R installed for a Surface Pro 7:
    X:\SurfaceDeploymentAccelerator\CreateSurfaceWindowsImage.ps1 -ISO "X:\en_windows_10_business_editions_version_2004_x64_dvd_d06ef8c5.iso" -OSSKU Pro -DestinationFolder C:\Output -Device SurfacePro7 -DotNet35 $false -Office365 $false

 - Once the script writes the image to the selected USB drive and has completed, take the resulting USB key to the device, and boot to it.  This will image the device and leave the device waiting in OOBE once complete.
 - To (re)image another device of the same type, simply use this USB key to (re)image that device as well.


# Known issues
**************************************************************************************************************************************************************************************************************************************************************************************************************
* - If you download this repository as a ZIP file from github.com on a Windows system, please right-click on the downloaded .zip file and select Properties, then check the "Unblock" box, and then click OK.  ONLY THEN IS IT SAFE TO UNZIP THE FILE - If you do not do this, your deployments *will* fail. *
**************************************************************************************************************************************************************************************************************************************************************************************************************
 - If an image fails to create, files may be held open in %temp%\Mount.  The script will detect files in this location on initial execution and fail if anything is found.  You will need to manually dismount the image (dism /unmount-wim /mountdir:%temp%\Mount\<folder>) to make certain that all folders under this location are unmounted.  Any failures to unmount will require manual cleanup before the script will execute successfully again.
 - If the path to the SDA folder itself, the Output folder parameter, or the LocalDriverPath parameter (if UseLocalDriverPath is set to $True) contains spaces, the script will abort.  Please make sure paths do not contain spaces.
 - Internet Explorer must be installed on the device used to create images, and first run wizard must be completed (you must successfully start Internet Explorer once) or file downloads may fail.
 - The Surface App from the Store is not sideloaded into the image at this time.  You will need to add the AppX to your own deployment or install once the imaging is complete.
 - An ISO containing English or Chinese (simplified) currently work with "Pro" or "Enterprise" string matching - other languages will be added over time (to the Languages.xml file).  Please file an issue here on Github if you have a language that we have not added as of yet.  This tool does not currently support LTSC or other SKUs (although you could edit the script and Languages.xml to add SKUs you'd like to use, they are not officially supported via this tool at this time).
 - WinPE is always in English, at this time.
 - This tool does NOT clean up images once created, only temporary files.  If you do not cleanup the Output folder, it will continue to consume more disk space for every image created.
 - On the very first execution, the \Logs folder sometimes is created in the location where CreateSurfaceWindowsImage.ps1 is executed from.  This only happens once, and logs do get created in the Output folder as well.  You can safely delete the Logs folder once script execution is complete.
 - The script supports Windows 10 1809 images or newer.  No support is planned for anything prior to 1809, and this minimum version will change over time as Windows releases new builds.  It is currently planned to be supported until 20H2 releases, at which point the minimum supported version of Windows 10 to use as an image source will move up to 1903.
 - Surface Hub 2 will support only Windows 10 Pro or Enterprise version 1903 and newer by default, and any future deviations will be documented here as well when possible.


# Full parameter documentation
The parameters that are supported to configure for the script are as follows:

 -ISO:                        Path to a Windows ISO file to use as the imaging source. (required)
 -DestinationFolder:          The folder used to place the resulting image files once complete. (required)
 -DotNet35:                   Install .NET 3.5 in the image, True or False.  True is the default.
 -ServicingStack:             Download/inject latest servicing stack update, True or False.  True is the default.
 -CumulativeUpdate:           Download/inject latest cumulative update, True or False.  True is the default.
 -CumulativeDotNetUpdate:     Download/inject latest cumulative update, True or False.  True is the default.
 -AdobeFlashUpdate:           Include latest Adobe Flash Player Security update, True or False.  True is the default.
 -Office365:                  Download and install the latest monthly C2R installation of Office 365, True or False.  True is the default.
 -Device:                     Enter Surface device type to download and inject latest drivers for.  Possible values: SurfacePro4, SurfacePro5, SurfacePro6, SurfacePro7, SurfaceLaptop, SurfaceLaptop2, SurfaceLaptop3, SurfaceBook, SurfaceBook2, SurfaceStudio, SurfaceStudio2, SurfaceGo, SurfaceGoLTE, and SurfaceHub2.  If this parameter is not specified, SurfacePro7 is used.
 -CreateUSB:                  Create bootable USB installation when finished, True or False.  True is the default.
 -CreateISO:                  Create bootable ISO file when finished, True or False.  False is the default.  This is useful for making imaging/scripting changes and testing quickly without needing USB keys and/or hardware to test.
 -WindowsKitInstall:          Enter target location of Windows ADK installation.  If not specified, the path "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit" will be used.
 -InstallWIM:                 Edit Install.wim file, True or False.  True is default.
 -BootWIM:                    Edit boot.wim, True or False.  True is default.  This can stop a boot.wim from being updated. If you're forcing it to use another boot.wim separate from one installed by ADK, you may want to set this to $false so the boot.wim isn't touched. If you choose $false, it’s expected you're going to use your own boot.wim.  This will also cause -CreateUSBKey to be forced to False, regardless of value passed in.
 -KeepOriginalWIM:            Keep customized unsplit WIM even if resulting image size is greater than 4GB (images over 4GB are split into SWM by default), True or False.  True is default.
 -UseLocalDriverPath          Use a local driver path and skip downloading the latest MSI for Device, True or False.  False is default.  If this parameter is set, you must also set LocalDriverPath to a valid path containing the extracted drivers you wish to inject, or this will fail.
 -LocalDriverPath             Filesystem (accessible to this PowerShell instance) path containing drivers to use.  Only read if UseLocalDriverPath is set to True.