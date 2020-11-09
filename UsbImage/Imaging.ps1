<#
.SYNOPSIS
    This script partitions the disk and applies a WIM or SWM files and sets recovery.

.NOTES
    Author:       Microsoft
    Last Update:  20th October 2020
    Version:      1.2.5.4

    Version 1.2.5.4
    - Match master package version

    Version 1.2.5.3
    - Match master package version

    Version 1.2.3
    - Changed all Get-WmiObject calls with Get-CimInstance calls to be more compatible with PowerShell Core
    
    Version 1.2.2
    - No changes

    Version 1.2.1
    - Performance improvements
    - Removed forced Bitlocker encryption - causes issues on non-Surface devices

    Version 1.2.0
    - Added support for running from NVME USB-attached drives
    - Bugfixes

    Version 1.1.0
    - Added support for running in a VM
    - Added support for running from an ISO or other read-only media

    Version 1.0.0
    - Initial release
#>



$SDAVersion = "1.2.5.4"
cls



Function New-RegKey
{
    param($key)
  
    $key = $key -replace ':',''
    $parts = $key -split '\\'
  
    $tempkey = ''
    $parts | ForEach-Object {
        $tempkey += ($_ + "\")
        if ( (Test-Path "Registry::$tempkey") -eq $false)  {
        New-Item "Registry::$tempkey" | Out-Null
        }
    }
}



Function ClearTPM
{
    $TPM = Get-CimInstance -ClassName "Win32_Tpm" -Namespace "ROOT\CIMV2\Security\MicrosoftTpm"

    Write-Output "Clearing TPM ownership....."
    $ClearRequest = $TPM.SetPhysicalPresenceRequest(14) | Out-Null
    If ($ClearRequest.ReturnValue -eq 0)
    {
        Write-Output "Successfully cleared the TPM chip. A reboot is required."
        Write-Output ""
    }
    Else
    {
        Write-Warning "Failed to clear TPM ownership..."
        Write-Output ""
    }
}



Function EnableBitlocker
{
    # Configure Bitlocker for XTS-AES 256Bit Encryption
    $FVERegKey = "HKLM:\SOFTWARE\Policies\Microsoft\FVE"
    New-RegKey $FVERegKey

    $EncryptionMethodWithXtsOsRegValue = Get-ItemProperty $FVERegKey EncryptionMethodWithXtsOs -ErrorAction SilentlyContinue
    $EncryptionMethodWithXtsFdvRegValue = Get-ItemProperty $FVERegKey EncryptionMethodWithXtsFdv -ErrorAction SilentlyContinue
    $OSEncryptionTypeRegValue = Get-ItemProperty $FVERegKey OSEncryptionType -ErrorAction SilentlyContinue

    If ($EncryptionMethodWithXtsOsRegValue -eq $null)
    {
        New-ItemProperty -Path $FVERegKey -Name EncryptionMethodWithXtsOs -PropertyType DWORD -Value 7 | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $FVERegKey -Name EncryptionMethodWithXtsOs -Value 7
    }

    If ($EncryptionMethodWithXtsFdvRegValue -eq $null)
    {
        New-ItemProperty -Path $FVERegKey -Name EncryptionMethodWithXtsFdv -PropertyType DWORD -Value 7 | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $FVERegKey -Name EncryptionMethodWithXtsFdv -Value 7
    }

    If ($OSEncryptionTypeRegValue -eq $null)
    {
        New-ItemProperty -Path $FVERegKey -Name OSEncryptionType -PropertyType DWORD -Value 1 | Out-Null
    }
    Else
    {
        Set-ItemProperty -Path $FVERegKey -Name OSEncryptionType -Value 1
    }
}



Function Clear-StoragePool
{
    $Pools = Get-StoragePool -IsPrimordial $false -ErrorAction SilentlyContinue
    If ($Pools)
    {
        Write-Output "Clearing storage pool: $($Pools.FriendlyName)"
        $Pools | Get-VirtualDisk | Remove-VirtualDisk -Confirm:$false
        $Pools | Remove-StoragePool -Confirm:$false
    }
}



Function Enable-SpacesBootSimple
{
    param(
        [Parameter (Mandatory=$false, Position=0)]
        [string] $DiskNumber
    )
    try
    {
        $PhysicalDisks = @(Get-PhysicalDisk | Where-Object { $_.BusType -ne "USB" -and $_.CanPool -eq $true })

        Write-Output "Creating Storage Pool with $($PhysicalDisks.Count) disks"
        $Pool = New-StoragePool -FriendlyName "Boot" -StorageSubsystemFriendlyName * -PhysicalDisks $PhysicalDisks
        
        Write-Output "Creating Virtual Disk"
        $VirtualDisk = New-VirtualDisk -FriendlyName Boot -StoragePoolFriendlyName $Pool.FriendlyName -UseMaximumSize -ResiliencySettingName Simple -WriteCacheSize 0

        Write-Output "Boot Storage Space created"
        
        $Disks = Get-Disk | Where-Object { $_.BusType -ne "USB" }
        ForEach ($Disk in $Disks)
        {
            If ($Disk.Model -eq "Storage Space")
            {
                $Size = $Disk.Size /1GB
                $Index = $Disk.Number
                $Name = $Disk.FriendlyName
                $Type = $Disk.BusType
                $Serial = $Disk.SerialNumber

                Write-Output "Chosen installation disk:"
                Write-Output "Disk Index:  $Index"
                Write-Output "Disk Name:   $Name"
                Write-Output "Disk Serial: $Serial"
                Write-Output "Disk Type:   $Type"
                Write-Output "Disk Size:   $Size"

                $DiskIndex = "$Index"
            }
        }
    }
    catch
    {
        $_ | Format-List -Force
        "ERROR enabling Storage Space!!!"
        Clear-StoragePool
        Exit
    }
}



Function Get-DiskIndex
{
    # Set Disk to image to
    Update-StorageProviderCache -DiscoveryLevel Full | Out-Null

    $SystemInformation = Get-CimInstance -ClassName MS_SystemInformation -Namespace root\wmi
    $Product = $SystemInformation.SystemSKU
    $Disks = Get-Disk | Where-Object { $_.BusType -ne "USB"}

    If ($Disks.Length -gt 1)
    {
        $Disks = Get-Disk | Where-Object { $_.BusType -ne "USB" -and $_.CanPool -eq $true }
        If ($Disks.Length -gt 1)
        {
            Enable-SpacesBootSimple
        }
    }

    If ($Product -like "Surface_Studio*")
    {
        ForEach ($Disk in $Disks)
        {
            # Everything seems ok here, even if RST is broken...
            If (($Disk.BusType -eq "RAID") -and ($Disk.Number -eq "0"))
            {
                $Size = $Disk.Size /1GB
                $Index = $Disk.Number
                $Name = $Disk.FriendlyName
                $Type = $Disk.BusType
                $Serial = $Disk.SerialNumber

                $DiskIndex = $Index

            }
        
            # Perhaps booted from disk 0 where it's not the RST - find the non-cache drive (will be >64 or 128GB depending on SKU)
            # No USB drives, and no drives smaller than 128GB should give us the proper RST mechanical disk:
            ElseIf (($Disk.BusType -ne "USB") -and ($Disk.Size -gt "130000000000"))
            {
                $Size = $Disk.Size /1GB
                $Index = $Disk.Number
                $Name = $Disk.FriendlyName
                $Type = $Disk.BusType
                $Serial = $Disk.SerialNumber

                $DiskIndex = $Index
            }
        }
    }
    ElseIf (($Product -like "Surface_Pro*") -or ($Product -like "Surface_Laptop*"))
    {
        ForEach ($Disk in $Disks)
        {
            If ($Disk.Model -like "*Storage Space*")
            {
                $Size = $Disk.Size /1GB
                $Index = $Disk.Number
                $Name = $Disk.FriendlyName
                $Type = $Disk.BusType
                $Serial = $Disk.SerialNumber

                $DiskIndex = $Index
            }
        }
    }
    Else
    {
        $DiskIndex = $Disks.Number
    }


    If (!($DiskIndex))
    {
        $DiskIndex = "0"
    }

    Return $DiskIndex
}



###########################
# Begin script processing #
###########################

If ($ENV:PROCESSOR_ARCHITECTURE -eq 'ARM64')
{
    try
    {
        # Replace with a custom Get-Volume for use with pwsh.exe
        Import-Module X:\windows\System32\WindowsPowershell\v1.0\Modules\Storage\Storage.psd1
    }
    catch {}
}

If ($ExecutionContext.SessionState.LanguageMode -eq "FullLanguage")
{
    # This can probably be reverted as new devices come along, but red on DarkBlue is unreadable on current devices
    $host.UI.RawUI.BackgroundColor = "Black"
    $Host.UI.RawUI.ForegroundColor = "White"
    $host.UI.RawUI.WindowTitle = "$(Get-Location)"
}

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition


Write-Output "SDA version:  $SDAVersion"
Start-Sleep 2
Write-Output ""
Write-Output ""
Write-Output "********************"
Write-Output "  OS IMAGE INSTALL  "
Write-Output "********************"

$UEFIVer = ($(& wmic bios get SMBIOSBIOSVersion /format:table)[2])
Write-Output "- UEFI Information:   $UEFIVer"
Write-Output ""
Write-Output "- WinPE Information"
$RegPath = "Registry::HKEY_LOCAL_MACHINE\Software"
$WinPEVersion = ""

$CurrentVersion = Get-ItemProperty -Path "$RegPath\Microsoft\Surface\OSImage" -ErrorAction SilentlyContinue
If ($CurrentVersion)
{
    try
    {
        Write-Output "   - ImageName        $($CurrentVersion.ImageName)"
        $WinPEVersion = $($CurrentVersion.ImageName)
    }
    catch {}
    try
    {
        Write-Output "   - RebasedImageName $($CurrentVersion.RebasedImageName)"
    }
    catch {}
}

$NTCurrentVersion = Get-ItemProperty -Path "$RegPath\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
If ($NTCurrentVersion)
{
    try
    {
        Write-Output "   - BuildLab         $($NTCurrentVersion.BuildLab)"
        Write-Output "   - BuildLabEx       $($NTCurrentVersion.BuildLabEx)"
        Write-Output "   - ProductName      $($NTCurrentVersion.ProductName)"
    }
    catch {}
}

Write-Output ""
Write-Output "- Hardware Information"
$SystemInformation = (Get-CimInstance -ClassName MS_SystemInformation -Namespace root\wmi)

If ($SystemInformation)
{
    try
    {
        Write-Output "   - Manufacturer     $($SystemInformation.BaseBoardManufacturer)"
        Write-Output "   - Product          $($SystemInformation.BaseBoardProduct)"
        Write-Output "   - SystemSKU        $($SystemInformation.SystemSKU)"
    }
    catch {}
}
Write-Output ""
Write-Output ""


# Make sure we have valid diskpart scripts and installation WIM/SWMs located before we go further
$diskpart = "$env:windir\System32\diskpart.exe"
$managebde = "$env:windir\System32\manage-bde.exe"
$bcdboot = "$env:windir\System32\bcdboot.exe"

$RamDrive = (Get-Location).Drive.Name
$DriveLetter = $RamDrive + ":\"
$SourceDrive = Get-ChildItem -Path "$DriveLetter" -Recurse | Where-Object { $_.Name -eq "Imaging.ps1" }

If ($SourceDrive)
{
    $DiskPartScript = Get-ChildItem -Path "X:\" -Recurse | Where-Object { $_.Name -eq "CreatePartitions-UEFI.txt" }
    $DiskPartScriptSource = Get-ChildItem -Path "X:\" -Recurse | Where-Object { $_.Name -eq "CreatePartitions-UEFI_Source.txt" }
    If ($DiskPartScript)
    {
        $DiskPartScriptPath = $DiskPartScript.FullName
        $DiskPartScriptSourcePath = $DiskPartScriptSource.FullName
    }
}

Write-Output "Finding all attached drives with recognized filesystems..."
Write-Output ""
$Drives = Get-CimInstance -ClassName Win32_LogicalDisk
If (!($Drives))
{
    Write-Output "No drives found, exiting."
    Write-Output ""
    Exit
}
Else
{
    Write-Output "Drives Found:"
    $Drives
    Write-Output ""
    Write-Output ""
    $WIMFound = $false
}

ForEach ($Drive in $Drives)
{
    $TempDrive = $Drive.DeviceID
    $TempPath = "$TempDrive\Sources"
    If (($Drive.Description -like "*Disc*") -and (!($Drive.FileSystem)))
    {
        # Drive does not have an ISO/disc inserted, skip
    }
    Else
    {
        If (!(Test-Path "$TempPath"))
        {
            # No \Sources folder found, thus no WIM/SWMs should be on this volume, skipping
        }
        Else
        {
            Write-Output "Checking drive $TempDrive for WIM/SWM files..."
            $WIMFile = Get-ChildItem -Path $TempPath -Recurse | Where-Object { $_.Name -like "*install*.wim" }
            $SWMFile = Get-ChildItem -Path $TempPath -Recurse | Where-Object { $_.Name -like "*install*--Split.swm" }
        }
        
    }

    If ($WIMFile)
    {
        $WIMFound = $true
        $WIMFilePath = $WIMFile.FullName
        Write-Output "Found file $WIMFilePath"
        Write-Output ""
        Break
    }
    ElseIf ($SWMFile)
    {
        $WIMFound = $true
        $SplitWIM = $true
        $SWMFilePath = $SWMFile.FullName
        Write-Output "Found file $SWMFilePath"
        Write-Output ""
        $SWMFilePattern = $SWMFile.DirectoryName + "\" + $SWMFile.BaseName + '*.swm'
        Break
    }
}

If ($WIMFound -eq $false)
{
    Write-Output "WIM/SWM file(s) not found.  Exiting..."
    Write-Output ""
    Exit
}


# Configure installation disk
$Result = Get-DiskIndex
Write-Output "Configuring disk $Result for imaging..."
Clear-Content -Path $DiskPartScriptPath
Add-Content -Path $DiskPartScriptPath -Value "select disk $Result"
Get-Content -Path $DiskPartScriptSourcePath | Add-Content -Path $DiskPartScriptPath
& $diskpart /s $DiskPartScriptPath


<#
# This isn't necessary on Modern Standby devices like Surface, but keeping in for custom device/custom deployment work
Enable Bitlocker
If ("$($SystemInformation.SystemFamily)" -like "*Virtual*")
{
    # VM, don't try to enable Bitlocker
}
Else
{
    #ClearTPM
    Write-Output "Enabling Bitlocker encryption"
    EnableBitlocker
    & $managebde -on W: -UsedSpaceOnly
    Write-Output ""
    Write-Output ""
}
#>


# Apply image
If ($SplitWIM -eq $true)
{
    Write-Output "Applying WIM $SWMFilePath using pattern $SWMFilePattern to W: ..."
    Expand-WindowsImage -ImagePath $SWMFilePath -SplitImageFilePattern $SWMFilePattern -ApplyPath "W:" -Index 1
    Write-Output ""
}
Else
{
    Write-Output "Applying WIM $WIMFilePath to W: ..."
    Expand-WindowsImage -ImagePath $WIMFilePath -ApplyPath "W:" -Index 1
    Write-Output ""
}


# Set System partition bootable
Write-Output "Marking system partition bootable..."
& $bcdboot W:\Windows /s S:
Write-Output ""


# Configure recovery
Write-Output "Configuring recovery..."
$RecoveryPath = "T:\Recovery"
$WinREPath = "T:\Recovery\WindowsRE"
$WinREWIM = "T:\Recovery\WindowsRE\WinRE.wim"
If (!(Test-Path "$RecoveryPath"))
{
    New-Item -Path "$RecoveryPath" -ItemType "directory" | Out-Null
}

If (!(Test-Path "$WinREPath"))
{
    New-Item -Path "$WinREPath" -ItemType "directory" | Out-Null
}

Write-Output "Copying W:\Windows\System32\Recovery\WinRE.WIM to $WinREPath..."
Copy-Item -Path "W:\Windows\System32\Recovery\WinRE.wim" -Destination $WinREPath
$reagentc = "W:\Windows\System32\reagentc.exe"
& $reagentc /setreimage /path $WinREWIM /target W:\Windows
Write-Output ""
Sleep 2


# Use MessageBox to prompt for reboot
Write-Output "Restart prompt..."
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
$result = [System.Windows.Forms.MessageBox]::Show("Reboot?", $PromptString, 'YesNo')
If ( $result -eq 'No' )
{
    try
    {
        $input = [int] ( Read-Host -Prompt "?" )
    }
    catch
    {
        Write-Output "Invalid Input"
        continue
    }
}
Else
{
    Restart-Computer -Force
    Exit
}