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
    $TPM = Get-WmiObject -Class "Win32_Tpm" -Namespace "ROOT\CIMV2\Security\MicrosoftTpm"
    
    Write-Output "Clearing TPM ownership....."
    $ClearRequest = $TPM.SetPhysicalPresenceRequest(14)
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

    $SystemInformation = Get-WmiObject -Namespace root\wmi -Class MS_SystemInformation
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
    # This can probably be reverted as new devices come along, but red on DarkBlue is unreadable on current DUT
    $host.UI.RawUI.BackgroundColor = "Black"
    $Host.UI.RawUI.ForegroundColor = "White"
    $host.UI.RawUI.WindowTitle = "$(Get-Location)"
}

$RemovableDisks = Get-Partition | Where-Object {$_.DiskId -like "*usbstor*"}
ForEach ($RemovableDisk in $RemovableDisks)
{
    $Drive = $RemovableDisk.DriveLetter
    $DriveLetter = $Drive + ":\"
    $SourceUSBKey = Get-ChildItem -Path "$DriveLetter" -Recurse | Where-Object { $_.Name -eq "Imaging.ps1" }
    If ($SourceUSBKey)
    {
        $Folder = Get-ChildItem -Path "$DriveLetter" -Recurse | Where-Object { $_.PSIsContainer -and $_.Name -like "Sources*" }
        $DiskPartScript = Get-ChildItem -Path "X:\" -Recurse | Where-Object { $_.Name -eq "CreatePartitions-UEFI.txt" }
        $DiskPartScriptSource = Get-ChildItem -Path "X:\" -Recurse | Where-Object { $_.Name -eq "CreatePartitions-UEFI_Source.txt" }
        $WIMFile = Get-ChildItem -Path $DriveLetter -Recurse | Where-Object { $_.Name -like "*install*.wim" }
        $SWMFile = Get-ChildItem -Path $DriveLetter -Recurse | Where-Object { $_.Name -like "*install*--Split.swm" }
        If ($DiskPartScript)
        {
            $DiskPartScriptPath = $DiskPartScript.FullName
            $DiskPartScriptSourcePath = $DiskPartScriptSource.FullName
        }
        If ($WIMFile)
        {
            [string]$WIMFilePath = $WIMFile.FullName
        }
        If ($SWMFile)
        {
            $SplitWIM = $true
            $SWMFilePath = $SWMFile.FullName
            $SWMFilePattern = $SWMFile.DirectoryName + "\" + $SWMFile.BaseName + '*.swm'
        }
    }
    Else
    {
        Break
    }
}

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$diskpart = "$env:windir\System32\diskpart.exe"
$managebde = "$env:windir\System32\manage-bde.exe"
$bcdboot = "$env:windir\System32\bcdboot.exe"



Write-Output "********************"
Write-Output "  OS IMAGE INSTALL  "
Write-Output "********************"

$UEFIVer = ($(& wmic bios get SMBIOSBIOSVersion /format:table)[2])
Write-Output "- UEFI Information: $UEFIVer"
Write-Output "- WinPE Information"
$RegPath = "Registry::HKEY_LOCAL_MACHINE\Software"
$WinPEVersion = ""
$CurrentVersion = Get-ItemProperty -Path "$RegPath\Microsoft\Surface\OSImage" -ErrorAction SilentlyContinue
If ($CurrentVersion)
{
    try
    {
        Write-Output "   - ImageName $($CurrentVersion.ImageName)"
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
        Write-Output "   - BuildLab $($NTCurrentVersion.BuildLab)"
        Write-Output "   - BuildLabEx $($NTCurrentVersion.BuildLabEx)"
        Write-Output "   - ProductName $($NTCurrentVersion.ProductName)"
    }
    catch {}
}


ClearTpm


# Configure installation disk
$Result = Get-DiskIndex
Write-Output "Configuring disk $Result for imaging..."
Clear-Content -Path $DiskPartScriptPath
Add-Content -Path $DiskPartScriptPath -Value "select disk $Result"
Get-Content -Path $DiskPartScriptSourcePath | Add-Content -Path $DiskPartScriptPath
& $diskpart /s $DiskPartScriptPath


# Enable XTS-AES 256bit cipher and bitlocker used space
Write-Output "Enabling XTS-AES 256Bit Bitlocker encryption"
EnableBitlocker
& $managebde -on W: -UsedSpaceOnly


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
#& $reagentc /info
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
