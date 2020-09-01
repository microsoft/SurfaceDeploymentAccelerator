Set-Location $PSScriptRoot
Start-Sleep 2
Get-Process sysprep -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep 2


Remove-Item -Path C:\Windows\Panther\unattend.xml -Force -ErrorAction SilentlyContinue
Remove-Item -Path C:\Windows\system32\sysprep\unattend.xml -Force -ErrorAction SilentlyContinue
Start-Sleep 2
Copy-Item -Path "$env:Temp\Reseal.xml" -Destination "$env:WINDIR\system32\sysprep\unattend.xml" -Force -ErrorAction SilentlyContinue
Start-Sleep 5


Start-Sleep 2
Get-Process sysprep -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep 5

$sysprep = "$env:WINDIR\System32\Sysprep\sysprep.exe"
$args = "/oobe /generalize /reboot /unattend:$env:WINDIR\system32\sysprep\unattend.xml"
Start-Process -FilePath $sysprep -ArgumentList $args -Wait
Start-Sleep 2