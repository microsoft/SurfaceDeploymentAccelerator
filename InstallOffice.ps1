Set-Location $PSScriptRoot

Set-Location ($env:Temp + '\Office365')
$Process = (Start-Process "setup.exe" -ArgumentList "/configure .\O365_configuration.xml" -Wait -PassThru)
$Process.WaitForExit()
exit(0)