Set-Location $PSScriptRoot
Start-Sleep 2

If (Test-path "C:\Program Files\Surface\TaskbarLayoutModification.xml")
{
    & reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v "LayoutXMLPath" /d "C:\Program Files\Surface\TaskbarLayoutModification.xml"
}