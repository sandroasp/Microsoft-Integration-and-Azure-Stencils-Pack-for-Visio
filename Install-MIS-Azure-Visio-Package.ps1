#########################################################
#                                                       #
# Install Microsoft Integration & Azure Stencils Pack   #
# Author: Sandro Pereira                                #
#                                                       #
#########################################################

[String]$location = Split-Path -Parent $PSCommandPath
[String]$destination = Get-ChildItem HKCU:\Software\Microsoft\Office\ -Recurse | Where-Object {$_.PSChildName -eq "Application"} | Get-ItemProperty -Name MyShapesPath | Select-Object -ExpandProperty MyShapesPath

$files = Get-ChildItem $location -recurse -force -Filter *.vssx
foreach($file in $files)
{
    if($file.PSPath.Contains("Previous Versions") -eq $false)
    {
        Copy-Item -Path $file.PSPath -Destination $destination -force
    }
}
