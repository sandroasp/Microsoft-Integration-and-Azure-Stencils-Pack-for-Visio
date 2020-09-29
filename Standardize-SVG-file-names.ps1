#########################################################
#                                                       #
# Standardize SVG file names (CamelCase without spaces) #
# Author: Sandro Pereira                                #
#                                                       #
#########################################################

[String]$location = Get-Location

$files = Get-ChildItem $location -recurse -force -Filter *.svg
foreach($file in $files)
{
    $textInfo = (Get-Culture).TextInfo
    $newname = $textInfo.ToTitleCase([String]$file.Name.ToLower()).Replace(" ", "-").Replace(".Svg",".svg")
    Try
    {
        Rename-Item -Path $file.PSPath $newname -ErrorAction 'Stop'
    }
    Catch
    {
        Write-Host $file.PSPath
    }
}
