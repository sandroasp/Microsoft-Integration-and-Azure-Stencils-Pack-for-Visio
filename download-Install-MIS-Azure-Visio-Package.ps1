######################################################################
#                                                                    #
# Download and Install Microsoft Integration & Azure Stencils Pack   #
# Author: Sandro Pereira                                             #
#                                                                    #
######################################################################

function DownloadGitHubRepository 
{ 
    [OutputType([String])]
    param( 
       [Parameter(Mandatory=$True)] 
       [string] $Name, 
        
       [Parameter(Mandatory=$True)] 
       [string] $Author, 
        
       [Parameter(Mandatory=$False)] 
       [string] $Branch = "master", 
        
       [Parameter(Mandatory=$False)] 
       [string] $Location = "c:\temp" 
    ) 
    
    # Force to create a zip file 
    $ZipFile = "$location\$Name.zip" 
    New-Item $ZipFile -ItemType File -Force

    #$RepositoryZipUrl = "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/archive/master.zip"
    $RepositoryZipUrl = "https://api.github.com/repos/$Author/$Name/zipball/$Branch"  
    # download the zip 
    Write-Host 'Starting downloading the GitHub Repository'
    Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile
    Write-Host 'Download finished'

    #Extract Zip File
    Write-Host 'Starting unziping the GitHub Repository locally'
    Expand-Archive -Path $ZipFile -DestinationPath $location -Force
    Write-Host 'Unzip finished'
    [String]$unzipFolder = (Get-ChildItem -Path $Location -Filter "sandroasp-Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio*" -Recurse -Directory).Fullname
    
    # remove zip file
    Remove-Item -Path $ZipFile -Force 

    return $unzipFolder
}

[String]$location = Split-Path -Parent $PSCommandPath
[String]$gitHubFolder = (DownloadGithubRepository -Name 'Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio' -Author 'sandroasp' -Location $location)[-1]

[String]$location = Get-Location
[String]$destination = Get-ChildItem HKCU:\Software\Microsoft\Office\ -Recurse | Where-Object {$_.PSChildName -eq "Application"} | Get-ItemProperty -Name MyShapesPath | Select-Object -ExpandProperty MyShapesPath

Write-Host 'Starting to install Microsoft Integration & Azure Stencils Pack locally'
$files = Get-ChildItem $gitHubFolder -recurse -force -Filter *.vssx
foreach($file in $files)
{
    if($file.PSPath.Contains("Previous Versions") -eq $false)
    {
        Copy-Item -Path $file.PSPath -Destination $destination -force
    }
}
Write-Host 'Microsoft Integration & Azure Stencils Pack installed'