$stencils = @(`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/AI%20and%20Machine%20Learning/MIS%20AI%20and%20Machine%20Learning%20Stencils.vssx", `
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Apps%20and%20Systems/MIS%20Apps%20and%20Systems%20Logo%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Azure/MIS%20Azure%20Additional%20or%20Support%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Azure/MIS%20Azure%20Mono%20Color.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Azure/MIS%20Azure%20Old%20Versions.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Azure/MIS%20Azure%20Others%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Azure/MIS%20Azure%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Databases%20and%20Analytics/MIS%20Databases%20and%20Analytics%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Enterprise%20Integration/MIS%20Additional%20or%20Support%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Enterprise%20Integration/MIS%20Integration%20Fun.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Enterprise%20Integration/MIS%20Integration%20Patterns%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Enterprise%20Integration/Microsoft%20Integration%20Stencils%20Old%20Version%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Enterprise%20Integration/Microsoft%20Integration%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/IoT/MIS%20IoT%20Devices%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Office%20365/MIS%20Office365.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Other%20Providers/MIS%20SAP%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Buildings%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Developer%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Devices%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Files%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Generic%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Infrastructure%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Servers%20(HEX)%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Others/MIS%20Users%20and%20Roles%20Stencils.vssx",`
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Power%20BI/MIS%20Power%20BI%20Stencils.vssx", `
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/PowerApps%20and%20Flow/MIS%20PowerApps%20and%20Flows%20Stencils.vssx", `
    "https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio/blob/master/Security%20and%20Governance/MIS%20Security%20and%20Governance.vssx")
    
$rootpath = "$([Environment]::GetFolderPath("MyDocuments"))\My Shapes\sandroasp"

foreach($stencil in $stencils){
    $filename = [System.IO.Path]::GetFileName($stencil).Replace('%20', ' ')
    $targetfile = [System.IO.Path]::Combine($rootpath, $filename)
    Write-Host "Downloading $filename"
    Invoke-WebRequest -Uri $stencil -OutFile "$targetfile"
}
