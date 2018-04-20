function Get-WsusReport() {
    <#
        .SYNOPSIS
        Generates a report from WSUS server.

        .DESCRIPTION
        The cmdlet can connect to WSUS server and generate small report about
        defined host groups - number of needed to install updates, last sync satus.
        Also it can show report in format for using by Zabbix sender utility.

        .PARAMETER Server
        IP address or DNS name of WSUS server. By default getting from registry:
        HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate

        .PARAMETER PortNumber
        Port which WSUS server listening to. 8530 by default and it can be get from
        registry - HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate.

        .PARAMETER ZabbixSenderFormat
        Print or save report in format using by zabbix_sender utility.
        https://www.zabbix.com/documentation/3.0/ru/manpages/zabbix_sender

        .PARAMETER ZabbixItemName
        Name of item in Zabbix hosts. Default value is "os.updates.count".

        .EXAMPLE
        Gets report about Workstations hosts group:
        Get-WsusReport -Targetgroup Workstations

        .EXAMPLE
        Same with manual sets servername and port number:
        Get-WsusReport -Targetgroup Workstations -Server wsus.domain.local -PortNumber 8358

        .NOTES
        Version: 0.2.3
        Author: Khatsayuk Alexandr
        Git: https://github.com/asand3r

    #>

    [CmdLetBinding()]
        param
        (
            [Parameter(Mandatory=$true,Position=0)][array]$TargetGroup,
            [string]$Server = $((Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate).wuserver -replace "^https?://" -replace ":\d{4}$"),
            [bool]$useSecureConnection = $False,
            [int]$PortNumber = $((Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate).wuserver -replace "^https?://" -replace "^.+:"),
            [switch]$ZabbixSenderFormat = $False,
            [string]$ZabbixItemName = "os.updates.count"
        )

    # Loading WSUS library
    try {
        Add-Type -Path "C:\Program Files\Update Services\Api\Microsoft.UpdateServices.Administration.dll"
    } catch {
        $_.Exception.Message
    }

    # Establishing connection with WSUS server.
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($Server,$useSecureConnection,$portNumber)
    # Getting computers list of the TargetGroup
    $tgClients = $wsus.GetComputerTargets() | Where-Object {$_.RequestedTargetGroupName -in $targetGroup}
    
    # Creating result array and fill it
    $report = @()
    for ($i = 0; $i -lt $tgClients.Count; $i++) {
        $tgClient = $tgClients[$i]
        $tgClientInfo = New-Object psobject
        Add-Member -InputObject $tgClientInfo -Name Name -MemberType NoteProperty -Value $($tgClient.FullDomainName -replace ".$($env:USERDNSDOMAIN)")
        Add-Member -InputObject $tgClientInfo -Name IPAddress -MemberType NoteProperty -Value $tgClient.IPAddress
        Add-Member -InputObject $tgClientInfo -Name OSVersion -MemberType NoteProperty -Value $tgClient.OSDescription
        Add-Member -InputObject $tgClientInfo -Name LastSyncResult -MemberType NoteProperty -Value $tgClient.LastSyncResult
        Add-Member -InputObject $tgClientInfo -Name NeedCount -MemberType NoteProperty -Value $tgClient.GetUpdateInstallationSummary().NotInstalledCount
        # Progress bar
        Write-Progress -Activity “Gathering information...” -status “Found Computers $i” -percentComplete (($i/$tgClients.count)*100)
        $report += $tgClientInfo
    }

    if ($ZabbixSenderFormat) {
        foreach ($comp in $report) {
            $('"' + $comp.Name.ToUpper() + '"'), $ZabbixItemName, $comp.NeedCount -join " "
        }
    } else {
        Write-Output $report
    }
}