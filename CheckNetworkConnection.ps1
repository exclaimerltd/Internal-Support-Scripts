$logTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = "C:\Temp\NetworkTest_$logTimeStamp.txt"
$Target = "8.8.8.8"

$DomainsToCheck = @(
    "au.outlooksignatures.exclaimer.net",
    "outlookclient.exclaimer.net"
)

if (!(Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - ### - Started Logging" | Out-File -FilePath $LogPath

Write-Host "Running: Check $LogPath" -ForegroundColor Yellow

$previousPingState = $null
$iterationCounter = 0

while ($true) {

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $ping = Test-Connection -ComputerName $Target -Count 1 -Quiet -ErrorAction SilentlyContinue

if ($ping) {
    if ($previousPingState -ne $true) {        
        "`n$timestamp - ########################################" | Out-File -FilePath $LogPath -Append
        "$timestamp - Ping to $Target SUCCESS" | Out-File -FilePath $LogPath -Append
        $state = "Success"


        $netConfigs = Get-NetIPConfiguration -ErrorAction SilentlyContinue
        if ($netConfigs) {
            foreach ($config in $netConfigs) {
                $adapter = $config.InterfaceAlias
                "$timestamp - Adapter: $adapter" | Out-File $LogPath -Append
                "$timestamp - IPv4: $($config.IPv4Address.IPAddress -join ', ')" | Out-File $LogPath -Append
                "$timestamp - Gateway: $($config.IPv4DefaultGateway.NextHop)" | Out-File $LogPath -Append
                "$timestamp - DNS: $($config.DnsServer.ServerAddresses -join ', ')" | Out-File $LogPath -Append
            }
        }
        foreach ($domain in $DomainsToCheck) {

        "$timestamp - # - DNS lookup for $domain" | Out-File $LogPath -Append

        try {
            $dnsResult = Resolve-DnsName -Name $domain -ErrorAction Stop

            foreach ($record in $dnsResult) {
                if ($record.IPAddress) {
                    "$timestamp - $domain -> $($record.IPAddress)" | Out-File $LogPath -Append
                }
            }
        }
        catch {
            "$timestamp - Resolve-DnsName failed for $domain, falling back to nslookup" | Out-File $LogPath -Append
            (nslookup $domain 2>&1) | Out-Null
        }
    }
        "$timestamp - ### Monitoring ###" | Out-File $LogPath -Append
    }
    $previousPingState = $true
}
else {
    "`n$timestamp - ########################################" | Out-File -FilePath $LogPath -Append
    "$timestamp - # - Ping to $Target FAILED" | Out-File -FilePath $LogPath -Append
    $state = "Failed"

    # Get adapters (modern method)
    $netConfigs = Get-NetIPConfiguration -ErrorAction SilentlyContinue

    if ($netConfigs) {

        foreach ($config in $netConfigs) {

            $adapter = $config.InterfaceAlias

            "$timestamp - # - Adapter: $adapter" | Out-File $LogPath -Append
            "$timestamp - IPv4: $($config.IPv4Address.IPAddress -join ', ')" | Out-File $LogPath -Append
            "$timestamp - Gateway: $($config.IPv4DefaultGateway.NextHop)" | Out-File $LogPath -Append
            "$timestamp - DNS: $($config.DnsServer.ServerAddresses -join ', ')" | Out-File $LogPath -Append
            "$timestamp - ---" | Out-File $LogPath -Append
        }

    }
    else {
        "$timestamp - No Adapter detected" | Out-File $LogPath -Append
    }

    foreach ($domain in $DomainsToCheck) {

        "$timestamp - # - DNS lookup for $domain" | Out-File $LogPath -Append

        try {
            $dnsResult = Resolve-DnsName -Name $domain -ErrorAction Stop

            foreach ($record in $dnsResult) {
                if ($record.IPAddress) {
                    "$timestamp - $domain -> $($record.IPAddress)" | Out-File $LogPath -Append
                }
            }
        }
        catch {
            "$timestamp - Resolve-DnsName failed for $domain, falling back to nslookup" | Out-File $LogPath -Append
            (nslookup $domain 2>&1) | Out-Null
        }
    }
    $previousPingState = $false
}

    $iterationCounter++
    Write-Host "`rIteration: $iterationCounter $state          " -NoNewline
    Start-Sleep -Seconds 5
}