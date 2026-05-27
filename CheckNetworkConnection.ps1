$logTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = "C:\Temp\NetworkTest_$logTimeStamp.txt"
$Target = "8.8.8.8"

$DomainsToCheck = @(
    @{ Host = "au.outlooksignatures.exclaimer.net"; Port = 443 },
    @{ Host = "outlookclient.exclaimer.net";        Port = 443 },
    @{ Host = "login.microsoftonline.com";         Port = 443 }  # NOTE: trailing 's' - verify if intentional
)

if (!(Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - ### - Started Logging" | Out-File -FilePath $LogPath

Write-Host "Running: Check $LogPath" -ForegroundColor Yellow

$maxAttempts = 200  # Set to 0 to run indefinitely)
$iterationCounter = 0

while ($true) {

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # --- Ping check ---
    $pingSuccess = Test-Connection -ComputerName $Target -Count 1 -Quiet -ErrorAction SilentlyContinue

    # --- Domain checks (run every iteration regardless of ping result) ---
    $domainResults = @()
    $anyDomainFailed = $false

    foreach ($item in $DomainsToCheck) {
        $domain = $item.Host
        $port   = $item.Port

        $entry = [PSCustomObject]@{
            Domain     = $domain
            Port       = $port
            DnsSuccess = $false
            IPs        = @()
            DnsError   = $null
            TcpSuccess = $false
            TcpError   = $null
        }

        # --- DNS check ---
        try {
            $dnsResult = Resolve-DnsName -Name $domain -ErrorAction Stop
            $ips = $dnsResult | Where-Object { $_.IPAddress } | ForEach-Object { $_.IPAddress }

            if ($ips) {
                $entry.DnsSuccess = $true
                $entry.IPs        = $ips
            }
            else {
                $entry.DnsError  = "No IP records returned"
                $anyDomainFailed = $true
            }
        }
        catch {
            # Resolve-DnsName failed - try nslookup as fallback
            $nsOutput  = (nslookup $domain 2>&1) -join "`n"
            $parsedIPs = $nsOutput | Select-String -Pattern '\b\d{1,3}(\.\d{1,3}){3}\b' -AllMatches |
                         ForEach-Object { $_.Matches.Value } |
                         Where-Object { $_ -notmatch '^(127\.|0\.)' }   # strip loopback/unspecified

            if ($parsedIPs) {
                $entry.DnsSuccess = $true
                $entry.IPs        = $parsedIPs
                $entry.DnsError   = "Resolve-DnsName failed; nslookup succeeded"
            }
            else {
                $entry.DnsError  = "Resolve-DnsName failed; nslookup also failed: $($_.Exception.Message)"
                $anyDomainFailed = $true
            }
        }

        # --- TCP connectivity check ---
        try {
            $tcpResult = Test-NetConnection -ComputerName $domain -Port $port -WarningAction SilentlyContinue -ErrorAction Stop

            if ($tcpResult.TcpTestSucceeded) {
                $entry.TcpSuccess = $true
            }
            else {
                $entry.TcpError  = "TCP connection refused or timed out"
                $anyDomainFailed = $true
            }
        }
        catch {
            $entry.TcpError  = "Test-NetConnection failed: $($_.Exception.Message)"
            $anyDomainFailed = $true
        }

        $domainResults += $entry
    }

    # --- Decide whether to log this iteration ---
    $shouldLog = (-not $pingSuccess) -or $anyDomainFailed

    if ($shouldLog) {

        "`n$timestamp - ########################################" | Out-File -FilePath $LogPath -Append

        if (-not $pingSuccess) {
            "$timestamp - # - Ping to $Target FAILED" | Out-File -FilePath $LogPath -Append
        }
        else {
            "$timestamp - # - Ping to $Target SUCCESS (domain failure triggered log)" | Out-File -FilePath $LogPath -Append
        }

        # Network adapter info
        $netConfigs = Get-NetIPConfiguration -ErrorAction SilentlyContinue

        if ($netConfigs) {
            foreach ($config in $netConfigs) {
                $adapter = $config.InterfaceAlias
                "$timestamp - # - Adapter: $adapter"                                          | Out-File $LogPath -Append
                "$timestamp -     IPv4:    $($config.IPv4Address.IPAddress -join ', ')"        | Out-File $LogPath -Append
                "$timestamp -     Gateway: $($config.IPv4DefaultGateway.NextHop)"             | Out-File $LogPath -Append
                "$timestamp -     DNS:     $($config.DnsServer.ServerAddresses -join ', ')"   | Out-File $LogPath -Append
                "$timestamp -     ---"                                                         | Out-File $LogPath -Append
            }
        }
        else {
            "$timestamp - # - No adapter detected" | Out-File $LogPath -Append
        }

        # Domain check results
        foreach ($result in $domainResults) {
            "$timestamp - # - Check: $($result.Domain) (Port $($result.Port))" | Out-File $LogPath -Append

            # DNS
            if ($result.DnsSuccess) {
                "$timestamp -     DNS:  OK -> $($result.IPs -join ', ')" | Out-File $LogPath -Append
                if ($result.DnsError) {
                    "$timestamp -     DNS Note: $($result.DnsError)" | Out-File $LogPath -Append
                }
            }
            else {
                "$timestamp -     DNS:  FAILED - $($result.DnsError)" | Out-File $LogPath -Append
            }

            # TCP
            if ($result.TcpSuccess) {
                "$timestamp -     TCP:  OK (port $($result.Port) reachable)" | Out-File $LogPath -Append
            }
            else {
                "$timestamp -     TCP:  FAILED - $($result.TcpError)" | Out-File $LogPath -Append
            }
        }

        "$timestamp - ########################################" | Out-File $LogPath -Append
    }

    # Console status line
    $iterationCounter++
    $pingLabel    = if ($pingSuccess)       { "Ping=OK"   } else { "Ping=FAIL" }
    $anyDnsFailed = $domainResults | Where-Object { -not $_.DnsSuccess }
    $anyTcpFailed = $domainResults | Where-Object { -not $_.TcpSuccess }
    $dnsLabel     = if (-not $anyDnsFailed) { "DNS=OK"    } else { "DNS=FAIL"  }
    $tcpLabel     = if (-not $anyTcpFailed) { "TCP=OK"    } else { "TCP=FAIL"  }
    $logLabel     = if ($shouldLog)         { "[LOGGED]"  } else { ""          }
    Write-Host "`rIteration: $iterationCounter  $pingLabel  $dnsLabel  $tcpLabel  $logLabel                               " -NoNewline

    if ($maxAttempts -gt 0 -and $iterationCounter -ge $maxAttempts) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "`n$timestamp - ### - Max attempts reached ($maxAttempts). Exiting." | Out-File -FilePath $LogPath -Append
        Write-Host "`nMax attempts reached ($maxAttempts). Exiting." -ForegroundColor Yellow
        break
    }

    Start-Sleep -Seconds 5
}