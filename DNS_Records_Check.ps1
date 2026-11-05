# <#
# .SYNOPSIS
#     Checks whether every authoritative nameserver for a customer's Microsoft 365 mail host (*.mail.protection.outlook.com) returns the host, to diagnose Exclaimer mail deferrals caused by DNS inconsistencies.
#
# .DESCRIPTION
#     This script prompts for the Mail Routing Domain configured in Exclaimer (Settings > Mail Flow), then checks that domain's MX record on two public resolvers (1.1.1.1 and 8.8.8.8) and reports any problems, such as no MX record, resolvers that disagree, or an MX host with no A record. It then tests every Microsoft 365 mail host (*.mail.protection.outlook.com) the MX record points to.
#     It then discovers the nameservers for the mail.protection.outlook.com zone from three sources: the zone's own NS records (asked of 1.1.1.1 and 8.8.8.8), the parent delegation (found by following referrals from a root server), and the NS set claimed by each nameserver already found.
#     Every nameserver IP address is queried directly (raw DNS over UDP port 53, no recursion) for the host's A record three times in a row, then its AAAA record three times in a row. Each address has one row, with one result per column (A (Try 1) to AAAA (Try 3)).
#     Results are colour coded in the console (red = NOT FOUND (NXDOMAIN) at least once, yellow = results changed between tries or NO RECORD (NODATA), TIMEOUT, SERVFAIL, REFUSED or another error, green = OK on every try) and saved to one CSV file "NsCheck_<domain>_<HHmmss>.csv", with the MX record check at the top and the nameserver results below it in the user's Downloads folder (C:\Temp if Downloads is not found).
#     Use it when a customer reports deferred mail with "422 4.4.1 Unable to connect" or "Name or service not known" to a Microsoft 365 mail host, to see whether any nameserver says the host does not exist.
#
# .NOTES
#     Date: 26th October 2026
#     Version: 1.3.1
#
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
#
# .REQUIREMENTS
#     - Windows PowerShell 5.1+ or PowerShell 7 on Windows (uses Resolve-DnsName for discovery)
#     - Administrator rights are not required
#     - Internet connectivity, including outbound UDP port 53 to public resolvers (1.1.1.1 and 8.8.8.8), DNS root servers and Microsoft nameservers
#     - Write access to the Downloads folder or C:\Temp for the CSV results file
#     - A wide console window (about 190 characters) so the results table does not wrap; the CSV has the same data
#
# .VERSION
#     1.3.1
#         - Tells the user to keep the window open while it runs, shows progress during the slowest steps, and says when it's safe to close
#     1.3.0
#         - One results file: the MX record check (one row per MX host, with whether it was tested, plus any findings) is at the top, with the nameserver results below it
#     1.2.1
#         - Reports each MX problem once: one finding per MX host, and one table row per host when both resolvers return the same answer (the CSV keeps every resolver)
#     1.2.0
#         - Checks the MX record of the domain entered on both public resolvers and reports issues: no MX record, resolvers that disagree, MX hosts with no A record, equal-preference hosts where one has no A record, and non-Microsoft hosts
#         - Tests the MX host(s) found, instead of working the host name out from a naming pattern, and no longer asks to confirm or override the host name
#         - A mail host name (*.mail.protection.outlook.com) can still be entered directly, which skips the MX check
#         - Saves the MX check to MxCheck_<domain>_<HHmmss>.csv alongside the nameserver results
#     1.1.3
#         - Asks to confirm or override the derived mail host name, and records the host in every CSV row
#         - Tests use raw DNS queries and read the response code, so results do not depend on Windows error text or language
#         - Looks up each nameserver's addresses several times on both resolvers to catch rotating address sets
#         - Reports NOT FOUND (NXDOMAIN), NO RECORD (NODATA), TIMEOUT, SERVFAIL, REFUSED and other response codes
#         - All hardcoded values are in one configuration block at the top
#     1.0.0
#         - Derives the mail host name from the Mail Routing Domain (custom domain, onmicrosoft.com or mail.onmicrosoft.com)
#         - Discovers nameservers from the zone's NS records, the parent delegation and each server's own NS set
#         - Queries each nameserver IP address for A and AAAA, three consecutive tries each, one result per column
#         - Saves the results to a timestamped CSV in Downloads (falls back to C:\Temp)
#
# .INSTRUCTIONS
#     1. Open PowerShell (Administrator is not required)
#     2. Set execution policy, e.g. `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`
#     3. Navigate to script folder, e.g. `cd c:\temp`
#     4. Execute: `.\DNS_Records_Check.ps1`
#     5. Enter the Mail Routing Domain configured in Exclaimer, e.g. contoso.onmicrosoft.com
#     6. Keep the window open until it says "Full results saved to" (about a minute or two), then review the MX record check first and the nameserver results below it
#     7. Attach the CSV file (NsCheck_...) to the ticket
# >

& {
  # ---------------- CONFIGURATION: the only values hardcoded in this script ----------------
  $resolvers   = @('1.1.1.1', '8.8.8.8')                          # public resolvers, used only to find nameserver names and addresses
  $rootServers = @('198.41.0.4', '192.5.5.241', '193.0.14.129')   # root servers a, f and k, used only to follow the delegation
  $tries       = 3                                                # consecutive tries per record type
  $nsLookups   = 2                                                # address lookups per resolver for each nameserver name
  $timeoutMs   = 3000                                             # wait for each DNS reply
  # -----------------------------------------------------------------------------------------

  $entry = (Read-Host "Enter the Mail Routing Domain configured in Exclaimer (Settings > Mail Flow), e.g. contoso.onmicrosoft.com").Trim().TrimEnd('.').ToLower()
  $entryIsHost = ($entry -match '\.protection\.outlook\.com$')   # a mail host name can be entered directly; the MX check is then skipped
  if ($entryIsHost) { Write-Host "Mail host:           $entry" -ForegroundColor Cyan } else { Write-Host "Mail Routing Domain: $entry" -ForegroundColor Cyan }
  Write-Host "This usually takes a minute or two. Please keep this window open until it says 'Full results saved to'." -ForegroundColor Yellow

  # Output path: Downloads folder, falling back to C:\Temp (created if needed)
  $outDir = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
  if (-not (Test-Path -Path $outDir)) {
    $outDir = 'C:\Temp'
    if (-not (Test-Path -Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
  }
  $stamp     = Get-Date -Format 'HHmmss'
  $label     = $entry.Split('.')[0]
  $csvPath   = Join-Path $outDir "NsCheck_${label}_$stamp.csv"
  $runTime = (Get-Date).ToUniversalTime().ToString("ddd dd MMMM yyyy, HH:mm 'UTC'")
  Write-Host "Results will be saved to: $csvPath" -ForegroundColor Cyan

  $state = @{ Next = 0 }
  function Resolve-Alt($name, $type) {
    $first = $resolvers[$state.Next % $resolvers.Count]; $state.Next++
    $order = @($first) + @($resolvers | Where-Object { $_ -ne $first })
    foreach ($srv in $order) {
      $r = Resolve-DnsName $name -Type $type -Server $srv -DnsOnly -ErrorAction SilentlyContinue
      if ($r) { return [pscustomobject]@{ Resolver = $srv; Records = @($r) } }
    }
  }
  $src = @{}
  function Add-NS($name, $source) {
    $k = $name.TrimEnd('.').ToLower()
    if (-not $src.ContainsKey($k)) { $src[$k] = @() }
    if ($src[$k] -notcontains $source) { $src[$k] += $source }
  }

  # Raw DNS client (UDP/53, no recursion) so response codes and referral sections can be read
  function Read-DnsName($b, $pos) {
    $labels = New-Object System.Collections.Generic.List[string]
    $p = $pos; $next = -1; $guard = 0
    while ($guard++ -lt 64) {
      $len = [int]$b[$p]
      if ($len -eq 0) { $p++; break }
      if (($len -band 0xC0) -eq 0xC0) {
        if ($next -lt 0) { $next = $p + 2 }
        $p = (($len -band 0x3F) -shl 8) -bor [int]$b[$p + 1]
        continue
      }
      $labels.Add([System.Text.Encoding]::ASCII.GetString($b, $p + 1, $len))
      $p += 1 + $len
    }
    if ($next -lt 0) { $next = $p }
    [pscustomobject]@{ Name = ($labels -join '.').ToLower(); Next = $next }
  }
  function Send-DnsQuery([string]$Name, [int]$Type, [string]$Server, [bool]$Recurse = $false) {
    try {
      $id  = Get-Random -Minimum 1 -Maximum 65535
      $rdFlag = if ($Recurse) { 1 } else { 0 }
      $hdr = [byte[]]@([byte]($id -shr 8), [byte]($id -band 255), [byte]$rdFlag, 0, 0, 1, 0, 0, 0, 0, 0, 0)
      $ms  = New-Object System.IO.MemoryStream
      $ms.Write($hdr, 0, $hdr.Length)
      foreach ($label in $Name.TrimEnd('.').Split('.')) {
        $lb = [System.Text.Encoding]::ASCII.GetBytes($label)
        $ms.WriteByte([byte]$lb.Length); $ms.Write($lb, 0, $lb.Length)
      }
      $ms.WriteByte(0)
      $ms.WriteByte([byte]($Type -shr 8)); $ms.WriteByte([byte]($Type -band 255))
      $ms.WriteByte(0); $ms.WriteByte(1)
      $q   = $ms.ToArray()
      $udp = New-Object System.Net.Sockets.UdpClient
      $udp.Client.ReceiveTimeout = $timeoutMs
      $udp.Connect($Server, 53)
      [void]$udp.Send($q, $q.Length)
      $ep  = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
      $b   = $udp.Receive([ref]$ep)
      $udp.Close()
    } catch { return $null }
    $sections = @{ Answer = @(); Authority = @(); Additional = @() }
    try {
      $counts = @{ Answer = ([int]$b[6] -shl 8) + $b[7]; Authority = ([int]$b[8] -shl 8) + $b[9]; Additional = ([int]$b[10] -shl 8) + $b[11] }
      $pos = 12
      for ($i = 0; $i -lt (([int]$b[4] -shl 8) + $b[5]); $i++) { $n = Read-DnsName $b $pos; $pos = $n.Next + 4 }
      foreach ($secName in @('Answer', 'Authority', 'Additional')) {
        $list = @()
        for ($i = 0; $i -lt $counts[$secName]; $i++) {
          if ($pos + 10 -gt $b.Length) { break }
          $n = Read-DnsName $b $pos; $pos = $n.Next
          $rtype = ([int]$b[$pos] -shl 8) + $b[$pos + 1]
          $rdlen = ([int]$b[$pos + 8] -shl 8) + $b[$pos + 9]
          $rd = $pos + 10; $data = $null
          if     ($rtype -eq 2) { $data = (Read-DnsName $b $rd).Name }
          elseif ($rtype -eq 1) { $data = ($b[$rd..($rd + 3)] -join '.') }
          elseif ($rtype -eq 15) { $data = ('{0} {1}' -f (([int]$b[$rd] -shl 8) + $b[$rd + 1]), (Read-DnsName $b ($rd + 2)).Name) }
          $list += [pscustomobject]@{ Name = $n.Name; Type = $rtype; Data = $data }
          $pos = $rd + $rdlen
        }
        $sections[$secName] = $list
      }
    } catch { }
    [pscustomobject]@{ Rcode = ([int]$b[3] -band 15); Answer = $sections.Answer; Authority = $sections.Authority; Additional = $sections.Additional }
  }

  # Ask a public resolver (recursion on) for one record type; Status says what came back
  function Get-DnsResult([string]$name, [int]$type, [string]$server) {
    $resp = $null
    foreach ($attempt in 1..2) { $resp = Send-DnsQuery $name $type $server $true; if ($resp) { break } }
    if (-not $resp) { return [pscustomobject]@{ Status = 'TIMEOUT'; Records = @() } }
    $recs = @($resp.Answer | Where-Object { $_.Type -eq $type })
    $status = switch ($resp.Rcode) {
      0 { if ($recs.Count) { 'OK' } else { 'NO RECORD (NODATA)' } }
      3 { 'NOT FOUND (NXDOMAIN)' }
      2 { 'SERVFAIL' }
      5 { 'REFUSED' }
      default { "RCODE $($resp.Rcode)" }
    }
    [pscustomobject]@{ Status = $status; Records = $recs }
  }

  # One results file: the MX record check first, then the nameserver results (each section has its own header row)
  function Save-Results {
    $lines = New-Object System.Collections.Generic.List[string]
    if ((-not $entryIsHost) -and $mxExport) {
      $lines.Add("MX record check for $entry")
      foreach ($l in @($mxExport | ConvertTo-Csv -NoTypeInformation)) { $lines.Add([string]$l) }
      $lines.Add('')
    }
    if ($rows) {
      $lines.Add("Nameserver test for $($hostsToTest -join ', ')")
      foreach ($l in @($rows | ConvertTo-Csv -NoTypeInformation)) { $lines.Add([string]$l) }
    }
    try {
      [System.IO.File]::WriteAllLines($csvPath, $lines, (New-Object System.Text.UTF8Encoding($true)))
      Write-Host "`nFull results saved to $csvPath" -ForegroundColor Cyan
      Write-Host "Done. You can close this window now." -ForegroundColor Green
    } catch {
      Write-Host "Couldn't save the results to $csvPath : $($_.Exception.Message)" -ForegroundColor Red
    }
  }

  # MX record check for the domain that was entered (skipped when a mail host name was entered)
  $mxFindings = New-Object System.Collections.Generic.List[object]
  $mxRows     = New-Object System.Collections.Generic.List[object]
  function Add-Finding($level, $text) { $mxFindings.Add([pscustomobject]@{ Level = $level; Finding = $text }) }
  $hostsToTest = @()
  if ($entryIsHost) {
    $hostsToTest = @($entry)
  } else {
    Write-Host "`nMX record check for ${entry}:" -ForegroundColor Cyan
    $mxHostSet = [ordered]@{}     # MX host -> preference
    $okSets    = @()
    foreach ($srv in $resolvers) {
      $res = Get-DnsResult $entry 15 $srv
      if ($res.Status -ne 'OK') { Add-Finding 'ISSUE' "The MX lookup for $entry via $srv returned $($res.Status)."; continue }
      $recs = @($res.Records | Where-Object { $_.Data } | ForEach-Object { $p = $_.Data -split ' ', 2; [pscustomobject]@{ Pref = [int]$p[0]; Host = $p[1] } })
      $okSets += (($recs | Sort-Object Pref, Host | ForEach-Object { "$($_.Pref) $($_.Host)" }) -join ';')
      foreach ($m in $recs) { if (-not $mxHostSet.Contains($m.Host)) { $mxHostSet[$m.Host] = $m.Pref } }
    }
    if (@($okSets | Sort-Object -Unique).Count -gt 1) { Add-Finding 'ISSUE' "The resolvers returned different MX records for $entry." }
    if (-not $mxHostSet.Count) {
      Add-Finding 'ISSUE' "No MX records could be read for $entry, so there's no mail host to test. Check the spelling of the domain, or enter a mail host name directly."
    } else {
      $prefCount = @{}
      foreach ($mh in $mxHostSet.Keys) { $prefCount[[string]$mxHostSet[$mh]] = 1 + [int]$prefCount[[string]$mxHostSet[$mh]] }
      foreach ($mh in $mxHostSet.Keys) {
        $aStat = @{}; $v6Stat = @{}
        foreach ($srv in $resolvers) {
          $a  = Get-DnsResult $mh 1 $srv
          $v6 = Get-DnsResult $mh 28 $srv
          $aTxt  = if ($a.Status  -eq 'OK') { "OK($($a.Records.Count))" }  else { $a.Status }
          $v6Txt = if ($v6.Status -eq 'OK') { "OK($($v6.Records.Count))" } else { $v6.Status }
          $aStat[$srv] = $aTxt; $v6Stat[$srv] = $v6Txt
          $mxRows.Add([pscustomobject]@{ Resolver = $srv; Preference = $mxHostSet[$mh]; MxHost = $mh; A = $aTxt; AAAA = $v6Txt })
        }
        $aBad = @($resolvers | Where-Object { $aStat[$_] -notlike 'OK(*' })
        if ($aBad.Count) {
          $aTxts = @($aBad | ForEach-Object { $aStat[$_] } | Sort-Object -Unique)
          $where = if ($aBad.Count -eq @($resolvers).Count -and $aTxts.Count -eq 1) { "both resolvers: $($aTxts[0])" } else { ($aBad | ForEach-Object { "${_}: $($aStat[$_])" }) -join '; ' }
          $msg  = "MX host $mh has no A record ($where)."
          $same = $prefCount[[string]$mxHostSet[$mh]] - 1
          if ($same -gt 0) { $msg += " It shares preference $($mxHostSet[$mh]) with $same other MX host(s), and a sender can pick any host with the same preference (standard MX behaviour), so this can cause intermittent failures." }
          $msg += " The nameserver test checks every nameserver found for this host."
          Add-Finding 'ISSUE' $msg
        } else {
          $v6Bad = @($resolvers | Where-Object { $v6Stat[$_] -notlike 'OK(*' })
          if ($v6Bad.Count) {
            $v6Txts = @($v6Bad | ForEach-Object { $v6Stat[$_] } | Sort-Object -Unique)
            $where6 = if ($v6Bad.Count -eq @($resolvers).Count -and $v6Txts.Count -eq 1) { "both resolvers: $($v6Txts[0])" } else { ($v6Bad | ForEach-Object { "${_}: $($v6Stat[$_])" }) -join '; ' }
            Add-Finding 'NOTE' "MX host $mh has no AAAA (IPv6) record ($where6). That's fine for a host with IPv4 addresses only."
          }
        }
        if ($mh -notlike '*.mail.protection.outlook.com') { Add-Finding 'NOTE' "MX host $mh isn't a Microsoft 365 mail host (*.mail.protection.outlook.com), so the nameserver test skips it." }
      }
      if ($entry -match '^[^.]+\.onmicrosoft\.com$') {
        $expected = $entry.Split('.')[0] + '.mail.protection.outlook.com'
        if (@($mxHostSet.Keys) -notcontains $expected) { Add-Finding 'NOTE' "None of the MX hosts is $expected, the host name this kind of domain normally has (an observed pattern, not a guarantee)." }
      }
    }

    Write-Host ("  {0,-9} {1,4}  {2,-58} {3,-24} {4}" -f 'Resolver', 'Pref', 'MX host', 'A', 'AAAA')
    $collapsed = $false
    foreach ($mh in $mxHostSet.Keys) {
      $hr = @($mxRows | Where-Object { $_.MxHost -eq $mh })
      if (@($hr | ForEach-Object { "$($_.A)|$($_.AAAA)" } | Sort-Object -Unique).Count -eq 1) {
        $c = if ($hr[0].A -like 'OK(*') { 'Green' } else { 'Red' }
        Write-Host ("  {0,-9} {1,4}  {2,-58} {3,-24} {4}" -f 'both', $hr[0].Preference, $mh, $hr[0].A, $hr[0].AAAA) -ForegroundColor $c
        $collapsed = $true
      } else {
        foreach ($m in $hr) {
          $c = if ($m.A -like 'OK(*') { 'Green' } else { 'Red' }
          Write-Host ("  {0,-9} {1,4}  {2,-58} {3,-24} {4}" -f $m.Resolver, $m.Preference, $m.MxHost, $m.A, $m.AAAA) -ForegroundColor $c
        }
      }
    }
    if ($collapsed) { Write-Host "  (both = $($resolvers -join ' and ') returned the same answer)" -ForegroundColor DarkGray }
    if ($mxFindings.Count) {
      Write-Host ''
      foreach ($f in $mxFindings) { Write-Host ("  [{0}] {1}" -f $f.Level, $f.Finding) -ForegroundColor $(if ($f.Level -eq 'ISSUE') { 'Red' } else { 'Yellow' }) }
    } else {
      Write-Host "`n  No issues found with the MX record." -ForegroundColor Green
    }

    $mxExport = @()
    foreach ($mh in $mxHostSet.Keys) {
      $hr   = @($mxRows | Where-Object { $_.MxHost -eq $mh })
      $same = (@($hr | ForEach-Object { "$($_.A)|$($_.AAAA)" } | Sort-Object -Unique).Count -eq 1)
      $use  = if ($same) { @($hr[0]) } else { $hr }
      foreach ($m in $use) {
        $mxExport += [pscustomobject]@{ Domain = $entry; Section = 'MX record'; Resolver = $(if ($same) { 'both' } else { $m.Resolver }); Preference = $m.Preference; MxHost = $mh; A = $m.A; AAAA = $m.AAAA; Tested = $(if ($mh -like '*.mail.protection.outlook.com') { 'Yes' } else { 'No' }); Level = ''; Finding = '' }
      }
    }
    foreach ($f in $mxFindings) { $mxExport += [pscustomobject]@{ Domain = $entry; Section = 'Finding'; Resolver = ''; Preference = ''; MxHost = ''; A = ''; AAAA = ''; Tested = ''; Level = $f.Level; Finding = $f.Finding } }
    $hostsToTest = @($mxHostSet.Keys | Where-Object { $_ -like '*.mail.protection.outlook.com' })
  }
  if (-not $hostsToTest.Count) {
    Write-Host "`nThere's no Microsoft 365 mail host to test, so the nameserver test was skipped." -ForegroundColor Yellow
    Save-Results
    return
  }
  $hostName = $hostsToTest[0]
  Write-Host "`nMail host(s) to test: $($hostsToTest -join ', ')" -ForegroundColor Cyan

  # Source 1: the zone's own NS records, asked of BOTH resolvers
  $labels = $hostName.Split('.'); $zone = $null; $zoneSets = @{}
  for ($i = 0; $i -lt $labels.Count - 1 -and -not $zone; $i++) {
    $cand = ($labels[$i..($labels.Count - 1)]) -join '.'
    foreach ($srv in $resolvers) {
      $r  = Resolve-DnsName $cand -Type NS -Server $srv -DnsOnly -ErrorAction SilentlyContinue
      $ns = @($r | Where-Object { $_.Type -eq 'NS' -and $_.Name.TrimEnd('.') -eq $cand })
      if ($ns.Count) {
        $zone = $cand
        $zoneSets[$srv] = @($ns | ForEach-Object { $_.NameHost.TrimEnd('.').ToLower() } | Sort-Object -Unique)
        $ns | ForEach-Object { Add-NS $_.NameHost "zone($srv)" }
      }
    }
  }
  if (-not $zone) { Write-Host "Couldn't find the zone for $hostName via $($resolvers -join ' or ')." -ForegroundColor Red; return }
  Write-Host "Zone: $zone" -ForegroundColor Cyan
  foreach ($srv in $resolvers) {
    $n = if ($zoneSets[$srv]) { $zoneSets[$srv].Count } else { 'none' }
    Write-Host "  Zone NS records returned by ${srv}: $n"
  }
  if (($zoneSets[$resolvers[0]] -join ',') -ne ($zoneSets[$resolvers[1]] -join ',')) {
    Write-Host "  Note: the two resolvers returned DIFFERENT NS sets for the zone." -ForegroundColor Yellow
  }

  # Source 2: the parent delegation, found by following referrals from a root server (like dig +trace)
  function Get-Delegation($zoneName) {
    foreach ($root in $rootServers) {
      $server = $root
      for ($hop = 1; $hop -le 8; $hop++) {
        $resp = Send-DnsQuery $zoneName 2 $server
        if (-not $resp) { Write-Host "  hop ${hop}: no reply from $server"; break }
        $auth = @($resp.Authority | Where-Object { $_.Type -eq 2 })
        $ans  = @($resp.Answer    | Where-Object { $_.Type -eq 2 })
        $forName = if ($auth.Count) { " for '" + $auth[0].Name + "'" } else { '' }
        Write-Host ("  hop {0}: {1} returned {2} answer NS and {3} referral NS{4}" -f $hop, $server, $ans.Count, $auth.Count, $forName)
        if ($auth.Count -and $auth[0].Name -eq $zoneName) {
          return @($auth | ForEach-Object { $_.Data } | Sort-Object -Unique)
        }
        if (-not $auth.Count) { break }
        $cand = $auth[0].Data
        $glue = ($resp.Additional | Where-Object { $_.Type -eq 1 -and $_.Name -eq $cand } | Select-Object -First 1).Data
        if (-not $glue) { $glue = ((Resolve-Alt $cand 'A').Records | Where-Object Type -eq 'A' | Select-Object -First 1).IPAddress }
        if (-not $glue) { Write-Host "  couldn't resolve $cand"; break }
        $server = $glue
      }
    }
    @()
  }
  Write-Host "`nFollowing the delegation for $zone from the root servers:" -ForegroundColor Cyan
  $parentNS = @(Get-Delegation $zone)
  $parentNS | ForEach-Object { Add-NS $_ 'parent' }
  Write-Host "Parent delegation: $($parentNS.Count) nameservers"
  if (-not $parentNS.Count) {
    Write-Host "Note: couldn't find the parent delegation, so servers listed only there won't be tested." -ForegroundColor Yellow
  }

  # Source 3: ask each nameserver found so far which NS set it claims for the zone, and add any new names
  $claims = @{}
  for ($pass = 1; $pass -le 3; $pass++) {
    foreach ($n in @($src.Keys)) {
      if ($claims.ContainsKey($n)) { continue }
      $ip = ((Resolve-Alt $n 'A').Records | Where-Object Type -eq 'A' | Select-Object -First 1).IPAddress
      if (-not $ip) { continue }
      $r   = Resolve-DnsName $zone -Type NS -Server $ip -DnsOnly -NoRecursion -ErrorAction SilentlyContinue
      $set = @($r | Where-Object { $_.Type -eq 'NS' -and $_.Name.TrimEnd('.') -eq $zone } |
               ForEach-Object { $_.NameHost.TrimEnd('.').ToLower() } | Sort-Object -Unique)
      $claims[$n] = $set
      foreach ($s in $set) { Add-NS $s 'claimed' }
    }
  }
  Write-Host "`nNS set claimed by each server (these should all be identical):" -ForegroundColor Cyan
  foreach ($k in ($claims.Keys | Sort-Object)) { "  {0,-46} claims {1} servers" -f $k, $claims[$k].Count }
  $distinctSets = @($claims.Values | ForEach-Object { $_ -join ',' } | Sort-Object -Unique).Count
  if ($distinctSets -gt 1) { Write-Host "  Note: the servers disagree about the zone's NS set ($distinctSets different sets)." -ForegroundColor Yellow }

  Write-Host "`nNameservers to test: $($src.Count)  (source: zone(resolver) = zone NS records, parent = parent delegation, claimed = listed by another server)" -ForegroundColor Cyan
  $src.Keys | Sort-Object | ForEach-Object { "  {0,-46} {1}" -f $_, ($src[$_] -join '+') }
  Write-Host "`nNow looking up each nameserver's addresses and testing every one of them. This is the slowest step, so please keep this window open and wait for the results." -ForegroundColor Yellow

  # Resolve every nameserver name to its IP addresses: several lookups on both resolvers, combined
  $nsNames = @($src.Keys | Sort-Object); $nsDone = 0
  $targets = foreach ($n in $nsNames) {
    Write-Progress -Activity 'Looking up nameserver addresses' -Status $n -PercentComplete ([int](100 * $nsDone / $nsNames.Count)); $nsDone++
    $ips = @()
    foreach ($srv in $resolvers) {
      foreach ($i in 1..$nsLookups) {
        $r = Resolve-DnsName $n -Type A -Server $srv -DnsOnly -ErrorAction SilentlyContinue
        $ips += @($r | Where-Object Type -eq 'A' | ForEach-Object { $_.IPAddress })
      }
    }
    $ips = @($ips | Sort-Object -Unique)
    if (-not $ips.Count) { Write-Host "  Couldn't resolve the address of $n, so it won't be tested." -ForegroundColor Yellow }
    foreach ($ip in $ips) { [pscustomobject]@{ NS = $n; IP = $ip } }
  }
  Write-Progress -Activity 'Looking up nameserver addresses' -Completed
  if (-not @($targets).Count) { Write-Host "No nameserver addresses to test." -ForegroundColor Red; return }

  # One query to one nameserver address; the result comes from the DNS response code
  function Q([string]$h, [int]$t, [string]$ip) {
    $resp = $null
    foreach ($attempt in 1..2) { $resp = Send-DnsQuery $h $t $ip; if ($resp) { break } }
    if (-not $resp) { return 'TIMEOUT' }
    switch ($resp.Rcode) {
      0 {
        $n = @($resp.Answer | Where-Object { $_.Type -eq $t }).Count
        if ($n) { return "OK($n)" } else { return 'NO RECORD (NODATA)' }
      }
      3 { return 'NOT FOUND (NXDOMAIN)' }
      2 { return 'SERVFAIL' }
      5 { return 'REFUSED' }
      default { return "RCODE $($resp.Rcode)" }
    }
  }

  # Each address: $tries tries in a row for A, then for AAAA; one result per column
  $typeNum    = @{ 'A' = 1; 'AAAA' = 28 }
  $cols  = @(foreach ($type in 'A', 'AAAA') { foreach ($n in 1..$tries) { "$type (Try $n)" } })
  $testTotal = $hostsToTest.Count * @($targets).Count; $testDone = 0
  $rows  = foreach ($h in $hostsToTest) { foreach ($t in $targets) {
    Write-Progress -Activity 'Testing each nameserver address' -Status "$h on $($t.NS) ($($t.IP))" -PercentComplete ([int](100 * $testDone / $testTotal)); $testDone++
    $r = [ordered]@{ Host = $h; Nameserver = $t.NS; IP = $t.IP }
    foreach ($type in 'A', 'AAAA') {
      $tn = $typeNum[$type]
      foreach ($n in 1..$tries) { $r["$type (Try $n)"] = Q $h $tn $t.IP }
    }
    $r['Source']    = ($src[$t.NS] -join '+')
    $r['TestedUTC'] = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    [pscustomobject]$r
  } }
  Write-Progress -Activity 'Testing each nameserver address' -Completed

  function Get-Cells($row) { @($cols | ForEach-Object { [string]$row.PSObject.Properties[$_].Value }) }
  $allCells = @(foreach ($r in $rows) { Get-Cells $r })
  $w   = [Math]::Max(12, ($allCells | Measure-Object -Property Length -Maximum).Maximum)
  $hdr = 'Nameserver'.PadRight(41) + ' ' + 'IP'.PadRight(15) + ' ' + (($cols | ForEach-Object { $_.PadRight($w) }) -join ' ')
  try {
    if ($hdr.Length -gt $Host.UI.RawUI.WindowSize.Width) {
      Write-Host "`nNote: the results table is $($hdr.Length) characters wide; widen or maximise the window to avoid wrapping. The CSV has the same data." -ForegroundColor Yellow
    }
  } catch { }
  foreach ($h in $hostsToTest) {
    $hostRows = @($rows | Where-Object { $_.Host -eq $h })
    "`nHost: $h   Run: $runTime   Tries per record type: $tries (back to back)"
    $hdr
    foreach ($r in ($hostRows | Sort-Object Nameserver, IP)) {
      $vals = Get-Cells $r
      $aSet = @($vals[0..($tries - 1)] | Sort-Object -Unique).Count
      $bSet = @($vals[$tries..(2 * $tries - 1)] | Sort-Object -Unique).Count
      $text = $vals -join ' '
      if     ($text -match 'NOT FOUND')   { $colour = 'Red' }
      elseif ($aSet -gt 1 -or $bSet -gt 1 -or $text -match 'NO RECORD|TIMEOUT|SERVFAIL|REFUSED|RCODE|ERR') { $colour = 'Yellow' }
      else   { $colour = 'Green' }
      $line = $r.Nameserver.PadRight(41) + ' ' + $r.IP.PadRight(15) + ' ' + (($vals | ForEach-Object { $_.PadRight($w) }) -join ' ')
      Write-Host $line -ForegroundColor $colour
    }
    $badRows  = @($hostRows | Where-Object { ((Get-Cells $_) -join ' ') -match 'NOT FOUND' })
    $badNames = @($badRows | Select-Object -ExpandProperty Nameserver -Unique).Count
    $allNames = @($hostRows | Select-Object -ExpandProperty Nameserver -Unique).Count
    Write-Host "`n$badNames of $allNames nameservers ($($badRows.Count) of $($hostRows.Count) addresses) returned NOT FOUND (NXDOMAIN) at least once for $h." -ForegroundColor Cyan
  }

  Save-Results
}