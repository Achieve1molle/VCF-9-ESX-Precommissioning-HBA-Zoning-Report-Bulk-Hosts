<#
.SYNOPSIS
ESXi Fibre Channel HBA / Storage Zoning Collector UI - Production

.DESCRIPTION
Production WPF UI for collecting Fibre Channel HBA, WWPN/WWNN, storage path, NAA/device,
adapter, raw command, and hardware information from a CSV list of ESXi hosts.

Design goals:
- Match the original black/dark-gray UI theme used by the VCF clone UI example.
- Direct-connect to each ESXi host using root credentials.
- Assume SSH may be disabled; enable SSH for collection only when needed and disable it afterward only if this script enabled it.
- Export a storage-admin-friendly Excel workbook.
- Preserve large WWN decimal values as text so Excel does not display values as 2.37817E+18.
- Sanitize Excel output to prevent sharedStrings.xml repair prompts.

.NOTES
Author: Michael Molle / Generated with M365 Copilot
Version: 1.4 Production
Date: 2026-05-14

.PREREQUISITES
- PowerShell 7+
- Windows PowerShell host capable of WPF/STA
- VCF.PowerCLI or VMware.PowerCLI
- Posh-SSH
- ImportExcel
- CSV file with a Host column

.EXAMPLE
pwsh -NoProfile -ExecutionPolicy Bypass -STA -File .\ESXi_HBA_Zoning_Collector_UI_v1.4_Production.ps1
#>
[CmdletBinding()]
param(
    [switch]$NoRelaunch,
    [switch]$SignedOk,
    [switch]$NoAutoSign
)

$Global:EsxiHbaUiVersion = '1.4 Production'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

function Ensure-SelfSigned {
    param([string]$TargetPath)
    try { $sig = Get-AuthenticodeSignature -FilePath $TargetPath -ErrorAction SilentlyContinue } catch { $sig = $null }
    if ($sig -and $sig.Status -eq 'Valid') { return $false }
    Write-Host '[SelfSign] Creating/trusting a local code-signing certificate and signing the script...'
    $subject = "CN=ESXiHBACollectorUI Local Code Signing ($env:USERNAME@$env:COMPUTERNAME)"
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -like 'CN=ESXiHBACollectorUI Local Code Signing*' } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
    if (-not $cert) {
        $cert = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject $subject `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -KeyAlgorithm RSA `
            -KeyLength 3072 `
            -HashAlgorithm SHA256 `
            -KeyExportPolicy Exportable `
            -NotAfter (Get-Date).AddYears(5)
    }
    foreach ($store in 'Cert:\CurrentUser\Root','Cert:\CurrentUser\TrustedPublisher') {
        try { $null = $cert | Copy-Item -Destination $store -Force -ErrorAction SilentlyContinue } catch {}
    }
    $null = Set-AuthenticodeSignature -FilePath $TargetPath -Certificate $cert -ErrorAction Stop
    Write-Host '[SelfSign] Script signed.'
    return $true
}

try { $pwsh = (Get-Process -Id $PID -ErrorAction SilentlyContinue).Path } catch { $pwsh = $null }
if (-not $pwsh) { $pwsh = 'pwsh.exe' }
if (-not $NoAutoSign -and -not $SignedOk -and $PSCommandPath) {
    try { $null = Ensure-SelfSigned -TargetPath $PSCommandPath } catch { Write-Host "[SelfSign] Skipped/failed: $($_.Exception.Message)" }
    & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "$PSCommandPath" -SignedOk -NoRelaunch
    exit $LASTEXITCODE
}
if (-not $NoRelaunch -and $PSCommandPath) {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File "$PSCommandPath" -NoRelaunch -SignedOk
        exit $LASTEXITCODE
    }
}

$script:RunDir = $null
$Global:LogFile = $null
$script:CancelRequested = $false
$script:IsExecuting = $false

function New-RunDir {
    param([string]$Base = (Get-Location).Path)
    if ([string]::IsNullOrWhiteSpace($Base) -or -not (Test-Path $Base)) { $Base = (Get-Location).Path }
    $dir = Join-Path $Base ("ESXi-HBA-Zoning-Run-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $Global:LogFile = Join-Path $dir ("ESXi-HBA-Zoning-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
    '' | Out-File -FilePath $Global:LogFile -Encoding utf8 -Force
    $script:RunDir = $dir
    return $dir
}

function Write-Log {
    param([Parameter(Mandatory)][string]$Message,[ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO')
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    try { if ($Global:LogFile) { Add-Content -Path $Global:LogFile -Value $line -Encoding utf8 -ErrorAction SilentlyContinue } } catch {}
    try {
        if ($script:txtLog) {
            $script:txtLog.AppendText($line + [Environment]::NewLine)
            $script:txtLog.ScrollToEnd()
        }
    } catch {}
    Write-Host $line
}

function Pump-Ui {
    try {
        $frame = New-Object Windows.Threading.DispatcherFrame
        $null = $script:window.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $frame.Continue = $false })
        [Windows.Threading.Dispatcher]::PushFrame($frame)
    } catch {}
}

function Update-ProgressUi {
    param([double]$Percent,[string]$Text,[switch]$Indeterminate)
    try {
        if ($script:pbExec) {
            $script:pbExec.IsIndeterminate = [bool]$Indeterminate
            if (-not $Indeterminate) {
                if ($Percent -lt 0) { $Percent = 0 }
                if ($Percent -gt 100) { $Percent = 100 }
                $script:pbExec.Value = $Percent
            }
        }
        if ($script:lblProgress) { $script:lblProgress.Text = $Text }
    } catch {}
}

function Has-Module { param([string]$Name) return [bool](Get-Module -ListAvailable -Name $Name | Select-Object -First 1) }
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (Has-Module $Name) { try { Import-Module $Name -ErrorAction Stop | Out-Null; return $true } catch { return $false } }
    try {
        $old = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue | Out-Null
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -AcceptLicense -ErrorAction Stop
        Import-Module $Name -ErrorAction Stop | Out-Null
        Write-Log "$Name installed/imported."
        return $true
    } catch {
        Write-Log "$Name install/import failed: $($_.Exception.Message)" 'ERROR'
        return $false
    } finally { $ProgressPreference = $old }
}
function Ensure-PowerCliModule {
    if (Ensure-Module -Name 'VCF.PowerCLI') { return $true }
    Write-Log 'VCF.PowerCLI was not available; trying VMware.PowerCLI...' 'WARN'
    return (Ensure-Module -Name 'VMware.PowerCLI')
}
function Set-PowerCliCertBehavior {
    try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope User | Out-Null } catch {
        try { Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null } catch {}
    }
}
function Set-StatusText {
    param([System.Windows.Controls.TextBlock]$Label,[string]$Text,[string]$State)
    if (-not $Label) { return }
    $Label.Text = $Text
    switch ($State) {
        'OK'   { $Label.Foreground = [Windows.Media.Brushes]::LightGreen }
        'WARN' { $Label.Foreground = [Windows.Media.Brushes]::Gold }
        'FAIL' { $Label.Foreground = [Windows.Media.Brushes]::Tomato }
        default { $Label.Foreground = [Windows.Media.Brushes]::White }
    }
}

function Convert-WwnDecimalToFullString {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try { return ([UInt64]$Value).ToString('0') } catch { return [string]$Value }
}
function Convert-WwnDecimalToHexColon {
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        $u = [UInt64]$Value
        $hex = $u.ToString('x16')
        return (($hex -split '(.{2})' | Where-Object { $_ }) -join ':')
    } catch { return [string]$Value }
}

function ConvertTo-ExcelSafeString {
    param([AllowNull()]$Value,[int]$MaxLength = 32000)
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    $s = $s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    $esc = [char]27
    $s = $s -replace ([regex]::Escape([string]$esc) + '\[[0-9;?]*[ -/]*[@-~]'), ''
    if ($s.Length -gt $MaxLength) {
        $suffix = "`r`n...[TRUNCATED for Excel cell safety. See Raw_Commands OutputFile column for full command output.]"
        $keep = [Math]::Max(0, $MaxLength - $suffix.Length)
        $s = $s.Substring(0, $keep) + $suffix
    }
    return $s
}
function ConvertTo-ExcelSafeData {
    param([Parameter(Mandatory)][object[]]$Data)
    $safe = @()
    foreach ($row in @($Data)) {
        if ($null -eq $row) { continue }
        $o = [ordered]@{}
        foreach ($p in $row.PSObject.Properties) {
            $v = $p.Value
            if ($null -eq $v) { $o[$p.Name] = $null }
            elseif ($v -is [string]) { $o[$p.Name] = ConvertTo-ExcelSafeString -Value $v }
            elseif ($v -is [System.Array]) { $o[$p.Name] = ConvertTo-ExcelSafeString -Value (($v | ForEach-Object { [string]$_ }) -join '; ') }
            elseif ($v -is [System.ValueType] -or $v -is [datetime]) { $o[$p.Name] = $v }
            else { $o[$p.Name] = ConvertTo-ExcelSafeString -Value $v }
        }
        $safe += [pscustomobject]$o
    }
    return @($safe)
}
function Get-SafeFileNamePart {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'blank' }
    $s = $Value -replace '[\\/:*?"<>|\s]+','_'
    $s = $s -replace '[^a-zA-Z0-9_.-]','_'
    if ($s.Length -gt 90) { $s = $s.Substring(0,90) }
    return $s.Trim('_')
}
function Save-RawCommandOutput {
    param([Parameter(Mandatory)][string]$HostName,[Parameter(Mandatory)][string]$Command,[AllowNull()][string]$Text)
    try {
        if (-not $script:RunDir) { return '' }
        $rawDir = Join-Path $script:RunDir 'RawCommandOutput'
        New-Item -ItemType Directory -Path $rawDir -Force | Out-Null
        $file = Join-Path $rawDir ("{0}-{1}-{2}.txt" -f (Get-SafeFileNamePart $HostName),(Get-SafeFileNamePart $Command),(Get-Date -Format 'yyyyMMdd-HHmmssfff'))
        Set-Content -Path $file -Value (ConvertTo-ExcelSafeString -Value $Text -MaxLength ([int]::MaxValue)) -Encoding utf8 -Force
        return $file
    } catch {
        Write-Log "[$HostName] Could not save raw command output for [$Command]: $($_.Exception.Message)" 'WARN'
        return ''
    }
}

function Invoke-SshCommandText {
    param([Parameter(Mandatory)][string]$HostName,[Parameter(Mandatory)][pscredential]$Credential,[Parameter(Mandatory)][string]$Command)
    $session = $null
    try {
        $session = New-SSHSession -ComputerName $HostName -Credential $Credential -AcceptKey -ConnectionTimeout 25 -ErrorAction Stop
        $out = Invoke-SSHCommand -SessionId $session.SessionId -Command $Command -TimeOut 180 -ErrorAction Stop
        return (($out.Output + $out.Error) -join [Environment]::NewLine)
    } finally {
        if ($session) { try { Remove-SSHSession -SessionId $session.SessionId -ErrorAction SilentlyContinue | Out-Null } catch {} }
    }
}
function Convert-EsxCliKeyValueOutput {
    param([string]$Text,[string]$HostName,[string]$CommandName)
    $rows = @(); if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $current = [ordered]@{ Host = $HostName; SourceCommand = $CommandName }
    foreach ($line in ($Text -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($current.Keys.Count -gt 2) { $rows += [pscustomobject]$current; $current = [ordered]@{ Host=$HostName; SourceCommand=$CommandName } }
            continue
        }
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)\s*$') {
            $key = ($Matches[1] -replace '[^a-zA-Z0-9_]+','_').Trim('_')
            if ([string]::IsNullOrWhiteSpace($key)) { $key = 'Field' }
            if ($current.Contains($key)) { $key = "$key`_2" }
            $current[$key] = $Matches[2].Trim()
        }
    }
    if ($current.Keys.Count -gt 2) { $rows += [pscustomobject]$current }
    return @($rows)
}
function Convert-AdapterListOutput {
    param([string]$Text,[string]$HostName)
    $rows = @(); if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    foreach ($line in @($Text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if ($line -match '^HBA\s+Name' -or $line -match '^-{3,}') { continue }
        $parts = @($line -split '\s{2,}' | Where-Object { $_ -ne '' })
        if ($parts.Count -ge 1 -and $parts[0] -match '^vmhba') {
            $rows += [pscustomobject]@{
                Host=$HostName; HBA_Name=$parts[0]
                Driver=$(if ($parts.Count -gt 1) { $parts[1] } else { '' })
                Link_State=$(if ($parts.Count -gt 2) { $parts[2] } else { '' })
                UID=$(if ($parts.Count -gt 3) { $parts[3] } else { '' })
                Capabilities=$(if ($parts.Count -gt 4) { $parts[4] } else { '' })
                Description=$(if ($parts.Count -gt 5) { ($parts[5..($parts.Count-1)] -join ' ') } else { '' })
                SourceCommand='esxcli storage core adapter list'
            }
        }
    }
    return @($rows)
}
function Read-HostCsv {
    param([string]$CsvPath,[string]$HostColumn)
    if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path $CsvPath)) { throw 'Select a valid CSV file.' }
    $rows = @(Import-Csv -Path $CsvPath)
    if ($rows.Count -lt 1) { throw 'CSV did not contain any rows.' }
    $columns = @($rows[0].PSObject.Properties.Name)
    if ([string]::IsNullOrWhiteSpace($HostColumn)) { $HostColumn = 'Host' }
    if ($columns -notcontains $HostColumn) {
        if ($columns -contains 'Hostname') { $HostColumn = 'Hostname' }
        elseif ($columns -contains 'IP') { $HostColumn = 'IP' }
        elseif ($columns -contains 'FQDN') { $HostColumn = 'FQDN' }
        else { throw "CSV must contain a [$HostColumn] column. Found: $($columns -join ', ')" }
    }
    return @($rows | ForEach-Object { ([string]$_.($HostColumn)).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-EsxiHbaDataForHost {
    param([Parameter(Mandatory)][string]$HostName,[Parameter(Mandatory)][pscredential]$Credential,[switch]$LeaveSshEnabled)
    $vi=$null; $vmhost=$null; $sshStartedByScript=$false
    $summary=@(); $adapters=@(); $adapterList=@(); $fcDetails=@(); $paths=@(); $devices=@(); $hardware=@(); $raw=@()
    try {
        Write-Log "[$HostName] Connecting directly with PowerCLI..."
        $vi = Connect-VIServer -Server $HostName -Credential $Credential -WarningAction SilentlyContinue -ErrorAction Stop
        $vmhost = Get-VMHost -Server $vi -ErrorAction Stop | Select-Object -First 1
        $svc = Get-VMHostService -VMHost $vmhost -ErrorAction SilentlyContinue | Where-Object { $_.Key -eq 'TSM-SSH' } | Select-Object -First 1
        if ($svc -and -not $svc.Running) {
            Write-Log "[$HostName] SSH is disabled; enabling TSM-SSH for collection."
            Start-VMHostService -HostService $svc -Confirm:$false -ErrorAction Stop | Out-Null
            $sshStartedByScript=$true; Start-Sleep -Seconds 2
        } elseif ($svc -and $svc.Running) { Write-Log "[$HostName] SSH already running; leaving current state unchanged." }
        else { Write-Log "[$HostName] Could not determine SSH service state; SSH commands may fail." 'WARN' }

        Write-Log "[$HostName] Collecting PowerCLI HBA and hardware data..."
        $vmhbas = @(Get-VMHostHba -VMHost $vmhost -Type FibreChannel -ErrorAction SilentlyContinue)
        if ($vmhbas.Count -eq 0) {
            $vmhbas = @(Get-VMHostHba -VMHost $vmhost -ErrorAction SilentlyContinue | Where-Object {
                ($_.Type -match 'Fibre|FibreChannel|FC') -or (($_.Device -match '^vmhba') -and ($_.PortWorldWideName -or $_.NodeWorldWideName))
            })
        }
        if ($vmhbas.Count -eq 0) { Write-Log "[$HostName] No Fibre Channel HBAs were returned by PowerCLI. SSH raw sheets may still contain adapter details." 'WARN' }
        foreach ($hba in $vmhbas) {
            $adapters += [pscustomobject]@{
                Host=$vmhost.Name; Adapter=$hba.Device; Type=$hba.Type; Status=$hba.Status; Model=$hba.Model; Driver=$hba.Driver; Pci=$hba.Pci
                WWPN_Decimal=(Convert-WwnDecimalToFullString $hba.PortWorldWideName); WWPN=(Convert-WwnDecimalToHexColon $hba.PortWorldWideName)
                WWNN_Decimal=(Convert-WwnDecimalToFullString $hba.NodeWorldWideName); WWNN=(Convert-WwnDecimalToHexColon $hba.NodeWorldWideName)
                SpeedGb=$hba.Speed
            }
        }

        $hw=$vmhost.ExtensionData.Hardware; $sum=$vmhost.ExtensionData.Summary; $cpuModel=''
        try { if ($hw.CpuPkg -and @($hw.CpuPkg).Count -gt 0) { $cpuModel=$hw.CpuPkg[0].Description } } catch {}
        $hardware += [pscustomobject]@{
            Host=$vmhost.Name; ConnectionState=$vmhost.ConnectionState; PowerState=$vmhost.PowerState
            Manufacturer=$hw.SystemInfo.Vendor; Model=$hw.SystemInfo.Model; SerialNumber=[string]$hw.SystemInfo.SerialNumber; UUID=[string]$hw.SystemInfo.Uuid
            BiosVersion=$hw.BiosInfo.BiosVersion; BiosReleaseDate=$hw.BiosInfo.ReleaseDate; CpuModel=$cpuModel
            CpuPackages=@($hw.CpuPkg).Count; CpuCores=$hw.CpuInfo.NumCpuCores; CpuThreads=$hw.CpuInfo.NumCpuThreads
            MemoryGB=[Math]::Round(($hw.MemorySize/1GB),2); ESXiVersion=$vmhost.Version; ESXiBuild=[string]$vmhost.Build; OverallStatus=$sum.OverallStatus
        }

        $commands=[ordered]@{
            AdapterList='esxcli storage core adapter list'; SanFcList='esxcli storage san fc list'; PathList='esxcli storage core path list'
            DeviceList='esxcli storage core device list'; ScsiDevsA='esxcfg-scsidevs -a'; VmwareVl='vmware -vl'; DmiInfo='vsish -e get /hardware/bios/dmiInfo'
        }
        foreach ($kv in $commands.GetEnumerator()) {
            try {
                Write-Log "[$HostName] SSH: $($kv.Value)"
                $cmdText = Invoke-SshCommandText -HostName $HostName -Credential $Credential -Command $kv.Value
                $rawPath = Save-RawCommandOutput -HostName $vmhost.Name -Command $kv.Value -Text $cmdText
                $raw += [pscustomobject]@{ Host=$vmhost.Name; Command=$kv.Value; OutputFile=$rawPath; OutputPreview=$cmdText }
                switch ($kv.Key) {
                    'AdapterList' { $adapterList += Convert-AdapterListOutput -Text $cmdText -HostName $vmhost.Name }
                    'SanFcList'   { $fcDetails += Convert-EsxCliKeyValueOutput -Text $cmdText -HostName $vmhost.Name -CommandName $kv.Value }
                    'PathList'    { $paths += Convert-EsxCliKeyValueOutput -Text $cmdText -HostName $vmhost.Name -CommandName $kv.Value }
                    'DeviceList'  { $devices += Convert-EsxCliKeyValueOutput -Text $cmdText -HostName $vmhost.Name -CommandName $kv.Value }
                }
            } catch {
                $errText="ERROR: $($_.Exception.Message)"; $rawPath=Save-RawCommandOutput -HostName $(if($vmhost){$vmhost.Name}else{$HostName}) -Command $kv.Value -Text $errText
                $raw += [pscustomobject]@{ Host=$(if($vmhost){$vmhost.Name}else{$HostName}); Command=$kv.Value; OutputFile=$rawPath; OutputPreview=$errText }
                Write-Log "[$HostName] SSH command failed [$($kv.Value)]: $($_.Exception.Message)" 'WARN'
            }
        }
        $summary += [pscustomobject]@{
            Host=$vmhost.Name; InputHost=$HostName; Status='Success'; FC_HBA_Count=@($adapters | Where-Object Host -eq $vmhost.Name).Count; SSH_Adapter_Count=@($adapterList | Where-Object Host -eq $vmhost.Name).Count
            WWPNs=(($adapters | Where-Object Host -eq $vmhost.Name | Select-Object -ExpandProperty WWPN) -join '; ')
            WWNNs=(($adapters | Where-Object Host -eq $vmhost.Name | Select-Object -ExpandProperty WWNN) -join '; ')
            Manufacturer=$hw.SystemInfo.Vendor; Model=$hw.SystemInfo.Model; SerialNumber=[string]$hw.SystemInfo.SerialNumber; ESXiVersion=$vmhost.Version; ESXiBuild=[string]$vmhost.Build
            SSHStartedByScript=$sshStartedByScript; Error=''
        }
    } catch {
        Write-Log "[$HostName] Failed: $($_.Exception.Message)" 'ERROR'
        $summary += [pscustomobject]@{ Host=''; InputHost=$HostName; Status='Failed'; FC_HBA_Count=0; SSH_Adapter_Count=0; WWPNs=''; WWNNs=''; Manufacturer=''; Model=''; SerialNumber=''; ESXiVersion=''; ESXiBuild=''; SSHStartedByScript=$sshStartedByScript; Error=$_.Exception.Message }
    } finally {
        if ($vi) {
            try {
                if ($sshStartedByScript -and -not $LeaveSshEnabled) {
                    $svc2 = if ($vmhost) { Get-VMHostService -VMHost $vmhost -ErrorAction SilentlyContinue | Where-Object { $_.Key -eq 'TSM-SSH' } | Select-Object -First 1 } else { $null }
                    if ($svc2 -and $svc2.Running) { Write-Log "[$HostName] Disabling SSH because this script enabled it."; Stop-VMHostService -HostService $svc2 -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
                } elseif ($sshStartedByScript -and $LeaveSshEnabled) { Write-Log "[$HostName] Leaving SSH enabled by user request." 'WARN' }
            } catch { Write-Log "[$HostName] SSH cleanup warning: $($_.Exception.Message)" 'WARN' }
            try { Disconnect-VIServer -Server $vi -Force -Confirm:$false | Out-Null } catch {}
        }
    }
    return [pscustomobject]@{ Summary=@($summary); Adapters=@($adapters); AdapterList=@($adapterList); FCDetails=@($fcDetails); Paths=@($paths); Devices=@($devices); Hardware=@($hardware); Raw=@($raw) }
}

function Export-HbaWorkbook {
    param([Parameter(Mandatory)]$Collected,[Parameter(Mandatory)][string]$OutputPath)
    if (Test-Path $OutputPath) { Remove-Item -Path $OutputPath -Force }
    $headerColor=[System.Drawing.Color]::FromArgb(42,42,44); $summaryColor=[System.Drawing.Color]::FromArgb(58,58,58); $white=[System.Drawing.Color]::White
    $sheets=@(
        @{Name='Summary'; Data=$Collected.Summary}, @{Name='HBA_Adapters_PowerCLI'; Data=$Collected.Adapters}, @{Name='Adapter_List_esxcli'; Data=$Collected.AdapterList},
        @{Name='FC_Details'; Data=$Collected.FCDetails}, @{Name='Storage_Paths'; Data=$Collected.Paths}, @{Name='Storage_Devices_NAA'; Data=$Collected.Devices},
        @{Name='Hardware'; Data=$Collected.Hardware}, @{Name='Raw_Commands'; Data=$Collected.Raw}, @{Name='Failures'; Data=@($Collected.Summary | Where-Object { $_.Status -ne 'Success' })}
    )
    $first=$true
    foreach ($sheet in $sheets) {
        $data=@($sheet.Data); if ($data.Count -eq 0) { $data=@([pscustomobject]@{ Message='No records' }) }
        $safeData=@(ConvertTo-ExcelSafeData -Data $data)
        $params=@{ Path=$OutputPath; WorksheetName=$sheet.Name; AutoSize=$true; FreezeTopRow=$true; BoldTopRow=$true; TableName=($sheet.Name -replace '[^A-Za-z0-9_]','_'); TableStyle='Medium2'; ErrorAction='Stop' }
        if (-not $first) { $params['Append']=$true }
        $safeData | Export-Excel @params
        $first=$false
    }
    $pkg=Open-ExcelPackage -Path $OutputPath
    try {
        foreach ($sheet in $sheets) {
            $ws=$pkg.Workbook.Worksheets[$sheet.Name]
            if ($ws -and $ws.Dimension) {
                $header=$ws.Cells[1,1,1,$ws.Dimension.End.Column]
                $header.Style.Fill.PatternType='Solid'; $header.Style.Fill.BackgroundColor.SetColor($(if($sheet.Name -eq 'Summary'){$summaryColor}else{$headerColor}))
                $header.Style.Font.Color.SetColor($white); $header.Style.Font.Bold=$true; $ws.View.ShowGridLines=$false
                for ($c=1; $c -le $ws.Dimension.End.Column; $c++) {
                    $headerText=[string]$ws.Cells[1,$c].Text
                    if ($headerText -match '(?i)decimal|wwpn|wwnn|naa|uid|serial|uuid|build') {
                        $ws.Column($c).Style.Numberformat.Format='@'
                        for ($r=2; $r -le $ws.Dimension.End.Row; $r++) { if ($null -ne $ws.Cells[$r,$c].Value) { $ws.Cells[$r,$c].Value=[string]$ws.Cells[$r,$c].Value } }
                    }
                }
                if ($sheet.Name -eq 'Raw_Commands' -and $ws.Dimension.End.Column -ge 4) { $ws.Column(3).Width=80; $ws.Column(4).Width=120; $ws.Column(4).Style.WrapText=$true }
            }
        }
    } finally { Close-ExcelPackage $pkg }
}

function Start-Collection {
    if ($script:IsExecuting) { return }
    $script:IsExecuting=$true; $script:CancelRequested=$false
    try {
        if (-not $script:RunDir) { $null=New-RunDir }
        $csv=($script:txtCsvPath.Text+'').Trim(); $outRoot=($script:txtOutputRoot.Text+'').Trim(); $hostColumn=($script:txtHostColumn.Text+'').Trim()
        if ([string]::IsNullOrWhiteSpace($hostColumn)) { $hostColumn='Host' }
        if ([string]::IsNullOrWhiteSpace($outRoot)) { $outRoot=$script:RunDir }
        if (-not (Test-Path $outRoot)) { New-Item -ItemType Directory -Path $outRoot -Force | Out-Null }
        $pass=($script:pbRootPass.Password+''); if ([string]::IsNullOrWhiteSpace($pass)) { throw 'Root password is required.' }
        Update-ProgressUi -Percent 0 -Text 'Checking prerequisites...'
        if ($PSVersionTable.PSVersion.Major -lt 7) { throw "PowerShell 7+ is required. Detected: $($PSVersionTable.PSVersion)" }
        if (-not (Ensure-PowerCliModule)) { throw 'VCF.PowerCLI or VMware.PowerCLI could not be installed/imported.' }
        if (-not (Ensure-Module -Name 'Posh-SSH')) { throw 'Posh-SSH could not be installed/imported.' }
        if (-not (Ensure-Module -Name 'ImportExcel')) { throw 'ImportExcel could not be installed/imported.' }
        Set-PowerCliCertBehavior
        $hosts=@(Read-HostCsv -CsvPath $csv -HostColumn $hostColumn); if ($hosts.Count -lt 1) { throw 'No hostnames/IPs found in CSV.' }
        Write-Log "Loaded $($hosts.Count) host(s) from CSV: $csv"
        $script:dgHosts.ItemsSource=@($hosts | ForEach-Object { [pscustomobject]@{ Host=$_; Status='Queued' } }); Pump-Ui
        $cred=[pscredential]::new('root',(ConvertTo-SecureString $pass -AsPlainText -Force))
        $all=[pscustomobject]@{ Summary=@(); Adapters=@(); AdapterList=@(); FCDetails=@(); Paths=@(); Devices=@(); Hardware=@(); Raw=@() }
        $i=0
        foreach ($h in $hosts) {
            if ($script:CancelRequested) { Write-Log 'Collection cancelled by user.' 'WARN'; break }
            $i++; Update-ProgressUi -Percent ([Math]::Round((($i-1)/$hosts.Count)*100,0)) -Text "Collecting $i/$($hosts.Count): $h"; Pump-Ui
            $res=Get-EsxiHbaDataForHost -HostName $h -Credential $cred -LeaveSshEnabled:([bool]$script:chkLeaveSsh.IsChecked)
            $all.Summary+=@($res.Summary); $all.Adapters+=@($res.Adapters); $all.AdapterList+=@($res.AdapterList); $all.FCDetails+=@($res.FCDetails); $all.Paths+=@($res.Paths); $all.Devices+=@($res.Devices); $all.Hardware+=@($res.Hardware); $all.Raw+=@($res.Raw)
            $script:dgHosts.ItemsSource=@($all.Summary | Select-Object InputHost,Host,Status,FC_HBA_Count,SSH_Adapter_Count,WWPNs,Error); Pump-Ui
        }
        $xlsx=Join-Path $outRoot ("ESXi-HBA-Zoning-Report-{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Update-ProgressUi -Percent 95 -Text 'Writing Excel workbook...'; Export-HbaWorkbook -Collected $all -OutputPath $xlsx
        Update-ProgressUi -Percent 100 -Text 'Complete.'; Write-Log "Excel report saved to: $xlsx"
        [System.Windows.MessageBox]::Show("Collection complete.`n`nReport:`n$xlsx", 'ESXi HBA Zoning Collector', 'OK', 'Information') | Out-Null
    } catch {
        Update-ProgressUi -Percent 0 -Text 'Failed.'; Write-Log "Collection failed: $($_.Exception.Message)" 'ERROR'
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'ESXi HBA Zoning Collector', 'OK', 'Error') | Out-Null
    } finally { $script:IsExecuting=$false }
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase -ErrorAction SilentlyContinue | Out-Null
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue | Out-Null
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue | Out-Null

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ESXi Fibre Channel HBA / Zoning Collector (v{#VER#})"
        Height="860" Width="1380" MinHeight="760" MinWidth="1180"
        WindowStartupLocation="CenterScreen"
        Background="#0f0f10" Foreground="#f3f3f3">
    <Window.Resources>
        <SolidColorBrush x:Key="Bg" Color="#0f0f10"/>
        <SolidColorBrush x:Key="PanelBg" Color="#1c1c1e"/>
        <SolidColorBrush x:Key="Fg" Color="#f3f3f3"/>
        <SolidColorBrush x:Key="Muted" Color="#bfbfbf"/>
        <SolidColorBrush x:Key="Border" Color="#3a3a3a"/>
        <SolidColorBrush x:Key="HeaderBg" Color="#2a2a2c"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.ControlTextBrushKey}" Color="#f3f3f3"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#1c1c1e"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#1c1c1e"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.WindowTextBrushKey}" Color="#f3f3f3"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#3a3a3a"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#f3f3f3"/>
        <Style TargetType="GroupBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="8"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Background" Value="{StaticResource Bg}"/></Style>
        <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Margin" Value="8,0,8,6"/></Style>
        <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Margin" Value="8,4,8,4"/></Style>
        <Style x:Key="InputTextBoxStyle" TargetType="TextBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="4"/><Setter Property="Height" Value="28"/><Setter Property="Background" Value="{StaticResource PanelBg}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/></Style>
        <Style TargetType="PasswordBox"><Setter Property="Margin" Value="8"/><Setter Property="Padding" Value="4"/><Setter Property="Height" Value="28"/><Setter Property="Background" Value="{StaticResource PanelBg}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="BorderBrush" Value="#565656"/></Style>
        <Style TargetType="Button"><Setter Property="Margin" Value="8,6,8,6"/><Setter Property="Padding" Value="8,4"/><Setter Property="Height" Value="30"/><Setter Property="Background" Value="#2a2a2c"/><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="BorderBrush" Value="#565656"/></Style>
        <Style TargetType="ProgressBar"><Setter Property="Margin" Value="8"/><Setter Property="Height" Value="20"/><Setter Property="Minimum" Value="0"/><Setter Property="Maximum" Value="100"/></Style>
        <Style TargetType="DataGrid"><Setter Property="Margin" Value="8"/><Setter Property="Background" Value="{StaticResource PanelBg}"/><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="GridLinesVisibility" Value="All"/><Setter Property="HeadersVisibility" Value="Column"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/><Setter Property="AlternationCount" Value="2"/><Setter Property="RowBackground" Value="#19191b"/><Setter Property="AlternatingRowBackground" Value="#151517"/><Setter Property="HorizontalGridLinesBrush" Value="#303034"/><Setter Property="VerticalGridLinesBrush" Value="#303034"/><Setter Property="SelectionUnit" Value="FullRow"/></Style>
        <Style TargetType="DataGridColumnHeader"><Setter Property="Foreground" Value="{StaticResource Fg}"/><Setter Property="Background" Value="{StaticResource HeaderBg}"/><Setter Property="BorderBrush" Value="{StaticResource Border}"/><Setter Property="FontWeight" Value="SemiBold"/></Style>
    </Window.Resources>
    <Grid Margin="8">
        <Grid.ColumnDefinitions><ColumnDefinition Width="1.05*"/><ColumnDefinition Width="1*"/></Grid.ColumnDefinitions>
        <Grid Grid.Column="0">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <GroupBox Header="Prerequisites" Grid.Row="0"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Grid.Column="0"><TextBlock x:Name="lblPS" Text="PowerShell: (checking...)"/><TextBlock x:Name="lblWPF" Text=".NET/WPF: (checking...)"/><TextBlock x:Name="lblPowerCLI" Text="VCF.PowerCLI / VMware.PowerCLI: (checking...)"/><TextBlock x:Name="lblPoshSSH" Text="Posh-SSH: (checking...)"/><TextBlock x:Name="lblImportExcel" Text="ImportExcel: (checking...)"/></StackPanel><StackPanel Grid.Column="1" VerticalAlignment="Center"><Button x:Name="btnRecheck" Content="Recheck" MinWidth="110"/></StackPanel><StackPanel Grid.Column="2" VerticalAlignment="Center"><Button x:Name="btnInstallModules" Content="Install Missing Modules" MinWidth="180"/></StackPanel></Grid></GroupBox>
            <GroupBox Header="Input" Grid.Row="1"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><Grid.ColumnDefinitions><ColumnDefinition Width="3*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="1*"/></Grid.ColumnDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="CSV Path"/><TextBlock Grid.Row="0" Grid.Column="2" Text="Host Column"/><TextBox Grid.Row="1" Grid.Column="0" x:Name="txtCsvPath" Style="{StaticResource InputTextBoxStyle}"/><Button Grid.Row="1" Grid.Column="1" x:Name="btnBrowseCsv" Content="Browse CSV" MinWidth="120"/><TextBox Grid.Row="1" Grid.Column="2" x:Name="txtHostColumn" Text="Host" Style="{StaticResource InputTextBoxStyle}"/><TextBlock Grid.Row="2" Grid.ColumnSpan="3" Text="CSV should contain a Host column with ESXi hostnames or IPs." Foreground="{StaticResource Muted}" FontSize="11"/></Grid></GroupBox>
            <GroupBox Header="Credentials and Output" Grid.Row="2"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><Grid.ColumnDefinitions><ColumnDefinition Width="1*"/><ColumnDefinition Width="3*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Grid.Row="0" Grid.Column="0" Text="Root Password"/><TextBlock Grid.Row="0" Grid.Column="1" Text="Output Folder"/><PasswordBox Grid.Row="1" Grid.Column="0" x:Name="pbRootPass"/><TextBox Grid.Row="1" Grid.Column="1" x:Name="txtOutputRoot" Style="{StaticResource InputTextBoxStyle}"/><Button Grid.Row="1" Grid.Column="2" x:Name="btnBrowseOutput" Content="Browse Output" MinWidth="130"/><CheckBox Grid.Row="2" Grid.ColumnSpan="3" x:Name="chkLeaveSsh" Content="Leave SSH enabled if this script had to enable it (not recommended)" IsChecked="False"/></Grid></GroupBox>
            <GroupBox Header="Actions" Grid.Row="3"><StackPanel><WrapPanel><Button x:Name="btnValidateCsv" Content="Validate CSV" MinWidth="120"/><Button x:Name="btnExecute" Content="Collect HBA Data" MinWidth="150"/><Button x:Name="btnCancel" Content="Cancel" MinWidth="100"/><Button x:Name="btnOpenRun" Content="Open Run Folder" MinWidth="140"/><Button x:Name="btnCloseWindow" Content="Close Window" MinWidth="120"/></WrapPanel><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><ProgressBar Grid.Column="0" x:Name="pbExec" Value="0"/><TextBlock Grid.Column="1" x:Name="lblProgress" Text="Idle" Margin="8,2,8,2" VerticalAlignment="Center"/></Grid></StackPanel></GroupBox>
            <GroupBox Header="Host Status" Grid.Row="4"><DataGrid x:Name="dgHosts" AutoGenerateColumns="True" IsReadOnly="True" CanUserAddRows="False"/></GroupBox>
        </Grid>
        <GroupBox Header="Log" Grid.Column="1"><Grid><TextBox x:Name="txtLog" Style="{x:Null}" Margin="8" Padding="6" AcceptsReturn="True" TextWrapping="NoWrap" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="#050505" Foreground="#f3f3f3" BorderBrush="#3a3a3a" FontFamily="Consolas" FontSize="12"/></Grid></GroupBox>
    </Grid>
</Window>
"@
$xaml=$xaml.Replace('{#VER#}',$Global:EsxiHbaUiVersion)
try { $script:window=[Windows.Markup.XamlReader]::Parse($xaml) } catch { [System.Windows.MessageBox]::Show("XAML parse failed:`r`n$($_.Exception.Message)",'ESXi HBA Zoning Collector','OK','Error') | Out-Null; throw }
foreach ($name in @('lblPS','lblWPF','lblPowerCLI','lblPoshSSH','lblImportExcel','btnRecheck','btnInstallModules','txtCsvPath','btnBrowseCsv','txtHostColumn','pbRootPass','txtOutputRoot','btnBrowseOutput','chkLeaveSsh','btnValidateCsv','btnExecute','btnCancel','btnOpenRun','btnCloseWindow','pbExec','lblProgress','dgHosts','txtLog')) { Set-Variable -Name $name -Scope Script -Value $script:window.FindName($name) }
function Prereq-Check {
    Set-StatusText -Label $script:lblPS -Text ("PowerShell {0}" -f $PSVersionTable.PSVersion) -State ($(if ($PSVersionTable.PSVersion.Major -ge 7) {'OK'} else {'FAIL'}))
    Set-StatusText -Label $script:lblWPF -Text '.NET/WPF: OK' -State 'OK'
    $pcli=(Has-Module 'VCF.PowerCLI') -or (Has-Module 'VMware.PowerCLI')
    Set-StatusText -Label $script:lblPowerCLI -Text ($(if($pcli){'VCF.PowerCLI / VMware.PowerCLI: Found'}else{'VCF.PowerCLI / VMware.PowerCLI: Not found'})) -State ($(if($pcli){'OK'}else{'WARN'}))
    Set-StatusText -Label $script:lblPoshSSH -Text ($(if(Has-Module 'Posh-SSH'){'Posh-SSH: Found'}else{'Posh-SSH: Not found'})) -State ($(if(Has-Module 'Posh-SSH'){'OK'}else{'WARN'}))
    Set-StatusText -Label $script:lblImportExcel -Text ($(if(Has-Module 'ImportExcel'){'ImportExcel: Found'}else{'ImportExcel: Not found'})) -State ($(if(Has-Module 'ImportExcel'){'OK'}else{'WARN'}))
}
$script:window.Add_ContentRendered({ try { if(-not $script:RunDir){$null=New-RunDir}; $script:txtOutputRoot.Text=$script:RunDir; Write-Log "==== ESXi Fibre Channel HBA / Zoning Collector started (v$Global:EsxiHbaUiVersion) ===="; Write-Log "Run folder: $script:RunDir"; Prereq-Check; Update-ProgressUi -Percent 0 -Text 'Idle' } catch {} })
$script:btnRecheck.Add_Click({ Prereq-Check })
$script:btnInstallModules.Add_Click({ Ensure-PowerCliModule | Out-Null; Ensure-Module -Name 'Posh-SSH' | Out-Null; Ensure-Module -Name 'ImportExcel' | Out-Null; Prereq-Check })
$script:btnBrowseCsv.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter='CSV files (*.csv)|*.csv|All files (*.*)|*.*'; if($dlg.ShowDialog() -eq 'OK'){ $script:txtCsvPath.Text=$dlg.FileName } })
$script:btnBrowseOutput.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog() -eq 'OK'){ $script:txtOutputRoot.Text=$dlg.SelectedPath } })
$script:btnValidateCsv.Add_Click({ try { $hosts=@(Read-HostCsv -CsvPath (($script:txtCsvPath.Text+'').Trim()) -HostColumn (($script:txtHostColumn.Text+'').Trim())); $script:dgHosts.ItemsSource=@($hosts | ForEach-Object { [pscustomobject]@{ Host=$_; Status='Ready' } }); Write-Log "CSV validation OK. Host count: $($hosts.Count)"; [System.Windows.MessageBox]::Show("CSV validation OK.`nHosts found: $($hosts.Count)",'ESXi HBA Zoning Collector','OK','Information') | Out-Null } catch { Write-Log "CSV validation failed: $($_.Exception.Message)" 'ERROR'; [System.Windows.MessageBox]::Show($_.Exception.Message,'ESXi HBA Zoning Collector','OK','Error') | Out-Null } })
$script:btnExecute.Add_Click({ Start-Collection })
$script:btnCancel.Add_Click({ $script:CancelRequested=$true; Write-Log 'Cancel requested by user.' 'WARN'; Update-ProgressUi -Percent $script:pbExec.Value -Text 'Cancel requested...' })
$script:btnOpenRun.Add_Click({ try { if($script:RunDir -and (Test-Path $script:RunDir)){ Start-Process $script:RunDir | Out-Null } } catch {} })
$script:btnCloseWindow.Add_Click({ if($script:IsExecuting){ $ans=[System.Windows.MessageBox]::Show('Collection is currently running. Close anyway?','Close Window','YesNo','Warning'); if($ans -ne 'Yes'){return} }; $script:window.Close() })
$null=$script:window.ShowDialog()

# SIG # Begin signature block
# MIIHmQYJKoZIhvcNAQcCoIIHijCCB4YCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqXsgOCfrVwZRW
# vU+nHDBkCsELoTtdQqCyufKbfqMapqCCBGIwggReMIICxqADAgECAhAVZPpanf6Q
# k0o3TfMU2tY9MA0GCSqGSIb3DQEBCwUAMEcxRTBDBgNVBAMMPEVTWGlIQkFDb2xs
# ZWN0b3JVSSBMb2NhbCBDb2RlIFNpZ25pbmcgKHhhZG1pbkBIT01FT0ZGSUNFTEFC
# KTAeFw0yNjA1MTQxOTUwMzJaFw0zMTA1MTQyMDAwMzJaMEcxRTBDBgNVBAMMPEVT
# WGlIQkFDb2xsZWN0b3JVSSBMb2NhbCBDb2RlIFNpZ25pbmcgKHhhZG1pbkBIT01F
# T0ZGSUNFTEFCKTCCAaIwDQYJKoZIhvcNAQEBBQADggGPADCCAYoCggGBAL9cYAuE
# Q3kGGsl46tjT/DWMCn55ZSY+3rrtQXRjPLdk9h/Aa57HgT346BxH8QegwynuUxXy
# D3nfv5ksW+HUeO1qlrd0BkTV4Obq7AuxZDhMnzpuks+z+QoM0X3WHmgPnlM8LDLs
# lXWpagDhXMkrpeOEe7RaQZ6tCmnkTeacyLqlJz2FrdxYbPXCXnIspfhHoujZGPA0
# jNb/7/o0fTCG7hJl2LXt3IcqrO0VS9f15WgFtmj5NB495XQ7RcTHMkT/fqwfZW95
# H3wI/quO0OUGGMx9WL6qnwOZRIfwk43+J4qgWiLdMVNsLe7Im1NM+7+359JqJfC8
# DRJE422wmQSdKLr9mgqajxGOkGjjN+cWzsGwvqrYYbsjnW8wXWB9PGd/LGFXAqxQ
# Lw1/cVUidSOwXjKwPYixltOX1hbTub2uN7wHZB9ezaNsyEGVNSRLZPLxmgVmD7tz
# N8w3JM8gjLqkfXQ+Q9gMmHy9HYamdwJU2aWh3VQiqtBnONhiS7Ia5askRQIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFOUd/tphOItaQZIanq/mdk2xc6E/MA0GCSqGSIb3DQEBCwUAA4IBgQBHfSHQ
# h5Vly/F1rJkEweqJyK28Jv8JtaqF8sbIF8sNVOoWr/X7a3l7RUsCHFfnFBqJKDoN
# nXjA/MGBmGfWckzQo1TcMqjIm7fYziAU9f1WC2BmR/a/QV3kO9X9pTjpuE69BSTr
# cim4BKuV+8z0VckaYxWlBI8M3KjCuXM2PjmL/IVsW6hzXjM+ONWvJhyjyC4JFPEL
# M6dJxxzVniBAHSWIxKxWkthcf1scwWDaXoNRCnxCK61raIVbglEUFJm9CGh+0WJR
# 7L5+SOxEdTifaV1lL/AiXpZafuffOcqFr+PGjk9Qvb4sJg+9i3o58yFAekgPtY8A
# iDh1BZPZ78pgtBDstI0NbhjVi4od5j/LkxQlNLL13glcZ2OohaVbv84D/4IV+LU6
# EmDSUrIw1WfwP2kOXeuNgFAh+fZqar+8x9SKeKZ45zH5iU0/aFwnwXvXyej3mYsd
# gEkn9I3nE1WXynS9LTrTSsj6aanyKP/zeO2pnehVxgqa3uUF3ivJtanKlR4xggKN
# MIICiQIBATBbMEcxRTBDBgNVBAMMPEVTWGlIQkFDb2xsZWN0b3JVSSBMb2NhbCBD
# b2RlIFNpZ25pbmcgKHhhZG1pbkBIT01FT0ZGSUNFTEFCKQIQFWT6Wp3+kJNKN03z
# FNrWPTANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDMVoIcAy9GPEuxwaNp271r5LCFOIkU
# dSpxdJGowYwuzDANBgkqhkiG9w0BAQEFAASCAYBDbDMsC3rr+tVebrlJKZBrgQwj
# kq2arLVsgR2BHh3ZHLXfj7oY0lLx/uEWRp53qjL5aJWj2RfjM0SCjQBkXaMd0XY7
# 6FcOC1jcqcjF1RjyTcXg8uN2FWATybibuMGMh5AtwRhljSungbfFX0jIb1auhBWo
# cHCCXlBB9cpeV3BDPaMgo1SIyPlEIxNkYvP5hB0Z09KX/nDClgl4ZqUL3FWKMNcu
# h4JUK77MQ+PcWCVuxmebutYqocJpyF8aSjpx65NhI62EYVw/YIlnBA4VpYzeCrum
# sj/p1VGCYlquqgvDSbwOrwfgwX32BNOvaY/dTkTUwdRrTlLwXuN8Po4Ua+Fig9kV
# K5h3xGEIMITZV391CZrQi77B+5CabM93jOePGml00jm70aZr60g/ojht8tS/5DAh
# oUIgrN+8NrkU7VJOxxrs6s9ZYiU23reEiY5pOscZWPHMyS7ftwR4x5zhpu8vtSRF
# B/LlDKivGMXuPByICDyCVwqIdmaxziTtfGaF+p8=
# SIG # End signature block
