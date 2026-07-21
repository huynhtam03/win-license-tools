# Check-LicenseStatus-Standalone.ps1 - Cong cu IT Helpdesk Modern Dark Dashboard All-In-One

[CmdletBinding()]
param (
    [switch]$ConsoleOnly,
    [switch]$RemoveKMS,
    [switch]$UninstallCracks,
    [switch]$UninstallOffice,
    [switch]$CleanTemp
)

# 1. THIET LAP THU VIEN WPF
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
} catch {}

# 2. TU DONG NANG QUYEN ADMINISTRATOR
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "CANH BAO: Cong cu yeu cau quyen Administrator!" -ForegroundColor Yellow
    Write-Host "Dang mo cua so PowerShell moi duoi quyen Administrator..." -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-sta", "-File", "`"$PSCommandPath`"" -Verb RunAs
    } catch {
        Write-Host "Loi: Khong the khoi chay duoi quyen Administrator!" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
    Exit
}

# 3. CAC HAM TRUY VAN HE THONG & IT HELPDESK
function Get-SystemInfoSummary {
    $ip = "127.0.0.1"
    $mac = "00:00:00:00:00:00"
    try {
        $net = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1
        if ($net) { $ip = $net.IPAddress }
        $nic = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1
        if ($nic) { $mac = $nic.MacAddress }
    } catch {}
    return @{ ComputerName = $env:COMPUTERNAME; IP = $ip; MAC = $mac }
}

function Get-FullComputerAssetInfo {
    $asset = [ordered]@{
        ComputerName            = $env:COMPUTERNAME
        SystemManufacturer      = "Khong xac dinh"
        SystemModel             = "Khong xac dinh"
        SystemUUID              = "Khong xac dinh"
        BIOSSerialNumber        = "Khong xac dinh"
        BIOSVersion             = "Khong xac dinh"
        MotherboardManufacturer = "Khong xac dinh"
        MotherboardProduct      = "Khong xac dinh"
        MotherboardSerial       = "Khong xac dinh"
        
        OSName                  = "Khong xac dinh"
        OSVersion               = "Khong xac dinh"
        OSArchitecture          = "Khong xac dinh"
        OSInstallDate           = "Khong xac dinh"
        CurrentUser             = $env:USERNAME
        UserDomain              = $env:USERDOMAIN
        
        CPU                     = "Khong xac dinh"
        CPUCoresThreads         = "Khong xac dinh"
        GPU                     = "Khong xac dinh"
        GPUVRAM_GB              = "Khong xac dinh"
        
        TotalRAM_GB             = 0
        RAMSlotsCount           = 0
        RAMSticksDetail         = "Khong xac dinh"
        
        TotalStorageCapacityGB  = 0
        TotalFreeStorageGB      = 0
        PhysicalDisksDetail     = "Khong xac dinh"
        LogicalDrivesDetail     = "Khong xac dinh"
        
        NetworkAdapter          = "Khong xac dinh"
        IPAddress               = "Khong xac dinh"
        MACAddress              = "Khong xac dinh"
        AuditTimestamp          = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    }

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $asset.SystemManufacturer = $cs.Manufacturer
            $asset.SystemModel = $cs.Model
        }

        $csp = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
        if ($csp) {
            $asset.SystemUUID = $csp.UUID
        }

        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $asset.BIOSSerialNumber = $bios.SerialNumber
            $asset.BIOSVersion = "$($bios.Manufacturer) $($bios.SMBIOSBIOSVersion)"
        }

        $baseboard = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue
        if ($baseboard) {
            $asset.MotherboardManufacturer = $baseboard.Manufacturer
            $asset.MotherboardProduct = $baseboard.Product
            $asset.MotherboardSerial = $baseboard.SerialNumber
        }

        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $asset.OSName = $os.Caption
            $asset.OSVersion = "$($os.Version) (Build $($os.BuildNumber))"
            $asset.OSArchitecture = $os.OSArchitecture
            $asset.OSInstallDate = $os.InstallDate.ToString('dd/MM/yyyy HH:mm:ss')
        }

        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpu) {
            $asset.CPU = $cpu.Name.Trim()
            $asset.CPUCoresThreads = "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"
        }

        $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
        if ($gpus) {
            $gpuNames = @()
            $vramGBs = @()
            foreach ($g in $gpus) {
                $gpuNames += $g.Name
                if ($g.AdapterRAM) {
                    $vramGBs += "$([math]::Round($g.AdapterRAM / 1GB, 2)) GB"
                }
            }
            $asset.GPU = $gpuNames -join " | "
            $asset.GPUVRAM_GB = if ($vramGBs.Count -gt 0) { $vramGBs -join " | " } else { "N/A" }
        }

        $ramChips = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
        if ($ramChips) {
            $totalBytes = ($ramChips | Measure-Object -Property Capacity -Sum).Sum
            $asset.TotalRAM_GB = [math]::Round($totalBytes / 1GB, 2)
            $asset.RAMSlotsCount = $ramChips.Count
            
            $ramDetailsList = @()
            foreach ($r in $ramChips) {
                $chipGB = [math]::Round($r.Capacity / 1GB, 2)
                $speed = if ($r.Speed) { "$($r.Speed)MHz" } else { "" }
                $mfg = if ($r.Manufacturer) { $r.Manufacturer.Trim() } else { "" }
                $bank = if ($r.DeviceLocator) { $r.DeviceLocator } else { "Slot" }
                $ramDetailsList += "[$($bank): $chipGB GB $speed $mfg]"
            }
            $asset.RAMSticksDetail = $ramDetailsList -join " | "
        }

        $physDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
        if ($physDisks) {
            $pDiskList = @()
            $totalDiskBytes = 0
            foreach ($d in $physDisks) {
                $gb = [math]::Round($d.Size / 1GB, 2)
                $totalDiskBytes += $d.Size
                $pDiskList += "$($d.Model) ($gb GB - SN: $($d.SerialNumber))"
            }
            $asset.TotalStorageCapacityGB = [math]::Round($totalDiskBytes / 1GB, 2)
            $asset.PhysicalDisksDetail = $pDiskList -join " | "
        }

        $logDrives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction SilentlyContinue
        if ($logDrives) {
            $lDriveList = @()
            $totalFreeBytes = 0
            foreach ($ld in $logDrives) {
                $totalGB = [math]::Round($ld.Size / 1GB, 2)
                $freeGB = [math]::Round($ld.FreeSpace / 1GB, 2)
                $usedGB = [math]::Round(($ld.Size - $ld.FreeSpace) / 1GB, 2)
                $totalFreeBytes += $ld.FreeSpace
                $percentFree = [math]::Round(($ld.FreeSpace / $ld.Size) * 100, 1)
                $lDriveList += "[$($ld.DeviceID) Tong: $totalGB GB | Dung: $usedGB GB | Trong: $freeGB GB ($percentFree% trong)]"
            }
            $asset.TotalFreeStorageGB = [math]::Round($totalFreeBytes / 1GB, 2)
            $asset.LogicalDrivesDetail = $lDriveList -join " | "
        }

        $net = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1
        if ($net) {
            $asset.IPAddress = $net.IPAddress
            $nic = Get-NetAdapter -InterfaceIndex $net.InterfaceIndex -ErrorAction SilentlyContinue
            if ($nic) {
                $asset.NetworkAdapter = $nic.InterfaceDescription
                $asset.MACAddress = $nic.MacAddress
            }
        }
    } catch {}

    return [PSCustomObject]$asset
}

function Get-WindowsActivation {
    $result = @{
        Status = "Unknown"
        StatusText = "Chua quet"
        ProductName = "Windows"
        Channel = "Khong xac dinh"
        KMSServer = "Khong phat hien"
        RawStatus = -1
        IsGenuine = $true
        Details = ""
    }
    try {
        $win = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not null" -ErrorAction SilentlyContinue
        if ($win) {
            $win = $win | Select-Object -First 1
            $result.ProductName = $win.Name
            $result.RawStatus = $win.LicenseStatus
            
            switch ($win.LicenseStatus) {
                0 { $result.StatusText = "Chua kich hoat (Unlicensed)"; $result.Status = "Unlicensed"; $result.IsGenuine = $false }
                1 { $result.StatusText = "Da kich hoat (Licensed)"; $result.Status = "Licensed"; $result.IsGenuine = $true }
                2 { $result.StatusText = "An han ban dau (OOB Grace)"; $result.Status = "Grace"; $result.IsGenuine = $false }
                3 { $result.StatusText = "An han het han (OOT Grace)"; $result.Status = "Grace"; $result.IsGenuine = $false }
                4 { $result.StatusText = "An han khong chinh hang (Non-genuine)"; $result.Status = "Warning"; $result.IsGenuine = $false }
                5 { $result.StatusText = "Cho kich hoat / Het han (Notification)"; $result.Status = "Notification"; $result.IsGenuine = $false }
                6 { $result.StatusText = "An han mo rong (Extended Grace)"; $result.Status = "Grace"; $result.IsGenuine = $false }
                default { $result.StatusText = "Khong xac dinh"; $result.Status = "Unknown" }
            }
            
            if ($win.Description -like "*RETAIL*") { $result.Channel = "Retail (Ban le)" }
            elseif ($win.Description -like "*OEM*") { $result.Channel = "OEM (Theo may)" }
            elseif ($win.Description -like "*VOLUME_KMSCLIENT*") { $result.Channel = "Volume KMS (KMS Doanh nghiep)" }
            elseif ($win.Description -like "*VOLUME_MAK*") { $result.Channel = "Volume MAK" }
            else { $result.Channel = "Khac/Tu cau hinh" }
            
            $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
            if ($licService -and $licService.KeyManagementServiceMachine) {
                $result.KMSServer = "$($licService.KeyManagementServiceMachine):$($licService.KeyManagementServicePort)"
            }
            
            $result.Details = "San pham: $($win.Name)`nPhan nhom: $($result.Channel)`nTrang thai: $($result.StatusText)`nMay chu KMS: $($result.KMSServer)"
        } else {
            $result.StatusText = "Khong co san pham hoat dong"
            $result.Details = "Khong the tim thay thong tin giay phep Windows hop le qua WMI."
        }
    } catch {
        $result.StatusText = "Loi truy van"
        $result.Details = "Loi khi lay thong tin ban quyen Windows: $($_.Exception.Message)"
    }
    return $result
}

function Get-OfficeActivation {
    $officeList = @()
    $offices = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
    if ($offices) {
        foreach ($off in $offices) {
            $statusText = switch ($off.LicenseStatus) {
                0 { "Chua kich hoat" }
                1 { "Da kich hoat" }
                2 { "An han ban dau" }
                3 { "An han het han" }
                4 { "An han khong chinh hang" }
                5 { "Cho kich hoat / Het han" }
                6 { "An han mo rong" }
                default { "Khong xac dinh" }
            }
            
            $channel = "Khong xac dinh"
            if ($off.Description -like "*RETAIL*") { $channel = "Retail (Ban le)" }
            elseif ($off.Description -like "*VOLUME_KMS*") { $channel = "Volume KMS" }
            elseif ($off.Description -like "*VOLUME_MAK*") { $channel = "Volume MAK" }
            elseif ($off.Description -like "*SUBSCRIPTION*") { $channel = "Subscription (Office 365)" }
            
            $officeList += [PSCustomObject]@{
                ProductName = $off.Name
                Status = $statusText
                Channel = $channel
                KMSServer = $off.KeyManagementServiceMachine
                GraceRemaining = $off.GracePeriodRemaining
                Source = "WMI"
            }
        }
    }
    
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            $output = cscript.exe //NoLogo "$ospp" /dstatus 2>&1
            $currentProduct = ""
            $status = ""
            $kms = ""
            $licenseName = ""
            
            foreach ($line in $output) {
                if ($line -like "*PRODUCT ID:*") {
                    if ($currentProduct) {
                        if (-not ($officeList | Where-Object { $_.ProductName -like "*$currentProduct*" -or $_.ProductName -like "*$licenseName*" })) {
                            $officeList += [PSCustomObject]@{
                                ProductName = "$licenseName"
                                Status = $status
                                Channel = "Volume Client (OSPP)"
                                KMSServer = $kms
                                GraceRemaining = "N/A"
                                Source = "OSPP"
                            }
                        }
                    }
                    $currentProduct = ($line -split "PRODUCT ID:")[1].Trim()
                    $status = ""
                    $kms = ""
                    $licenseName = ""
                }
                elseif ($line -like "*LICENSE NAME:*") {
                    $licenseName = ($line -split "LICENSE NAME:")[1].Trim()
                }
                elseif ($line -like "*LICENSE STATUS:*") {
                    $statusLine = ($line -split "LICENSE STATUS:")[1].Trim()
                    if ($statusLine -like "*---LICENSED---*") { $status = "Da kich hoat" }
                    elseif ($statusLine -like "*---NOTIFIED---*") { $status = "Cho kich hoat / Het han" }
                    else { $status = "Chua kich hoat ($statusLine)" }
                }
                elseif ($line -like "*KMS MACHINE NAME:*") {
                    $kms = ($line -split "KMS MACHINE NAME:")[1].Trim()
                }
            }
            if ($currentProduct) {
                if (-not ($officeList | Where-Object { $_.ProductName -like "*$currentProduct*" -or $_.ProductName -like "*$licenseName*" })) {
                    $officeList += [PSCustomObject]@{
                        ProductName = "$licenseName"
                        Status = $status
                        Channel = "Volume Client (OSPP)"
                        KMSServer = $kms
                        GraceRemaining = "N/A"
                        Source = "OSPP"
                    }
                }
            }
        }
    }
    return $officeList
}

function Get-CrackDetection {
    $risks = @()
    $details = @()

    $crackTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { 
        $_.TaskName -like "*AutoKMS*" -or 
        $_.TaskName -like "*KMSAuto*" -or 
        $_.TaskName -like "*KMSConnection*" -or 
        $_.TaskName -like "*AutoPico*" -or 
        $_.TaskName -like "*HEU_KMS*" -or
        $_.TaskName -like "*KMS-Activator*"
    }
    if ($crackTasks) {
        foreach ($t in $crackTasks) {
            $risks += "Phat hien Tac vu lap lich be khoa (Scheduled Task): $($t.TaskName)"
            $details += "Tac vu ngam '$($t.TaskName)' dang duoc len lich chay tai '$($t.TaskPath)'."
        }
    }

    $crackServices = Get-Service -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*AutoKMS*" -or 
        $_.Name -like "*KMSpico*" -or 
        $_.Name -like "*KMSConnection*" -or 
        $_.Name -like "*Service_KMS*" 
    }
    if ($crackServices) {
        foreach ($s in $crackServices) {
            $risks += "Phat hien Dich vu be khoa chay ngam (Service): $($s.DisplayName)"
            $details += "Dich vu '$($s.DisplayName)' ($($s.Name)) dang o trang thai $($s.Status)."
        }
    }

    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        $hostsContent = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $kmsBlocks = $hostsContent | Where-Object { 
            $_ -and (-not $_.StartsWith("#")) -and (
                $_ -like "*kms*" -or 
                $_ -like "*microsoft.com*" -or 
                $_ -like "*activation*" -or
                $_ -like "*adobe*" -or
                $_ -like "*autodesk*"
            )
        }
        if ($kmsBlocks) {
            $risks += "File Hosts bi can thiep de chan xac minh ban quyen"
            $details += "Phat hien $($kmsBlocks.Count) dong cau hinh Redirect/Block trong file Hosts ($hostsPath)."
        }
    }

    $suspiciousPaths = @(
        "$env:windir\AutoKMS\AutoKMS.exe",
        "$env:ProgramFiles\KMSpico\KMSpico.exe",
        "${env:ProgramFiles(x86)}\KMSpico\KMSpico.exe",
        "$env:windir\KMSAuto Net.exe",
        "C:\ProgramData\KMSAutoS\KMSAuto X64.exe"
    )
    foreach ($path in $suspiciousPaths) {
        if (Test-Path $path) {
            $risks += "Phat hien tep tin cong cu be khoa tren o đia: $path"
            $details += "Tep tin thuc thi be khoa nguy hiem ton tai tai '$path'."
        }
    }

    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    $crackKeywords = @("AutoKMS", "KMSpico", "KMSAuto", "MiniTool", "CCleaner Pro Crack", "Adobe GenP", "Patch", "Crack", "Keygen")
    
    foreach ($app in $installedApps) {
        foreach ($kw in $crackKeywords) {
            if ($app.DisplayName -like "*$kw*") {
                $risks += "Phan mem thuong mai bi Crack/Be khoa: $($app.DisplayName)"
                $details += "Ung dung '$($app.DisplayName)' duoc ghi nhan trong Registry voi InstallLocation: $($app.InstallLocation)."
            }
        }
    }

    $officeRegKms = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    if ($officeRegKms -and $officeRegKms.KeyManagementServiceMachine) {
        $kmsMachineOffice = $officeRegKms.KeyManagementServiceMachine
        if ($kmsMachineOffice -eq "127.0.0.1" -or $kmsMachineOffice -eq "localhost" -or $kmsMachineOffice -eq "::1") {
            $risks += "Office cau hinh may chu KMS noi bo (Localhost)"
            $details += "Ban quyen Office tro may chu kich hoat ve chinh may tinh thong qua Registry."
        }
    }

    try {
        $kmsPort = Get-NetTCPConnection -LocalPort 1688 -State Listen -ErrorAction SilentlyContinue
        if ($kmsPort) {
            $risks += "Phat hien cong dich vu KMS (1688) dang mo"
            $details += "Cong TCP 1688 dang lang nghe (Listen) cuc bo. Dau hieu cua KMS gia lap chay an."
        }
    } catch {
        $netstat = netstat -ano 2>&1 | Select-String ":1688\s+.*LISTENING"
        if ($netstat) {
            $risks += "Phat hien cong dich vu KMS (1688) dang mo (netstat)"
            $details += "Cong 1688 dang mo o trang thai LISTENING qua lenh netstat."
        }
    }

    $crackRegs = @("HKLM:\SOFTWARE\KMSAuto", "HKLM:\SOFTWARE\KMSAutoS", "HKLM:\SOFTWARE\KMSpico", "HKCU:\Software\KMSAuto", "HKCU:\Software\KMSAutoS")
    foreach ($reg in $crackRegs) {
        if (Test-Path $reg) {
            $risks += "Phat hien khoa Registry be khoa: $reg"
            $details += "Khoa cau hinh cua KMSAuto/KMSpico duoc phat hien tai '$reg'."
        }
    }
    
    return [PSCustomObject]@{
        Risks = $risks
        Details = $details
    }
}

# 4. THUC THI CAC HAM XU LY CONG VIEC (LOGIC)
function Start-KMSRemovalProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("        BAT DAU QUA TRINH GO BO BAN QUYEN KMS LAU VA LAM SACH HE THONG            ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    # Buoc 1
    $log.AppendLine("[BUOC 1] Dang go bo cau hinh may chu KMS Windows (slmgr /ckms)...") | Out-Null
    try {
        $ckmsResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ckms 2>&1
        $log.AppendLine(" -> Ket qua slmgr /ckms: $($ckmsResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /ckms: $($_.Exception.Message)") | Out-Null
    }

    # Buoc 2
    $log.AppendLine("[BUOC 2] Dang go bo Product Key KMS Windows (slmgr /upk)...") | Out-Null
    try {
        $upkResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /upk 2>&1
        $log.AppendLine(" -> Ket qua slmgr /upk: $($upkResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /upk: $($_.Exception.Message)") | Out-Null
    }

    # Buoc 3
    $log.AppendLine("[BUOC 3] Dang xoa thong tin Product Key Windows khoi Registry (slmgr /cpky)...") | Out-Null
    try {
        $cpkyResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /cpky 2>&1
        $log.AppendLine(" -> Ket qua slmgr /cpky: $($cpkyResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /cpky: $($_.Exception.Message)") | Out-Null
    }

    # Buoc 4
    $log.AppendLine("[BUOC 4] Dang kiem tra va lam sach tep tin Hosts...") | Out-Null
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        try {
            $lines = Get-Content $hostsPath -ErrorAction SilentlyContinue
            $cleanLines = @()
            $removedCount = 0
            
            foreach ($line in $lines) {
                $clean = $line.Trim()
                if ($clean -and -not $clean.StartsWith("#") -and ($clean -like "*microsoft*" -or $clean -like "*activation*" -or $clean -like "*kms*")) {
                    $log.AppendLine(" -> Da go bo dong chan: $clean") | Out-Null
                    $removedCount++
                } else {
                    $cleanLines += $line
                }
            }
            if ($removedCount -gt 0) {
                $cleanLines | Set-Content $hostsPath -Encoding utf8 -Force
                $log.AppendLine(" -> Da lam sach $removedCount dong cau hinh chan trong file Hosts!") | Out-Null
            } else {
                $log.AppendLine(" -> Tep Hosts khong chua cau hinh chan may chu Microsoft.") | Out-Null
            }
        } catch {
            $log.AppendLine(" -> Loi khi chinh sua file Hosts: $($_.Exception.Message)") | Out-Null
        }
    }

    # Buoc 5
    $log.AppendLine("[BUOC 5] Dang don dep cac tac vu be khoa tu dong (Scheduled Tasks)...") | Out-Null
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    $removedTasks = 0
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($t in $tasks) {
            try {
                Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
                $log.AppendLine(" -> Da go bo tac vu be khoa: $($t.TaskName) ($($t.TaskPath))") | Out-Null
                $removedTasks++
            } catch {}
        }
    }
    if ($removedTasks -eq 0) {
        $log.AppendLine(" -> Khong phat hien tac vu chay ngam be khoa nao.") | Out-Null
    }

    # Buoc 6
    $log.AppendLine("[BUOC 6] Dang don dep cac dich vu KMS lau (Services)...") | Out-Null
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    $removedSrvs = 0
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            try {
                Stop-Service -Name $srvName -Force -ErrorAction SilentlyContinue
                sc.exe delete $srvName | Out-Null
                $log.AppendLine(" -> Da dung va go bo dich vu KMS lau: $($srv.DisplayName) ($srvName)") | Out-Null
                $removedSrvs++
            } catch {}
        }
    }
    if ($removedSrvs -eq 0) {
        $log.AppendLine(" -> Khong phat hien dich vu KMS lau chay ngam nao.") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Buoc 7
    $log.AppendLine("[BUOC 7] Dang go bo ban quyen KMS cua Office/Project/Visio qua ospp.vbs va WMI...") | Out-Null
    
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    
    $osppFound = $false
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            $osppFound = $true
            $log.AppendLine(" -> Phat hien OSPP tool tai: $ospp") | Out-Null
            try { cscript.exe //NoLogo "$ospp" /remkms 2>&1 | Out-Null } catch {}
            try {
                $dstatusOut = cscript.exe //NoLogo "$ospp" /dstatus 2>&1
                $keysToUninstall = @()
                foreach ($line in $dstatusOut) {
                    if ($line -match "Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})") {
                        $key5 = $matches[1]
                        if ($key5 -and $keysToUninstall -notcontains $key5) {
                            $keysToUninstall += $key5
                        }
                    }
                }
                foreach ($k5 in $keysToUninstall) {
                    $log.AppendLine(" -> Dang go Product Key OSPP (5 ky tu cuoi: $k5)...") | Out-Null
                    cscript.exe //NoLogo "$ospp" /unpkey:$k5 2>&1 | Out-Null
                }
            } catch {}
        }
    }
    
    if (-not $osppFound) {
        $log.AppendLine(" -> Khong tim thay tep tin ospp.vbs tren cac duong dan mac dinh.") | Out-Null
    }

    $log.AppendLine(" -> Dang quet va go toan bo khoa san pham Office/Project/Visio trong WMI...") | Out-Null
    try {
        $wmiProds = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue | Where-Object { 
            ($_.PartialProductKey -or $_.LicenseStatus -eq 1 -or $_.LicenseStatus -eq 5) -and 
            ($_.Name -like "*Office*" -or $_.Name -like "*Project*" -or $_.Name -like "*Visio*" -or $_.Description -like "*Office*" -or $_.Description -like "*Project*" -or $_.Description -like "*Visio*")
        }
        foreach ($wp in $wmiProds) {
            $log.AppendLine(" -> Dang go khoa WMI san pham: $($wp.Name)...") | Out-Null
            try { Invoke-CimMethod -InputObject $wp -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
    } catch {}

    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    $log.AppendLine() | Out-Null

    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("                    HOAN TAT QUA TRINH GO BO BAN QUYEN KMS LAU                    ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return $log.ToString()
}

function Start-CrackedAppsUninstallProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("        BAT DAU QUA TRINH GO CAI DAT HANG LOAT CAC APP CRACK BEN THU BA            ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    $crackResult = Get-CrackDetection
    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    
    $crackedAppsFound = @()
    foreach ($risk in $crackResult.Risks) {
        if ($risk -like "*bi Crack*") {
            foreach ($app in $installedApps) {
                if ($risk -like "*$($app.DisplayName)*") {
                    if (-not ($crackedAppsFound | Where-Object { $_.DisplayName -eq $app.DisplayName })) {
                        $crackedAppsFound += $app
                    }
                }
            }
        }
    }
    
    if ($crackedAppsFound.Count -eq 0) {
        $log.AppendLine(" -> [THONG BAO] Khong phat hien thay phan mem ben thu ba nao dang bi be khoa tren he thong.") | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        return @{ Log = $log.ToString(); Count = 0 }
    }

    $log.AppendLine(" -> Phat hien $($crackedAppsFound.Count) phan mem bi be khoa can go hang loat:") | Out-Null
    foreach ($a in $crackedAppsFound) {
        $log.AppendLine("    * $($a.DisplayName)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    $uninstalledCount = 0
    foreach ($app in $crackedAppsFound) {
        $unCmd = $app.QuietUninstallString
        $isQuiet = $true
        if ([string]::IsNullOrWhiteSpace($unCmd)) { 
            $unCmd = $app.UninstallString 
            $isQuiet = $false
        }
        
        if (-not [string]::IsNullOrWhiteSpace($unCmd)) {
            $log.AppendLine("[+] Dang go tu dong ($($uninstalledCount + 1)/$($crackedAppsFound.Count)): $($app.DisplayName)...") | Out-Null
            
            try {
                $exe = ""
                $rawArgs = ""
                $cleanUnCmd = $unCmd.Trim()

                if ($cleanUnCmd -match '^\s*"([^"]+\.[eE][xX][eE])"\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } elseif ($cleanUnCmd -match '^\s*([^\s]+\.[eE][xX][eE])\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } else {
                    $exe = $cleanUnCmd
                    $rawArgs = ""
                }

                if ($exe -and (Test-Path $exe -ErrorAction SilentlyContinue)) {
                    $argList = if ([string]::IsNullOrWhiteSpace($rawArgs)) { "" } else { $rawArgs.Trim() }
                    if (-not $isQuiet) {
                        if ($exe -like "*unins000.exe*" -or $exe -like "*uninstall.exe*") {
                            if ($argList -notlike "*/SILENT*") {
                                $argList = ($argList + " /SILENT /VERYSILENT /SUPPRESSMSGBOXES /NORESTART").Trim()
                            }
                        } elseif ($exe -like "*msiexec.exe*") {
                            if ($argList -notlike "*/qn*") {
                                $argList = ($argList + " /qn /norestart").Trim()
                            }
                        }
                    }
                    $argList = $argList.Trim()

                    $p = if ([string]::IsNullOrWhiteSpace($argList)) {
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    
                    if ($p) { $p.WaitForExit(60000) }
                } else {
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(60000) }
                }
                
                $log.AppendLine(" -> DA GO THANH CONG: $($app.DisplayName)") | Out-Null
                $uninstalledCount++
            } catch {
                $log.AppendLine(" -> Loi khi go $($app.DisplayName): $($_.Exception.Message)") | Out-Null
            }
        }
    }
    
    $log.AppendLine() | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("       HOAN TAT GO CAI DAT HANG LOAT $uninstalledCount PHAN MEM BE KHOA           ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return @{ Log = $log.ToString(); Count = $uninstalledCount }
}

function Start-OfficeUninstallProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("     BAT DAU QUA TRINH GO CAI DAT & LAM SACH OFFICE / PROJECT / VISIO             ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    $scanList = Get-OfficeActivation
    if ($scanList.Count -gt 0) {
        $log.AppendLine(" -> Phat hien $($scanList.Count) phien ban/giay phep Office/Project/Visio tu ket qua quet:") | Out-Null
        foreach ($item in $scanList) {
            $log.AppendLine("    * San pham : $($item.ProductName)") | Out-Null
            $log.AppendLine("      Trang thai: $($item.Status)") | Out-Null
            $log.AppendLine("      Kenh      : $($item.Channel)") | Out-Null
            $log.AppendLine("      Nguon     : $($item.Source)") | Out-Null
            $log.AppendLine("      --------------------------------------------------") | Out-Null
        }
    } else {
        $log.AppendLine(" -> Khong tim thay giay phep Office/Project/Visio nao trong ket qua quet.") | Out-Null
    }
    $log.AppendLine() | Out-Null

    $c2rPaths = @(
        "${env:ProgramFiles}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    )
    foreach ($c2rPath in $c2rPaths) {
        if (Test-Path $c2rPath) {
            $log.AppendLine(" -> Dang chay trinh go sach Click-To-Run ($c2rPath)...") | Out-Null
            try {
                $p = Start-Process -FilePath $c2rPath -ArgumentList "scenario=uninstall forceuninstall=True displaylevel=False" -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                if ($p) {
                    $p.WaitForExit(300000)
                    $log.AppendLine(" -> Da thuc thi trinh go Click-To-Run. Ma thoat: $($p.ExitCode)") | Out-Null
                }
            } catch {
                $log.AppendLine(" -> Loi khi chay Click-To-Run: $($_.Exception.Message)") | Out-Null
            }
        }
    }

    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    $officeRegApps = @()
    foreach ($app in $installedApps) {
        $name = $app.DisplayName
        if ((($name -like "*Microsoft Office*") -or ($name -like "*Microsoft 365*") -or ($name -like "*Microsoft Project*") -or ($name -like "*Microsoft Visio*")) -and 
            ($name -notlike "*Proof*") -and 
            ($name -notlike "*Language*") -and 
            ($name -notlike "*Database Engine*") -and 
            ($name -notlike "*Visual Studio*") -and 
            ($name -notlike "*Viewer*") -and
            ($name -notlike "*Office Web Apps*") -and
            ($name -notlike "*OneDrive*") -and
            ($name -notlike "*Teams*") -and
            ($name -notlike "*Update*") -and
            ($name -notlike "*Desktop App Support*")) {
            $officeRegApps += $app
        }
    }

    foreach ($app in $officeRegApps) {
        $unCmd = $app.QuietUninstallString
        $isQuiet = $true
        if ([string]::IsNullOrWhiteSpace($unCmd)) {
            $unCmd = $app.UninstallString
            $isQuiet = $false
        }
        if (-not [string]::IsNullOrWhiteSpace($unCmd)) {
            $log.AppendLine(" -> Dang go goi Registry: $($app.DisplayName)...") | Out-Null
            try {
                $exe = ""
                $rawArgs = ""
                $cleanUnCmd = $unCmd.Trim()
                if ($cleanUnCmd -match '^\s*"([^"]+\.[eE][xX][eE])"\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } elseif ($cleanUnCmd -match '^\s*([^\s]+\.[eE][xX][eE])\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } else {
                    $exe = $cleanUnCmd
                    $rawArgs = ""
                }
                if ($exe -and (Test-Path $exe -ErrorAction SilentlyContinue)) {
                    $argList = if ([string]::IsNullOrWhiteSpace($rawArgs)) { "" } else { $rawArgs.Trim() }
                    if (-not $isQuiet -and $exe -like "*msiexec.exe*") {
                        if ($argList -notlike "*/qn*") { $argList = ($argList + " /qn /norestart").Trim() }
                    }
                    $p = if ([string]::IsNullOrWhiteSpace($argList)) {
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    if ($p) { $p.WaitForExit(180000) }
                } else {
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(180000) }
                }
                $log.AppendLine("    Da hoan tat go Registry: $($app.DisplayName)") | Out-Null
            } catch {}
        }
    }
    $log.AppendLine() | Out-Null

    $log.AppendLine("[BUOC 3] Dang xoa sach tat ca Product Key va thong tin KMS cua Office / Project / Visio...") | Out-Null
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            $log.AppendLine(" -> Dang thuc thi OSPP tool ($ospp)...") | Out-Null
            try { cscript.exe //NoLogo "$ospp" /remkms 2>&1 | Out-Null } catch {}
            try {
                $dstatusOut = cscript.exe //NoLogo "$ospp" /dstatus 2>&1
                $keysToUninstall = @()
                foreach ($line in $dstatusOut) {
                    if ($line -match "Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})") {
                        $key5 = $matches[1]
                        if ($key5 -and $keysToUninstall -notcontains $key5) {
                            $keysToUninstall += $key5
                        }
                    }
                }
                foreach ($k5 in $keysToUninstall) {
                    $log.AppendLine(" -> Dang go Product Key OSPP (5 ky tu cuoi: $k5)...") | Out-Null
                    cscript.exe //NoLogo "$ospp" /unpkey:$k5 2>&1 | Out-Null
                }
            } catch {}
        }
    }

    try {
        $wmiProducts = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue | Where-Object { 
            ($_.PartialProductKey -or $_.LicenseStatus -eq 1 -or $_.LicenseStatus -eq 5) -and 
            ($_.Name -like "*Office*" -or $_.Name -like "*Project*" -or $_.Name -like "*Visio*" -or $_.Description -like "*Office*" -or $_.Description -like "*Project*" -or $_.Description -like "*Visio*")
        }
        foreach ($p in $wmiProducts) {
            $log.AppendLine(" -> Dang go khoa WMI san pham: $($p.Name)...") | Out-Null
            try {
                Invoke-CimMethod -InputObject $p -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue | Out-Null
            } catch {}
        }
    } catch {}

    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue

    $log.AppendLine() | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("       HOAN TAT GO CAI DAT VA DONG BO LAM SACH HE THONG                           ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return @{ Log = $log.ToString(); Count = ($scanList.Count + $officeRegApps.Count) }
}

function Start-SystemAndNetworkCleanupProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("     BAT DAU QUA TRINH DON DEP TEMP, PREFETCH VA RESET CACHE MANG (IT HELPDESK)    ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    $tempFolders = @(
        $env:TEMP,
        "$env:windir\Temp",
        "$env:windir\Prefetch",
        "$env:windir\SoftwareDistribution\Download"
    )
    $deletedFiles = 0
    $deletedBytes = 0

    foreach ($folder in $tempFolders) {
        if (Test-Path $folder) {
            $log.AppendLine(" -> Dang don dep thu muc: $folder...") | Out-Null
            $files = Get-ChildItem -Path $folder -Recurse -File -ErrorAction SilentlyContinue
            foreach ($f in $files) {
                try {
                    $size = $f.Length
                    Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                    $deletedFiles++
                    $deletedBytes += $size
                } catch {}
            }
        }
    }
    $freedMB = [math]::Round($deletedBytes / 1MB, 2)
    $log.AppendLine(" -> DA XOA THANH CONG: $deletedFiles tep tin rac (Giai phong ~$freedMB MB dung luong o C).") | Out-Null
    $log.AppendLine() | Out-Null

    $log.AppendLine("[BUOC 2] Dang reset cache he thong mang...") | Out-Null
    try {
        $flushRes = ipconfig /flushdns 2>&1
        $log.AppendLine(" -> Xoa Cache DNS (ipconfig /flushdns): $($flushRes -join ' ')") | Out-Null
    } catch {}

    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        $log.AppendLine(" -> Clear-DnsClientCache: Thanh cong.") | Out-Null
    } catch {}

    try {
        $winsockRes = netsh winsock reset 2>&1
        $log.AppendLine(" -> Reset Winsock catalog (netsh winsock reset): Thanh cong.") | Out-Null
    } catch {}

    try {
        $ipResetRes = netsh int ip reset 2>&1
        $log.AppendLine(" -> Reset IP stack (netsh int ip reset): Thanh cong.") | Out-Null
    } catch {}

    $log.AppendLine() | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("             HOAN TAT QUA TRINH DON DEP HE THONG VA RESET MANG                    ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return $log.ToString()
}

# 5. XUAT BAO CAO HTML DANG IT AUDIT REPORT
function Export-HTMLReport {
    param ($winInfo, $officeList, $crackInfo, $logText)
    
    $sys = Get-SystemInfoSummary
    $reportPath = Join-Path $env:USERPROFILE "Desktop\BaoCao_BanQuyen_ITHelpdesk.html"
    
    $officeHtml = ""
    if ($officeList.Count -eq 0) {
        $officeHtml = "<tr><td colspan='4'>Khong phat hien san pham Office / Project / Visio nao.</td></tr>"
    } else {
        foreach ($o in $officeList) {
            $officeHtml += "<tr><td>$($o.ProductName)</td><td>$($o.Status)</td><td>$($o.Channel)</td><td>$($o.KMSServer)</td></tr>"
        }
    }

    $riskHtml = ""
    if ($crackInfo.Risks.Count -eq 0) {
        $riskHtml = "<tr style='background-color:#d1fae5;color:#065f46;'><td colspan='2'><strong>[AN TOAN]</strong> Khong phat hien dau hieu be khoa.</td></tr>"
    } else {
        for ($i = 0; $i -lt $crackInfo.Risks.Count; $i++) {
            $riskHtml += "<tr style='background-color:#fee2e2;color:#991b1b;'><td>$($crackInfo.Risks[$i])</td><td>$($crackInfo.Details[$i])</td></tr>"
        }
    }

    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>BAO CAO KIEM TRA BAN QUYEN IT HELPDESK</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8fafc; color: #1e293b; margin: 20px; }
        .container { max-width: 900px; margin: auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
        h1 { color: #0f172a; border-bottom: 3px solid #3b82f6; padding-bottom: 10px; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; background: #f1f5f9; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 25px; }
        th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: left; }
        th { background-color: #3b82f6; color: white; }
        .log-box { background: #0f172a; color: #38bdf8; padding: 15px; border-radius: 8px; font-family: monospace; white-space: pre-wrap; max-height: 300px; overflow-y: auto; }
    </style>
</head>
<body>
    <div class="container">
        <h1>BAO CAO KIEM TRA BAN QUYEN IT HELPDESK</h1>
        <p>Thoi gian xuat bao cao: <strong>$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</strong></p>
        
        <div class="info-grid">
            <div><strong>Ten may tinh:</strong> $($sys.ComputerName)</div>
            <div><strong>Dia chi IP:</strong> $($sys.IP)</div>
            <div><strong>Dia chi MAC:</strong> $($sys.MAC)</div>
        </div>

        <h2>1. Ban Quyen Windows</h2>
        <table>
            <tr><th>Phien ban</th><td>$($winInfo.ProductName)</td></tr>
            <tr><th>Trang thai</th><td>$($winInfo.StatusText)</td></tr>
            <tr><th>Kenh giay phep</th><td>$($winInfo.Channel)</td></tr>
            <tr><th>May chu KMS</th><td>$($winInfo.KMSServer)</td></tr>
        </table>

        <h2>2. Ban Quyen Office / Project / Visio</h2>
        <table>
            <thead>
                <tr><th>Ten san pham</th><th>Trang thai</th><th>Kenh</th><th>KMS Server</th></tr>
            </thead>
            <tbody>
                $officeHtml
            </tbody>
        </table>

        <h2>3. Nguy Co Crack / Be Khoa</h2>
        <table>
            <thead>
                <tr><th>Dau hieu</th><th>Chi tiet mo ta</th></tr>
            </thead>
            <tbody>
                $riskHtml
            </tbody>
        </table>

        <h2>4. Nhat Ky Xu Ly Chi Tiet</h2>
        <div class="log-box">$logText</div>
    </div>
</body>
</html>
"@

    $htmlContent | Set-Content $reportPath -Encoding utf8 -Force
    return $reportPath
}

# 6. CHE DO DONG LENH (CLI)
if ($ConsoleOnly -or $RemoveKMS -or $UninstallCracks -or $UninstallOffice -or $CleanTemp) {
    if ($RemoveKMS) {
        $logStr = Start-KMSRemovalProcess
        Write-Host $logStr
    } elseif ($UninstallCracks) {
        $res = Start-CrackedAppsUninstallProcess
        Write-Host $res.Log
    } elseif ($UninstallOffice) {
        $res = Start-OfficeUninstallProcess
        Write-Host $res.Log
    } elseif ($CleanTemp) {
        $logStr = Start-SystemAndNetworkCleanupProcess
        Write-Host $logStr
    } else {
        Write-Host "`n==========================================================================" -ForegroundColor Cyan
        Write-Host "             CONG CU KIEM TRA BAN QUYEN WINDOWS, OFFICE & CRACK            " -ForegroundColor White -BackgroundColor DarkBlue
        Write-Host "==========================================================================" -ForegroundColor Cyan
        
        $win = Get-WindowsActivation
        Write-Host "`n[1] BAN QUYEN WINDOWS:" -ForegroundColor Yellow
        Write-Host " - Phien ban  : $($win.ProductName)"
        Write-Host " - Trang thai : $($win.StatusText)"
        
        $offList = Get-OfficeActivation
        Write-Host "`n[2] BAN QUYEN OFFICE / PROJECT / VISIO:" -ForegroundColor Yellow
        foreach ($o in $offList) {
            Write-Host " - $($o.ProductName) | Trang thai: $($o.Status) | Kenh: $($o.Channel)"
        }
        
        $crack = Get-CrackDetection
        Write-Host "`n[3] PHAT HIEN CRACK:" -ForegroundColor Yellow
        if ($crack.Risks.Count -eq 0) {
            Write-Host " -> [AN TOAN] Khong phat hien dau hieu be khoa." -ForegroundColor Green
        } else {
            foreach ($r in $crack.Risks) { Write-Host " -> [CANH BAO] $r" -ForegroundColor Red }
        }
    }
    Exit
}

# 7. GIAO DIEN DO HOA WPF MODERN DARK DASHBOARD
$sysInfo = Get-SystemInfoSummary

$inputXML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IT Helpdesk - Cong cu Kiem tra &amp; Go Ban Quyen" Height="820" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#0F172A" FontFamily="Segoe UI">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- HEADER DASHBOARD CARD -->
        <Border Grid.Row="0" Background="#1E293B" CornerRadius="12" Padding="20" Margin="0,0,0,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="IT HELPDESK DASHBOARD" Foreground="#38BDF8" FontSize="13" FontWeight="Bold"/>
                    <TextBlock Text="Cong Cu Kiem Tra Ban Quyen &amp; Don Dep He Thong" Foreground="#F8FAFC" FontSize="22" FontWeight="Bold" Margin="0,4,0,8"/>
                    <StackPanel Orientation="Horizontal">
                        <Border Background="#334155" CornerRadius="6" Padding="8,4" Margin="0,0,10,0">
                            <TextBlock Text="May: $($sysInfo.ComputerName)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold"/>
                        </Border>
                        <Border Background="#334155" CornerRadius="6" Padding="8,4" Margin="0,0,10,0">
                            <TextBlock Text="IP: $($sysInfo.IP)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold"/>
                        </Border>
                        <Border Background="#334155" CornerRadius="6" Padding="8,4">
                            <TextBlock Text="MAC: $($sysInfo.MAC)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold"/>
                        </Border>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="1" VerticalAlignment="Center" HorizontalAlignment="Right">
                    <Border Background="#10B981" CornerRadius="8" Padding="12,6">
                        <TextBlock Text="[ADMINISTRATOR]" Foreground="White" FontWeight="Bold" FontSize="12"/>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>

        <!-- KPI STATUS CARDS -->
        <Grid Grid.Row="1" Margin="0,0,0,15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="15"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="15"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- CARD 1: WINDOWS STATUS -->
            <Border Grid.Column="0" Background="#1E293B" CornerRadius="10" Padding="15">
                <StackPanel>
                    <TextBlock Text="WINDOWS LICENSE" Foreground="#94A3B8" FontSize="11" FontWeight="Bold"/>
                    <TextBlock x:Name="txtCardWin" Text="Chua quet..." Foreground="#38BDF8" FontSize="16" FontWeight="Bold" Margin="0,6,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>

            <!-- CARD 2: OFFICE STATUS -->
            <Border Grid.Column="2" Background="#1E293B" CornerRadius="10" Padding="15">
                <StackPanel>
                    <TextBlock Text="OFFICE / PROJECT STATUS" Foreground="#94A3B8" FontSize="11" FontWeight="Bold"/>
                    <TextBlock x:Name="txtCardOffice" Text="Chua quet..." Foreground="#38BDF8" FontSize="16" FontWeight="Bold" Margin="0,6,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>

            <!-- CARD 3: RISK WARNING -->
            <Border Grid.Column="4" Background="#1E293B" CornerRadius="10" Padding="15">
                <StackPanel>
                    <TextBlock Text="SECURITY RISK LEVEL" Foreground="#94A3B8" FontSize="11" FontWeight="Bold"/>
                    <TextBlock x:Name="txtCardRisk" Text="Chua kiem tra" Foreground="#10B981" FontSize="16" FontWeight="Bold" Margin="0,6,0,0" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- ACTION TOOLBAR GROUPS -->
        <Border Grid.Row="2" Background="#1E293B" CornerRadius="12" Padding="15" Margin="0,0,0,15">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- ROW 1: BAN QUYEN & CRACK -->
                <StackPanel Grid.Row="0" Margin="0,0,0,10">
                    <TextBlock Text="1. QUAN LY BAN QUYEN &amp; CRACK" Foreground="#94A3B8" FontSize="11" FontWeight="Bold" Margin="0,0,0,8"/>
                    <WrapPanel>
                        <Button x:Name="btnScan" Content="BAT DAU QUET" Width="150" Height="38" Background="#3B82F6" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnRemoveKms" Content="GO BO KMS LAU" Width="150" Height="38" Background="#EF4444" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnUninstallCracks" Content="GO APP CRACK" Width="150" Height="38" Background="#F59E0B" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnUninstallOffice" Content="GO BO OFFICE/PROJECT" Width="170" Height="38" Background="#DC2626" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                    </WrapPanel>
                </StackPanel>

                <!-- ROW 2: CONG CU IT HELPDESK -->
                <StackPanel Grid.Row="1">
                    <TextBlock Text="2. CONG CU HO TRO IT HELPDESK &amp; QUAN LY TAI SAN" Foreground="#94A3B8" FontSize="11" FontWeight="Bold" Margin="0,0,0,8"/>
                    <WrapPanel>
                        <Button x:Name="btnCleanTemp" Content="DON DEP TEMP &amp; DNS" Width="160" Height="38" Background="#10B981" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnActivateKey" Content="KICH HOAT KEY MOI" Width="150" Height="38" Background="#8B5CF6" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnAssetInfo" Content="THONG TIN TAI SAN" Width="160" Height="38" Background="#EC4899" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnExportAsset" Content="XUAT TAI SAN (CSV)" Width="150" Height="38" Background="#14B8A6" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                        <Button x:Name="btnExportReport" Content="XUAT BAO CAO HTML" Width="150" Height="38" Background="#06B6D4" Foreground="White" FontWeight="Bold" BorderThickness="0" Margin="0,0,10,5" Cursor="Hand"/>
                    </WrapPanel>
                </StackPanel>
            </Grid>
        </Border>

        <!-- LOG TERMINAL CONSOLE -->
        <Grid Grid.Row="3">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Grid Grid.Row="0" Margin="0,0,0,6">
                <TextBlock Text="NHAT KY XU LY (TERMINAL LOG)" Foreground="#94A3B8" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnCopyLog" Content="SAO CHEP LOG" Width="110" Height="26" Background="#334155" Foreground="White" FontSize="11" FontWeight="SemiBold" BorderThickness="0" Margin="0,0,8,0" Cursor="Hand"/>
                    <Button x:Name="btnClearLog" Content="XOA LOG" Width="90" Height="26" Background="#334155" Foreground="White" FontSize="11" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
            </Grid>

            <Border Grid.Row="1" Background="#090D16" CornerRadius="8" Padding="12">
                <TextBox x:Name="txtLog" Foreground="#38BDF8" Background="Transparent" BorderThickness="0" FontFamily="Consolas" FontSize="12.5" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/>
            </Border>
        </Grid>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader ([xml]$inputXML))
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Lay Control
$btnScan = $Window.FindName("btnScan")
$btnRemoveKms = $Window.FindName("btnRemoveKms")
$btnUninstallCracks = $Window.FindName("btnUninstallCracks")
$btnUninstallOffice = $Window.FindName("btnUninstallOffice")
$btnCleanTemp = $Window.FindName("btnCleanTemp")
$btnActivateKey = $Window.FindName("btnActivateKey")
$btnAssetInfo = $Window.FindName("btnAssetInfo")
$btnExportAsset = $Window.FindName("btnExportAsset")
$btnExportReport = $Window.FindName("btnExportReport")
$btnCopyLog = $Window.FindName("btnCopyLog")
$btnClearLog = $Window.FindName("btnClearLog")

$txtCardWin = $Window.FindName("txtCardWin")
$txtCardOffice = $Window.FindName("txtCardOffice")
$txtCardRisk = $Window.FindName("txtCardRisk")
$txtLog = $Window.FindName("txtLog")

# Biến lưu trữ kết quả
$global:lastWinInfo = $null
$global:lastOfficeList = $null
$global:lastCrackInfo = $null

# HAM CAP NHAT CARD DASHBOARD
function Update-DashboardCards {
    param ($win, $officeList, $crack)
    
    if ($win) {
        $txtCardWin.Text = "$($win.ProductName)`n$($win.StatusText)"
        if ($win.Status -eq "Licensed" -and $win.KMSServer -eq "Khong phat hien") {
            $txtCardWin.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
        } else {
            $txtCardWin.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
        }
    }

    if ($officeList) {
        if ($officeList.Count -eq 0) {
            $txtCardOffice.Text = "Khong phat hien Office"
            $txtCardOffice.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#94A3B8")
        } else {
            $txtCardOffice.Text = "Tim thấy $($officeList.Count) giay phep Office/Project"
            $txtCardOffice.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#38BDF8")
        }
    }

    if ($crack) {
        if ($crack.Risks.Count -eq 0) {
            $txtCardRisk.Text = "[AN TOAN] Khong phat hien crack"
            $txtCardRisk.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
        } else {
            $txtCardRisk.Text = "[CANH BAO] $($crack.Risks.Count) dau hieu nguy hiem!"
            $txtCardRisk.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
        }
    }
}

# HELPER ASYNC UI DISPATCHER
function Invoke-UIAsync {
    param ([scriptblock]$Action)
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{
        & $Action
    }, [System.Windows.Threading.DispatcherPriority]::Background)
}

# EVENT BINDING
$btnScan.Add_Click({
    $btnScan.IsEnabled = $false
    $txtLog.Text = "Dang thuc hien quet he thong xin vui long cho...`n"
    Invoke-UIAsync {
        $win = Get-WindowsActivation
        $offices = Get-OfficeActivation
        $crack = Get-CrackDetection
        
        $global:lastWinInfo = $win
        $global:lastOfficeList = $offices
        $global:lastCrackInfo = $crack

        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("==================================================================================") | Out-Null
        $sb.AppendLine("                     KET QUA KIEM TRA BAN QUYEN & HE THONG                        ") | Out-Null
        $sb.AppendLine("==================================================================================") | Out-Null
        $sb.AppendLine("Thoi gian quet : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
        $sb.AppendLine("Ten may tinh   : $env:COMPUTERNAME") | Out-Null
        $sb.AppendLine("----------------------------------------------------------------------------------") | Out-Null
        $sb.AppendLine() | Out-Null
        
        $sb.AppendLine("[1] KET QUA KIEM TRA WINDOWS:") | Out-Null
        $sb.AppendLine($win.Details) | Out-Null
        $sb.AppendLine() | Out-Null
        
        $sb.AppendLine("[2] KET QUA KIEM TRA MICROSOFT OFFICE / PROJECT / VISIO:") | Out-Null
        if ($offices.Count -eq 0) {
            $sb.AppendLine(" -> Khong tim thay san pham Office / Project / Visio nao duoc dang ky.") | Out-Null
        } else {
            foreach ($o in $offices) {
                $sb.AppendLine(" -> Ten san pham : $($o.ProductName)") | Out-Null
                $sb.AppendLine("    Trang thai   : $($o.Status)") | Out-Null
                $sb.AppendLine("    Kenh         : $($o.Channel)") | Out-Null
                if ($o.KMSServer) { $sb.AppendLine("    KMS Server   : $($o.KMSServer)") | Out-Null }
                $sb.AppendLine("    --------------------------------------------------") | Out-Null
            }
        }
        $sb.AppendLine() | Out-Null
        
        $sb.AppendLine("[3] KET QUA QUET DAU VET CRACK / BE KHOA:") | Out-Null
        if ($crack.Risks.Count -eq 0) {
            $sb.AppendLine(" -> [AN TOAN] Khong phat hien bat ky dau hieu be khoa he thong nao.") | Out-Null
        } else {
            $sb.AppendLine(" -> [CANH BAO] Phat hien $($crack.Risks.Count) dau hieu nguy hiem:") | Out-Null
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                $sb.AppendLine("    * Dau hieu: $($crack.Risks[$i])") | Out-Null
                $sb.AppendLine("      Chi tiet : $($crack.Details[$i])") | Out-Null
            }
        }
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("==================================================================================") | Out-Null
        
        $txtLog.Text = $sb.ToString()
        Update-DashboardCards -win $win -officeList $offices -crack $crack
        $btnScan.IsEnabled = $true
    }
})

$btnRemoveKms.Add_Click({
    $btnRemoveKms.IsEnabled = $false
    $txtLog.Text = "Dang thuc hien go bo ban quyen KMS lau xin vui long cho...`n"
    Invoke-UIAsync {
        $logOut = Start-KMSRemovalProcess
        $txtLog.Text = $logOut
        $btnRemoveKms.IsEnabled = $true
    }
})

$btnUninstallCracks.Add_Click({
    $btnUninstallCracks.IsEnabled = $false
    $txtLog.Text = "Dang thuc hien go bo phan mem crack xin vui long cho...`n"
    Invoke-UIAsync {
        $res = Start-CrackedAppsUninstallProcess
        $txtLog.Text = $res.Log
        $btnUninstallCracks.IsEnabled = $true
    }
})

$btnUninstallOffice.Add_Click({
    $btnUninstallOffice.IsEnabled = $false
    $txtLog.Text = "Dang thuc hien go bo Office / Project / Visio xin vui long cho...`n"
    Invoke-UIAsync {
        $res = Start-OfficeUninstallProcess
        $txtLog.Text = $res.Log
        $btnUninstallOffice.IsEnabled = $true
    }
})

$btnCleanTemp.Add_Click({
    $btnCleanTemp.IsEnabled = $false
    $txtLog.Text = "Dang don dep he thong va reset cache mang xin vui long cho...`n"
    Invoke-UIAsync {
        $logOut = Start-SystemAndNetworkCleanupProcess
        $txtLog.Text = $logOut
        $btnCleanTemp.IsEnabled = $true
    }
})

$btnActivateKey.Add_Click({
    $keyInput = [Microsoft.VisualBasic.Interaction]::InputBox("Nhap Product Key chinh hang (Windows hoac Office):`n(Vi du: VK7JG-NPHTM-C97JM-9MPGT-3V66T)", "Kich Hoat Ban Quyen Chinh Hang", "")
    if (-not [string]::IsNullOrWhiteSpace($keyInput)) {
        $cleanKey = $keyInput.Trim()
        $txtLog.Text = "Dang tien hanh kich hoat khoa ban quyen: $cleanKey ...`n"
        $btnActivateKey.IsEnabled = $false
        Invoke-UIAsync {
            $log = New-Object System.Text.StringBuilder
            $log.AppendLine("Dang cai dat Product Key Windows: $cleanKey ...") | Out-Null
            try {
                $res1 = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ipk $cleanKey 2>&1
                $log.AppendLine(" -> Ket qua slmgr /ipk: $($res1 -join ' ')") | Out-Null
                $res2 = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ato 2>&1
                $log.AppendLine(" -> Ket qua slmgr /ato: $($res2 -join ' ')") | Out-Null
            } catch {
                $log.AppendLine(" -> Loi khi kich hoat Windows: $($_.Exception.Message)") | Out-Null
            }
            $txtLog.Text = $log.ToString()
            $btnActivateKey.IsEnabled = $true
        }
    }
})

$btnExportReport.Add_Click({
    if (-not $global:lastWinInfo) {
        $global:lastWinInfo = Get-WindowsActivation
        $global:lastOfficeList = Get-OfficeActivation
        $global:lastCrackInfo = Get-CrackDetection
    }
    
    try {
        $filePath = Export-HTMLReport -winInfo $global:lastWinInfo -officeList $global:lastOfficeList -crackInfo $global:lastCrackInfo -logText $txtLog.Text
        [System.Windows.MessageBox]::Show("Da xuat bao cao HTML thanh cong tai:`n$filePath", "Xuat Bao Cao Thanh Cong", "OK", "Information")
        Start-Process $filePath -ErrorAction SilentlyContinue
    } catch {
        [System.Windows.MessageBox]::Show("Loi khi xuat bao cao HTML: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})

$btnAssetInfo.Add_Click({
    $btnAssetInfo.IsEnabled = $false
    $txtLog.Text = "Dang thu thap thong tin tai san IT chi tiet xin vui long cho...`n"
    Invoke-UIAsync {
        $asset = Get-FullComputerAssetInfo
        $global:lastAssetInfo = $asset

        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine("==================================================================================") | Out-Null
        $sb.AppendLine("         THONG TIN TAI SAN THIET BI IT CHI TIET (DEEP IT ASSET AUDIT REPORT)       ") | Out-Null
        $sb.AppendLine("==================================================================================") | Out-Null
        $sb.AppendLine("Thoi gian thu thap : $($asset.AuditTimestamp)") | Out-Null
        $sb.AppendLine("Ten may tinh       : $($asset.ComputerName)") | Out-Null
        $sb.AppendLine("----------------------------------------------------------------------------------") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("[1] HE THONG & BO MACH CHU (SYSTEM & MAINBOARD):") | Out-Null
        $sb.AppendLine(" -> Ten may tinh     : $($asset.ComputerName)") | Out-Null
        $sb.AppendLine(" -> Nha san xuat     : $($asset.SystemManufacturer)") | Out-Null
        $sb.AppendLine(" -> Model thiet bi   : $($asset.SystemModel)") | Out-Null
        $sb.AppendLine(" -> Serial (Service): $($asset.BIOSSerialNumber)") | Out-Null
        $sb.AppendLine(" -> System UUID      : $($asset.SystemUUID)") | Out-Null
        $sb.AppendLine(" -> Mainboard Hang   : $($asset.MotherboardManufacturer)") | Out-Null
        $sb.AppendLine(" -> Mainboard Model  : $($asset.MotherboardProduct)") | Out-Null
        $sb.AppendLine(" -> Mainboard Serial : $($asset.MotherboardSerial)") | Out-Null
        $sb.AppendLine(" -> BIOS Version     : $($asset.BIOSVersion)") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("[2] HE DIEU HANH & NGUOI DUNG:") | Out-Null
        $sb.AppendLine(" -> Ten he dieu hanh : $($asset.OSName) ($($asset.OSArchitecture))") | Out-Null
        $sb.AppendLine(" -> Phien ban OS     : $($asset.OSVersion)") | Out-Null
        $sb.AppendLine(" -> Ngay cai dat OS  : $($asset.OSInstallDate)") | Out-Null
        $sb.AppendLine(" -> Nguoi dung       : $($asset.UserDomain)\$($asset.CurrentUser)") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("[3] BO NHO RAM & CARD DO HOA (MEMORY & GPU):") | Out-Null
        $sb.AppendLine(" -> Tong RAM he thong: $($asset.TotalRAM_GB) GB ($($asset.RAMSlotsCount) thanh RAM)") | Out-Null
        $sb.AppendLine(" -> Chi tiet thanh RAM: $($asset.RAMSticksDetail)") | Out-Null
        $sb.AppendLine(" -> Card do hoa (GPU): $($asset.GPU) (VRAM: $($asset.GPUVRAM_GB))") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("[4] CHI TIET OU DIA & PHAN VUNG (STORAGE & DRIVES):") | Out-Null
        $sb.AppendLine(" -> Tong dung luong o: $($asset.TotalStorageCapacityGB) GB (Trong: $($asset.TotalFreeStorageGB) GB)") | Out-Null
        $sb.AppendLine(" -> O đia vat ly      : $($asset.PhysicalDisksDetail)") | Out-Null
        $sb.AppendLine(" -> Phan vung o đia   : $($asset.LogicalDrivesDetail)") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("[5] THONG TIN MANG (NETWORK):") | Out-Null
        $sb.AppendLine(" -> Card mang (NIC)  : $($asset.NetworkAdapter)") | Out-Null
        $sb.AppendLine(" -> Dia chi IP       : $($asset.IPAddress)") | Out-Null
        $sb.AppendLine(" -> Dia chi MAC      : $($asset.MACAddress)") | Out-Null
        $sb.AppendLine() | Out-Null
        $sb.AppendLine("==================================================================================") | Out-Null

        $txtLog.Text = $sb.ToString()
        $btnAssetInfo.IsEnabled = $true
    }
})

$btnExportAsset.Add_Click({
    if (-not $global:lastAssetInfo) {
        $global:lastAssetInfo = Get-FullComputerAssetInfo
    }
    try {
        $csvPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($global:lastAssetInfo.ComputerName).csv"
        $jsonPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($global:lastAssetInfo.ComputerName).json"
        
        $global:lastAssetInfo | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8 -Force
        $global:lastAssetInfo | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding utf8 -Force
        
        [System.Windows.MessageBox]::Show("Da xuat thong tin tai san thanh cong ra Desktop:`n* CSV: $csvPath`n* JSON: $jsonPath", "Xuat Tai San Thanh Cong", "OK", "Information")
    } catch {
        [System.Windows.MessageBox]::Show("Loi khi xuat file tai san: $($_.Exception.Message)", "Loi", "OK", "Error")
    }
})

$btnCopyLog.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($txtLog.Text)) {
        [System.Windows.Clipboard]::SetText($txtLog.Text)
        [System.Windows.MessageBox]::Show("Da sao chep toan bo nhat ky log vao Clipboard!", "Thong Bao", "OK", "Information")
    }
})

$btnClearLog.Add_Click({
    $txtLog.Text = ""
})

# CHO HIEN THI CUA SO GIAO DIEN
$Window.ShowDialog() | Out-Null
