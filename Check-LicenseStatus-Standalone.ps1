<#
    Check-LicenseStatus-Standalone.ps1 - Cong cu TAT-CA-TRONG-MOT (All-In-One Edition)
    Bao gom:
    1. Kiem tra ban quyen Windows & Microsoft Office.
    2. Quet dong cac phan mem ben thu ba bi crack (MiniTool, Adobe, IDM, AutoCAD...).
    3. GO CAI DAT TU DONG HANG LOAT (Bulk/Batch Uninstall) cac phan mem bi crack.
    4. Go bo ban quyen KMS lau (slmgr /ckms, /upk, /cpky), lam sach file Hosts, go bo tac vu/dich vu be khoa.
#>

[CmdletBinding()]
param (
    [switch]$ConsoleOnly,
    [switch]$RemoveKMS,
    [switch]$UninstallCracks,
    [switch]$UninstallOffice
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

# 3. CAC HAM QUET BAN QUYEN VA PHAT HIEN BE KHOA
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
                0 { $result.StatusText = "Chua kich hoat (Unlicensed)"; $result.Status = "Unlicensed" }
                1 { $result.StatusText = "Da kich hoat (Licensed)"; $result.Status = "Licensed" }
                2 { $result.StatusText = "An han ban dau (OOB Grace)"; $result.Status = "Grace" }
                3 { $result.StatusText = "An han het han (OOT Grace)"; $result.Status = "Grace" }
                4 { $result.StatusText = "An han khong chinh hang (Non-genuine)"; $result.Status = "Warning" }
                5 { $result.StatusText = "Cho kich hoat / Het han (Notification)"; $result.Status = "Notification" }
                6 { $result.StatusText = "An han mo rong (Extended Grace)"; $result.Status = "Grace" }
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
            
            $result.Details = "San pham: $($win.Name)`nPhan nhom: $($result.Channel)`nTrang thai: $($result.StatusText)`nMay chu kich hoat KMS: $($result.KMSServer)"
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
            break
        }
    }
    
    $c2rInstalled = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Microsoft 365*" }
    if ($c2rInstalled -and ($officeList.Count -eq 0 -or -not ($officeList | Where-Object { $_.ProductName -like "*365*" -or $_.ProductName -like "*Subscription*" }))) {
        $officeList += [PSCustomObject]@{
            ProductName = "Microsoft 365 Apps (Dang ky thue bao)"
            Status = "Da cai dat (Kich hoat qua Tai khoan Microsoft)"
            Channel = "Subscription (ClickToRun)"
            KMSServer = ""
            GraceRemaining = "N/A"
            Source = "Registry"
        }
    }
    return $officeList
}

function Get-CrackDetection {
    $risks = @()
    $details = @()
    
    # 1. Kiem tra Hook DLL
    $hookPaths = @(
        "$env:windir\System32\SppExtComObjHook.dll",
        "$env:windir\SysWOW64\SppExtComObjHook.dll",
        "$env:windir\SppExtComObjHook.dll",
        "$env:windir\System32\sppextcomobj.exe.local",
        "$env:windir\System32\sppextcomobj_hook.dll"
    )
    foreach ($path in $hookPaths) {
        if (Test-Path $path) {
            $risks += "Phat hien file hook KMS: $path"
            $details += "Tep tin '$path' thuong do KMSpico hoac KMSAuto cai dat de bypass may chu ban quyen cua Microsoft."
        }
    }
    
    # 2. Kiem tra file Hosts
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        $hostsLines = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $redirects = @()
        $thirdPartyDomains = @{
            "minitool.com" = "MiniTool Partition Wizard / Data Recovery"
            "adobe.com" = "Adobe Creative Cloud (Photoshop, Illustrator...)"
            "adobelogin.com" = "Adobe Creative Cloud"
            "autodesk.com" = "Autodesk (AutoCAD, 3ds Max...)"
            "corel.com" = "CorelDraw"
            "registeridm.com" = "Internet Download Manager (IDM)"
            "internetdownloadmanager.com" = "Internet Download Manager (IDM)"
            "sketchup.com" = "Trimble SketchUp"
            "ccleaner.com" = "CCleaner"
            "chaosgroup.com" = "Chaos Group V-Ray"
            "solidworks.com" = "SolidWorks"
        }
        $blockedApps = @()
        foreach ($line in $hostsLines) {
            $clean = $line.Trim()
            if ($clean -and -not $clean.StartsWith("#")) {
                if ($clean -like "*microsoft*" -or $clean -like "*activation*" -or $clean -like "*kms*") {
                    if ($clean -match "(127\.0\.0\.1|0\.0\.0\.0)\s+(.*)") { $redirects += $matches[2] }
                }
                foreach ($domain in $thirdPartyDomains.Keys) {
                    if ($clean -like "*$domain*" -and $clean -match "(127\.0\.0\.1|0\.0\.0\.0)") {
                        $appName = $thirdPartyDomains[$domain]
                        if ($blockedApps -notcontains $appName) { $blockedApps += $appName }
                    }
                }
            }
        }
        if ($redirects.Count -gt 0) {
            $risks += "Sua doi tep tin Hosts (Dinh tuyen chan Microsoft)"
            $details += "Tep tin Hosts chan ket noi xac minh cua Microsoft (vi du: $($redirects -join ', '))."
        }
        if ($blockedApps.Count -gt 0) {
            $risks += "Chan may chu xac thuc ban quyen phan mem trong file Hosts"
            $details += "Tep Hosts chan ket noi den hang phan mem: $($blockedApps -join ', '). Cho thay cac phan mem nay dang dung ban crack."
        }
    }
    
    # 3. Kiem tra Scheduled Tasks
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($task in $tasks) {
            $risks += "Tac vu chay ngam be khoa: $($task.TaskName)"
            $details += "Tac vu tu dong gia han KMS lau chay ngam tai '$($task.TaskPath)'."
        }
    }
    
    # 4. Kiem tra Services
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            $risks += "Dich vu KMS lau chay nen: $($srv.DisplayName) ($srvName)"
            $details += "Dich vu gia lap KMS noi bo dang hoat dong."
        }
    }

    # 5. Kiem tra Processes
    $suspiciousProcessNames = @("KMSpico", "AutoKMS", "KMSAuto", "KMSConnectionMonitor", "HEU_KMS", "HEU_KMS_Activator")
    foreach ($procName in $suspiciousProcessNames) {
        $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($proc) {
            $risks += "Tien trinh be khoa dang chay: $procName"
            $details += "Tien trinh '$procName' (PID: $($proc.Id)) dang chay ngam thoi gian thuc."
        }
    }
    
    # 6. Kiem tra cac thu muc chua cong cu be khoa he thong
    $crackFolders = @(
        "${env:ProgramFiles}\KMSpico",
        "${env:ProgramFiles(x86)}\KMSpico",
        "${env:ProgramData}\KMSAuto",
        "${env:ProgramData}\KMSAutoS",
        "${env:ProgramFiles(x86)}\Microsoft Toolkit",
        "C:\Windows\Setup\Scripts"
    )
    foreach ($folder in $crackFolders) {
        if (Test-Path $folder) {
            if ($folder -eq "C:\Windows\Setup\Scripts") {
                $scripts = Get-ChildItem -Path $folder -Filter "*.bat" -ErrorAction SilentlyContinue
                if ($scripts) {
                    $risks += "Thu muc he thong chua Script kich hoat lau: $folder"
                    $details += "Thu muc chua script *.bat cau hinh ban quyen tu dong khi cai Win."
                }
            } else {
                $risks += "Phat hien thu muc phan mem be khoa he thong: $folder"
                $details += "Thu muc cai dat ung dung be khoa ton tai o '$folder'."
            }
        }
    }
    
    # 7. QUET DONG PHAN MEM BEN THU 3 (Dynamic Third-Party Crack Scan)
    $targetKeywords = @("MiniTool", "Adobe", "Autodesk", "AutoCAD", "Corel", "Camtasia", "TechSmith", "Internet Download Manager", "IDM", "Wondershare", "CCleaner", "SketchUp", "SolidWorks", "WinZip", "WinRAR")
    $targetFilesMap = @{
        "MiniTool" = @("partitionwizard.dll", "partitionwizard.exe", "PowerDataRecoveryCore.dll", "PowerDataRecovery.exe")
        "Internet Download Manager" = @("idm.core.dll", "IDMan.exe", "IDMGrHlp.exe")
        "IDM" = @("idm.core.dll", "IDMan.exe", "IDMGrHlp.exe")
        "Adobe" = @("amtlib.dll", "Acrobat.exe", "Photoshop.exe", "Illustrator.exe")
        "Autodesk" = @("acad.exe", "AdskLicensingAgent.exe")
        "Corel" = @("CorelDRW.exe", "PASMUTILITY.dll")
        "SketchUp" = @("SketchUp.exe")
        "Camtasia" = @("CamtasiaStudio.exe", "Camtasia.exe")
        "Wondershare" = @("WondershareHelper.exe", "Filmora.exe")
        "CCleaner" = @("CCleaner.exe", "CCleaner64.exe")
        "SolidWorks" = @("sldworks.exe")
    }

    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName } |
        Select-Object DisplayName, InstallLocation, UninstallString, QuietUninstallString
        
    foreach ($app in $installedApps) {
        $matchedKeyword = $null
        foreach ($kw in $targetKeywords) {
            if ($app.DisplayName -like "*$kw*") {
                $matchedKeyword = $kw
                break
            }
        }
        if ($matchedKeyword) {
            $installDir = $app.InstallLocation
            if (-not $installDir -and $app.UninstallString) {
                if ($app.UninstallString -match '^["]?([^"]+\.[eE][xX][eE])["]?') {
                    $exePath = $matches[1]
                    if (Test-Path $exePath) { $installDir = Split-Path $exePath }
                }
            }
            if ($installDir -and (Test-Path $installDir)) {
                $mapKey = $null
                foreach ($k in $targetFilesMap.Keys) {
                    if ($matchedKeyword -like "*$k*" -or $k -like "*$matchedKeyword*") {
                        $mapKey = $k
                        break
                    }
                }
                if ($mapKey) {
                    $filesToAudit = $targetFilesMap[$mapKey]
                    foreach ($fileName in $filesToAudit) {
                        $foundFiles = Get-ChildItem -Path $installDir -Filter $fileName -Recurse -Depth 2 -ErrorAction SilentlyContinue
                        foreach ($file in $foundFiles) {
                            $sig = Get-AuthenticodeSignature -FilePath $file.FullName -ErrorAction SilentlyContinue
                            if ($sig -and $sig.Status -ne 'Valid') {
                                $risks += "Phat hien $($app.DisplayName) bi Crack (Mat chu ky so)"
                                $details += "Tep tin '$($file.FullName)' cua '$($app.DisplayName)' mat chu ky so hop le (Trang thai: $($sig.Status)), chung to tep tin nay da bi crack (patch)."
                            }
                        }
                    }
                }
            }
        }
    }
    
    # 8. Cau hinh KMS cua Windows & Office
    $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
    if ($licService -and $licService.KeyManagementServiceMachine) {
        $kmsMachine = $licService.KeyManagementServiceMachine
        $knownKmsServers = @("kms8.msguides.com", "kms.digiboy.ir", "kms.lotro.cc", "kms.chinancce.com", "zh.us.to", "kms.msguides.com", "kms.spacespaces.xyz")
        if ($kmsMachine -eq "127.0.0.1" -or $kmsMachine -eq "localhost" -or $kmsMachine -eq "::1") {
            $risks += "Windows cau hinh may chu KMS noi bo (Localhost)"
            $details += "Ban quyen Windows tro ve chinh may tinh ($kmsMachine). Dau hieu be khoa gia lap cuc bo."
        } else {
            foreach ($srv in $knownKmsServers) {
                if ($kmsMachine -like "*$srv*") {
                    $risks += "Windows su dung may chu KMS be khoa cong cong: $kmsMachine"
                    $details += "Ban quyen he thong tro toi may chu KMS mien phi tren Internet."
                    break
                }
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

    # 9. Kiem tra cong KMS 1688
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

    # 10. Kiem tra Registry cua phan mem be khoa
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

# 4. HAM GO CAI DAT HANG LOAT (BULK UNINSTALL) CAC PHAN MEM CRACK BEN THU 3
function Start-CrackedAppsUninstallProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("          BAT DAU QUA TRINH GO CAI DAT HANG LOAT (BULK UNINSTALL) CAC APP CRACK   ") | Out-Null
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
                        $log.AppendLine("    Lenh thuc thi: `"$exe`"") | Out-Null
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        $log.AppendLine("    Lenh thuc thi: `"$exe`" $argList") | Out-Null
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    
                    if ($p) { $p.WaitForExit(60000) }
                } else {
                    $log.AppendLine("    Thuc thi lenh CMD: $cleanUnCmd") | Out-Null
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(60000) }
                }
                
                $log.AppendLine(" -> DA GO THANH CONG: $($app.DisplayName)") | Out-Null
                $uninstalledCount++
            } catch {
                $log.AppendLine(" -> Loi khi go $($app.DisplayName): $($_.Exception.Message)") | Out-Null
            }
        } else {
            $log.AppendLine(" -> Bo qua $($app.DisplayName): Khong tim thay UninstallString trong Registry.") | Out-Null
        }
        $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    }
    
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("       HOAN TAT GO CAI DAT HANG LOAT $uninstalledCount / $($crackedAppsFound.Count) PHAN MEM BE KHOA           ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return @{ Log = $log.ToString(); Count = $uninstalledCount }
}

# 5. HAM GO BO KMS LAU VA DON DEP HE THONG
function Start-KMSRemovalProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("              BAT DAU QUA TRINH GO BO BAN QUYEN KMS LAU VA BAO TRI SYSTEM         ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    # Buoc 1: Xoa cau hinh may chu KMS (slmgr /ckms)
    $log.AppendLine("[BUOC 1] Dang go bo cau hinh may chu KMS (Clear KMS Server)...") | Out-Null
    try {
        $ckmsResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ckms 2>&1
        $log.AppendLine(" -> Ket qua slmgr /ckms: $($ckmsResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /ckms: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Buoc 2: Go bo khoa san pham KMS hien tai (slmgr /upk)
    $log.AppendLine("[BUOC 2] Dang go bo Product Key KMS hien tai (Uninstall Product Key)...") | Out-Null
    try {
        $upkResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /upk 2>&1
        $log.AppendLine(" -> Ket qua slmgr /upk: $($upkResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /upk: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Buoc 3: Xoa thong tin Product Key khoi Registry (slmgr /cpky)
    $log.AppendLine("[BUOC 3] Dang xoa thong tin Product Key khoi Registry (Clear Product Key from Registry)...") | Out-Null
    try {
        $cpkyResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /cpky 2>&1
        $log.AppendLine(" -> Ket qua slmgr /cpky: $($cpkyResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Loi khi chay slmgr /cpky: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Buoc 4: Lam sach tep tin Hosts (Go chan cac may chu xac minh Microsoft)
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
                $log.AppendLine(" -> Da lam sach $removedCount dong cau hinh chan trong file Hosts thanh cong!") | Out-Null
            } else {
                $log.AppendLine(" -> Tep Hosts khong chua cau hinh chan may chu Microsoft.") | Out-Null
            }
        } catch {
            $log.AppendLine(" -> Loi khi chinh sua file Hosts: $($_.Exception.Message)") | Out-Null
        }
    }
    $log.AppendLine() | Out-Null

    # Buoc 5: Tim va go bo Tac vu chay ngam be khoa (Scheduled Tasks)
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
    $log.AppendLine() | Out-Null

    # Buoc 6: Tim va dung/go bo Dich vu KMS lau (Services)
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

    # Buoc 7: Go bo khoa san pham Office/Project/Visio KMS con sot lai
    $log.AppendLine("[BUOC 7] Dang go bo cac Product Key Office/Project/Visio KMS lau...") | Out-Null
    
    # 7.1 Dung WMI class OfficeSoftwareProtectionProduct (chuyen dung cho Office)
    try {
        $products = Get-CimInstance -ClassName OfficeSoftwareProtectionProduct -Filter "PartialProductKey is not null" -ErrorAction SilentlyContinue
        $removedCount = 0
        foreach ($p in $products) {
            $isKMS = ($p.Description -like "*KMS*" -or $p.Description -like "*VOLUME*" -or $p.Name -like "*VL*" -or $p.Name -like "*KMS*")
            $isNotRetail = ($p.Description -notlike "*RETAIL*" -and $p.Description -notlike "*SUBSCRIPTION*")
            if ($isKMS -and $isNotRetail) {
                $log.AppendLine(" -> [WMI Office] Dang go khoa san pham: $($p.Name) (Key: $($p.PartialProductKey))") | Out-Null
                Invoke-CimMethod -InputObject $p -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue
                $removedCount++
            }
        }
    } catch {
        $log.AppendLine(" -> Loi khi go qua WMI Office: $($_.Exception.Message)") | Out-Null
    }

    # 7.2 Dung WMI class SoftwareLicensingProduct (du phong)
    try {
        $products = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
        foreach ($p in $products) {
            $isKMS = ($p.Description -like "*KMS*" -or $p.Description -like "*VOLUME*" -or $p.Name -like "*VL*" -or $p.Name -like "*KMS*")
            $isNotRetail = ($p.Description -notlike "*RETAIL*" -and $p.Description -notlike "*SUBSCRIPTION*")
            if ($isKMS -and $isNotRetail) {
                $log.AppendLine(" -> [WMI Windows] Dang go khoa san pham: $($p.Name) (Key: $($p.PartialProductKey))") | Out-Null
                Invoke-CimMethod -InputObject $p -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue
            }
        }
    } catch {}

    # 7.3 Dung cscript.exe va ospp.vbs (Du phong cuoi cung cuc ky tin cay)
    try {
        $keysToUninstall = @()
        $p1 = Get-CimInstance -ClassName OfficeSoftwareProtectionProduct -Filter "PartialProductKey is not null" -ErrorAction SilentlyContinue
        $p2 = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
        
        foreach ($p in ($p1 + $p2)) {
            if ($p.PartialProductKey) {
                $isKMS = ($p.Description -like "*KMS*" -or $p.Description -like "*VOLUME*" -or $p.Name -like "*VL*" -or $p.Name -like "*KMS*")
                $isNotRetail = ($p.Description -notlike "*RETAIL*" -and $p.Description -notlike "*SUBSCRIPTION*")
                if ($isKMS -and $isNotRetail) {
                    if ($keysToUninstall -notcontains $p.PartialProductKey) {
                        $keysToUninstall += $p.PartialProductKey
                    }
                }
            }
        }
        
        if ($keysToUninstall.Count -gt 0) {
            $officePaths = @(
                "${env:ProgramFiles}\Microsoft Office\Office16",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
                "${env:ProgramFiles}\Microsoft Office\Office15",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
            )
            foreach ($path in $officePaths) {
                $ospp = Join-Path $path "ospp.vbs"
                if (Test-Path $ospp) {
                    foreach ($key in $keysToUninstall) {
                        $log.AppendLine(" -> [OSPP Script] Dang go khoa du phong: $key bang file ospp.vbs...") | Out-Null
                        $res = cscript.exe //NoLogo "$ospp" /unpkey:$key 2>&1
                        $log.AppendLine("    Ket qua: $($res -join ' ')") | Out-Null
                    }
                }
            }
        }
    } catch {
        $log.AppendLine(" -> Loi khi go qua OSPP: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("                    HOAN TAT QUA TRINH GO BO BAN QUYEN KMS LAU                    ") | Out-Null
    $log.AppendLine(" - Trang thai Windows/Office hien tai da duoc lam sach cac thong tin KMS lau.") | Out-Null
    $log.AppendLine(" - Ban co the nhap khoa ban quyen chinh hang de kich hoat lai.") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return $log.ToString()
}

# 5.5 HAM GO BO MICROSOFT OFFICE
function Start-OfficeUninstallProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("              BAT DAU QUA TRINH GO CAI DAT MICROSOFT OFFICE                       ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    
    $officeApps = @()
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
            
            $officeApps += $app
        }
    }

    if ($officeApps.Count -eq 0) {
        $log.AppendLine(" -> [THONG BAO] Khong phat hien thay phien ban Microsoft Office nao duoc cai dat.") | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        return @{ Log = $log.ToString(); Count = 0 }
    }

    $log.AppendLine(" -> Phat hien $($officeApps.Count) san pham Office can go:") | Out-Null
    foreach ($o in $officeApps) {
        $log.AppendLine("    * $($o.DisplayName)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    $uninstalledCount = 0
    
    # Click-To-Run uninstaller
    $c2rPath = "${env:ProgramFiles}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    if (-not (Test-Path $c2rPath)) {
        $c2rPath = "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    }

    $hasC2R = $false
    foreach ($app in $officeApps) {
        if ($app.UninstallString -like "*ClickToRun*" -or (Test-Path $c2rPath)) {
            $hasC2R = $true
        }
    }

    if ($hasC2R -and (Test-Path $c2rPath)) {
        $log.AppendLine("[+] Dang chay trinh go sach Click-To-Run tu dong...") | Out-Null
        try {
            $p = Start-Process -FilePath $c2rPath -ArgumentList "scenario=uninstall forceuninstall=True displaylevel=False" -PassThru -NoNewWindow -ErrorAction SilentlyContinue
            if ($p) {
                $p.WaitForExit(300000)
                $log.AppendLine(" -> Da thuc thi trinh go Click-To-Run. Ma thoat: $($p.ExitCode)") | Out-Null
                $uninstalledCount++
            } else {
                $log.AppendLine(" -> Khong the khoi chay trinh go Click-To-Run.") | Out-Null
            }
        } catch {
            $log.AppendLine(" -> Loi khi chay Click-To-Run uninstaller: $($_.Exception.Message)") | Out-Null
        }
    }

    # MSI and other uninstallation strings
    foreach ($app in $officeApps) {
        $stillExists = Get-ItemProperty -Path $app.PSPath -ErrorAction SilentlyContinue
        if (-not $stillExists) {
            $uninstalledCount++
            continue
        }

        $unCmd = $app.QuietUninstallString
        $isQuiet = $true
        if ([string]::IsNullOrWhiteSpace($unCmd)) {
            $unCmd = $app.UninstallString
            $isQuiet = $false
        }

        if (-not [string]::IsNullOrWhiteSpace($unCmd)) {
            $log.AppendLine("[+] Dang go thu cong ($($app.DisplayName))...") | Out-Null
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
                        if ($exe -like "*msiexec.exe*") {
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
                    if ($p) { $p.WaitForExit(180000) }
                } else {
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(180000) }
                }
                $log.AppendLine(" -> Da hoan tat go: $($app.DisplayName)") | Out-Null
                $uninstalledCount++
            } catch {
                $log.AppendLine(" -> Loi khi go $($app.DisplayName): $($_.Exception.Message)") | Out-Null
            }
        }
    }

    $log.AppendLine() | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("       HOAN TAT GO CAI DAT MICROSOFT OFFICE                                       ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return @{ Log = $log.ToString(); Count = $uninstalledCount }
}

# 6. CHAY CHE DO DONG LENH (CLI)
if ($ConsoleOnly) {
    if ($RemoveKMS) {
        $res = Start-KMSRemovalProcess
        Write-Host $res
    } elseif ($UninstallCracks) {
        $res = Start-CrackedAppsUninstallProcess
        Write-Host $res.Log
    } elseif ($UninstallOffice) {
        $res = Start-OfficeUninstallProcess
        Write-Host $res.Log
    } else {
        Write-Host "`n==========================================================================" -ForegroundColor Cyan
        Write-Host "             CONG CU KIEM TRA BAN QUYEN WINDOWS & OFFICE                   " -ForegroundColor White -BackgroundColor DarkBlue
        Write-Host "==========================================================================" -ForegroundColor Cyan
        
        Write-Host "`n[+] Dang quet ban quyen Windows..." -ForegroundColor Yellow
        $win = Get-WindowsActivation
        Write-Host " - Phien ban    : $($win.ProductName)"
        Write-Host " - Phan loai    : $($win.Channel)"
        Write-Host " - Trang thai   : $($win.StatusText)"
        if ($win.KMSServer -and $win.KMSServer -ne "Khong phat hien") {
            Write-Host " - May chu KMS  : $($win.KMSServer)" -ForegroundColor Red
        }
        
        Write-Host "`n[+] Dang quet ban quyen Office..." -ForegroundColor Yellow
        $officeList = Get-OfficeActivation
        if ($officeList.Count -eq 0) {
            Write-Host " - Khong tim thay ban Office nao da dang ky khoa ban quyen." -ForegroundColor Gray
        } else {
            foreach ($off in $officeList) {
                Write-Host " - Ten san pham : $($off.ProductName)"
                Write-Host "   Trang thai   : $($off.Status)"
                Write-Host "   Phan nhom    : $($off.Channel)"
                if ($off.KMSServer) {
                    Write-Host "   May chu KMS  : $($off.KMSServer)" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "`n[+] Dang quet cac dau vet cong cu Crack..." -ForegroundColor Yellow
        $crack = Get-CrackDetection
        if ($crack.Risks.Count -eq 0) {
            Write-Host " - [AN TOAN] Khong phat hien thay bat ky dau hieu phan mem crack/be khoa nao." -ForegroundColor Green
        } else {
            Write-Host " - [CANH BAO] Phat hien $($crack.Risks.Count) dau hieu crack/be khoa he thong:" -ForegroundColor Red
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                Write-Host "   * Dau hieu: $($crack.Risks[$i])" -ForegroundColor Red
                Write-Host "     Mo ta   : $($crack.Details[$i])" -ForegroundColor Gray
            }
        }
        Write-Host "`n==========================================================================" -ForegroundColor Cyan
    }
    Exit
}

# 7. DINH NGHIA XAML GIAO DIEN UNIFIED GUI (TAT-CA-TRONG-MOT VOI 3 NUT TAC VU)
$xamlContent = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Cong cu kiem tra &amp; Quan ly ban quyen Windows - Tat ca trong mot" Height="720" Width="1060" 
        WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize"
        Background="#0B0F19" BorderBrush="#1E293B" BorderThickness="1">
    
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="FontFamily" Value="Segoe UI, Inter, Arial"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
        </Style>
    </Window.Resources>

    <Grid Margin="25">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header chua 3 nut chinh -->
        <Grid Grid.Row="0" Margin="0,0,0,25">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Text="CONG CU QUAN LY BAN QUYEN &amp; DON CRACK HE THONG" FontSize="18.5" FontWeight="Bold" Foreground="#38BDF8"/>
                <TextBlock Text="Quet ban quyen Windows/Office, go bo ung dung crack va don may chu KMS lau" FontSize="12" Foreground="#94A3B8" Margin="0,5,0,0"/>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnScan" Content="BAT DAU QUET" Width="140" Height="42" Background="#0284C7" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
                <Button Name="btnUninstallCracks" Content="GO CRACK HANG LOAT" Width="175" Height="42" Background="#D97706" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
                <Button Name="btnRemoveKMS" Content="GO BO KMS LAU" Width="145" Height="42" Background="#DC2626" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
                <Button Name="btnUninstallOffice" Content="GO BO OFFICE" Width="135" Height="42" Background="#4B5563" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
            </StackPanel>
        </Grid>

        <!-- 3 Cards thong so -->
        <UniformGrid Grid.Row="1" Columns="3" Margin="0,0,0,25">
            <Border Name="cardWin" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="0,0,12,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="HE DIEU HANH WINDOWS" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtWinStatus" Text="Chua kiem tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtWinDetails" Text="Nhan nut Quet de nhan thong tin chi tiet ve phien ban va kenh kich hoat." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>

            <Border Name="cardOffice" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="6,0,6,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="MICROSOFT OFFICE" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtOfficeStatus" Text="Chua kiem tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtOfficeDetails" Text="Nhan nut Quet de phat hien cac phien ban Office duoc cai dat tren he thong." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>

            <Border Name="cardCrack" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="12,0,0,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="PHAN MEM BE KHOA / CRACK" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtCrackStatus" Text="Chua kiem tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtCrackDetails" Text="Quet chu ky so ung dung cai dat, tep Hosts va tac vu ngam de phat hien crack..." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>
        </UniformGrid>

        <!-- Log Console -->
        <Grid Grid.Row="2" Margin="0,0,0,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Text="KET QUA QUET / THUC THI CHI TIET" FontSize="12" FontWeight="Bold" Foreground="#94A3B8" Margin="0,0,0,8"/>
            <TextBox Name="txtLogs" Grid.Row="1" Background="#070A13" Foreground="#38BDF8" BorderBrush="#1E293B" BorderThickness="1.5" FontFamily="Consolas" FontSize="12.5" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" AcceptsReturn="True" Padding="15"/>
        </Grid>

        <!-- Footer -->
        <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Trang thai he thong: San sang | Dang chay voi quyen Administrator" FontSize="11" Foreground="#475569" VerticalAlignment="Center"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnExport" Content="Xuat Bao Cao (.txt)" Width="160" Height="32" Background="#1E293B" Foreground="#E2E8F0" FontWeight="SemiBold" FontSize="12" BorderThickness="0" Cursor="Hand" IsEnabled="False">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="4"/>
                        </Style>
                    </Button.Resources>
                </Button>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
'@

# 8. KHOI CHAY VA RANG BUOC SU KIEN GIAO DIEN
try {
    $xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xamlContent)
    $Window = [Windows.Markup.XamlReader]::Load($xmlReader)

    $xamlObj = [xml]$xamlContent
    $xamlObj.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name "UI_$($_.Name)" -Value $Window.FindName($_.Name)
    }

    $BrushGreen = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
    $BrushRed = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
    $BrushYellow = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F59E0B")
    $BrushWhite = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F1F5F9")
    $BrushSlate = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#475569")
    $BorderRed = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#7F1D1D")
    $BorderGreen = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#064E3B")
    $BorderNormal = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1F2937")

    $globalReportText = ""

    # SU KIEN CLICK: QUET HE THONG
    $UI_btnScan.Add_Click({
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnUninstallOffice.IsEnabled = $false
        $UI_btnScan.Content = "DANG QUET..."
        
        $UI_txtWinStatus.Text = "Dang xu ly..."
        $UI_txtOfficeStatus.Text = "Dang xu ly..."
        $UI_txtCrackStatus.Text = "Dang xu ly..."
        
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $win = Get-WindowsActivation
        if ($win.RawStatus -eq 1) {
            $UI_txtWinStatus.Text = "Da kich hoat [OK]"
            $UI_txtWinStatus.Foreground = $BrushGreen
            $UI_cardWin.BorderBrush = $BorderGreen
        } elseif ($win.RawStatus -eq 5 -or $win.RawStatus -eq 0) {
            $UI_txtWinStatus.Text = "Chua kich hoat [!]"
            $UI_txtWinStatus.Foreground = $BrushYellow
            $UI_cardWin.BorderBrush = $BorderNormal
        } else {
            $UI_txtWinStatus.Text = "Trang thai khac [!]"
            $UI_txtWinStatus.Foreground = $BrushYellow
            $UI_cardWin.BorderBrush = $BorderNormal
        }
        $winServerInfo = ""
        if ($win.KMSServer -and $win.KMSServer -ne "Khong phat hien") {
            $winServerInfo = "`n[!] May chu KMS: $($win.KMSServer)"
        }
        $UI_txtWinDetails.Text = "Kenh kich hoat: $($win.Channel)`nPhien ban: $($win.ProductName)$winServerInfo"
        
        $officeList = Get-OfficeActivation
        if ($officeList.Count -eq 0) {
            $UI_txtOfficeStatus.Text = "Chua cai dat / Khong tim thay "
            $UI_txtOfficeStatus.Foreground = $BrushSlate
            $UI_txtOfficeDetails.Text = "Khong tim thay phien ban Microsoft Office nao da dang ky khoa ban quyen."
            $UI_cardOffice.BorderBrush = $BorderNormal
        } else {
            $isOfficeActivated = $true
            $officeBrief = ""
            foreach ($off in $officeList) {
                $briefName = $off.ProductName
                if ($briefName.Length -gt 35) { $briefName = $briefName.Substring(0, 32) + "..." }
                $officeBrief += "- $briefName : $($off.Status)`n"
                if ($off.Status -ne "Da kich hoat" -and $off.Status -notlike "*Kich hoat qua*") {
                    $isOfficeActivated = $false
                }
            }
            $UI_txtOfficeDetails.Text = $officeBrief.Trim()
            
            if ($isOfficeActivated) {
                $UI_txtOfficeStatus.Text = "Da kich hoat [OK]"
                $UI_txtOfficeStatus.Foreground = $BrushGreen
                $UI_cardOffice.BorderBrush = $BorderGreen
            } else {
                $UI_txtOfficeStatus.Text = "Co phien ban chua kich hoat [!]"
                $UI_txtOfficeStatus.Foreground = $BrushYellow
                $UI_cardOffice.BorderBrush = $BorderNormal
            }
        }
        
        $crack = Get-CrackDetection
        if ($crack.Risks.Count -eq 0) {
            $UI_txtCrackStatus.Text = "He thong An toan [OK]"
            $UI_txtCrackStatus.Foreground = $BrushGreen
            $UI_cardCrack.BorderBrush = $BorderGreen
            $UI_txtCrackDetails.Text = "Khong phat hien bat ky tien trinh, tac vu chay ngam, file hoac tep tin hosts nao bi can thiep boi cong cu be khoa."
        } else {
            $UI_txtCrackStatus.Text = "Phat hien vet be khoa ❌"
            $UI_txtCrackStatus.Foreground = $BrushRed
            $UI_cardCrack.BorderBrush = $BorderRed
            $UI_txtCrackDetails.Text = "Phat hien $($crack.Risks.Count) diem canh bao be khoa phan mem he thong. Xem chi tiet ben duoi!"
        }
        
        $log = New-Object System.Text.StringBuilder
        $log.AppendLine("==================================================================================") | Out-Null
        $log.AppendLine("                   BAO CAO KIEM TRA BAN QUYEN VA PHAT HIEN CRACK                  ") | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        $log.AppendLine("Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
        $log.AppendLine("Ten may tinh        : $env:COMPUTERNAME") | Out-Null
        $log.AppendLine("He dieu hanh        : $( (Get-CimInstance Win32_OperatingSystem).Caption ) (Phien ban: $((Get-CimInstance Win32_OperatingSystem).Version))") | Out-Null
        $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[1] KET QUA KIEM TRA WINDOWS:") | Out-Null
        $log.AppendLine(" - San pham          : $($win.ProductName)") | Out-Null
        $log.AppendLine(" - Kenh phan phoi    : $($win.Channel)") | Out-Null
        $log.AppendLine(" - Trang thai        : $($win.StatusText)") | Out-Null
        if ($win.KMSServer -and $win.KMSServer -ne "Khong phat hien") {
            $log.AppendLine(" - May chu KMS (Lau) : $($win.KMSServer) (Phat hien cau hinh ban quyen huong ngoai)") | Out-Null
        }
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[2] KET QUA KIEM TRA OFFICE:") | Out-Null
        if ($officeList.Count -eq 0) {
            $log.AppendLine(" - Khong tim thay thong tin san pham Microsoft Office.") | Out-Null
        } else {
            foreach ($off in $officeList) {
                $log.AppendLine(" - Ten phien ban   : $($off.ProductName)") | Out-Null
                $log.AppendLine("   Trang thai      : $($off.Status)") | Out-Null
                $log.AppendLine("   Kenh giay phep  : $($off.Channel)") | Out-Null
                $log.AppendLine("   Nguon phat hien : $($off.Source)") | Out-Null
                if ($off.KMSServer) {
                    $log.AppendLine("   May chu KMS     : $($off.KMSServer)") | Out-Null
                }
                $log.AppendLine("   --------------------------------------") | Out-Null
            }
        }
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[3] DANH GIA VA PHAT HIEN CRACK/BE KHOA:") | Out-Null
        if ($crack.Risks.Count -eq 0) {
            $log.AppendLine(" -> [AN TOAN] Khong phat hien bat ky tien trinh, tac vu, tep tin hosts hay thu muc be khoa nao.") | Out-Null
        } else {
            $log.AppendLine(" -> [CANH BAO NGUY HIEM] Phat hien $($crack.Risks.Count) canh bao be khoa phan mem:") | Out-Null
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                $log.AppendLine("   * Canh bao: $($crack.Risks[$i])") | Out-Null
                $log.AppendLine("     Mo ta   : $($crack.Details[$i])") | Out-Null
                $log.AppendLine("     ----------------------------------------------------------------") | Out-Null
            }
            $log.AppendLine() | Out-Null
            $log.AppendLine("KHUYEN NGHI BAO MAT:") | Out-Null
            $log.AppendLine(" - Cac phan mem be khoa co nguy co bi cai cam ma doc lay cap du lieu doanh nghiep.") | Out-Null
            $log.AppendLine(" - Nen go bo ban crack va su dung cac phien ban chinh hang hoac ma nguon mo thay the.") | Out-Null
        }
        $log.AppendLine() | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        
        $script:globalReportText = $log.ToString()
        $UI_txtLogs.Text = $script:globalReportText
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnUninstallOffice.IsEnabled = $true
        $UI_btnScan.Content = "BAT DAU QUET"
        $UI_btnExport.IsEnabled = $true
    })

    # SU KIEN CLICK: GO CAI DAT HANG LOAT UNG DUNG CRACK
    $UI_btnUninstallCracks.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show("Ban co chac chan muon TU DONG GO CAI DAT HANG LOAT tat ca cac phan mem thuong mai bi phat hien be khoa tren may tinh khong?", "Xac nhan go crack hang loat", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnUninstallOffice.IsEnabled = $false
        $UI_btnUninstallCracks.Content = "DANG GO HANG LOAT..."
        
        $UI_txtLogs.Text = "Dang tien hanh go cai dat tu dong hang loat tat ca phan mem be khoa... Vui long doi."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $res = Start-CrackedAppsUninstallProcess
        $script:globalReportText = $res.Log
        $UI_txtLogs.Text = $res.Log
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnUninstallOffice.IsEnabled = $true
        $UI_btnUninstallCracks.Content = "GO CRACK HANG LOAT"
        $UI_btnExport.IsEnabled = $true
        
        if ($res.Count -gt 0) {
            [System.Windows.MessageBox]::Show("Da hoan tat qua trinh go cai dat hang loat $($res.Count) phan mem bi crack!", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            [System.Windows.MessageBox]::Show("Khong tim thay phan mem ben thu ba nao dang bi crack tren he thong.", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })

    # SU KIEN CLICK: GO BO KMS LAU
    $UI_btnRemoveKMS.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show("Ban co chac chan muon go bo may chu KMS lau va lam sach he thong khong?`n`nHanh dong nay se go Product Key KMS lau va dua Windows tro ve trang thai chua kich hoat chuan.", "Xac nhan go KMS", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnUninstallOffice.IsEnabled = $false
        $UI_btnRemoveKMS.Content = "DANG GO BO..."
        
        $UI_txtLogs.Text = "Dang tien hanh go bo ban quyen KMS lau va don dep he thong... Vui long doi trong giay lat."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $resultText = Start-KMSRemovalProcess
        $script:globalReportText = $resultText
        $UI_txtLogs.Text = $resultText
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnUninstallOffice.IsEnabled = $true
        $UI_btnRemoveKMS.Content = "GO BO KMS LAU"
        $UI_btnExport.IsEnabled = $true
        [System.Windows.MessageBox]::Show("Da hoan tat qua trinh go bo ban quyen KMS lau va lam sach he thong!", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    # SU KIEN CLICK: GO BO OFFICE
    $UI_btnUninstallOffice.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show("Ban co chac chan muon GO CAI DAT toan bo cac phien ban Microsoft Office/Microsoft 365 tren may tinh nay khong?", "Xac nhan go Office", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnUninstallOffice.IsEnabled = $false
        $UI_btnUninstallOffice.Content = "DANG GO BO..."
        
        $UI_txtLogs.Text = "Dang tien hanh go bo sach toan bo phien ban Microsoft Office... Vui long doi trong giay lat."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $res = Start-OfficeUninstallProcess
        $script:globalReportText = $res.Log
        $UI_txtLogs.Text = $res.Log
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnUninstallOffice.IsEnabled = $true
        $UI_btnUninstallOffice.Content = "GO BO OFFICE"
        $UI_btnExport.IsEnabled = $true
        
        if ($res.Count -gt 0) {
            [System.Windows.MessageBox]::Show("Da hoan tat qua trinh go bo cac phien ban Microsoft Office!", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            [System.Windows.MessageBox]::Show("Khong tim thay phien ban Microsoft Office nao de go hoac da duoc go truoc do.", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })

    # SU KIEN CLICK: XUAT BAO CAO
    $UI_btnExport.Add_Click({
        if (-not $script:globalReportText) { return }
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
        $saveFileDialog.FileName = "BaoCao_BanQuyen_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        if ($saveFileDialog.ShowDialog() -eq $true) {
            try {
                $script:globalReportText | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8
                [System.Windows.MessageBox]::Show("Xuat bao cao thanh cong tai:`n$($saveFileDialog.FileName)", "Thong bao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } catch {
                [System.Windows.MessageBox]::Show("Loi khi ghi tep: $($_.Exception.Message)", "Loi", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

    # HIEN THI CUA SO
    $Window.ShowDialog() | Out-Null

} catch {
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "LOI THUC THI GIAO DIEN:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "Nhan phim bat ky de dong..." -ForegroundColor Yellow
    [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
