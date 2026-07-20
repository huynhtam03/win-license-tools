<#
    Check-LicenseStatus-Standalone.ps1 - Công cụ TẤT-CẢ-TRONG-MỘT (All-In-One Edition)
    Bao gồm:
    1. Kiểm tra bản quyền Windows & Microsoft Office.
    2. Quét động các phần mềm bên thứ ba bị crack (MiniTool, Adobe, IDM, AutoCAD...).
    3. GỠ CÀI ĐẶT TỰ ĐỘNG HÀNG LOẠT (Bulk/Batch Uninstall) các phần mềm bị crack.
    4. Gỡ bỏ bản quyền KMS lậu (slmgr /ckms, /upk, /cpky), làm sạch file Hosts, gỡ bỏ tác vụ/dịch vụ bẻ khóa.
#>

[CmdletBinding()]
param (
    [switch]$ConsoleOnly,
    [switch]$RemoveKMS,
    [switch]$UninstallCracks
)

# 1. THIẾT LẬP THƯ VIỆN WPF
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
} catch {}

# 2. TỰ ĐỘNG NÂNG QUYỀN ADMINISTRATOR
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "CẢNH BÁO: Công cụ yêu cầu quyền Administrator!" -ForegroundColor Yellow
    Write-Host "Đang mở cửa sổ PowerShell mới dưới quyền Administrator..." -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-sta", "-File", "`"$PSCommandPath`"" -Verb RunAs
    } catch {
        Write-Host "Lỗi: Không thể khởi chạy dưới quyền Administrator!" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
    Exit
}

# 3. CÁC HÀM QUÉT BẢN QUYỀN VÀ PHÁT HIỆN BẺ KHÓA
function Get-WindowsActivation {
    $result = @{
        Status = "Unknown"
        StatusText = "Chưa quét"
        ProductName = "Windows"
        Channel = "Không xác định"
        KMSServer = "Không phát hiện"
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
                0 { $result.StatusText = "Chưa kích hoạt (Unlicensed)"; $result.Status = "Unlicensed" }
                1 { $result.StatusText = "Đã kích hoạt (Licensed)"; $result.Status = "Licensed" }
                2 { $result.StatusText = "Ân hạn ban đầu (OOB Grace)"; $result.Status = "Grace" }
                3 { $result.StatusText = "Ân hạn hết hạn (OOT Grace)"; $result.Status = "Grace" }
                4 { $result.StatusText = "Ân hạn không chính hãng (Non-genuine)"; $result.Status = "Warning" }
                5 { $result.StatusText = "Chờ kích hoạt / Hết hạn (Notification)"; $result.Status = "Notification" }
                6 { $result.StatusText = "Ân hạn mở rộng (Extended Grace)"; $result.Status = "Grace" }
                default { $result.StatusText = "Không xác định"; $result.Status = "Unknown" }
            }
            
            if ($win.Description -like "*RETAIL*") { $result.Channel = "Retail (Bán lẻ)" }
            elseif ($win.Description -like "*OEM*") { $result.Channel = "OEM (Theo máy)" }
            elseif ($win.Description -like "*VOLUME_KMSCLIENT*") { $result.Channel = "Volume KMS (KMS Doanh nghiệp)" }
            elseif ($win.Description -like "*VOLUME_MAK*") { $result.Channel = "Volume MAK" }
            else { $result.Channel = "Khác/Tự cấu hình" }
            
            $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
            if ($licService -and $licService.KeyManagementServiceMachine) {
                $result.KMSServer = "$($licService.KeyManagementServiceMachine):$($licService.KeyManagementServicePort)"
            }
            
            $result.Details = "Sản phẩm: $($win.Name)`nPhân nhóm: $($result.Channel)`nTrạng thái: $($result.StatusText)`nMáy chủ kích hoạt KMS: $($result.KMSServer)"
        } else {
            $result.StatusText = "Không có sản phẩm hoạt động"
            $result.Details = "Không thể tìm thấy thông tin giấy phép Windows hợp lệ qua WMI."
        }
    } catch {
        $result.StatusText = "Lỗi truy vấn"
        $result.Details = "Lỗi khi lấy thông tin bản quyền Windows: $($_.Exception.Message)"
    }
    return $result
}

function Get-OfficeActivation {
    $officeList = @()
    $offices = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
    if ($offices) {
        foreach ($off in $offices) {
            $statusText = switch ($off.LicenseStatus) {
                0 { "Chưa kích hoạt" }
                1 { "Đã kích hoạt" }
                2 { "Ân hạn ban đầu" }
                3 { "Ân hạn hết hạn" }
                4 { "Ân hạn không chính hãng" }
                5 { "Chờ kích hoạt / Hết hạn" }
                6 { "Ân hạn mở rộng" }
                default { "Không xác định" }
            }
            
            $channel = "Không xác định"
            if ($off.Description -like "*RETAIL*") { $channel = "Retail (Bán lẻ)" }
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
                    if ($statusLine -like "*---LICENSED---*") { $status = "Đã kích hoạt" }
                    elseif ($statusLine -like "*---NOTIFIED---*") { $status = "Chờ kích hoạt / Hết hạn" }
                    else { $status = "Chưa kích hoạt ($statusLine)" }
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
            ProductName = "Microsoft 365 Apps (Đăng ký thuê bao)"
            Status = "Đã cài đặt (Kích hoạt qua Tài khoản Microsoft)"
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
    
    # 1. Kiểm tra Hook DLL
    $hookPaths = @(
        "$env:windir\System32\SppExtComObjHook.dll",
        "$env:windir\SysWOW64\SppExtComObjHook.dll",
        "$env:windir\SppExtComObjHook.dll",
        "$env:windir\System32\sppextcomobj.exe.local",
        "$env:windir\System32\sppextcomobj_hook.dll"
    )
    foreach ($path in $hookPaths) {
        if (Test-Path $path) {
            $risks += "Phát hiện file hook KMS: $path"
            $details += "Tệp tin '$path' thường do KMSpico hoặc KMSAuto cài đặt để bypass máy chủ bản quyền của Microsoft."
        }
    }
    
    # 2. Kiểm tra file Hosts
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
            $risks += "Sửa đổi tệp tin Hosts (Định tuyến chặn Microsoft)"
            $details += "Tệp tin Hosts chặn kết nối xác minh của Microsoft (ví dụ: $($redirects -join ', '))."
        }
        if ($blockedApps.Count -gt 0) {
            $risks += "Chặn máy chủ xác thực bản quyền phần mềm trong file Hosts"
            $details += "Tệp Hosts chặn kết nối đến hãng phần mềm: $($blockedApps -join ', '). Cho thấy các phần mềm này đang dùng bản crack."
        }
    }
    
    # 3. Kiểm tra Scheduled Tasks
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($task in $tasks) {
            $risks += "Tác vụ chạy ngầm bẻ khóa: $($task.TaskName)"
            $details += "Tác vụ tự động gia hạn KMS lậu chạy ngầm tại '$($task.TaskPath)'."
        }
    }
    
    # 4. Kiểm tra Services
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            $risks += "Dịch vụ KMS lậu chạy nền: $($srv.DisplayName) ($srvName)"
            $details += "Dịch vụ giả lập KMS nội bộ đang hoạt động."
        }
    }

    # 5. Kiểm tra Processes
    $suspiciousProcessNames = @("KMSpico", "AutoKMS", "KMSAuto", "KMSConnectionMonitor", "HEU_KMS", "HEU_KMS_Activator")
    foreach ($procName in $suspiciousProcessNames) {
        $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($proc) {
            $risks += "Tiến trình bẻ khóa đang chạy: $procName"
            $details += "Tiến trình '$procName' (PID: $($proc.Id)) đang chạy ngầm thời gian thực."
        }
    }
    
    # 6. Kiểm tra các thư mục chứa công cụ bẻ khóa hệ thống
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
                    $risks += "Thư mục hệ thống chứa Script kích hoạt lậu: $folder"
                    $details += "Thư mục chứa script *.bat cấu hình bản quyền tự động khi cài Win."
                }
            } else {
                $risks += "Phát hiện thư mục phần mềm bẻ khóa hệ thống: $folder"
                $details += "Thư mục cài đặt ứng dụng bẻ khóa tồn tại ở '$folder'."
            }
        }
    }
    
    # 7. QUÉT ĐỘNG PHẦN MỀM BÊN THỨ 3 (Dynamic Third-Party Crack Scan)
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
                                $risks += "Phát hiện $($app.DisplayName) bị Crack (Mất chữ ký số)"
                                $details += "Tệp tin '$($file.FullName)' của '$($app.DisplayName)' mất chữ ký số hợp lệ (Trạng thái: $($sig.Status)), chứng tỏ tệp tin này đã bị crack (patch)."
                            }
                        }
                    }
                }
            }
        }
    }
    
    # 8. Cấu hình KMS của Windows & Office
    $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
    if ($licService -and $licService.KeyManagementServiceMachine) {
        $kmsMachine = $licService.KeyManagementServiceMachine
        $knownKmsServers = @("kms8.msguides.com", "kms.digiboy.ir", "kms.lotro.cc", "kms.chinancce.com", "zh.us.to", "kms.msguides.com", "kms.spacespaces.xyz")
        if ($kmsMachine -eq "127.0.0.1" -or $kmsMachine -eq "localhost" -or $kmsMachine -eq "::1") {
            $risks += "Windows cấu hình máy chủ KMS nội bộ (Localhost)"
            $details += "Bản quyền Windows trỏ về chính máy tính ($kmsMachine). Dấu hiệu bẻ khóa giả lập cục bộ."
        } else {
            foreach ($srv in $knownKmsServers) {
                if ($kmsMachine -like "*$srv*") {
                    $risks += "Windows sử dụng máy chủ KMS bẻ khóa công cộng: $kmsMachine"
                    $details += "Bản quyền hệ thống trỏ tới máy chủ KMS miễn phí trên Internet."
                    break
                }
            }
        }
    }

    $officeRegKms = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    if ($officeRegKms -and $officeRegKms.KeyManagementServiceMachine) {
        $kmsMachineOffice = $officeRegKms.KeyManagementServiceMachine
        if ($kmsMachineOffice -eq "127.0.0.1" -or $kmsMachineOffice -eq "localhost" -or $kmsMachineOffice -eq "::1") {
            $risks += "Office cấu hình máy chủ KMS nội bộ (Localhost)"
            $details += "Bản quyền Office trỏ máy chủ kích hoạt về chính máy tính thông qua Registry."
        }
    }

    # 9. Kiểm tra cổng KMS 1688
    try {
        $kmsPort = Get-NetTCPConnection -LocalPort 1688 -State Listen -ErrorAction SilentlyContinue
        if ($kmsPort) {
            $risks += "Phát hiện cổng dịch vụ KMS (1688) đang mở"
            $details += "Cổng TCP 1688 đang lắng nghe (Listen) cục bộ. Dấu hiệu của KMS giả lập chạy ẩn."
        }
    } catch {
        $netstat = netstat -ano 2>&1 | Select-String ":1688\s+.*LISTENING"
        if ($netstat) {
            $risks += "Phát hiện cổng dịch vụ KMS (1688) đang mở (netstat)"
            $details += "Cổng 1688 đang mở ở trạng thái LISTENING qua lệnh netstat."
        }
    }

    # 10. Kiểm tra Registry của phần mềm bẻ khóa
    $crackRegs = @("HKLM:\SOFTWARE\KMSAuto", "HKLM:\SOFTWARE\KMSAutoS", "HKLM:\SOFTWARE\KMSpico", "HKCU:\Software\KMSAuto", "HKCU:\Software\KMSAutoS")
    foreach ($reg in $crackRegs) {
        if (Test-Path $reg) {
            $risks += "Phát hiện khóa Registry bẻ khóa: $reg"
            $details += "Khóa cấu hình của KMSAuto/KMSpico được phát hiện tại '$reg'."
        }
    }
    
    return [PSCustomObject]@{
        Risks = $risks
        Details = $details
    }
}

# 4. HÀM GỠ CÀI ĐẶT HÀNG LOẠT (BULK UNINSTALL) CÁC PHẦN MỀM CRACK BÊN THỨ 3
function Start-CrackedAppsUninstallProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("          BẮT ĐẦU QUÁ TRÌNH GỠ CÀI ĐẶT HÀNG LOẠT (BULK UNINSTALL) CÁC APP CRACK   ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thời gian thực hiện : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Tên máy tính        : $env:COMPUTERNAME") | Out-Null
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
        if ($risk -like "*bị Crack*") {
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
        $log.AppendLine(" -> [THÔNG BÁO] Không phát hiện thấy phần mềm bên thứ ba nào đang bị bẻ khóa trên hệ thống.") | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        return @{ Log = $log.ToString(); Count = 0 }
    }

    $log.AppendLine(" -> Phát hiện $($crackedAppsFound.Count) phần mềm bị bẻ khóa cần gỡ hàng loạt:") | Out-Null
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
            $log.AppendLine("[+] Đang gỡ tự động ($($uninstalledCount + 1)/$($crackedAppsFound.Count)): $($app.DisplayName)...") | Out-Null
            
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
                        $log.AppendLine("    Lệnh thực thi: `"$exe`"") | Out-Null
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        $log.AppendLine("    Lệnh thực thi: `"$exe`" $argList") | Out-Null
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    
                    if ($p) { $p.WaitForExit(60000) }
                } else {
                    $log.AppendLine("    Thực thi lệnh CMD: $cleanUnCmd") | Out-Null
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(60000) }
                }
                
                $log.AppendLine(" -> ĐÃ GỠ THÀNH CÔNG: $($app.DisplayName)") | Out-Null
                $uninstalledCount++
            } catch {
                $log.AppendLine(" -> Lỗi khi gỡ $($app.DisplayName): $($_.Exception.Message)") | Out-Null
            }
        } else {
            $log.AppendLine(" -> Bỏ qua $($app.DisplayName): Không tìm thấy UninstallString trong Registry.") | Out-Null
        }
        $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    }
    
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("       HOÀN TẤT GỠ CÀI ĐẶT HÀNG LOẠT $uninstalledCount / $($crackedAppsFound.Count) PHẦN MỀM BỆ KHÓA           ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return @{ Log = $log.ToString(); Count = $uninstalledCount }
}

# 5. HÀM GỠ BỎ KMS LẬU VÀ DỌN DẸP HỆ THỐNG
function Start-KMSRemovalProcess {
    $log = New-Object System.Text.StringBuilder
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("              BẮT ĐẦU QUÁ TRÌNH GỠ BỎ BẢN QUYỀN KMS LẬU VÀ BẢO TRÌ SYSTEM         ") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("Thời gian thực hiện : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
    $log.AppendLine("Tên máy tính        : $env:COMPUTERNAME") | Out-Null
    $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
    $log.AppendLine() | Out-Null

    # Bước 1: Xóa cấu hình máy chủ KMS (slmgr /ckms)
    $log.AppendLine("[BƯỚC 1] Đang gỡ bỏ cấu hình máy chủ KMS (Clear KMS Server)...") | Out-Null
    try {
        $ckmsResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ckms 2>&1
        $log.AppendLine(" -> Kết quả slmgr /ckms: $($ckmsResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Lỗi khi chạy slmgr /ckms: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Bước 2: Gỡ bỏ khóa sản phẩm KMS hiện tại (slmgr /upk)
    $log.AppendLine("[BƯỚC 2] Đang gỡ bỏ Product Key KMS hiện tại (Uninstall Product Key)...") | Out-Null
    try {
        $upkResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /upk 2>&1
        $log.AppendLine(" -> Kết quả slmgr /upk: $($upkResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Lỗi khi chạy slmgr /upk: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Bước 3: Xóa thông tin Product Key khỏi Registry (slmgr /cpky)
    $log.AppendLine("[BƯỚC 3] Đang xóa thông tin Product Key khỏi Registry (Clear Product Key from Registry)...") | Out-Null
    try {
        $cpkyResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /cpky 2>&1
        $log.AppendLine(" -> Kết quả slmgr /cpky: $($cpkyResult -join ' ')") | Out-Null
    } catch {
        $log.AppendLine(" -> Lỗi khi chạy slmgr /cpky: $($_.Exception.Message)") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Bước 4: Làm sạch tệp tin Hosts (Gỡ chặn các máy chủ xác minh Microsoft)
    $log.AppendLine("[BƯỚC 4] Đang kiểm tra và làm sạch tệp tin Hosts...") | Out-Null
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        try {
            $lines = Get-Content $hostsPath -ErrorAction SilentlyContinue
            $cleanLines = @()
            $removedCount = 0
            
            foreach ($line in $lines) {
                $clean = $line.Trim()
                if ($clean -and -not $clean.StartsWith("#") -and ($clean -like "*microsoft*" -or $clean -like "*activation*" -or $clean -like "*kms*")) {
                    $log.AppendLine(" -> Đã gỡ bỏ dòng chặn: $clean") | Out-Null
                    $removedCount++
                } else {
                    $cleanLines += $line
                }
            }
            if ($removedCount -gt 0) {
                $cleanLines | Set-Content $hostsPath -Encoding utf8 -Force
                $log.AppendLine(" -> Đã làm sạch $removedCount dòng cấu hình chặn trong file Hosts thành công!") | Out-Null
            } else {
                $log.AppendLine(" -> Tệp Hosts không chứa cấu hình chặn máy chủ Microsoft.") | Out-Null
            }
        } catch {
            $log.AppendLine(" -> Lỗi khi chỉnh sửa file Hosts: $($_.Exception.Message)") | Out-Null
        }
    }
    $log.AppendLine() | Out-Null

    # Bước 5: Tìm và gỡ bỏ Tác vụ chạy ngầm bẻ khóa (Scheduled Tasks)
    $log.AppendLine("[BƯỚC 5] Đang dọn dẹp các tác vụ bẻ khóa tự động (Scheduled Tasks)...") | Out-Null
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    $removedTasks = 0
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($t in $tasks) {
            try {
                Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
                $log.AppendLine(" -> Đã gỡ bỏ tác vụ bẻ khóa: $($t.TaskName) ($($t.TaskPath))") | Out-Null
                $removedTasks++
            } catch {}
        }
    }
    if ($removedTasks -eq 0) {
        $log.AppendLine(" -> Không phát hiện tác vụ chạy ngầm bẻ khóa nào.") | Out-Null
    }
    $log.AppendLine() | Out-Null

    # Bước 6: Tìm và dừng/gỡ bỏ Dịch vụ KMS lậu (Services)
    $log.AppendLine("[BƯỚC 6] Đang dọn dẹp các dịch vụ KMS lậu (Services)...") | Out-Null
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    $removedSrvs = 0
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            try {
                Stop-Service -Name $srvName -Force -ErrorAction SilentlyContinue
                sc.exe delete $srvName | Out-Null
                $log.AppendLine(" -> Đã dừng và gỡ bỏ dịch vụ KMS lậu: $($srv.DisplayName) ($srvName)") | Out-Null
                $removedSrvs++
            } catch {}
        }
    }
    if ($removedSrvs -eq 0) {
        $log.AppendLine(" -> Không phát hiện dịch vụ KMS lậu chạy ngầm nào.") | Out-Null
    }
    $log.AppendLine() | Out-Null

    $log.AppendLine("==================================================================================") | Out-Null
    $log.AppendLine("                    HOÀN TẤT QUÁ TRÌNH GỠ BỎ BẢN QUYỀN KMS LẬU                    ") | Out-Null
    $log.AppendLine(" - Trạng thái Windows hiện tại đã trở về chưa kích hoạt (Default / Unlicensed).") | Out-Null
    $log.AppendLine(" - Bạn có thể nhập khóa bản quyền chính hãng (Retail / OEM / MAK) để kích hoạt lại.") | Out-Null
    $log.AppendLine("==================================================================================") | Out-Null

    return $log.ToString()
}

# 6. CHẠY CHẾ ĐỘ DÒNG LỆNH (CLI)
if ($ConsoleOnly) {
    if ($RemoveKMS) {
        $res = Start-KMSRemovalProcess
        Write-Host $res
    } elseif ($UninstallCracks) {
        $res = Start-CrackedAppsUninstallProcess
        Write-Host $res.Log
    } else {
        Write-Host "`n==========================================================================" -ForegroundColor Cyan
        Write-Host "             CÔNG CỤ KIỂM TRA BẢN QUYỀN WINDOWS & OFFICE                   " -ForegroundColor White -BackgroundColor DarkBlue
        Write-Host "==========================================================================" -ForegroundColor Cyan
        
        Write-Host "`n[+] Đang quét bản quyền Windows..." -ForegroundColor Yellow
        $win = Get-WindowsActivation
        Write-Host " - Phiên bản    : $($win.ProductName)"
        Write-Host " - Phân loại    : $($win.Channel)"
        Write-Host " - Trạng thái   : $($win.StatusText)"
        if ($win.KMSServer -and $win.KMSServer -ne "Không phát hiện") {
            Write-Host " - Máy chủ KMS  : $($win.KMSServer)" -ForegroundColor Red
        }
        
        Write-Host "`n[+] Đang quét bản quyền Office..." -ForegroundColor Yellow
        $officeList = Get-OfficeActivation
        if ($officeList.Count -eq 0) {
            Write-Host " - Không tìm thấy bản Office nào đã đăng ký khóa bản quyền." -ForegroundColor Gray
        } else {
            foreach ($off in $officeList) {
                Write-Host " - Tên sản phẩm : $($off.ProductName)"
                Write-Host "   Trạng thái   : $($off.Status)"
                Write-Host "   Phân nhóm    : $($off.Channel)"
                if ($off.KMSServer) {
                    Write-Host "   Máy chủ KMS  : $($off.KMSServer)" -ForegroundColor Red
                }
            }
        }
        
        Write-Host "`n[+] Đang quét các dấu vết công cụ Crack..." -ForegroundColor Yellow
        $crack = Get-CrackDetection
        if ($crack.Risks.Count -eq 0) {
            Write-Host " - [AN TOÀN] Không phát hiện thấy bất kỳ dấu hiệu phần mềm crack/bẻ khóa nào." -ForegroundColor Green
        } else {
            Write-Host " - [CẢNH BÁO] Phát hiện $($crack.Risks.Count) dấu hiệu crack/bẻ khóa hệ thống:" -ForegroundColor Red
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                Write-Host "   * Dấu hiệu: $($crack.Risks[$i])" -ForegroundColor Red
                Write-Host "     Mô tả   : $($crack.Details[$i])" -ForegroundColor Gray
            }
        }
        Write-Host "`n==========================================================================" -ForegroundColor Cyan
    }
    Exit
}

# 7. ĐỊNH NGHĨA XAML GIAO DIỆN UNIFIED GUI (TẤT-CẢ-TRONG-MỘT VỚI 3 NÚT TÁC VỤ)
$xamlContent = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Công cụ kiểm tra &amp; Quản lý bản quyền Windows - Tất cả trong một" Height="720" Width="1060" 
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

        <!-- Header chứa 3 nút chính -->
        <Grid Grid.Row="0" Margin="0,0,0,25">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Text="CÔNG CỤ QUẢN LÝ BẢN QUYỀN &amp; DỌN CRACK HỆ THỐNG" FontSize="18.5" FontWeight="Bold" Foreground="#38BDF8"/>
                <TextBlock Text="Quét bản quyền Windows/Office, gỡ bỏ ứng dụng crack và dọn máy chủ KMS lậu" FontSize="12" Foreground="#94A3B8" Margin="0,5,0,0"/>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnScan" Content="🔍 BẮT ĐẦU QUÉT" Width="140" Height="42" Background="#0284C7" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
                <Button Name="btnUninstallCracks" Content="⚡ GỠ CRACK HÀNG LOẠT" Width="175" Height="42" Background="#D97706" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
                <Button Name="btnRemoveKMS" Content="🗑️ GỠ BỎ KMS LẬU" Width="145" Height="42" Background="#DC2626" Foreground="White" FontWeight="Bold" FontSize="12" BorderThickness="0" Cursor="Hand">
                    <Button.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="6"/>
                        </Style>
                    </Button.Resources>
                </Button>
            </StackPanel>
        </Grid>

        <!-- 3 Cards thông số -->
        <UniformGrid Grid.Row="1" Columns="3" Margin="0,0,0,25">
            <Border Name="cardWin" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="0,0,12,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="🖥️" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="HỆ ĐIỀU HÀNH WINDOWS" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtWinStatus" Text="Chưa kiểm tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtWinDetails" Text="Nhấn nút Quét để nhận thông tin chi tiết về phiên bản và kênh kích hoạt." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>

            <Border Name="cardOffice" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="6,0,6,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="📂" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="MICROSOFT OFFICE" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtOfficeStatus" Text="Chưa kiểm tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtOfficeDetails" Text="Nhấn nút Quét để phát hiện các phiên bản Office được cài đặt trên hệ thống." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>

            <Border Name="cardCrack" Background="#111827" BorderBrush="#1F2937" BorderThickness="1.5" CornerRadius="10" Margin="12,0,0,0" Padding="18">
                <StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="🛡️" FontSize="16" Margin="0,0,8,0"/>
                        <TextBlock Text="PHẦN MỀM BẺ KHÓA / CRACK" FontSize="11" FontWeight="Bold" Foreground="#94A3B8" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBlock Name="txtCrackStatus" Text="Chưa kiểm tra" FontSize="18" FontWeight="Bold" Foreground="#F1F5F9" Margin="0,15,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="txtCrackDetails" Text="Quét chữ ký số ứng dụng cài đặt, tệp Hosts và tác vụ ngầm để phát hiện crack..." FontSize="11.5" Foreground="#64748B" TextWrapping="Wrap" LineHeight="16"/>
                </StackPanel>
            </Border>
        </UniformGrid>

        <!-- Log Console -->
        <Grid Grid.Row="2" Margin="0,0,0,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Text="KẾT QUẢ QUÉT / THỰC THI CHI TIẾT" FontSize="12" FontWeight="Bold" Foreground="#94A3B8" Margin="0,0,0,8"/>
            <TextBox Name="txtLogs" Grid.Row="1" Background="#070A13" Foreground="#38BDF8" BorderBrush="#1E293B" BorderThickness="1.5" FontFamily="Consolas" FontSize="12.5" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" AcceptsReturn="True" Padding="15"/>
        </Grid>

        <!-- Footer -->
        <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Trạng thái hệ thống: Sẵn sàng | Đang chạy với quyền Administrator" FontSize="11" Foreground="#475569" VerticalAlignment="Center"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnExport" Content="Xuất Báo Cáo (.txt)" Width="160" Height="32" Background="#1E293B" Foreground="#E2E8F0" FontWeight="SemiBold" FontSize="12" BorderThickness="0" Cursor="Hand" IsEnabled="False">
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

# 8. KHỞI CHẠY VÀ RÀNG BUỘC SỰ KIỆN GIAO DIỆN
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

    # SỰ KIỆN CLICK: QUÉT HỆ THỐNG
    $UI_btnScan.Add_Click({
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnScan.Content = "ĐANG QUÉT..."
        
        $UI_txtWinStatus.Text = "Đang xử lý..."
        $UI_txtOfficeStatus.Text = "Đang xử lý..."
        $UI_txtCrackStatus.Text = "Đang xử lý..."
        
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $win = Get-WindowsActivation
        if ($win.RawStatus -eq 1) {
            $UI_txtWinStatus.Text = "Đã kích hoạt ✅"
            $UI_txtWinStatus.Foreground = $BrushGreen
            $UI_cardWin.BorderBrush = $BorderGreen
        } elseif ($win.RawStatus -eq 5 -or $win.RawStatus -eq 0) {
            $UI_txtWinStatus.Text = "Chưa kích hoạt ⚠️"
            $UI_txtWinStatus.Foreground = $BrushYellow
            $UI_cardWin.BorderBrush = $BorderNormal
        } else {
            $UI_txtWinStatus.Text = "Trạng thái khác ⚠️"
            $UI_txtWinStatus.Foreground = $BrushYellow
            $UI_cardWin.BorderBrush = $BorderNormal
        }
        $winServerInfo = ""
        if ($win.KMSServer -and $win.KMSServer -ne "Không phát hiện") {
            $winServerInfo = "`n⚠️ Máy chủ KMS: $($win.KMSServer)"
        }
        $UI_txtWinDetails.Text = "Kênh kích hoạt: $($win.Channel)`nPhiên bản: $($win.ProductName)$winServerInfo"
        
        $officeList = Get-OfficeActivation
        if ($officeList.Count -eq 0) {
            $UI_txtOfficeStatus.Text = "Chưa cài đặt / Không tìm thấy 📂"
            $UI_txtOfficeStatus.Foreground = $BrushSlate
            $UI_txtOfficeDetails.Text = "Không tìm thấy phiên bản Microsoft Office nào đã đăng ký khóa bản quyền."
            $UI_cardOffice.BorderBrush = $BorderNormal
        } else {
            $isOfficeActivated = $true
            $officeBrief = ""
            foreach ($off in $officeList) {
                $briefName = $off.ProductName
                if ($briefName.Length -gt 35) { $briefName = $briefName.Substring(0, 32) + "..." }
                $officeBrief += "• $briefName : $($off.Status)`n"
                if ($off.Status -ne "Đã kích hoạt" -and $off.Status -notlike "*Kích hoạt qua*") {
                    $isOfficeActivated = $false
                }
            }
            $UI_txtOfficeDetails.Text = $officeBrief.Trim()
            
            if ($isOfficeActivated) {
                $UI_txtOfficeStatus.Text = "Đã kích hoạt ✅"
                $UI_txtOfficeStatus.Foreground = $BrushGreen
                $UI_cardOffice.BorderBrush = $BorderGreen
            } else {
                $UI_txtOfficeStatus.Text = "Có phiên bản chưa kích hoạt ⚠️"
                $UI_txtOfficeStatus.Foreground = $BrushYellow
                $UI_cardOffice.BorderBrush = $BorderNormal
            }
        }
        
        $crack = Get-CrackDetection
        if ($crack.Risks.Count -eq 0) {
            $UI_txtCrackStatus.Text = "Hệ thống An toàn ✅"
            $UI_txtCrackStatus.Foreground = $BrushGreen
            $UI_cardCrack.BorderBrush = $BorderGreen
            $UI_txtCrackDetails.Text = "Không phát hiện bất kỳ tiến trình, tác vụ chạy ngầm, file hoặc tệp tin hosts nào bị can thiệp bởi công cụ bẻ khóa."
        } else {
            $UI_txtCrackStatus.Text = "Phát hiện vết bẻ khóa ❌"
            $UI_txtCrackStatus.Foreground = $BrushRed
            $UI_cardCrack.BorderBrush = $BorderRed
            $UI_txtCrackDetails.Text = "Phát hiện $($crack.Risks.Count) điểm cảnh báo bẻ khóa phần mềm hệ thống. Xem chi tiết bên dưới!"
        }
        
        $log = New-Object System.Text.StringBuilder
        $log.AppendLine("==================================================================================") | Out-Null
        $log.AppendLine("                   BÁO CÁO KIỂM TRA BẢN QUYỀN VÀ PHÁT HIỆN CRACK                  ") | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        $log.AppendLine("Thời gian thực hiện : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')") | Out-Null
        $log.AppendLine("Tên máy tính        : $env:COMPUTERNAME") | Out-Null
        $log.AppendLine("Hệ điều hành        : $( (Get-CimInstance Win32_OperatingSystem).Caption ) (Phiên bản: $((Get-CimInstance Win32_OperatingSystem).Version))") | Out-Null
        $log.AppendLine("----------------------------------------------------------------------------------") | Out-Null
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[1] KẾT QUẢ KIỂM TRA WINDOWS:") | Out-Null
        $log.AppendLine(" - Sản phẩm          : $($win.ProductName)") | Out-Null
        $log.AppendLine(" - Kênh phân phối    : $($win.Channel)") | Out-Null
        $log.AppendLine(" - Trạng thái        : $($win.StatusText)") | Out-Null
        if ($win.KMSServer -and $win.KMSServer -ne "Không phát hiện") {
            $log.AppendLine(" - Máy chủ KMS (Lậu) : $($win.KMSServer) (Phát hiện cấu hình bản quyền hướng ngoại)") | Out-Null
        }
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[2] KẾT QUẢ KIỂM TRA OFFICE:") | Out-Null
        if ($officeList.Count -eq 0) {
            $log.AppendLine(" - Không tìm thấy thông tin sản phẩm Microsoft Office.") | Out-Null
        } else {
            foreach ($off in $officeList) {
                $log.AppendLine(" - Tên phiên bản   : $($off.ProductName)") | Out-Null
                $log.AppendLine("   Trạng thái      : $($off.Status)") | Out-Null
                $log.AppendLine("   Kênh giấy phép  : $($off.Channel)") | Out-Null
                $log.AppendLine("   Nguồn phát hiện : $($off.Source)") | Out-Null
                if ($off.KMSServer) {
                    $log.AppendLine("   Máy chủ KMS     : $($off.KMSServer)") | Out-Null
                }
                $log.AppendLine("   --------------------------------------") | Out-Null
            }
        }
        $log.AppendLine() | Out-Null
        
        $log.AppendLine("[3] ĐÁNH GIÁ VÀ PHÁT HIỆN CRACK/BẺ KHÓA:") | Out-Null
        if ($crack.Risks.Count -eq 0) {
            $log.AppendLine(" -> [AN TOÀN] Không phát hiện bất kỳ tiến trình, tác vụ, tệp tin hosts hay thư mục bẻ khóa nào.") | Out-Null
        } else {
            $log.AppendLine(" -> [CẢNH BÁO NGUY HIỂM] Phát hiện $($crack.Risks.Count) cảnh báo bẻ khóa phần mềm:") | Out-Null
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                $log.AppendLine("   * Cảnh báo: $($crack.Risks[$i])") | Out-Null
                $log.AppendLine("     Mô tả   : $($crack.Details[$i])") | Out-Null
                $log.AppendLine("     ----------------------------------------------------------------") | Out-Null
            }
            $log.AppendLine() | Out-Null
            $log.AppendLine("KHUYẾN NGHỊ BẢO MẬT:") | Out-Null
            $log.AppendLine(" - Các phần mềm bẻ khóa có nguy cơ bị cài cắm mã độc lấy cắp dữ liệu doanh nghiệp.") | Out-Null
            $log.AppendLine(" - Nên gỡ bỏ bản crack và sử dụng các phiên bản chính hãng hoặc mã nguồn mở thay thế.") | Out-Null
        }
        $log.AppendLine() | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        
        $script:globalReportText = $log.ToString()
        $UI_txtLogs.Text = $script:globalReportText
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnScan.Content = "🔍 BẮT ĐẦU QUÉT"
        $UI_btnExport.IsEnabled = $true
    })

    # SỰ KIỆN CLICK: GỠ CÀI ĐẶT HÀNG LOẠT ỨNG DỤNG CRACK
    $UI_btnUninstallCracks.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show("Bạn có chắc chắn muốn TỰ ĐỘNG GỠ CÀI ĐẶT HÀNG LOẠT tất cả các phần mềm thương mại bị phát hiện bẻ khóa trên máy tính không?", "Xác nhận gỡ crack hàng loạt", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnUninstallCracks.Content = "ĐANG GỠ HÀNG LOẠT..."
        
        $UI_txtLogs.Text = "Đang tiến hành gỡ cài đặt tự động hàng loạt tất cả phần mềm bẻ khóa... Vui lòng đợi."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $res = Start-CrackedAppsUninstallProcess
        $script:globalReportText = $res.Log
        $UI_txtLogs.Text = $res.Log
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnUninstallCracks.Content = "⚡ GỠ CRACK HÀNG LOẠT"
        $UI_btnExport.IsEnabled = $true
        
        if ($res.Count -gt 0) {
            [System.Windows.MessageBox]::Show("Đã hoàn tất quá trình gỡ cài đặt hàng loạt $($res.Count) phần mềm bị crack!", "Thông báo", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            [System.Windows.MessageBox]::Show("Không tìm thấy phần mềm bên thứ ba nào đang bị crack trên hệ thống.", "Thông báo", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })

    # SỰ KIỆN CLICK: GỠ BỎ KMS LẬU
    $UI_btnRemoveKMS.Add_Click({
        $confirm = [System.Windows.MessageBox]::Show("Bạn có chắc chắn muốn gỡ bỏ máy chủ KMS lậu và làm sạch hệ thống không?`n`nHành động này sẽ gỡ Product Key KMS lậu và đưa Windows trở về trạng thái chưa kích hoạt chuẩn.", "Xác nhận gỡ KMS", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        
        $UI_btnScan.IsEnabled = $false
        $UI_btnUninstallCracks.IsEnabled = $false
        $UI_btnRemoveKMS.IsEnabled = $false
        $UI_btnRemoveKMS.Content = "ĐANG GỠ BỎ..."
        
        $UI_txtLogs.Text = "Đang tiến hành gỡ bỏ bản quyền KMS lậu và dọn dẹp hệ thống... Vui lòng đợi trong giây lát."
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $resultText = Start-KMSRemovalProcess
        $script:globalReportText = $resultText
        $UI_txtLogs.Text = $resultText
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnUninstallCracks.IsEnabled = $true
        $UI_btnRemoveKMS.IsEnabled = $true
        $UI_btnRemoveKMS.Content = "🗑️ GỠ BỎ KMS LẬU"
        $UI_btnExport.IsEnabled = $true
        [System.Windows.MessageBox]::Show("Đã hoàn tất quá trình gỡ bỏ bản quyền KMS lậu và làm sạch hệ thống!", "Thông báo", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    # SỰ KIỆN CLICK: XUẤT BÁO CÁO
    $UI_btnExport.Add_Click({
        if (-not $script:globalReportText) { return }
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
        $saveFileDialog.FileName = "BaoCao_BanQuyen_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        
        if ($saveFileDialog.ShowDialog() -eq $true) {
            try {
                $script:globalReportText | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8
                [System.Windows.MessageBox]::Show("Xuất báo cáo thành công tại:`n$($saveFileDialog.FileName)", "Thông báo", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } catch {
                [System.Windows.MessageBox]::Show("Lỗi khi ghi tệp: $($_.Exception.Message)", "Lỗi", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

    # HIỂN THỊ CỬA SỔ
    $Window.ShowDialog() | Out-Null

} catch {
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "LỖI THỰC THI GIAO DIỆN:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host "========================================================" -ForegroundColor Red
    Write-Host "Nhấn phím bất kỳ để đóng..." -ForegroundColor Yellow
    [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
