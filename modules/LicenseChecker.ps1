<#
    LicenseChecker.ps1 - Module chứa các hàm quét bản quyền hệ thống và phát hiện bẻ khóa.
#>

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
        # Lấy thông tin từ lớp SoftwareLicensingProduct
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
            
            # Phân tích kênh bản quyền
            if ($win.Description -like "*RETAIL*") {
                $result.Channel = "Retail (Bán lẻ)"
            }
            elseif ($win.Description -like "*OEM*") {
                $result.Channel = "OEM (Theo máy)"
            }
            elseif ($win.Description -like "*VOLUME_KMSCLIENT*") {
                $result.Channel = "Volume KMS (KMS Doanh nghiệp)"
            }
            elseif ($win.Description -like "*VOLUME_MAK*") {
                $result.Channel = "Volume MAK (Khóa kích hoạt nhiều máy)"
            }
            else {
                $result.Channel = "Khác/Tự cấu hình"
            }
            
            # Lấy địa chỉ máy chủ KMS từ dịch vụ cấp phép nếu có
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
    
    # 1. Truy vấn các giấy phép Office qua CIM/WMI
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
    
    # 2. Truy vấn bổ sung qua tệp cấu hình ospp.vbs
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
    
    # 3. Kiểm tra Office 365 ClickToRun trong Registry
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
    
    # 1. Kiểm tra các tệp tin Hook DLL (Can thiệp sâu vào tệp tin bản quyền hệ thống)
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
            $details += "Tệp tin '$path' thường do KMSpico hoặc KMSAuto cài đặt vào để giả lập dịch vụ cấp phép của Microsoft, tự động xác thực bản quyền giả mạo tại chỗ."
        }
    }
    
    # 2. Kiểm tra tệp tin hosts (Phát hiện chặn máy chủ bản quyền của Microsoft & Các hãng phần mềm)
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        $hostsLines = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $redirects = @()
        
        # Danh sách các miền thường bị chặn khi crack phần mềm bên thứ 3
        $thirdPartyDomains = @{
            "minitool.com" = "MiniTool Partition Wizard / Data Recovery"
            "adobe.com" = "Adobe Creative Cloud (Photoshop, Illustrator, Acrobat...)"
            "adobelogin.com" = "Adobe Creative Cloud"
            "autodesk.com" = "Autodesk (AutoCAD, Maya, 3ds Max...)"
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
                # Kiểm tra máy chủ Microsoft
                if ($clean -like "*microsoft*" -or $clean -like "*activation*" -or $clean -like "*kms*") {
                    if ($clean -match "(127\.0\.0\.1|0\.0\.0\.0)\s+(.*)") {
                        $redirects += $matches[2]
                    }
                }
                # Kiểm tra hãng phần mềm khác
                foreach ($domain in $thirdPartyDomains.Keys) {
                    if ($clean -like "*$domain*" -and $clean -match "(127\.0\.0\.1|0\.0\.0\.0)") {
                        $appName = $thirdPartyDomains[$domain]
                        if ($blockedApps -notcontains $appName) {
                            $blockedApps += $appName
                        }
                    }
                }
            }
        }
        if ($redirects.Count -gt 0) {
            $risks += "Sửa đổi tệp tin Hosts (Định tuyến chặn Microsoft)"
            $details += "Tệp tin Hosts cấu hình chuyển hướng chặn máy chủ Microsoft (ví dụ: $($redirects -join ', ')) về IP nội bộ. Việc này thường nhằm chặn Microsoft quét bản quyền lậu."
        }
        if ($blockedApps.Count -gt 0) {
            $risks += "Chặn máy chủ xác thực bản quyền phần mềm trong file Hosts"
            $details += "Tệp Hosts chứa cấu hình chặn kết nối đến hãng phần mềm: $($blockedApps -join ', '). Đây là dấu hiệu chắc chắn hệ thống đang chạy bản bẻ khóa của các phần mềm này."
        }
    }
    
    # 3. Kiểm tra các tác vụ chạy định kỳ (Scheduled Tasks) - Đã tối ưu hóa truy vấn trực tiếp (Nhanh gấp 10 lần)
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($task in $tasks) {
            $risks += "Tác vụ chạy ngầm bẻ khóa: $($task.TaskName)"
            $details += "Phát hiện tác vụ tự động tại đường dẫn '$($task.TaskPath)'. Đây là công cụ thường kỳ chạy lại để cập nhật ngày gia hạn kích hoạt KMS lậu."
        }
    }
    
    # 4. Kiểm tra dịch vụ chạy nền (Services)
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            $risks += "Dịch vụ KMS lậu chạy nền: $($srv.DisplayName) ($srvName)"
            $details += "Phát hiện dịch vụ '$($srv.DisplayName)' đang hoạt động trên hệ thống. Dịch vụ này giả lập một máy chủ KMS ngay trên máy để liên tục cấp bản quyền giả."
        }
    }

    # 5. Kiểm tra các tiến trình (Processes) bẻ khóa đang chạy ngầm
    $suspiciousProcessNames = @("KMSpico", "AutoKMS", "KMSAuto", "KMSConnectionMonitor", "HEU_KMS", "HEU_KMS_Activator")
    foreach ($procName in $suspiciousProcessNames) {
        $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($proc) {
            $risks += "Tiến trình bẻ khóa đang chạy: $procName"
            $details += "Tiến trình '$procName' (PID: $($proc.Id)) đang chạy ngầm trên hệ thống, cho thấy phần mềm bẻ khóa đang hoạt động thời gian thực."
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
                    $risks += "Thư mục cài đặt hệ thống chứa Script kích hoạt lậu: $folder"
                    $details += "Thư mục '$folder' chứa các tệp *.bat cấu hình bản quyền tự động. Thường xuất hiện trên các phiên bản Windows cài sẵn (bản ghost, bản cài đặt tải từ nguồn không uy tín)."
                }
            } else {
                $risks += "Phát hiện thư mục phần mềm bẻ khóa hệ thống: $folder"
                $details += "Thư mục cài đặt ứng dụng crack tồn tại ở '$folder'."
            }
        }
    }
    
    # 7. QUÉT ĐỘNG PHẦN MỀM BÊN THỨ 3 ĐANG CÀI ĐẶT TRÊN MÁY (Dynamic Third-Party Crack Scanning)
    $targetKeywords = @("MiniTool", "Adobe", "Autodesk", "AutoCAD", "Corel", "Camtasia", "TechSmith", "Internet Download Manager", "IDM", "Wondershare", "CCleaner", "SketchUp", "SolidWorks", "WinZip", "WinRAR")
    
    # Bản đồ chứa các tệp tin lõi lưu trữ bản quyền của từng hãng (phải có chữ ký số hợp lệ mới là hàng xịn)
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
    
    # Đọc danh sách cài đặt từ Registry
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName } |
        Select-Object DisplayName, InstallLocation, UninstallString, Publisher
        
    foreach ($app in $installedApps) {
        $matchedKeyword = $null
        foreach ($kw in $targetKeywords) {
            if ($app.DisplayName -like "*$kw*") {
                $matchedKeyword = $kw
                break
            }
        }
        
        if ($matchedKeyword) {
            # Giải quyết thư mục cài đặt thực tế của ứng dụng
            $installDir = $app.InstallLocation
            if (-not $installDir -and $app.UninstallString) {
                # Trích xuất thư mục từ chuỗi UninstallString (ví dụ: "C:\Program Files\...\unins000.exe")
                if ($app.UninstallString -match '^["]?([^"]+\.[eE][xX][eE])["]?') {
                    $exePath = $matches[1]
                    if (Test-Path $exePath) {
                        $installDir = Split-Path $exePath
                    }
                }
            }
            
            # Thực hiện kiểm tra nếu tìm thấy thư mục cài đặt hợp lệ
            if ($installDir -and (Test-Path $installDir)) {
                # Xác định danh sách tệp tin cần kiểm tra chữ ký số dựa theo từ khóa
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
                        # Tìm kiếm file tương ứng (tìm ở thư mục cài đặt gốc và sâu tối đa 2 cấp thư mục con)
                        $foundFiles = Get-ChildItem -Path $installDir -Filter $fileName -Recurse -Depth 2 -ErrorAction SilentlyContinue
                        foreach ($file in $foundFiles) {
                            $sig = Get-AuthenticodeSignature -FilePath $file.FullName -ErrorAction SilentlyContinue
                            
                            # Nếu tệp tin tồn tại nhưng KHÔNG có chữ ký số hợp lệ -> Rất có thể đã bị patch/crack
                            if ($sig -and $sig.Status -ne 'Valid') {
                                $risks += "Phát hiện $($app.DisplayName) bị Crack (Mất chữ ký số)"
                                $details += "Tệp tin bản quyền cốt lõi '$($file.FullName)' của phần mềm '$($app.DisplayName)' bị phát hiện mất chữ ký số hợp lệ (Trạng thái chữ ký: $($sig.Status)). Bản gốc chính hãng luôn được nhà sản xuất ký số; việc mất chữ ký cho thấy tệp tin này đã bị chỉnh sửa (patched) hoặc thay thế bằng tệp crack để vượt qua lớp kiểm định bản quyền."
                            }
                        }
                    }
                }
            }
        }
    }
    
    # 8. Kiểm tra xem máy chủ KMS của Windows & Office
    # A. Cấu hình Windows KMS
    $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
    if ($licService -and $licService.KeyManagementServiceMachine) {
        $kmsMachine = $licService.KeyManagementServiceMachine
        $knownKmsServers = @("kms8.msguides.com", "kms.digiboy.ir", "kms.lotro.cc", "kms.chinancce.com", "zh.us.to", "kms.msguides.com", "kms.spacespaces.xyz")
        
        if ($kmsMachine -eq "127.0.0.1" -or $kmsMachine -eq "localhost" -or $kmsMachine -eq "::1") {
            $risks += "Windows cấu hình máy chủ KMS nội bộ (Localhost)"
            $details += "Bản quyền Windows trỏ về chính máy tính ($kmsMachine). Đây là dấu hiệu của trình bẻ khóa giả lập cục bộ (như KMSpico, KMSAuto, MAS)."
        } else {
            foreach ($srv in $knownKmsServers) {
                if ($kmsMachine -like "*$srv*") {
                    $risks += "Windows sử dụng máy chủ KMS bẻ khóa công cộng: $kmsMachine"
                    $details += "Bản quyền hệ thống được kích hoạt bằng cách trỏ tới máy chủ KMS cộng đồng miễn phí ngoài Internet ($kmsMachine) để kích hoạt lậu."
                    break
                }
            }
        }
    }

    # B. Cấu hình Office KMS trong Registry
    $officeRegKms = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    if ($officeRegKms -and $officeRegKms.KeyManagementServiceMachine) {
        $kmsMachineOffice = $officeRegKms.KeyManagementServiceMachine
        if ($kmsMachineOffice -eq "127.0.0.1" -or $kmsMachineOffice -eq "localhost" -or $kmsMachineOffice -eq "::1") {
            $risks += "Office cấu hình máy chủ KMS nội bộ (Localhost)"
            $details += "Cấu hình Registry của Office đang trỏ máy chủ kích hoạt về chính máy tính ($kmsMachineOffice). Đây là dấu hiệu Office được bẻ khóa bằng công cụ KMS giả lập."
        }
    }

    # 9. Kiểm tra xem cổng KMS 1688 có đang lắng nghe cục bộ (Chỉ dấu cực kỳ mạnh mẽ của trình giả lập KMS)
    try {
        $kmsPort = Get-NetTCPConnection -LocalPort 1688 -State Listen -ErrorAction SilentlyContinue
        if ($kmsPort) {
            $risks += "Phát hiện cổng dịch vụ KMS (1688) đang mở"
            $details += "Cổng TCP 1688 (cổng KMS chuẩn) đang ở trạng thái LẮNG NGHE (Listen) trên máy tính này. Đây là dấu hiệu chắc chắn của một phần mềm KMS giả lập (như KMSpico, KMSAuto, HEU) đang chạy nền để kích hoạt bản quyền."
        }
    } catch {
        # Fallback bằng lệnh netstat nếu chạy trên phiên bản hệ điều hành cũ
        $netstat = netstat -ano 2>&1 | Select-String ":1688\s+.*LISTENING"
        if ($netstat) {
            $risks += "Phát hiện cổng dịch vụ KMS (1688) đang mở (netstat)"
            $details += "Phát hiện cổng 1688 đang mở ở trạng thái LISTENING qua lệnh netstat. Cho thấy dịch vụ giả lập KMS nội bộ đang chạy trên máy tính."
        }
    }

    # 10. Kiểm tra khóa Registry của phần mềm bẻ khóa
    $crackRegs = @(
        "HKLM:\SOFTWARE\KMSAuto",
        "HKLM:\SOFTWARE\KMSAutoS",
        "HKLM:\SOFTWARE\KMSpico",
        "HKCU:\Software\KMSAuto",
        "HKCU:\Software\KMSAutoS"
    )
    foreach ($reg in $crackRegs) {
        if (Test-Path $reg) {
            $risks += "Phát hiện khóa Registry bẻ khóa: $reg"
            $details += "Khóa Registry liên quan đến công cụ KMSAuto/KMSpico được phát hiện tại '$reg', cho thấy phần mềm bẻ khóa đã từng hoặc đang được cài đặt."
        }
    }
    
    return [PSCustomObject]@{
        Risks = $risks
        Details = $details
    }
}
