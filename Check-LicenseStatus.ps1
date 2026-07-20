<#
.SYNOPSIS
    Công cụ kiểm tra bản quyền Windows, Office và phát hiện phần mềm crack.
.DESCRIPTION
    Phiên bản đã cấu trúc hóa dạng mô-đun (chia nhỏ mã nguồn thành XAML và mã quét riêng).
#>

[CmdletBinding()]
param (
    [switch]$ConsoleOnly
)

# 1. NẠP THƯ VIỆN WPF & THIẾT LẬP STA
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
} catch {
    Write-Warning "Không thể nạp thư viện WPF. Giao diện đồ họa có thể không hoạt động."
}

# 2. Tự động kiểm tra và nâng quyền Administrator (Ép chạy chế độ -sta)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "CẢNH BÁO: Công cụ yêu cầu quyền Administrator để kiểm tra." -ForegroundColor Yellow
    Write-Host "Đang mở cửa sổ PowerShell mới dưới quyền Administrator..." -ForegroundColor Yellow
    Write-Host "Vui lòng chọn 'Yes' (Có) trên màn hình Windows UAC." -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-sta", "-File", "`"$PSCommandPath`"" -Verb RunAs
    } catch {
        Write-Host "Lỗi: Không thể khởi chạy dưới quyền Administrator hoặc bị từ chối nâng quyền!" -ForegroundColor Red
        Start-Sleep -Seconds 3
    }
    Exit
}

# 3. Nạp module quét License chứa các hàm kiểm tra từ thư mục con
$CheckerModulePath = Join-Path $PSScriptRoot "modules\LicenseChecker.ps1"
if (Test-Path $CheckerModulePath) {
    . $CheckerModulePath
} else {
    Write-Error "Không tìm thấy tệp tin logic tại: $CheckerModulePath"
    Write-Host "Vui lòng đảm bảo thư mục 'modules' nằm cùng thư mục với script này." -ForegroundColor Red
    Start-Sleep -Seconds 5
    Exit
}

try {
    # 4. CHẾ ĐỘ DÒNG LỆNH (CLI MODE)
    if ($ConsoleOnly) {
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
        Exit
    }

    # 5. CHẾ ĐỘ GIAO DIỆN ĐỒ HỌA (GUI MODE)
    $XamlPath = Join-Path $PSScriptRoot "modules\GUI.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "Không tìm thấy tệp giao diện: $XamlPath"
    }

    # Đọc và phân tích XAML
    $xamlContent = Get-Content -Path $XamlPath -Raw -Encoding utf8
    $xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xamlContent)
    $Window = [Windows.Markup.XamlReader]::Load($xmlReader)

    # Tự động trích xuất các thành phần có Name trong XAML sang biến PowerShell
    $xamlObj = [xml]$xamlContent
    $xamlObj.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name "UI_$($_.Name)" -Value $Window.FindName($_.Name)
    }

    # Định nghĩa các Brush màu sắc hiển thị
    $BrushGreen = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10B981")
    $BrushRed = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#EF4444")
    $BrushYellow = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F59E0B")
    $BrushWhite = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F1F5F9")
    $BrushSlate = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#475569")
    $BorderRed = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#7F1D1D")
    $BorderGreen = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#064E3B")
    $BorderNormal = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1F2937")

    $globalReportText = ""

    # Xử lý sự kiện click nút Quét Hệ Thống
    $UI_btnScan.Add_Click({
        $UI_btnScan.IsEnabled = $false
        $UI_btnScan.Content = "ĐANG QUÉT..."
        
        $UI_txtWinStatus.Text = "Đang xử lý..."
        $UI_txtOfficeStatus.Text = "Đang xử lý..."
        $UI_txtCrackStatus.Text = "Đang xử lý..."
        
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        # Gọi hàm kiểm tra Windows
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
        
        # Gọi hàm kiểm tra Office
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
        
        # Gọi hàm quét bẻ khóa/crack
        $crack = Get-CrackDetection
        if ($crack.Risks.Count -eq 0) {
            $UI_txtCrackStatus.Text = "Hệ thống An toàn ✅"
            $UI_txtCrackStatus.Foreground = $BrushGreen
            $UI_cardCrack.BorderBrush = $BorderGreen
            $UI_txtCrackDetails.Text = "Không phát hiện bất kỳ dấu hiệu, tác vụ chạy ngầm, file hoặc tệp tin hosts nào bị can thiệp bởi công cụ bẻ khóa."
        } else {
            $UI_txtCrackStatus.Text = "Phát hiện vết bẻ khóa ❌"
            $UI_txtCrackStatus.Foreground = $BrushRed
            $UI_cardCrack.BorderBrush = $BorderRed
            $UI_txtCrackDetails.Text = "Phát hiện $($crack.Risks.Count) dấu vết/công cụ bẻ khóa hệ thống trên máy tính. Xem chi tiết bên dưới!"
        }
        
        # Tổng hợp báo cáo chi tiết
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
            $log.AppendLine(" -> [NGUY HIỂM] Phát hiện $($crack.Risks.Count) điểm cảnh báo về việc bẻ khóa:") | Out-Null
            for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
                $log.AppendLine("   * Điểm cảnh báo: $($crack.Risks[$i])") | Out-Null
                $log.AppendLine("     Chi tiết      : $($crack.Details[$i])") | Out-Null
                $log.AppendLine("     ----------------------------------------------------------------") | Out-Null
            }
            $log.AppendLine() | Out-Null
            $log.AppendLine("KHUYẾN NGHỊ BẢO MẬT:") | Out-Null
            $log.AppendLine(" - Các phần mềm crack như KMSpico, KMSAuto có nguy cơ cài cắm mã độc (RAT, Coin Miner).") | Out-Null
            $log.AppendLine(" - Nên gỡ bỏ công cụ crack bằng cách chạy lệnh xóa KMS: 'slmgr /ckms', 'slmgr /upk'.") | Out-Null
            $log.AppendLine(" - Khuyên dùng Windows & Office chính hãng để đảm bảo an toàn cho dữ liệu.") | Out-Null
        }
        $log.AppendLine() | Out-Null
        $log.AppendLine("==================================================================================") | Out-Null
        
        $script:globalReportText = $log.ToString()
        $UI_txtLogs.Text = $script:globalReportText
        
        $UI_btnScan.IsEnabled = $true
        $UI_btnScan.Content = "BẮT ĐẦU QUÉT"
        $UI_btnExport.IsEnabled = $true
    })

    # Xử lý sự kiện click nút Xuất Báo Cáo
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

    # Hiển thị cửa sổ
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
