<#
    Remove-KMSActivation-Standalone.ps1 - Công cụ gỡ bỏ bản quyền KMS lậu và làm sạch hệ thống Windows.
    Tác dụng:
    1. Gỡ bỏ địa chỉ máy chủ KMS lậu (slmgr /ckms).
    2. Gỡ bỏ Product Key Volume KMS lậu đang cài trên Windows (slmgr /upk).
    3. Xóa thông tin Product Key khỏi Registry (slmgr /cpky).
    4. Tự động dọn dẹp các dòng chặn máy chủ Microsoft trong tệp tin Hosts.
    5. Xóa bỏ các tác vụ ngầm (Scheduled Tasks) và dịch vụ (Services) bẻ khóa (AutoKMS, KMSAuto...).
#>

[CmdletBinding()]
param (
    [switch]$ConsoleOnly
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
    Write-Host "CẢNH BÁO: Công cụ gỡ KMS yêu cầu quyền Administrator!" -ForegroundColor Yellow
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

# 3. HÀM XỬ LÝ GỠ BỎ KMS VÀ CÁC DẤU VẾT BẺ KHÓA
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

# 4. CHẾ ĐỘ DÒNG LỆNH (CLI)
if ($ConsoleOnly) {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "             CÔNG CỤ GỠ BỎ BẢN QUYỀN WINDOWS KMS LẬU                    " -ForegroundColor White -BackgroundColor DarkRed
    Write-Host "==========================================================================" -ForegroundColor Cyan
    $result = Start-KMSRemovalProcess
    Write-Host $result
    Exit
}

# 5. ĐỊNH NGHĨA GIAO DIỆN WPF (GUI)
$xamlContent = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Công cụ gỡ bỏ bản quyền KMS lậu &amp; Dọn dẹp hệ thống" Height="600" Width="850" 
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
        </Grid.RowDefinitions>

        <!-- Header -->
        <Grid Grid.Row="0" Margin="0,0,0,20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
                <TextBlock Text="CÔNG CỤ GỠ BỎ BẢN QUYỀN KMS LẬU" FontSize="18" FontWeight="Bold" Foreground="#EF4444"/>
                <TextBlock Text="Gỡ bỏ hoàn toàn máy chủ KMS giả lập, gỡ Key lậu và dọn dẹp các tác vụ bẻ khóa ngầm" FontSize="12" Foreground="#94A3B8" Margin="0,4,0,0"/>
            </StackPanel>
            <Button Name="btnRemove" Grid.Column="1" Content="THỰC HIỆN GỠ KMS" Width="180" Height="40" Background="#DC2626" Foreground="White" FontWeight="Bold" FontSize="12.5" BorderThickness="0" Cursor="Hand">
                <Button.Resources>
                    <Style TargetType="Border">
                        <Setter Property="CornerRadius" Value="6"/>
                    </Style>
                </Button.Resources>
            </Button>
        </Grid>

        <!-- Card cảnh báo -->
        <Border Grid.Row="1" Background="#18181B" BorderBrush="#27272A" BorderThickness="1" CornerRadius="8" Padding="15" Margin="0,0,0,20">
            <TextBlock Text="⚠️ LƯU Ý: Công cụ này sẽ gỡ bỏ khóa bản quyền KMS lậu đang cài trên Windows và đưa máy tính về trạng thái chưa kích hoạt chuẩn. Sau khi gỡ, bạn có thể nhập khóa bản quyền chính hãng (Retail/OEM) để kích hoạt lại." FontSize="11.5" Foreground="#FBBF24" TextWrapping="Wrap" LineHeight="18"/>
        </Border>

        <!-- Khu vực Log Console -->
        <Grid Grid.Row="2">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Text="NHẬT KÝ THỰC THI GỠ BỎ" FontSize="11.5" FontWeight="Bold" Foreground="#94A3B8" Margin="0,0,0,8"/>
            <TextBox Name="txtLogs" Grid.Row="1" Background="#070A13" Foreground="#F87171" BorderBrush="#1E293B" BorderThickness="1.5" FontFamily="Consolas" FontSize="12" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" AcceptsReturn="True" Padding="12"/>
        </Grid>
    </Grid>
</Window>
'@

# 6. KHỞI CHẠY GIAO DIỆN WPF
try {
    $xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xamlContent)
    $Window = [Windows.Markup.XamlReader]::Load($xmlReader)

    $UI_btnRemove = $Window.FindName("btnRemove")
    $UI_txtLogs = $Window.FindName("txtLogs")

    $UI_btnRemove.Add_Click({
        $UI_btnRemove.IsEnabled = $false
        $UI_btnRemove.Content = "ĐANG GỠ BỎ..."
        $UI_txtLogs.Text = "Đang tiến hành gỡ bỏ bản quyền KMS lậu và dọn dẹp hệ thống... Vui lòng đợi trong giây lát."
        
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        $resultText = Start-KMSRemovalProcess
        $UI_txtLogs.Text = $resultText
        
        $UI_btnRemove.IsEnabled = $true
        $UI_btnRemove.Content = "THỰC HIỆN GỠ KMS"
        [System.Windows.MessageBox]::Show("Đã hoàn tất quá trình gỡ bỏ bản quyền KMS lậu và làm sạch hệ thống!", "Thông báo", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })

    $Window.ShowDialog() | Out-Null
} catch {
    Write-Host "Lỗi thực thi giao diện: $($_.Exception.Message)" -ForegroundColor Red
    [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
