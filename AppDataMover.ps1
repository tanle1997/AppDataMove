<#
.SYNOPSIS
App Data Mover - Move application data folders to a new location with optional symbolic links.
.DESCRIPTION
This PowerShell script provides a GUI tool to move application data folders for selected applications to a new destination. It supports automatic scanning of known applications, manual folder selection, and the creation of symbolic links to maintain application functionality. The tool can run with or without administrator privileges, with symlink support requiring admin rights.
.NOTES
Author: Tan Le
Version: 0.1
Date: Dec 2025
Email: pro.ngoctan@gmail.com
Website: https://www.co-workspace.io.vn
GitHub: https://github.com/tanle1997
LICENSE: MIT License
.LINK
https://www.co-workspace.io.vn/coding/self-help-tools/app-data-mover
#>

#region 1. LANGUAGE DATA
$LangData = @{
    EN = @{
        Title = "App Data Mover";
        ModeAuto = "MODE: AUTO SCAN"; ModeManual = "MODE: MANUAL SELECT";
        TipMode = "Click to toggle between Database Scan and Manual Path.";
        
        GrpAuto = "Select Application:"; GrpFolder = "Select Folder(s):";
        GrpManual = "Source Folder:"; GrpDest = "Destination Folder:";
        OptSymlink = "Create Symbolic Link"; 
        TipSymlink = "If checked: Original folder becomes .OLD, and a Link is created pointing to Dest.";
        OptForce = "Force Close Apps (Risk)";
        TipForce = "If checked: Apps will be terminated immediately. Uncheck to close manually.";
        BtnStart = "START MOVE"; BtnAdmin = "Run as Administrator"; BtnLog = "Open Log File";
        LblProgress = "Progress:"; LblStatus = "Ready."; LogHeader = "Live Log";
        LangName = "English";
        
        MsgWaitUser = "Waiting for user action...";
        MsgAborted = "Operation aborted by user.";
        MsgAdminReady = "Running as Administrator. Symlinks Enabled.";
        MsgUserWarn = "Running as User. Please restart as Admin for Symlink support.";
        MsgScanTarget = "Scanning target: "; MsgChecking = "Checking: ";
        MsgLinked = "Linked"; MsgFound = "Found"; MsgError = "Access Denied"; MsgMissing = "Missing";
        MsgNewRun = "=== NEW RUN STARTED ===";
        MsgSymlinkWarn = "WARNING: You selected 'Create Symbolic Link'.`n`nOriginal folders will be RENAMED to .OLD.`nAre you sure you want to proceed?";
        MsgDestEmpty = "Error: Please select a Destination folder!";
        MsgSourceInvalid = "Error: Invalid Source Folder."; MsgNoSelect = "Error: No items selected!";
        MsgCalcSize = "Calculating stats (files/size)...";
        MsgDiskFull = "Error: Not enough free space on destination!";
        MsgProcessRun = "Process running: "; MsgForceKill = "Force closing: ";
        MsgManualKill = "App '{0}' is running. Please close it manually.";
        MsgKillDone = "Closed app: "; MsgTask = "--- Task {0}: {1} ---";
        MsgCopyOK = "Copy Success."; MsgLinkOK = "Symlink Created."; MsgLinkFail = "Link Failed: ";
        MsgRollback = "Rolling back..."; MsgRollbackDone = "Rollback Done.";
        MsgAllDone = "ALL TASKS COMPLETED."; MsgCleanup = "Cleanup: Deleted ";
        MsgAskCleanup = "Finished! Created {0} backups (.OLD).`nDelete them now to free space?";
        MsgStatLog = "Stats: {0} Folders, {1} Files ready to move.";
        MsgSummaryLog = "SUMMARY: Moved {0} folders, {1} files.";
        MsgScanDone = "Scan completed.`n";
        Robo0="[Code: 0] OK (No changes)"; Robo1="[Code: 1] OK (Copy success)";
        Robo8="[Code: 8] FAIL (Retry limit - File locked?)"; Robo16="[Code: 16] FATAL (Invalid Path/Access Denied)";
    };
    VI = @{
        Title = "Chuyển Dữ Liệu Ứng Dụng";
        ModeAuto = "CHẾ ĐỘ: TỰ ĐỘNG QUÉT"; ModeManual = "CHẾ ĐỘ: CHỌN THỦ CÔNG";
        TipMode = "Bấm để chuyển đổi giữa Tự động quét và Chọn thủ công.";

        GrpAuto = "Chọn Ứng Dụng:"; GrpFolder = "Danh sách thư mục:";
        GrpManual = "Thư mục nguồn:"; GrpDest = "Thư mục đích:";
        OptSymlink = "Tạo liên kết (Symlink)"; 
        TipSymlink = "Nếu chọn: Thư mục gốc sẽ đổi tên thành .OLD và tạo Link trỏ đến Đích.";
        OptForce = "Buộc tắt App (Cẩn thận)";
        TipForce = "Nếu chọn: App sẽ bị tắt ngay lập tức (có thể mất dữ liệu). Bỏ chọn để tắt thủ công.";
        BtnStart = "BẮT ĐẦU"; BtnAdmin = "Chạy quyền Admin"; BtnLog = "Mở File Log";
        LblProgress = "Tiến độ:"; LblStatus = "Sẵn sàng."; LogHeader = "Nhật ký hoạt động";
        LangName = "Tiếng Việt";
        
        MsgWaitUser = "Đang chờ thao tác...";
        MsgAborted = "Đã hủy bỏ bởi người dùng.";
        MsgAdminReady = "Đang chạy quyền Admin. Sẵn sàng tạo Link.";
        MsgUserWarn = "Đang chạy quyền User. Cần Admin để tạo Link.";
        MsgScanTarget = "Đang quét: "; MsgChecking = "Kiểm tra: ";
        MsgLinked = "Đã Liên Kết"; MsgFound = "Tìm Thấy"; MsgError = "Lỗi Truy Cập"; MsgMissing = "Không Thấy";
        MsgNewRun = "=== PHIÊN CHẠY MỚI ===";
        MsgSymlinkWarn = "CẢNH BÁO: Bạn chọn 'Tạo Symlink'.`nThư mục gốc sẽ bị ĐỔI TÊN thành .OLD.`nBạn có chắc chắn?";
        MsgDestEmpty = "Lỗi: Chưa chọn thư mục Đích!";
        MsgSourceInvalid = "Lỗi: Thư mục Nguồn không hợp lệ."; MsgNoSelect = "Lỗi: Chưa chọn mục nào!";
        MsgCalcSize = "Đang tính toán số lượng...";
        MsgDiskFull = "Lỗi: Ổ đĩa đích không đủ chỗ trống!";
        MsgProcessRun = "App đang chạy: "; MsgForceKill = "Buộc tắt: ";
        MsgManualKill = "App '{0}' đang chạy. Vui lòng tắt thủ công.";
        MsgKillDone = "Đã tắt: "; MsgTask = "--- Tác vụ {0}: {1} ---";
        MsgCopyOK = "Copy Thành công."; MsgLinkOK = "Tạo Link Thành công."; MsgLinkFail = "Lỗi tạo Link: ";
        MsgRollback = "Đang khôi phục (Rollback)..."; MsgRollbackDone = "Khôi phục xong.";
        MsgAllDone = "HOÀN TẤT TOÀN BỘ."; MsgCleanup = "Dọn dẹp: Đã xóa ";
        MsgAskCleanup = "TỔNG KẾT:`n- Thư mục đã chuyển: {0}`n- Tệp tin đã chuyển: {1}`n- Bản backup đã tạo: {2}`n`nBạn có muốn XÓA các file backup (.OLD) ngay bây giờ?";
        MsgStatLog = "Thống kê: {0} Thư mục, {1} File sẽ được chuyển.";
        MsgSummaryLog = "TỔNG KẾT: Đã chuyển {0} thư mục, {1} tệp tin.";
        MsgScanDone = "Quét hoàn tất.`n";
        Robo0="[Mã: 0] OK (Không thay đổi)"; Robo1="[Mã: 1] OK (Copy thành công)";
        Robo8="[Mã: 8] LỖI COPY (File đang mở?)"; Robo16="[Mã: 16] LỖI NGHIÊM TRỌNG (Sai đường dẫn/Quyền)";
    }
}
$CurrentLang = $LangData.EN # DEFAULT ENGLISH
#endregion

#region 2. MODERN FOLDER EXPLORER (C#)
$ModernFolderPickerCode = @"
using System;
using System.Runtime.InteropServices;

namespace Win32
{
    public class FolderPicker
    {
        [DllImport("shell32.dll")]
        private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, ref Guid riid, out IShellItem ppv);

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        private const string IID_IFileDialog = "42f85136-db7e-439c-85f1-e4075d135fc8";
        private const string IID_IShellItem  = "43826d1e-e718-42ee-bc55-a1e261c37bfe";
        private const string CLSID_FileOpenDialog = "DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7";

        [ComImport, Guid(IID_IFileDialog), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IFileDialog
        {
            [PreserveSig] int Show(IntPtr parent);
            void SetFileTypes();
            void SetFileTypeIndex();
            void GetFileTypeIndex();
            void Advise();
            void Unadvise();
            void SetOptions(uint fos);
            void GetOptions();
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder();
            void GetCurrentSelection();
            void SetFileName();
            void GetFileName();
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
            void SetOkButtonLabel();
            void SetFileNameLabel();
            void GetResult(out IShellItem ppsi);
        }

        [ComImport, Guid(IID_IShellItem), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellItem
        {
            void BindToHandler();
            void GetParent();
            void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        }

        [ComImport, Guid(CLSID_FileOpenDialog), ClassInterface(ClassInterfaceType.None)]
        internal class FileOpenDialogRCW { }

        public static string Show(string initialDirectory)
        {
            try
            {
                IFileDialog dialog = (IFileDialog)new FileOpenDialogRCW();
                // FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM
                dialog.SetOptions((uint)0x00000020 | (uint)0x00000040);
                dialog.SetTitle("Select Folder");

                if (!string.IsNullOrEmpty(initialDirectory))
                {
                    Guid guid = new Guid(IID_IShellItem);
                    IShellItem item;
                    if (SHCreateItemFromParsingName(initialDirectory, IntPtr.Zero, ref guid, out item) == 0)
                    {
                        dialog.SetFolder(item);
                    }
                }

                if (dialog.Show(GetActiveWindow()) == 0)
                {
                    IShellItem result;
                    dialog.GetResult(out result);
                    string path;
                    // SIGDN_FILESYSPATH
                    result.GetDisplayName(0x80058000, out path);
                    return path;
                }
            }
            catch { }
            return null;
        }
    }
}
"@
#endregion

#region LIBRARY
$Library = @(
    @{
        App = 'Zalo'
        Components = @(
            @{ Label="Zalo PC Data";            Path="$env:LocalAppData\ZaloPC";          Process="Zalo" },
            @{ Label="Zalo Settings";           Path="$env:AppData\ZaloData";             Process="Zalo" },
            @{ Label="Received Files";          Path="$env:UserProfile\Documents\Zalo Received Files"; Process="Zalo" },
            @{ Label="Zalo Program Local";      Path="$env:LocalAppData\Programs\Zalo";   Process="Zalo" },
            @{ Label="Zalo Program x86";        Path="${env:ProgramFiles(x86)}\Zalo";     Process="Zalo" },
            @{ Label="Zalo Updater Roaming";    Path="$env:AppData\zalo-updater";         Process="Zalo" },
            @{ Label="Zalo Updater Local";      Path="$env:LocalAppData\zalo-updater";    Process="Zalo" }
        )
    }
    @{
        App = 'Telegram'
        Components = @(
            @{ Label="Telegram Desktop Data";   Path="$env:AppData\Telegram Desktop";     Process="Telegram" },
            @{ Label="Telegram Downloads";      Path="$env:UserProfile\Downloads\Telegram Desktop"; Process="Telegram" }
        )
    }
    @{
        App = 'Browser Data'
        Components = @(
            @{ Label="Chrome User Data";        Path="$env:LocalAppData\Google\Chrome\User Data"; Process="chrome" },
            @{ Label="Edge User Data";          Path="$env:LocalAppData\Microsoft\Edge\User Data"; Process="msedge" },
            @{ Label="CocCoc User Data";        Path="$env:LocalAppData\CocCoc\Browser\User Data"; Process="browser" }
        )
    }
    @{
        App = 'Viber'
        Components = @(
            @{ Label="Viber PC Data";           Path="$env:LocalAppData\ViberPC";         Process="Viber" },
            @{ Label="Viber Downloads";         Path="$env:UserProfile\Documents\ViberDownloads"; Process="Viber" }
        )
    }
)
$Config = @{ LogFile = "$env:TEMP\AppDataMover.log" }
#endregion

#region INIT
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Clear-Host
# RESTART AS STA IF NEEDED
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne "STA") {
    $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"
    $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $proc.Verb = "RunAs"
    [System.Diagnostics.Process]::Start($proc); Exit
}
# LOAD WPF ASSEMBLIES
try {
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Xml, WindowsBase -ErrorAction Stop

    # Load C# Code only if not already loaded
    if (-not ([System.Management.Automation.PSTypeName]'Win32.FolderPicker').Type) {
        Add-Type -TypeDefinition $ModernFolderPickerCode -Language CSharp
    }

    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match "PresentationCore" })) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("PresentationCore")
    }
} catch { # MISSING DEPENDENCY
    $wshell = New-Object -ComObject WScript.Shell # Create popup
    $answer = $wshell.Popup("Tool needs .NET Framework (WPF). Download now?", 0, "Missing Dependency", 4 + 16) # Yes/No + Warning icon
    if ($answer -eq 6) { Start-Process "https://go.microsoft.com/fwlink/?linkid=2088631" } # Open .NET download link
    Exit
}
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
#endregion

#region XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="MainWindow"
        Title="App Data Mover"
        Width="1280" Height="900" WindowStartupLocation="CenterScreen"
        Background="#F7F9FC" FontFamily="Segoe UI" FontSize="14"
        SnapsToDevicePixels="True" UseLayoutRounding="True"
        ResizeMode="CanMinimize">
    
    <Window.Resources>
        <Style x:Key="BlueButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#2278c7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#1A5C99"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="2">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3385D6"/> 
                                <Setter TargetName="border" Property="BorderBrush" Value="Black"/> 
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1A5C99"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#CCCCCC"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#AAAAAA"/>
                                <Setter Property="Foreground" Value="#888888"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="StdButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#F0F0F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#999999"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="2">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="White"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="Black"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ToggleStyle" TargetType="ToggleButton">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="#555"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2" Padding="10,0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E6F4FF"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#2278c7"/>
                                <Setter Property="Foreground" Value="#2278c7"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#999"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <Grid Grid.Row="0" Margin="0,0,0,20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="TitleBlock" Text="App Data Mover" FontWeight="Bold" FontSize="28" Foreground="#2278c7" HorizontalAlignment="Left"/>
            
            <ToggleButton x:Name="BtnLangToggle" Grid.Column="1" Content="English" Height="32" Style="{StaticResource ToggleStyle}" FocusVisualStyle="{x:Null}"/>
        </Grid>
        
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="420"/> 
                <ColumnDefinition Width="*"/>   
            </Grid.ColumnDefinitions>

            <Grid Grid.Column="0" Margin="0,0,24,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/> 
                    <RowDefinition Height="Auto"/> 
                    <RowDefinition Height="*"/>    
                    <RowDefinition Height="Auto"/> 
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,20" HorizontalAlignment="Center">
                    <ToggleButton x:Name="BtnModeToggle" Content="MODE: AUTO SCAN" IsChecked="True" Height="40" Width="390" Style="{StaticResource ToggleStyle}" FocusVisualStyle="{x:Null}"/>
                </StackPanel>

                <StackPanel Grid.Row="1">
                    <StackPanel x:Name="AutoPanel" Visibility="Visible">
                        <TextBlock x:Name="LblGrpAuto" Text="Select App:" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <ComboBox x:Name="AppCombo" VerticalContentAlignment="Center" Margin="0,0,0,12" Height="32" Padding="6,4"/>
                        <TextBlock x:Name="LblGrpFolder" Text="Folders:" FontWeight="Normal" Margin="0,0,0,6"/>
                        <Border BorderBrush="#ccc" BorderThickness="1" CornerRadius="4" Background="White" Height="250">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel x:Name="CompPanel" Orientation="Vertical" Margin="8"/>
                            </ScrollViewer>
                        </Border>
                    </StackPanel>

                    <StackPanel x:Name="ManualPanel" Visibility="Collapsed">
                        <TextBlock x:Name="LblGrpManual" Text="Source:" FontWeight="SemiBold" Margin="0,5,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                            <TextBox x:Name="SourcePathBox" Grid.Column="0" Height="32" IsReadOnly="True" VerticalContentAlignment="Center"/>
                            <Button x:Name="BrowseSourceBtn" Grid.Column="1" Content="..." Width="45" Margin="8,0,0,0" Style="{StaticResource StdButtonStyle}"/>
                        </Grid>
                    </StackPanel>

                    <TextBlock x:Name="LblGrpDest" Text="Destination:" FontWeight="SemiBold" Margin="0,15,0,6"/>
                    <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBox x:Name="DestPathBox" Grid.Column="0" Height="32" IsReadOnly="True" VerticalContentAlignment="Center"/>
                        <Button x:Name="BrowseDestBtn" Grid.Column="1" Content="..." Width="45" Margin="8,0,0,0" Style="{StaticResource StdButtonStyle}"/>
                    </Grid>
                </StackPanel>

                <StackPanel Grid.Row="3" VerticalAlignment="Bottom">
                    <CheckBox x:Name="SymlinkBox" Content="Symlink" FontWeight="SemiBold" Margin="0,10,0,8" ToolTip="Original folder becomes .OLD" VerticalContentAlignment="Center" Padding="6,0,0,0" FocusVisualStyle="{x:Null}"/>
                    <CheckBox x:Name="ForceKillBox" Content="Force Kill" Foreground="Red" Margin="0,0,0,20" ToolTip="Terminate apps immediately" VerticalContentAlignment="Center" Padding="6,0,0,0" FocusVisualStyle="{x:Null}"/>
                    
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,15">
                        <Button x:Name="StartBtn" Content="START" Width="140" Height="40" Margin="0,0,10,0" Style="{StaticResource BlueButtonStyle}" FocusVisualStyle="{x:Null}"/>
                        <Button x:Name="LogBtn" Content="Open Log" Width="120" Height="40" Style="{StaticResource StdButtonStyle}" FocusVisualStyle="{x:Null}"/>
                    </StackPanel>

                    <TextBlock x:Name="LblProgress" Text="Progress:" FontWeight="SemiBold" Margin="0,5,0,6"/>
                    <ProgressBar x:Name="ProgBar" Height="14" Minimum="0" Maximum="100" Value="0" Background="#eee" BorderBrush="#ccc"/>
                    <TextBlock x:Name="StatusBlock" Margin="0,8,0,15" Foreground="#555" TextWrapping="Wrap" FontSize="13" Text="..."/>
                    
                    <Button x:Name="RestartAdminBtn" Content="Run as Admin" Width="200" Height="35" HorizontalAlignment="Center" Style="{StaticResource StdButtonStyle}" FocusVisualStyle="{x:Null}"/>
                </StackPanel>
            </Grid>

            <GroupBox x:Name="GrpLog" Grid.Column="1" Header="Log" BorderBrush="#ccc" BorderThickness="1">
                <RichTextBox x:Name="LogBox" VerticalScrollBarVisibility="Auto" IsReadOnly="True" 
                             FontFamily="Consolas" FontSize="13" Background="White" BorderThickness="0" Padding="10">
                    <FlowDocument>
                        <Paragraph><Run Text="..." Foreground="Gray"/></Paragraph>
                    </FlowDocument>
                </RichTextBox>
            </GroupBox>
        </Grid>
    </Grid>
</Window>
"@
#endregion

#region LOGIC
$StringReader = New-Object System.IO.StringReader($xaml)
$XmlReader    = [System.Xml.XmlReader]::Create($StringReader)
$window       = [System.Windows.Markup.XamlReader]::Load($XmlReader)

# Bind Controls
$BtnLangToggle = $window.FindName('BtnLangToggle')
$TitleBlock = $window.FindName('TitleBlock')
$BtnModeToggle = $window.FindName('BtnModeToggle')
$AutoPanel = $window.FindName('AutoPanel'); $ManualPanel = $window.FindName('ManualPanel')
$LblGrpAuto = $window.FindName('LblGrpAuto'); $AppCombo = $window.FindName('AppCombo')
$LblGrpFolder = $window.FindName('LblGrpFolder'); $CompPanel = $window.FindName('CompPanel')
$LblGrpManual = $window.FindName('LblGrpManual'); $SourcePathBox = $window.FindName('SourcePathBox')
$BrowseSourceBtn = $window.FindName('BrowseSourceBtn'); $LblGrpDest = $window.FindName('LblGrpDest')
$DestPathBox = $window.FindName('DestPathBox'); $BrowseDestBtn = $window.FindName('BrowseDestBtn')
$SymlinkBox = $window.FindName('SymlinkBox'); $ForceKillBox = $window.FindName('ForceKillBox')
$StartBtn = $window.FindName('StartBtn'); $RestartAdminBtn = $window.FindName('RestartAdminBtn')
$LogBtn = $window.FindName('LogBtn'); $LblProgress = $window.FindName('LblProgress')
$ProgBar = $window.FindName('ProgBar'); $StatusBlock = $window.FindName('StatusBlock')
$GrpLog = $window.FindName('GrpLog'); $LogBox = $window.FindName('LogBox')

function Write-LogUI {
    param([string]$Msg, [string]$Type="INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    $LogBox.Dispatcher.Invoke({
        $para = New-Object System.Windows.Documents.Paragraph
        $para.Margin = New-Object System.Windows.Thickness(0,0,0,4)
        $runTime = New-Object System.Windows.Documents.Run("[$ts] ")
        $runTime.Foreground = [System.Windows.Media.Brushes]::DimGray
        $para.Inlines.Add($runTime)
        # Set color based on type
        $brush = [System.Windows.Media.Brushes]::Black
        $prefix = ""
        switch ($Type) { # Set color and prefix
            "OK"    { $brush = [System.Windows.Media.Brushes]::Green;       $prefix = "OK: " }
            "WARN"  { $brush = [System.Windows.Media.Brushes]::DarkOrange;  $prefix = "WARN: " }
            "ERR"   { $brush = [System.Windows.Media.Brushes]::Red;         $prefix = "ERR: " }
            "SCAN"  { $brush = [System.Windows.Media.Brushes]::Blue;        $prefix = "SCAN: " }
            "TRACE" { $brush = [System.Windows.Media.Brushes]::Gray;        $prefix = "  > " }
            "INFO"  { $brush = [System.Windows.Media.Brushes]::Black;       $prefix = "" }
        }
        if ($prefix) { # Add prefix run
            $runP = New-Object System.Windows.Documents.Run($prefix)
            $runP.FontWeight = "Bold"; $runP.Foreground = $brush; $para.Inlines.Add($runP)
        }
        $runMsg = New-Object System.Windows.Documents.Run($Msg)
        $runMsg.Foreground = $brush
        if ($Type -eq "TRACE") { $runMsg.FontStyle = "Italic" }
        $para.Inlines.Add($runMsg)
        $LogBox.Document.Blocks.Add($para)
        $LogBox.ScrollToEnd()
    })
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-UILanguage {
    param([string]$Code)
    if ($Code -eq 'VI') { $Lang = $LangData.VI } else { $Lang = $LangData.EN }
    
    $window.Title = $Lang.Title; $TitleBlock.Text = $Lang.Title
    if ($BtnModeToggle.IsChecked) { $BtnModeToggle.Content = $Lang.ModeAuto } else { $BtnModeToggle.Content = $Lang.ModeManual }
    $BtnModeToggle.ToolTip = $Lang.TipMode

    $LblGrpAuto.Text = $Lang.GrpAuto; $LblGrpFolder.Text = $Lang.GrpFolder
    $LblGrpManual.Text = $Lang.GrpManual; $LblGrpDest.Text = $Lang.GrpDest
    $SymlinkBox.Content = $Lang.OptSymlink; $ForceKillBox.Content = $Lang.OptForce
    $SymlinkBox.ToolTip = $Lang.TipSymlink; $ForceKillBox.ToolTip = $Lang.TipForce
    $StartBtn.Content = $Lang.BtnStart; $RestartAdminBtn.Content = $Lang.BtnAdmin
    $LogBtn.Content = $Lang.BtnLog; $LblProgress.Text = $Lang.LblProgress
    $StatusBlock.Text = $Lang.LblStatus; $GrpLog.Header = $Lang.LogHeader
    $BtnLangToggle.Content = $Lang.LangName

    $Script:CurrentLang = $Lang
    
    if ($isAdmin) { Write-LogUI $Lang.MsgAdminReady "OK" } else { Write-LogUI $Lang.MsgUserWarn "WARN" }

    # update folder panel contents without changing selection
    foreach ($k in $CompPanel.Children) {
        $c = $k.Tag
        $suffix = $Lang.MsgMissing; $color = "Gray"
        # Determine status based on current state or quick check
        if (Test-Path $c.Path) {
            $isLinked = (Get-Item $c.Path -Force).LinkType -match "SymbolicLink|Junction"
            if ($isLinked) { $suffix = $Lang.MsgLinked; $color = "DarkOrange" }
            else { $suffix = $Lang.MsgFound; $color = "DarkBlue" }
        } else { $color = "Gray" }
        $k.Content = "$($c.Label) [$($suffix)]"
        $k.Foreground = $color
    }
}

$BtnLangToggle.Add_Click({
    if ($BtnLangToggle.IsChecked) { Update-UILanguage "VI" } else { Update-UILanguage "EN" }
})

$BtnModeToggle.Add_Click({
    if ($BtnModeToggle.IsChecked) {
        $AutoPanel.Visibility = "Visible"; $ManualPanel.Visibility = "Collapsed"
        $BtnModeToggle.Content = $CurrentLang.ModeAuto
    } else {
        $AutoPanel.Visibility = "Collapsed"; $ManualPanel.Visibility = "Visible"
        $BtnModeToggle.Content = $CurrentLang.ModeManual
    }
})

function Get-RoboMsg {
    param([int]$Code)
    if ($Code -eq 0) { return $CurrentLang.Robo0 }
    if ($Code -eq 1) { return $CurrentLang.Robo1 }
    if ($Code -eq 2) { return $CurrentLang.Robo2 }
    if ($Code -eq 3) { return $CurrentLang.Robo3 }
    if ($Code -eq 4) { return $CurrentLang.Robo4 }
    if ($Code -eq 8) { return $CurrentLang.Robo8 }
    if ($Code -ge 16) { return $CurrentLang.Robo16 }
    return ($CurrentLang.RoboUnknown -f $Code)
}

if ($isAdmin) { $RestartAdminBtn.Visibility = "Collapsed"; $SymlinkBox.IsEnabled = $true } else { $RestartAdminBtn.Visibility = "Visible"; $SymlinkBox.IsEnabled = $false }

$RestartAdminBtn.Add_Click({
    $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"
    $safePath = $PSCommandPath.Replace("'", "''") 
    $cmd = "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; & '$safePath'"
    
    $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -Command `"$cmd`""
    $proc.Verb = "RunAs"
    [System.Diagnostics.Process]::Start($proc); $window.Close(); Exit
})

$LogBtn.Add_Click({
    if (Test-Path $Config.LogFile) { Invoke-Item $Config.LogFile } else { [System.Windows.MessageBox]::Show("No log file found.", "Info") }
})

# CLEANUP ON CLOSE
$window.Add_Closing({
    if (Test-Path $Config.LogFile) { Remove-Item $Config.LogFile -ErrorAction SilentlyContinue }
})

foreach ($lib in $Library) { $AppCombo.Items.Add($lib.App) | Out-Null}
$AppCombo.SelectedIndex = -1

function UpdateCompPanel {
    if (-not $BtnModeToggle.IsChecked) { return }
    $CompPanel.Children.Clear()
    $lib = if ($AppCombo.SelectedIndex -ge 0) { $Library[$AppCombo.SelectedIndex] } else { $null }
    if ($null -eq $lib) { return }
    
    Write-LogUI ($CurrentLang.MsgScanTarget + $lib.App) "SCAN"
    foreach ($comp in $lib.Components) {
        $chk = New-Object System.Windows.Controls.CheckBox
        $realPath = $comp.Path
        $chk.Tag = [PSCustomObject]@{ Path = $comp.Path; Process = $comp.Process; Label = $comp.Label }
        $chk.Margin = "0,4,0,4"
        
        $chk.VerticalContentAlignment = "Center"; $chk.HorizontalAlignment = "Left"; $chk.Padding = "6,0,0,0"
        $chk.FocusVisualStyle = $null; $chk.ToolTip = $realPath
        
        Write-LogUI ($CurrentLang.MsgChecking + $realPath) "TRACE"
        if (Test-Path $realPath) {
            try {
                $itemObj = Get-Item -LiteralPath $realPath -ErrorAction SilentlyContinue
                $isLink = $false; $target = ""
                if ($null -ne $itemObj.LinkType -and $itemObj.LinkType -match "SymbolicLink|Junction") {
                    $isLink = $true
                    $rawTarget = $itemObj.Target
                    if ($rawTarget -is [System.Array]) { $target = $rawTarget | Select-Object -First 1 } else { $target = $rawTarget }
                } elseif ($itemObj.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { $isLink = $true; $target = "(Reparse Point)" }
                
                if ($isLink) {
                    $chk.Content = "$($comp.Label) [$($CurrentLang.MsgLinked)]"; $chk.Foreground = "DarkOrange"; $chk.IsEnabled = $false
                    Write-LogUI ("{0} -> {1}" -f $CurrentLang.MsgLinked, $target) "WARN"
                } else {
                    $chk.Content = "$($comp.Label) [$($CurrentLang.MsgFound)]"; $chk.Foreground = "DarkBlue"; $chk.IsEnabled = $true; $chk.IsChecked = $true
                    Write-LogUI ("{0}: {1}" -f $CurrentLang.MsgFound, $comp.Label) "OK"
                }
            } catch {
                $chk.Content = "$($comp.Label) [$($CurrentLang.MsgError)]"; $chk.Foreground = "Red"; $chk.IsEnabled = $false
                Write-LogUI ("{0}: {1}" -f $CurrentLang.MsgError, $comp.Label) "ERR"
            }
        } else {
            $chk.Content = "$($comp.Label) [$($CurrentLang.MsgMissing)]"; $chk.Foreground = "Gray"; $chk.IsEnabled = $false
            Write-LogUI ("{0}: {1}" -f $CurrentLang.MsgMissing, $comp.Label) "INFO"
        }
        $CompPanel.Children.Add($chk)
    }
    Write-LogUI ($CurrentLang.MsgScanDone) "SCAN"
}
$AppCombo.Add_SelectionChanged({ UpdateCompPanel })

function Get-FolderSelect($initPath) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($null -ne $initPath -and $initPath -ne "" -and (Test-Path $initPath)) { $dlg.SelectedPath = $initPath }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    return ""
}

$BrowseSourceBtn.Add_Click({ $p = [Win32.FolderPicker]::Show($SourcePathBox.Text); if($p){$SourcePathBox.Text=$p} })
$BrowseDestBtn.Add_Click({ $p = [Win32.FolderPicker]::Show($DestPathBox.Text); if($p){$DestPathBox.Text=$p} })

function Stop-TargetProcesses($procNames) {
    $uniqueProcs = $procNames | Select-Object -Unique
    $force = $ForceKillBox.IsChecked
    foreach ($name in $uniqueProcs) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $running = Get-Process -Name $name -ErrorAction SilentlyContinue
        if ($running) {
            $StatusBlock.Text = $CurrentLang.MsgProcessRun + $name
            Write-LogUI ($CurrentLang.MsgProcessRun + $name) "WARN"
            if ($force) {
                Stop-Process -Name $name -Force -ErrorAction SilentlyContinue
                Write-LogUI ($CurrentLang.MsgForceKill + $name) "OK"
            } else {
                $retry = $true
                while ($retry -and (Get-Process -Name $name -ErrorAction SilentlyContinue)) {
                    $msg = $CurrentLang.MsgManualKill -f $name
                    $result = [System.Windows.MessageBox]::Show($msg, "Action Required", [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Warning)
                    if ($result -eq "Cancel") { throw "User aborted." }
                    Start-Sleep -Seconds 1
                }
                Write-LogUI ($CurrentLang.MsgKillDone + $name) "OK"
            }
        }
    }
}

function Format-Bytes { param([long]$Bytes) if ($Bytes -gt 1GB) { "{0:N2} GB" -f ($Bytes/1GB) } elseif ($Bytes -gt 1MB) { "{0:N2} MB" -f ($Bytes/1MB) } else { "{0:N0} B" -f $Bytes } }

function Invoke-AppDataMove {
    $StartBtn.IsEnabled = $false; $StatusBlock.Text = $CurrentLang.MsgWaitUser; $ProgBar.Value = 0
    Write-LogUI "========================================" "INFO"
    Write-LogUI $CurrentLang.MsgNewRun "INFO"
    
    if ($SymlinkBox.IsChecked) {
        $warnRes = [System.Windows.MessageBox]::Show($CurrentLang.MsgSymlinkWarn, "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($warnRes -eq "No") { 
            Write-LogUI "Aborted by user." "WARN"; 
            $StatusBlock.Text = $CurrentLang.MsgAborted; 
            $StartBtn.IsEnabled = $true; return 
        }
    }

    $destRoot = $DestPathBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($destRoot)) { 
        Write-LogUI $CurrentLang.MsgDestEmpty "ERR"; 
        $StatusBlock.Text = $CurrentLang.MsgDestEmpty; 
        $StartBtn.IsEnabled = $true; return 
    }

    if (-not (Test-Path $destRoot)) {
        try { New-Item -Path $destRoot -ItemType Directory -Force | Out-Null; Write-LogUI ("Created: " + $destRoot) "OK" }
        catch { 
            Write-LogUI "Error: Cannot create destination folder." "ERR"; 
            $StatusBlock.Text = "Error: Create Destination Failed"; 
            $StartBtn.IsEnabled = $true; return 
        }
    }

    $batch = @(); $processList = @()
    if ($BtnModeToggle.IsChecked) { # AUTO
        foreach ($chk in $CompPanel.Children) {
            if ($chk.IsChecked) {
                $tag = $chk.Tag
                $destPath = Join-Path $destRoot (Split-Path $tag.Path -Leaf)
                $batch += @{ Source = $tag.Path; Dest = $destPath; Process = $tag.Process }
                if ($tag.Process) { $processList += $tag.Process }
            }
        }
    } else { # MANUAL
        $src = $SourcePathBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($src) -or -not (Test-Path $src)) { 
            Write-LogUI $CurrentLang.MsgSourceInvalid "ERR"; 
            $StatusBlock.Text = $CurrentLang.MsgSourceInvalid; 
            $StartBtn.IsEnabled = $true; return 
        }
        $batch += @{ Source = $src; Dest = Join-Path $destRoot (Split-Path $src -Leaf); Process = "" }
    }

    if ($batch.Count -eq 0) { 
        Write-LogUI $CurrentLang.MsgNoSelect "ERR"; 
        $StatusBlock.Text = $CurrentLang.MsgNoSelect; 
        $StartBtn.IsEnabled = $true; return 
    }

    $StatusBlock.Text = $CurrentLang.MsgCalcSize
    [long]$totalSize = 0; [int]$countFiles = 0; [int]$countDirs = 0
    foreach ($b in $batch) {
        $fileStats = Get-ChildItem $b.Source -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
        $dirStats = Get-ChildItem $b.Source -Recurse -Force -Directory -ErrorAction SilentlyContinue | Measure-Object
        $totalSize += ($fileStats.Sum); $countFiles += ($fileStats.Count); $countDirs += ($dirStats.Count)
    }
    Write-LogUI ($CurrentLang.MsgStatLog -f $countDirs, $countFiles) "INFO"
    
    try {
        $drive = Get-PSDrive | Where-Object { $_.Root -eq ([System.IO.Path]::GetPathRoot($destRoot)) }
        if ($null -ne $drive) {
            $freeSpace = $drive.Free
            Write-LogUI ("Size: {0} | Free: {1}" -f (Format-Bytes $totalSize), (Format-Bytes $freeSpace)) "INFO"
            if ($freeSpace -lt $totalSize) { 
                Write-LogUI $CurrentLang.MsgDiskFull "ERR"; 
                $StatusBlock.Text = $CurrentLang.MsgDiskFull;
                $StartBtn.IsEnabled = $true; return 
            }
        }
    } catch { Write-LogUI "Skip size check." "WARN" }

    $ProgBar.Maximum = $batch.Count
    try {
        Stop-TargetProcesses $processList
        $backupList = @()
        $i = 0
        foreach ($item in $batch) {
            $i++
            $ProgBar.Value = $i
            $folderName = Split-Path $item.Source -Leaf
            $StatusBlock.Text = ("{0} ({1}/{2}): {3}" -f $CurrentLang.LblProgress, $i, $batch.Count, $folderName)
            Write-LogUI ($CurrentLang.MsgTask -f $i, $folderName) "INFO"
            
            if (-not (Test-Path $item.Dest)) { New-Item $item.Dest -ItemType Directory -Force | Out-Null }
            
            $cleanSource = $item.Source.TrimEnd('\')
            $cleanDest = $item.Dest.TrimEnd('\')
            $roboArgsString = '"{0}" "{1}" /E /NP /R:1 /W:1 /MT:8' -f $cleanSource, $cleanDest
            if ($isAdmin) { $roboArgsString += ' /ZB /COPYALL' } else { $roboArgsString += ' /COPY:DAT' }
            $roboArgsString += (' /LOG+:"{0}"' -f $Config.LogFile)
            
            $p = Start-Process -FilePath "robocopy.exe" -ArgumentList $roboArgsString -Wait -PassThru -WindowStyle Hidden
            $code = $p.ExitCode
            
            $roboMsg = Get-RoboMsg $code
            if ($code -le 7) {
                Write-LogUI $roboMsg "OK"
                if ($SymlinkBox.IsChecked -and $isAdmin) {
                    $ts = Get-Date -f 'yyyyMMdd_HHmmss'
                    $backupName = "{0}.OLD_{1}" -f $item.Source, $ts
                    try {
                        Rename-Item -LiteralPath $item.Source -NewName $backupName -ErrorAction Stop
                        New-Item -Path $item.Source -ItemType SymbolicLink -Target $item.Dest -ErrorAction Stop | Out-Null
                        if (Test-Path $item.Source) { Write-LogUI $CurrentLang.MsgLinkOK "OK"; $backupList += $backupName } else { throw "Link check failed" }
                    } catch {
                        Write-LogUI ($CurrentLang.MsgLinkFail + $_.Exception.Message) "ERR"
                        if (Test-Path $item.Source) { Remove-Item $item.Source -Force -ErrorAction SilentlyContinue }
                        if (Test-Path $backupName) { Rename-Item $backupName -NewName $item.Source -ErrorAction Stop; Write-LogUI $CurrentLang.MsgRollbackDone "OK" }
                    }
                }
            } else { Write-LogUI $roboMsg "ERR" }
        }
        
        $StatusBlock.Text = $CurrentLang.MsgAllDone
        Write-LogUI $CurrentLang.MsgAllDone "OK"
        # SUMMARY LOG
        Write-LogUI ($CurrentLang.MsgSummaryLog -f $countDirs, $countFiles) "INFO"
        
        if ($backupList.Count -gt 0) {
            $askMsg = $CurrentLang.MsgAskCleanup -f $backupList.Count
            $res = [System.Windows.MessageBox]::Show($askMsg, "Cleanup", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($res -eq "Yes") {
                foreach ($b in $backupList) { 
                    Remove-Item -LiteralPath $b -Recurse -Force -ErrorAction SilentlyContinue 
                    Write-LogUI ($CurrentLang.MsgCleanup + (Split-Path $b -Leaf)) "OK"
                }
            }
        }
    } catch {
        Write-LogUI ("Fatal Error: " + $_.Exception.Message) "ERR"
        $StatusBlock.Text = "Fatal Error"
    } finally { $StartBtn.IsEnabled = $true }
}
$StartBtn.Add_Click({ Invoke-AppDataMove })
#endregion

# TRIGGER DEFAULT LANG
Update-UILanguage "EN"

$window.ShowDialog() | Out-Null