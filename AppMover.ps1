<#
.SYNOPSIS
Robocopy UI Tool for Moving App Data with Symbolic Links
.DESCRIPTION
This PowerShell script provides a graphical user interface (GUI) for moving application data folders to a new location using Robocopy. It supports automatic scanning of predefined application data paths, manual folder selection, and the option to create symbolic links from the original location to the new location. The script ensures that required dependencies are met, checks for administrative privileges, and provides real-time progress updates during the move operation.
.NOTES
#>

# --- 0. CONSOLE INIT ---
Clear-Host
Write-Host "===========================" -ForegroundColor Cyan
Write-Host "=      APP DATA MOVER     =" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan
Write-Host ""

# --- 1. FORCE STA MODE ---
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "[INIT] Switching to STA Mode..." -ForegroundColor DarkGray
    $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"
    $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $proc.Verb = "RunAs"
    [System.Diagnostics.Process]::Start($proc)
    Exit
}

# --- 2. LOAD LIBRARIES ---
Write-Host "[INIT] Loading System Libraries... " -NoNewline -ForegroundColor Gray
try {
    # DPI Awareness Fix
    $code = @"
    using System;
    using System.Runtime.InteropServices;
    public class DPIFixer {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@
    Add-Type -TypeDefinition $code -PassThru | Out-Null
    [DPIFixer]::SetProcessDPIAware() | Out-Null
    
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Xml, WindowsBase -ErrorAction Stop
    
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match "PresentationCore" })) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("PresentationCore")
    }
    Write-Host "[OK]" -ForegroundColor Green
} catch {
    Write-Host "[FAIL]" -ForegroundColor Red
    # Native Popup for .NET Check
    $wshell = New-Object -ComObject WScript.Shell
    $answer = $wshell.Popup("This tool requires .NET Framework (WPF) to run.`n`nDo you want to download it now?", 0, "Missing Dependency", 4 + 16) # Yes/No + Warning Icon
    if ($answer -eq 6) { Start-Process "https://go.microsoft.com/fwlink/?linkid=2088631" } # 6 = Yes
    Exit
}
# --- ADMIN CHECK ---
Write-Host "[CHECK] Checking Admin Rights... " -NoNewline -ForegroundColor Gray
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if ($isAdmin) { Write-Host "[CHECK] Admin Rights: GRANTED" -ForegroundColor Green } 
else { Write-Host "[CHECK] Admin Rights: RESTRICTED" -ForegroundColor Yellow }

# --- 3. DATA LIBRARY ---
$AppLibrary = @{
    "Zalo (Full Suite)" = @(
        @{ Label="[Data]    Zalo PC (Images/Files)";            Path="%LocalAppData%\ZaloPC";                     Process="Zalo" },
        @{ Label="[Config]  Zalo AppData (Settings)";           Path="%AppData%\ZaloData";                        Process="Zalo" },
        @{ Label="[Docs]    Received Files";                    Path="%UserProfile%\Documents\Zalo Received Files";Process="Zalo" },
        @{ Label="[Exec]    Zalo Program (Local)";              Path="%LocalAppData%\Programs\Zalo";              Process="Zalo" },
        @{ Label="[Exec]    Zalo Program (x86)";                Path="${env:ProgramFiles(x86)}\Zalo";                Process="Zalo" },
        @{ Label="[Update]  Zalo Updater (Roaming)";            Path="%AppData%\zalo-updater";                    Process="Zalo" },
        @{ Label="[Update]  Zalo Updater (Local)";              Path="%LocalAppData%\zalo-updater";               Process="Zalo" }
    )
    "Telegram" = @(
        @{ Label="Telegram Desktop Data"; Path="%AppData%\Telegram Desktop"; Process="Telegram" },
        @{ Label="Telegram Downloads";    Path="%UserProfile%\Downloads\Telegram Desktop"; Process="Telegram" }
    )
    "Browser Data" = @(
        @{ Label="Chrome User Data"; Path="%LocalAppData%\Google\Chrome\User Data"; Process="chrome" },
        @{ Label="Edge User Data";   Path="%LocalAppData%\Microsoft\Edge\User Data"; Process="msedge" },
        @{ Label="CocCoc User Data"; Path="%LocalAppData%\CocCoc\Browser\User Data"; Process="browser" }
    )
    "Viber" = @(
        @{ Label="Viber PC Data"; Path="%LocalAppData%\ViberPC"; Process="Viber" },
        @{ Label="Viber Downloads"; Path="%UserProfile%\Documents\ViberDownloads"; Process="Viber" }
    )
}
# --- SCRIPT GLOBALS ---
$script:DestAutoState = ""
$script:DestManualState = ""
$script:LastLogFile = "$env:TEMP\RoboCopy_BatchLog.txt"

# --- GLOBAL SYNC HASH ---
$SyncHash = [hashtable]::Synchronized(@{
    ProgressMsg = "Ready"
    Results = New-Object System.Collections.ArrayList
    Stats = @{ Files=0; Dirs=0 }
    BatchList = @()
    Link = $false
    LogPath = $script:LastLogFile
    IsAdmin = $isAdmin
})

# --- XAML UI ---
$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="MainWindow"
    Title="App Data Mover"
    Width="1280" Height="800"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize"
    UseLayoutRounding="True"
    SnapsToDevicePixels="True"
    TextOptions.TextFormattingMode="Display"
    TextOptions.TextRenderingMode="ClearType">
    <Window.Resources>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderBrush" Value="#888"/>
        </Style>
        <Style TargetType="TextBlock" x:Key="StepHeader">
            <Setter Property="Foreground" Value="#0066CC"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Margin" Value="0,0,0,5"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <GroupBox Grid.Row="0" Header="STEP 1: SELECT MODE">
            <StackPanel Orientation="Horizontal">
                <RadioButton x:Name="ModeScanRadio" Content="Auto Scan (Batch Move)" GroupName="SourceMode" IsChecked="True" Margin="0,0,20,0" Cursor="Hand" FontSize="13"/>
                <RadioButton x:Name="ModeManualRadio" Content="Manual Folder Browse" GroupName="SourceMode" Cursor="Hand" FontSize="13"/>
            </StackPanel>
        </GroupBox>

        <GroupBox Grid.Row="1" Header="STEP 2: SELECT SOURCE DATA">
            <Grid>
                <Grid x:Name="ScanPanel">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="A. Select Application Suite:" Style="{StaticResource StepHeader}"/>
                    <ComboBox Grid.Row="1" x:Name="AppSelectorCombo" Width="400" HorizontalAlignment="Left" IsEditable="False" Height="25" Margin="0,0,0,15"/>
                    <Grid Grid.Row="2">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <DockPanel Grid.Row="0" LastChildFill="False" Margin="0,0,0,5">
                            <TextBlock Text="B. Check Components (Orange = Already Linked):" Style="{StaticResource StepHeader}" VerticalAlignment="Center"/>
                            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                                <Button x:Name="BtnSelectAll" Content="All" Width="40" Margin="0,0,5,0" FontSize="10"/>
                                <Button x:Name="BtnSelectNone" Content="None" Width="40" FontSize="10"/>
                            </StackPanel>
                        </DockPanel>
                        <Border Grid.Row="1" BorderBrush="Gray" BorderThickness="1" Padding="5" CornerRadius="3" Background="#FAFAFA" Margin="0,0,0,2">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="200">
                                <StackPanel x:Name="PathResultPanel"/>
                            </ScrollViewer>
                        </Border>
                    </Grid>
                </Grid>
                <Grid x:Name="ManualPanel" Visibility="Collapsed">
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="Browse for a specific folder:" Style="{StaticResource StepHeader}"/>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Margin="0,0,5,0" Height="25" VerticalContentAlignment="Center" x:Name="SourcePathTextBox" IsReadOnly="True"/>
                        <Button Grid.Column="1" Width="80" Content="Browse..." x:Name="BrowseSourceButton"/>
                    </Grid>
                </Grid>
            </Grid>
        </GroupBox>

        <GroupBox Grid.Row="2" Header="STEP 3: SELECT DESTINATION ROOT">
            <StackPanel>
                <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                    <TextBox Grid.Column="0" Margin="0,0,5,0" Height="25" VerticalContentAlignment="Center" x:Name="DestPathTextBox" IsReadOnly="True"/>
                    <Button Grid.Column="1" Width="80" Content="Browse..." x:Name="BrowseDestButton"/>
                </Grid>
                <TextBlock Text="* Note: App folders will be auto-created inside here." FontSize="11" Foreground="Gray" Margin="2,2,0,0" FontStyle="Italic"/>
            </StackPanel>
        </GroupBox>

        <Grid Grid.Row="3" Margin="0,5,0,5">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <CheckBox Grid.Column="0" x:Name="SymlinkCheckBox" Content="Create Symbolic Links (Rename Source -> Link to New Location)" VerticalContentAlignment="Center" FontWeight="Bold"/>
            <Button Grid.Column="1" x:Name="RestartAdminButton" Width="200" Height="30" Background="#FFCCCC" Visibility="Collapsed" Cursor="Hand" BorderBrush="Red">
                <TextBlock Text="[ADMIN] Restart as Admin" FontWeight="Bold" Foreground="Red" HorizontalAlignment="Center"/>
            </Button>
        </Grid>

        <Grid Grid.Row="4" Margin="0,10,0,5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" x:Name="StatusLabel" Text="Ready." FontWeight="Bold" Foreground="Black" TextWrapping="Wrap" VerticalAlignment="Center"/>
            <Button Grid.Column="1" x:Name="BtnOpenLog" Content="VIEW LOG" Width="80" Height="25" Visibility="Collapsed" Margin="10,0,0,0" FontSize="11"/>
        </Grid>

        <ProgressBar Grid.Row="5" Height="10" x:Name="MyProgressBar" Visibility="Hidden" IsIndeterminate="False" Maximum="100"/>

        <Button Grid.Row="7" Margin="0,15,0,0" Height="50" x:Name="CopyButton" Content="START MOVE PROCESS" FontWeight="Bold" FontSize="16"/>
    </Grid>
</Window>
"@

# --- 4. LOGIC ---

try {
    Write-Host "[UI] Parsing XAML..." -ForegroundColor Gray
    $reader = [System.XML.XmlNodeReader]::new([xml]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    Write-Host "[UI] Loaded Successfully." -ForegroundColor Green
} catch { [System.Windows.Forms.MessageBox]::Show("XAML Error: $($_.Exception.Message)"); Exit }

# Bindings
$ModeScanRadio=$window.FindName('ModeScanRadio'); $ModeManualRadio=$window.FindName('ModeManualRadio')
$ScanPanel=$window.FindName('ScanPanel'); $ManualPanel=$window.FindName('ManualPanel')
$AppSelectorCombo=$window.FindName('AppSelectorCombo'); $PathResultPanel=$window.FindName('PathResultPanel')
$SourcePathTextBox=$window.FindName('SourcePathTextBox'); $BrowseSourceButton=$window.FindName('BrowseSourceButton')
$DestPathTextBox=$window.FindName('DestPathTextBox'); $BrowseDestButton=$window.FindName('BrowseDestButton')
$SymlinkCheckBox=$window.FindName('SymlinkCheckBox'); $StatusLabel=$window.FindName('StatusLabel')
$MyProgressBar=$window.FindName('MyProgressBar'); $CopyButton=$window.FindName('CopyButton')
$RestartAdminButton=$window.FindName('RestartAdminButton'); $BtnOpenLog=$window.FindName('BtnOpenLog')
$BtnSelectAll=$window.FindName('BtnSelectAll'); $BtnSelectNone=$window.FindName('BtnSelectNone')

# Init
if ($isAdmin) {
    $window.Title = "App Data Mover (ADMIN)"
    $RestartAdminButton.Visibility = "Collapsed"
    $StatusLabel.Text = "Ready (Admin)."
} else {
    $window.Title = "App Data Mover (USER)"
    $RestartAdminButton.Visibility = "Visible"
    $SymlinkCheckBox.IsEnabled = $false
    $SymlinkCheckBox.Content = "Create Symbolic Links (Disabled: Requires Admin)"
    $StatusLabel.Text = "Standard Mode. Restart as Admin for full features."
}
# Show Window
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
# Populate App Selector
foreach ($key in $AppLibrary.Keys) { $AppSelectorCombo.Items.Add($key) | Out-Null }
# --- TIMER TICK EVENT ---
function Select-Folder {
    param([string]$Initial)
    try { $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; if (-not [string]::IsNullOrWhiteSpace($Initial) -and (Test-Path $Initial)) { $dlg.SelectedPath = $Initial }; if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath } } catch {} return $null
}
# --- UPDATE FOLDER LIST FUNCTION ---
function Update-FolderList {
    try { # Clear Existing
        $selectedApp = $AppSelectorCombo.SelectedItem
        if (-not $selectedApp) { return }
        $PathResultPanel.Children.Clear()
        $pathsToCheck = $AppLibrary[$selectedApp]

        Write-Host "[SCAN] Scanning folders for: $selectedApp" -ForegroundColor Yellow
        #  Check Each Path
        foreach ($item in $pathsToCheck) {
            $realPath = [Environment]::ExpandEnvironmentVariables($item.Path)
            $chk = New-Object System.Windows.Controls.CheckBox
            $chk.Margin = "0,2,0,2"
            $chk.Cursor = "Hand"

            Write-Host "   - Checking: $realPath" -NoNewline
            # Check Existence
            if (Test-Path $realPath -PathType Container) {
                Write-Host " [FOUND]" -ForegroundColor Green
                $itemObj = Get-Item -LiteralPath $realPath -Force
                $isLink = $false; $target = ""
                if ($itemObj.LinkType -match "SymbolicLink|Junction") { $isLink = $true; $target = $itemObj.Target }
                elseif ($itemObj.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { $isLink = $true; $target = "(Reparse Point)" }
                # Update Checkbox
                if ($isLink) { # Already a Link
                    Write-Host " [ALREADY LINKED]" -ForegroundColor DarkYellow
                    $chk.Content = "$($item.Label) `n  [ALREADY LINKED] -> $target"
                    $chk.Foreground = "DarkOrange"; $chk.FontWeight = "Bold"; $chk.IsChecked = $false; $chk.IsEnabled = $false
                } else { # Normal Found
                    $chk.Content = "$($item.Label) `n  [FOUND]: $realPath"
                    $chk.Foreground = "DarkBlue"; $chk.FontWeight = "Bold"; $chk.IsChecked = $true; $chk.IsEnabled = $true
                }
                $itemClone = $item.Clone(); $itemClone.Path = $realPath; $chk.Tag = $itemClone # Store Real Path in Tag
            } else { # Not Found
                Write-Host " [MISSING]" -ForegroundColor Gray
                $chk.Content = "$($item.Label) `n  [NOT FOUND]: $realPath"
                $chk.Foreground = "Gray"; $chk.FontWeight = "Normal"; $chk.IsChecked = $false; $chk.IsEnabled = $false
            }
            $PathResultPanel.Children.Add($chk) | Out-Null # Add to Panel
        }
    } catch { [System.Windows.Forms.MessageBox]::Show("Scan Error: " + $_.Exception.Message) }
}

$RestartAdminButton.Add_Click({ try { $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"; $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -File `"$PSCommandPath`""; $proc.Verb = "RunAs"; [System.Diagnostics.Process]::Start($proc); $window.Close() } catch {} })
$AppSelectorCombo.Add_SelectionChanged({ if ($ModeManualRadio.IsChecked) { return }; Update-FolderList })
$BtnSelectAll.Add_Click({ foreach ($c in $PathResultPanel.Children) { if ($c -is [System.Windows.Controls.CheckBox] -and $c.IsEnabled) { $c.IsChecked = $true } } })
$BtnSelectNone.Add_Click({ foreach ($c in $PathResultPanel.Children) { if ($c -is [System.Windows.Controls.CheckBox] -and $c.IsEnabled) { $c.IsChecked = $false } } })
$ModeScanRadio.Add_Checked({ $ScanPanel.Visibility = "Visible"; $ManualPanel.Visibility = "Collapsed"; if ($AppSelectorCombo.SelectedItem) { Update-FolderList } }); $ModeManualRadio.Add_Checked({ $ScanPanel.Visibility = "Collapsed"; $ManualPanel.Visibility = "Visible" })
$BrowseSourceButton.Add_Click({ $path = Select-Folder $SourcePathTextBox.Text; if ($path) { $SourcePathTextBox.Text = $path } }); $BrowseDestButton.Add_Click({ $path = Select-Folder $DestPathTextBox.Text; if ($path) { $DestPathTextBox.Text = $path } })
$BtnOpenLog.Add_Click({ if (Test-Path $script:LastLogFile) { Invoke-Item $script:LastLogFile } })

# --- THREAD LOGIC ---
$BackgroundScript = {
    param($Sync)
    #--- PROCESSING LOGIC ---
    $list = $Sync.BatchList
    $log = $Sync.LogPath
    $link = $Sync.Link
    $amAdmin = $Sync.IsAdmin
    # Initialize Counters
    $i = 0
    $total = $list.Count
    $totFiles = 0; $totDirs = 0
    # Initialize Log
    Set-Content -Path $log -Value "--- Log Start ---" -Force
    #--- PROCESS EACH ITEM ---
    foreach ($item in $list) {
        $i++
        $s = $item.Source; $d = $item.Dest
        # Update Progress
        $Sync.ProgressMsg = "Processing $($i)/$($total): $(Split-Path $s -Leaf)"
        # Create Destination Folder
        if (-not (Test-Path $d)) { New-Item $d -Type Directory -Force | Out-Null }
        # Temporary Log File
        $tLog = "$log" + "_$i.tmp"
        
        # --- CONSTRUCT ARGS ARRAY (EXPLICIT) ---
        $argsArr = @("$s", "$d")
        
        # Robocopy Parameters:
        # /E  : Copy subdirectories, including empty ones
        # /NP : No Progress - don't display percentage copied (keeps log clean)
        # /R:1: Retry 1 time on failed copies
        # /W:1: Wait 1 second between retries
        # /MT:8 : Multi-threaded copy (speed boost)
        $argsArr += @("/E", "/NP", "/R:1", "/W:1", "/MT:8")
        
        # Admin/User Specific Flags
        if ($amAdmin) { 
            # /ZB : Use Restartable mode; if access denied, use Backup mode (Admin only)
            # /COPYALL : Copy all file info (Data, Attributes, Time stamps, Security, Owner, Auditing)
            $argsArr += @("/ZB", "/COPYALL") 
        } else { 
            # /COPY:DAT : Copy Data, Attributes, Time stamps only (Safe for User)
            $argsArr += "/COPY:DAT" 
        }
        
        # Exclusions & Logging
        # /XD : Exclude Directories (System Volume Info, Recycle Bin)
        # /LOG: : Output status to log file (redirects output, prevents console pollution)
        # /TEE : Output to console window as well (KEPT in Thread for Debug Visibility)
        $argsArr += @("/XD", "System Volume Information", "`$RECYCLE.BIN", "/LOG:$tLog", "/TEE")
        
        # START PROCESS
        $p = Start-Process -FilePath "robocopy.exe" -ArgumentList $argsArr -WindowStyle Hidden -Wait -PassThru
        $code = $p.ExitCode
        $p = $null
        
        # Parse Stats
        try {
            if (Test-Path $tLog) { # Read Temp Log
                $txt = Get-Content $tLog
                Add-Content -Path $log -Value "`n--- FOLDER: $s ---"
                Add-Content -Path $log -Value $txt
                # Get Last File/Dir Lines
                $fL = $txt | Where-Object { $_ -match "^\s*Files :" } | Select-Object -Last 1
                $dL = $txt | Where-Object { $_ -match "^\s*Dirs :" } | Select-Object -Last 1
                # Extract Numbers
                if ($fL -match "Files :\s+\d+\s+(\d+)") { $totFiles += [int]$matches[1] }
                if ($dL -match "Dirs :\s+\d+\s+(\d+)")   { $totDirs += [int]$matches[1] }

                Remove-Item $tLog -Force -ErrorAction SilentlyContinue
            }
        } catch {}
        # Determine Status
        $status = "FAILED ($code)"
        if ($code -le 7) {
            $status = "Success"
            if ($link) {
                try { # Create Symlink
                    $t = Get-Date -Format "yyMMdd_HHmm"
                    Rename-Item "$s" -NewName "$s.OLD_$t" -ErrorAction Stop | Out-Null
                    New-Item "$s" -Type SymbolicLink -Target "$d" -ErrorAction Stop | Out-Null
                    $status = "MOVED & LINKED"
                } catch { $status = "COPIED (Link Failed)" }
            }
        }
        # Log Result
        $Sync.Results.Add("[RESULT] $(Split-Path $s -Leaf) : $status")
    }
    # Update Stats
    $Sync.Stats.Files = $totFiles
    $Sync.Stats.Dirs = $totDirs
}

# START BUTTON
$CopyButton.Add_Click({
    try {
        $BatchList = @(); $ProcessList = @(); $DestRoot = $DestPathTextBox.Text.Trim()
        # Build Batch List
        if ($ModeManualRadio.IsChecked) {
            $src = $SourcePathTextBox.Text.Trim()
            if (-not (Test-Path $src)) { 
                $StatusLabel.Text = "Warning: Source Invalid."; $StatusLabel.Foreground = "Red"; return 
            }
            # Link Check
            $itemObj = Get-Item -LiteralPath $src -Force
            if ($itemObj.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { 
                $StatusLabel.Text = "Error: Folder is already a Link!"; $StatusLabel.Foreground = "Red"; return 
            }
            $BatchList += @{ Source=$src; Dest=Join-Path $DestRoot (Split-Path $src -Leaf); Process="" }
        } else { # Auto Scan Mode
            foreach ($child in $PathResultPanel.Children) {
                if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked) {
                    $tag = $child.Tag
                    $BatchList += @{ Source=$tag.Path; Dest=Join-Path $DestRoot (Split-Path $tag.Path -Leaf); Process=$tag.Process }
                    if ($tag.Process -and $ProcessList -notcontains $tag.Process) { $ProcessList += $tag.Process }
                }
            }
        }
        # Validations
        if ($BatchList.Count -eq 0) { $StatusLabel.Text = "Warning: Select items to move."; $StatusLabel.Foreground = "Red"; return }
        if ([string]::IsNullOrWhiteSpace($DestRoot)) { $StatusLabel.Text = "Warning: Select Destination!"; $StatusLabel.Foreground = "Red"; return }
        # Check Destination
        foreach ($pName in $ProcessList) {
            if (Get-Process -Name $pName -ErrorAction SilentlyContinue) {
                $c = [System.Windows.Forms.MessageBox]::Show("App '$pName' is running.`nForce Close?", "Process Check", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($c -eq "Yes") { Stop-Process -Name $pName -Force -ErrorAction SilentlyContinue; Start-Sleep 1 } else { return }
            }
        }
        # Prepare SyncHash
        $SyncHash.Link = $SymlinkCheckBox.IsChecked
        $SyncHash.IsAdmin = $isAdmin
        $SyncHash.BatchList = $BatchList
        $SyncHash.Results.Clear()
        $SyncHash.Stats.Files = 0
        $SyncHash.Stats.Dirs = 0
        # START UI UPDATES
        $CopyButton.IsEnabled = $false; $MyProgressBar.Visibility = "Visible"; $MyProgressBar.IsIndeterminate = $true; $BtnOpenLog.Visibility = "Collapsed"
        $StatusLabel.Foreground = "Black"
        Write-Host "[JOB] Starting Background Thread..." -ForegroundColor Cyan
        
        # LAUNCH THREAD
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = "STA"
        $rs.ThreadOptions = "ReuseThread"
        $rs.Open()
        # Create PowerShell Instance
        $Global:psInstance = [PowerShell]::Create()
        $Global:psInstance.Runspace = $rs
        $Global:psInstance.AddScript($BackgroundScript).AddArgument($SyncHash)
        # Begin Async Invoke
        $Global:asyncHandle = $Global:psInstance.BeginInvoke()
        $timer.Start()
    } catch { [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message) }
})

# TIMER MONITOR
$timer.Add_Tick({
    if ($Global:asyncHandle -and $Global:asyncHandle.IsCompleted) {
        $timer.Stop(); $MyProgressBar.Visibility = "Hidden"
        # Finalize Thread
        try {
            $Global:psInstance.EndInvoke($Global:asyncHandle)
            if ($Global:psInstance.Streams.Error.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Thread Error: $($Global:psInstance.Streams.Error[0])", "Error")
            }
        } catch { [System.Windows.Forms.MessageBox]::Show("Thread Crash: $($_.Exception.Message)", "Critical") }
        # Summarize Results
        $sCount = 0; $fCount = 0
        foreach ($res in $SyncHash.Results) { if ($res -match "Success|MOVED") { $sCount++ } else { $fCount++ } }
        # Update Status
        $StatusLabel.Text = "Finished: $sCount Success, $fCount Failed (Moved $($SyncHash.Stats.Files) Files)."
        if ($fCount -gt 0) { $StatusLabel.Foreground = "Red" } else { $StatusLabel.Foreground = "Green" }
        $BtnOpenLog.Visibility = "Visible"; $CopyButton.IsEnabled = $true
        $Global:psInstance.Dispose()
        Write-Host "[JOB] Finished." -ForegroundColor Green
    }
    else { # Update Progress
        $StatusLabel.Text = $SyncHash.ProgressMsg
    }
})
#--- CLEANUP ON CLOSE ---
$window.Add_Closed({ try { if (Test-Path $script:LastLogFile) { Remove-Item $script:LastLogFile -Force -ErrorAction SilentlyContinue } } catch {} })
$window.ShowDialog() | Out-Null