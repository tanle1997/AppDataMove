<#
.SYNOPSIS
Robocopy UI Tool v35 - DEPENDENCY CHECKER.
Feature: Checks for required .NET Assemblies at startup.
Action: If missing, auto-opens the Microsoft Download page for .NET Framework 4.8.
#>

# --- 1. FORCE STA MODE ---
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"
    $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $proc.Verb = "RunAs"
    [System.Diagnostics.Process]::Start($proc)
    Exit
}

# --- 2. ROBUST LIBRARY LOADING & DEPENDENCY CHECK ---
Write-Host "Checking System Requirements..." -ForegroundColor Cyan
try {
    # Try to load required WPF/WinForms libraries
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, System.Xml, WindowsBase -ErrorAction Stop
    
    # Ensure PresentationCore is loaded (sometimes tricky on older systems)
    if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match "PresentationCore" })) {
        [void][System.Reflection.Assembly]::LoadWithPartialName("PresentationCore")
    }
    
    Write-Host "Libraries Loaded Successfully." -ForegroundColor Green
} catch {
    # --- IF .NET IS MISSING ---
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host " CRITICAL ERROR: MISSING .NET FRAMEWORK " -ForegroundColor Red
    Write-Host "==========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "This tool requires .NET Framework (WPF) to run." -ForegroundColor Yellow
    Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Action: Opening download page for .NET Framework 4.8..." -ForegroundColor Cyan
    
    # Open Browser to Download Link
    try { Start-Process "https://go.microsoft.com/fwlink/?linkid=2088631" } catch {} 
    
    Write-Host ""
    Read-Host "Please install .NET and try again. Press Enter to Exit..."
    Exit
}

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

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

# --- GLOBAL VARS ---
$script:DestAutoState = ""
$script:DestManualState = ""
$script:LastLogFile = "$env:TEMP\RoboCopy_BatchLog.txt"

# --- XAML UI ---
$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="MainWindow"
    Title="App Data Mover (V35 - Dependency Check)"
    Width="900" Height="750"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize">
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
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
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="A. Select Application Suite:" Style="{StaticResource StepHeader}"/>
                    <ComboBox Grid.Row="1" x:Name="AppSelectorCombo" Width="400" HorizontalAlignment="Left" IsEditable="False" Height="25" Margin="0,0,0,10"/>

                    <Grid Grid.Row="2">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <DockPanel Grid.Row="0" LastChildFill="False" Margin="0,0,0,5">
                            <TextBlock Text="B. Check Components (Orange = Already Linked):" Style="{StaticResource StepHeader}" VerticalAlignment="Center"/>
                            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                                <Button x:Name="BtnSelectAll" Content="All" Width="40" Margin="0,0,5,0" FontSize="10"/>
                                <Button x:Name="BtnSelectNone" Content="None" Width="40" FontSize="10"/>
                            </StackPanel>
                        </DockPanel>

                        <Border Grid.Row="1" BorderBrush="Gray" BorderThickness="1" Padding="5" CornerRadius="3" Background="#FAFAFA">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel x:Name="PathResultPanel"/>
                            </ScrollViewer>
                        </Border>
                    </Grid>
                </Grid>

                <Grid x:Name="ManualPanel" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="Browse for a specific folder:" Style="{StaticResource StepHeader}"/>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Margin="0,0,5,0" Height="25" VerticalContentAlignment="Center" x:Name="SourcePathTextBox" IsReadOnly="True"/>
                        <Button Grid.Column="1" Width="80" Content="Browse..." x:Name="BrowseSourceButton"/>
                    </Grid>
                </Grid>
            </Grid>
        </GroupBox>

        <GroupBox Grid.Row="2" Header="STEP 3: SELECT DESTINATION ROOT">
            <StackPanel>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Grid.Column="0" Margin="0,0,5,0" Height="25" VerticalContentAlignment="Center" x:Name="DestPathTextBox" IsReadOnly="True"/>
                    <Button Grid.Column="1" Width="80" Content="Browse..." x:Name="BrowseDestButton"/>
                </Grid>
                <TextBlock Text="* Note: App folders (e.g. \ZaloPC) will be created inside this folder." FontSize="11" Foreground="Gray" Margin="2,2,0,0" FontStyle="Italic"/>
            </StackPanel>
        </GroupBox>

        <Grid Grid.Row="3" Margin="0,5,0,5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <CheckBox Grid.Column="0" x:Name="SymlinkCheckBox" Content="Create Symbolic Links (Rename Source -> Link to New Location)" VerticalContentAlignment="Center" FontWeight="Bold"/>
            <Button Grid.Column="1" x:Name="RestartAdminButton" Width="200" Height="30" Background="#FFCCCC" Visibility="Collapsed" Cursor="Hand" BorderBrush="Red">
                <TextBlock Text="[ADMIN] Restart as Admin" FontWeight="Bold" Foreground="Red" HorizontalAlignment="Center"/>
            </Button>
        </Grid>

        <TextBlock Grid.Row="4" Margin="0,10,0,5" x:Name="StatusLabel" Text="Ready." FontWeight="Bold" Foreground="Black" TextWrapping="Wrap"/>

        <ProgressBar Grid.Row="5" Height="10" x:Name="MyProgressBar" Visibility="Hidden" IsIndeterminate="False" Maximum="100"/>

        <Button Grid.Row="6" Margin="0,15,0,0" Height="50" x:Name="CopyButton" Content="START MOVE PROCESS" FontWeight="Bold" FontSize="16"/>
    </Grid>
</Window>
"@

# --- 4. LOGIC ---

try {
    $reader = [System.XML.XmlNodeReader]::new([xml]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Bindings
    $ModeScanRadio      = $window.FindName('ModeScanRadio')
    $ModeManualRadio    = $window.FindName('ModeManualRadio')
    $ScanPanel          = $window.FindName('ScanPanel')
    $ManualPanel        = $window.FindName('ManualPanel')
    $AppSelectorCombo   = $window.FindName('AppSelectorCombo')
    $PathResultPanel    = $window.FindName('PathResultPanel')
    $SourcePathTextBox  = $window.FindName('SourcePathTextBox')
    $BrowseSourceButton = $window.FindName('BrowseSourceButton')
    $DestPathTextBox    = $window.FindName('DestPathTextBox')
    $BrowseDestButton   = $window.FindName('BrowseDestButton')
    $SymlinkCheckBox    = $window.FindName('SymlinkCheckBox')
    $StatusLabel        = $window.FindName('StatusLabel')
    $MyProgressBar      = $window.FindName('MyProgressBar')
    $CopyButton         = $window.FindName('CopyButton')
    $RestartAdminButton = $window.FindName('RestartAdminButton')
    $BtnSelectAll       = $window.FindName('BtnSelectAll')
    $BtnSelectNone      = $window.FindName('BtnSelectNone')

    # Init State
    if ($isAdmin) {
        $window.Title = "App Mover (ADMIN MODE)"
        $RestartAdminButton.Visibility = "Collapsed"
        $StatusLabel.Text = "Ready (Admin)."
    } else {
        $window.Title = "App Mover (USER MODE)"
        $RestartAdminButton.Visibility = "Visible"
        $SymlinkCheckBox.IsEnabled = $false
        $SymlinkCheckBox.Content = "Create Symbolic Links (Disabled: Requires Admin)"
        $StatusLabel.Text = "Standard Mode. Restart as Admin for full features."
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)

    foreach ($key in $AppLibrary.Keys) { $AppSelectorCombo.Items.Add($key) | Out-Null }

    function Select-Folder {
        param([string]$Initial)
        try {
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            if (-not [string]::IsNullOrWhiteSpace($Initial) -and (Test-Path $Initial)) { $dlg.SelectedPath = $Initial }
            if ($dlg.ShowDialog() -eq "OK") { return $dlg.SelectedPath }
        } catch {}
        return $null
    }

    function Refresh-ScanResults {
        try {
            $selectedApp = $AppSelectorCombo.SelectedItem
            if (-not $selectedApp) { return }
            $PathResultPanel.Children.Clear()
            $pathsToCheck = $AppLibrary[$selectedApp]
            
            Write-Host "Scanning for: $selectedApp" -ForegroundColor Yellow

            foreach ($item in $pathsToCheck) {
                $realPath = [Environment]::ExpandEnvironmentVariables($item.Path)
                $chk = New-Object System.Windows.Controls.CheckBox
                $chk.Margin = "0,2,0,2"
                $chk.Cursor = "Hand"
                
                if (Test-Path $realPath -PathType Container) {
                    # Check Link
                    $itemObj = Get-Item -LiteralPath $realPath -Force
                    $isLink = $false; $target = ""
                    if ($itemObj.LinkType -match "SymbolicLink|Junction") { $isLink = $true; $target = $itemObj.Target }
                    elseif ($itemObj.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { $isLink = $true; $target = "(Reparse Point)" }

                    if ($isLink) {
                        $chk.Content = "$($item.Label) `n  [ALREADY LINKED] -> $target"
                        $chk.Foreground = "DarkOrange"; $chk.FontWeight = "Bold"; $chk.IsChecked = $false; $chk.IsEnabled = $false
                    } else {
                        $chk.Content = "$($item.Label) `n  [FOUND]: $realPath"
                        $chk.Foreground = "DarkBlue"; $chk.FontWeight = "Bold"; $chk.IsChecked = $true; $chk.IsEnabled = $true
                    }
                    $itemClone = $item.Clone(); $itemClone.Path = $realPath; $chk.Tag = $itemClone
                } else {
                    $chk.Content = "$($item.Label) `n  [NOT FOUND]: $realPath"
                    $chk.Foreground = "Gray"; $chk.FontWeight = "Normal"; $chk.IsChecked = $false; $chk.IsEnabled = $false
                }
                $PathResultPanel.Children.Add($chk) | Out-Null
            }
        } catch { [System.Windows.MessageBox]::Show("Scan Error: " + $_.Exception.Message) }
    }

    # Events
    $RestartAdminButton.Add_Click({
        try {
            $currentScript = $PSCommandPath
            if (-not $currentScript) { [System.Windows.MessageBox]::Show("Save script first."); return }
            $proc = New-Object System.Diagnostics.ProcessStartInfo "powershell"
            $proc.Arguments = "-NoProfile -Sta -ExecutionPolicy Bypass -File `"$currentScript`""
            $proc.Verb = "RunAs"
            [System.Diagnostics.Process]::Start($proc)
            $window.Close()
        } catch {}
    })

    $AppSelectorCombo.Add_SelectionChanged({ if ($ModeManualRadio.IsChecked) { return }; Refresh-ScanResults })
    $BtnSelectAll.Add_Click({ foreach ($c in $PathResultPanel.Children) { if ($c -is [System.Windows.Controls.CheckBox] -and $c.IsEnabled -and $c.Foreground -ne "DarkOrange") { $c.IsChecked = $true } } })
    $BtnSelectNone.Add_Click({ foreach ($c in $PathResultPanel.Children) { if ($c -is [System.Windows.Controls.CheckBox] -and $c.IsEnabled) { $c.IsChecked = $false } } })

    $ModeScanRadio.Add_Checked({ 
        $script:DestManualState = $DestPathTextBox.Text; $ScanPanel.Visibility = "Visible"; $ManualPanel.Visibility = "Collapsed"; $DestPathTextBox.Text = $script:DestAutoState
        if ($AppSelectorCombo.SelectedItem) { Refresh-ScanResults } 
    })
    $ModeManualRadio.Add_Checked({ 
        $script:DestAutoState = $DestPathTextBox.Text; $ScanPanel.Visibility = "Collapsed"; $ManualPanel.Visibility = "Visible"; $DestPathTextBox.Text = $script:DestManualState 
    })
    
    $BrowseSourceButton.Add_Click({ $path = Select-Folder $SourcePathTextBox.Text; if ($path) { $SourcePathTextBox.Text = $path } })
    $BrowseDestButton.Add_Click({ $path = Select-Folder $DestPathTextBox.Text; if ($path) { $DestPathTextBox.Text = $path } })

    # START BUTTON
    $CopyButton.Add_Click({
        try {
            $BatchList = @()
            $ProcessList = @()
            $DestRoot = $DestPathTextBox.Text.Trim()

            if ($ModeManualRadio.IsChecked) {
                $src = $SourcePathTextBox.Text.Trim()
                if (-not (Test-Path $src)) { $StatusLabel.Text = "Error: Source Invalid"; $StatusLabel.Foreground="Red"; return }
                
                $itemObj = Get-Item -LiteralPath $src -Force
                if ($itemObj.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)) { [System.Windows.MessageBox]::Show("Folder is ALREADY a Link!", "Warning"); return }
                $folderName = Split-Path $src -Leaf; $targetDest = Join-Path $DestRoot $folderName
                $BatchList += @{ Source=$src; Dest=$targetDest; Process="" }
            } else {
                foreach ($child in $PathResultPanel.Children) {
                    if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked) {
                        $tag = $child.Tag; $src = $tag.Path; $folderName = Split-Path $src -Leaf; $targetDest = Join-Path $DestRoot $folderName
                        $BatchList += @{ Source=$src; Dest=$targetDest; Process=$tag.Process }
                        if ($tag.Process -and $ProcessList -notcontains $tag.Process) { $ProcessList += $tag.Process }
                    }
                }
            }

            if ($BatchList.Count -eq 0) { [System.Windows.MessageBox]::Show("Select folders to move."); return }
            if ([string]::IsNullOrWhiteSpace($DestRoot)) { $StatusLabel.Text = "Error: Empty Destination"; $StatusLabel.Foreground="Red"; return }
            
            foreach ($pName in $ProcessList) {
                if (Get-Process -Name $pName -ErrorAction SilentlyContinue) {
                    $c = [System.Windows.MessageBox]::Show("App '$pName' is running.`nForce Close?", "App Running", 4, 48)
                    if ($c -eq "Yes") { Stop-Process -Name $pName -Force -ErrorAction SilentlyContinue; Start-Sleep 1 } else { return }
                }
            }

            $DoSymlink = $SymlinkCheckBox.IsChecked
            $roboArgs = if ($isAdmin) { "/E /ZB /NP /R:1 /W:1 /COPYALL" } else { "/E /NP /R:1 /W:1 /COPY:DAT" }
            
            $CopyButton.IsEnabled = $false; $MyProgressBar.Visibility = "Visible"; $MyProgressBar.IsIndeterminate = $true; $StatusLabel.Text = "Moving..."; $StatusLabel.Foreground = "Blue"

            $sb = {
                param($list, $link, $log, $argsStr)
                $results = @(); $i = 0; $total = $list.Count
                foreach ($item in $list) {
                    $i++; $s = $item.Source; $d = $item.Dest
                    Write-Output "STATUS:Processing ${i}/${total}: $(Split-Path $s -Leaf)"
                    if (-not (Test-Path $d)) { New-Item $d -Type Directory -Force | Out-Null }
                    $cmd = "robocopy.exe `"$s`" `"$d`" $argsStr /XD `"System Volume Information`" `"`$RECYCLE.BIN`" /LOG+:`"$log`""
                    Invoke-Expression $cmd
                    $code = $LASTEXITCODE
                    $status = "Failed ($code)"
                    if ($code -le 7) {
                        $status = "Success"
                        if ($link) {
                            try {
                                $t = Get-Date -Format "yyMMdd_HHmm"
                                Rename-Item "$s" -NewName "$s.OLD_$t" -ErrorAction Stop
                                New-Item "$s" -Type SymbolicLink -Target "$d" -ErrorAction Stop
                                $status = "Moved & Linked"
                            } catch { $status = "Copied (Link Failed)" }
                        }
                    }
                    $results += "$s -> $status"
                }
                return $results
            }

            Set-Content -Path $script:LastLogFile -Value "--- Batch Log Started ---"
            $Global:currentJob = Start-Job -ScriptBlock $sb -ArgumentList $BatchList, $DoSymlink, $script:LastLogFile, $roboArgs
            $timer.Start()
        } catch { [System.Windows.MessageBox]::Show("Error: " + $_.Exception.Message) }
    })

    # TIMER
    $timer.Add_Tick({
        if ($Global:currentJob.State -eq 'Running') {
            $msgs = Receive-Job $Global:currentJob
            foreach ($m in $msgs) { if ($m -match "STATUS:(.*)") { $StatusLabel.Text = $matches[1] } }
        }
        elseif ($Global:currentJob.State -in 'Completed','Failed') {
            $timer.Stop(); $MyProgressBar.Visibility = "Hidden"
            $finalRes = Receive-Job $Global:currentJob -Keep; Remove-Job $Global:currentJob
            $summary = "Batch Complete!`n"; foreach ($line in $finalRes) { if ($line -is [string] -and $line -notmatch "STATUS:") { $summary += "`n$line" } }
            $StatusLabel.Text = "Operation Finished."; $StatusLabel.Foreground = "Green"
            [System.Windows.MessageBox]::Show($summary, "Report"); $CopyButton.IsEnabled = $true
        }
    })
    
    $window.ShowDialog() | Out-Null
} catch { Write-Error $_.Exception.Message; Read-Host "Exit..." }