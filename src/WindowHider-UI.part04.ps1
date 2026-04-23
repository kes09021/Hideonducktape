    Unregister-Hotkeys -Target $Target

    try {
        Register-Hotkeys -Target $Target -Settings $NewSettings
    }
    catch {
        if ($RollbackSettings) {
            try {
                Register-Hotkeys -Target $Target -Settings $RollbackSettings
            }
            catch {
            }
        }

        throw
    }
}

function Add-ManualTarget {
    $rawValue = $script:AddProgramTextBox.Text
    $normalized = Normalize-ExecutableName -Name $rawValue

    if ($null -eq $normalized) {
        [System.Windows.Forms.MessageBox]::Show('Enter a program name like chrome.exe or Code.exe.', 'Window Hider') | Out-Null
        return
    }

    $currentChecks = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @(Get-CheckedTargetNames)) {
        $currentChecks.Add($name) | Out-Null
    }

    if (-not $currentChecks.Contains($normalized)) {
        $currentChecks.Add($normalized) | Out-Null
    }

    $script:AddProgramTextBox.Clear()
    Refresh-TargetList -CheckedTargets $currentChecks
    Set-StatusMessage "Added $normalized to the target list. Click Save Settings to keep it."
}

function Save-UiChanges {
    $selectedTargets = @(Get-CheckedTargetNames)
    if ($selectedTargets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Choose at least one target program before saving.', 'Window Hider') | Out-Null
        return
    }

    try {
        $newSettings = [pscustomobject]@{
            ConfigPath   = $script:Settings.ConfigPath
            Hotkey       = Parse-Hotkey -Text $script:ToggleHotkeyTextBox.Text
            ExitHotkey   = if ([string]::IsNullOrWhiteSpace($script:ExitHotkeyTextBox.Text)) { $null } else { Parse-Hotkey -Text $script:ExitHotkeyTextBox.Text }
            Targets      = New-TargetSet -Names $selectedTargets
            WindowRules  = Copy-WindowRules -Rules @(Get-WindowRulesArray)
            LogToConsole = $script:Settings.LogToConsole
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Window Hider') | Out-Null
        return
    }

    try {
        Apply-Hotkeys -Target $script:HotkeySink -NewSettings $newSettings -RollbackSettings $script:Settings
        Save-Settings -Settings $newSettings
        $script:Settings = $newSettings
        Initialize-SettingsCollections -Settings $script:Settings
        Update-RuntimeTargets
        Update-UiStatus
        Set-StatusMessage 'Settings saved.'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Window Hider') | Out-Null
    }
}

function Show-MainForm {
    $script:MainForm.ShowInTaskbar = $true
    $script:MainForm.Show()
    $script:MainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $script:MainForm.Activate()
}

function New-TrayAppIcon {
    $bitmap = New-Object System.Drawing.Bitmap 16, 16
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $backgroundBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(32, 144, 140))
    $font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $stringFormat = New-Object System.Drawing.StringFormat
    $stringFormat.Alignment = [System.Drawing.StringAlignment]::Center
    $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Center

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)
        $graphics.FillEllipse($backgroundBrush, 0, 0, 15, 15)
        $graphics.DrawString('H', $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, 0, 16, 16)), $stringFormat)
        $iconHandle = $bitmap.GetHicon()
        try {
            return [System.Drawing.Icon]([System.Drawing.Icon]::FromHandle($iconHandle).Clone())
        }
        finally {
            [void][Win32WindowTools]::DestroyIcon($iconHandle)
        }
    }
    finally {
        $stringFormat.Dispose()
        $font.Dispose()
        $backgroundBrush.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $scriptRoot -ChildPath 'window-hider.config.json'
}

$script:Settings = Read-Settings -Path $ConfigPath
Initialize-SettingsCollections -Settings $script:Settings
$script:HiddenWindows = [System.Collections.Generic.List[object]]::new()
$script:RuntimeTargets = New-TargetSet -Names ($script:Settings.Targets | Sort-Object)
$script:LatestInventory = @()
$script:TargetEntries = @()
$script:ToggleHotkeyId = 1
$script:ExitHotkeyId = 2
$script:HotkeysRegistered = $false
$script:TrayHintShown = $false
$script:IsRefreshingTargetList = $false
$script:TrayIcon = $null
$script:IsShuttingDown = $false

if ($ValidateOnly) {
    $validateExitText = if ($script:Settings.ExitHotkey) { $script:Settings.ExitHotkey.Display } else { 'not set' }
    Write-Host 'UI script OK'
    Write-Host "Config file    : $($script:Settings.ConfigPath)"
    Write-Host "Toggle hotkey  : $($script:Settings.Hotkey.Display)"
    Write-Host ("Exit hotkey    : {0}" -f $validateExitText)
    Write-Host "Saved targets  : $((@($script:Settings.Targets | Sort-Object) -join ', '))"
    Write-Host "Window rules   : $(@(Get-WindowRulesArray).Count)"
    exit 0
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:HotkeySink = New-Object HotkeyMessageWindow
$script:MainForm = New-Object HotkeyAwareForm
$script:MainForm.Text = 'Window Hider'
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.Size = New-Object System.Drawing.Size(900, 640)
$script:MainForm.MinimumSize = New-Object System.Drawing.Size(900, 640)
$script:MainForm.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$targetsGroup = New-Object System.Windows.Forms.GroupBox
$targetsGroup.Text = 'Target Programs'
$targetsGroup.Location = New-Object System.Drawing.Point(12, 12)
$targetsGroup.Size = New-Object System.Drawing.Size(320, 560)
$targetsGroup.Anchor = 'Top,Bottom,Left'
$script:MainForm.Controls.Add($targetsGroup)

$targetsHelpLabel = New-Object System.Windows.Forms.Label
$targetsHelpLabel.Text = 'Pick the apps you want to hide with the hotkey.'
$targetsHelpLabel.Location = New-Object System.Drawing.Point(15, 28)
$targetsHelpLabel.Size = New-Object System.Drawing.Size(285, 20)
$targetsGroup.Controls.Add($targetsHelpLabel)

$script:TargetListBox = New-Object System.Windows.Forms.CheckedListBox
$script:TargetListBox.Location = New-Object System.Drawing.Point(15, 55)
$script:TargetListBox.Size = New-Object System.Drawing.Size(287, 327)
$script:TargetListBox.CheckOnClick = $true
$script:TargetListBox.Anchor = 'Top,Bottom,Left,Right'
$targetsGroup.Controls.Add($script:TargetListBox)

$addProgramLabel = New-Object System.Windows.Forms.Label
$addProgramLabel.Text = 'Add exe manually'
$addProgramLabel.Location = New-Object System.Drawing.Point(15, 400)
$addProgramLabel.Size = New-Object System.Drawing.Size(140, 20)
$addProgramLabel.Anchor = 'Bottom,Left'
$targetsGroup.Controls.Add($addProgramLabel)

$script:AddProgramTextBox = New-Object System.Windows.Forms.TextBox
$script:AddProgramTextBox.Location = New-Object System.Drawing.Point(15, 424)
$script:AddProgramTextBox.Size = New-Object System.Drawing.Size(180, 25)
$script:AddProgramTextBox.Anchor = 'Bottom,Left'
$targetsGroup.Controls.Add($script:AddProgramTextBox)

$addProgramButton = New-Object System.Windows.Forms.Button
$addProgramButton.Text = 'Add'
$addProgramButton.Location = New-Object System.Drawing.Point(203, 422)
$addProgramButton.Size = New-Object System.Drawing.Size(99, 30)
$addProgramButton.Anchor = 'Bottom,Right'
$targetsGroup.Controls.Add($addProgramButton)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = 'Refresh List'
$refreshButton.Location = New-Object System.Drawing.Point(15, 468)
$refreshButton.Size = New-Object System.Drawing.Size(135, 34)
$refreshButton.Anchor = 'Bottom,Left'
$targetsGroup.Controls.Add($refreshButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = 'Clear Checks'
$clearButton.Location = New-Object System.Drawing.Point(167, 468)
$clearButton.Size = New-Object System.Drawing.Size(135, 34)
$clearButton.Anchor = 'Bottom,Right'
$targetsGroup.Controls.Add($clearButton)

$selectionHintLabel = New-Object System.Windows.Forms.Label
$selectionHintLabel.Text = 'Selections take effect right away. Save Settings to keep them.'
$selectionHintLabel.Location = New-Object System.Drawing.Point(15, 515)
$selectionHintLabel.Size = New-Object System.Drawing.Size(287, 32)
$selectionHintLabel.Anchor = 'Bottom,Left,Right'
$targetsGroup.Controls.Add($selectionHintLabel)

$statusGroup = New-Object System.Windows.Forms.GroupBox
$statusGroup.Text = 'Status And Controls'
$statusGroup.Location = New-Object System.Drawing.Point(346, 12)
$statusGroup.Size = New-Object System.Drawing.Size(526, 220)
$statusGroup.Anchor = 'Top,Left,Right'
$script:MainForm.Controls.Add($statusGroup)

$script:CurrentStateLabel = New-Object System.Windows.Forms.Label
$script:CurrentStateLabel.Text = 'Current status: checking'
$script:CurrentStateLabel.Location = New-Object System.Drawing.Point(18, 30)
$script:CurrentStateLabel.Size = New-Object System.Drawing.Size(480, 28)
$script:CurrentStateLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
$statusGroup.Controls.Add($script:CurrentStateLabel)

$script:SummaryLabel = New-Object System.Windows.Forms.Label
$script:SummaryLabel.Location = New-Object System.Drawing.Point(20, 68)
$script:SummaryLabel.Size = New-Object System.Drawing.Size(480, 20)
$statusGroup.Controls.Add($script:SummaryLabel)

$script:HotkeyStatusLabel = New-Object System.Windows.Forms.Label
$script:HotkeyStatusLabel.Location = New-Object System.Drawing.Point(20, 92)
$script:HotkeyStatusLabel.Size = New-Object System.Drawing.Size(480, 20)
$statusGroup.Controls.Add($script:HotkeyStatusLabel)

$script:HelperLabel = New-Object System.Windows.Forms.Label
$script:HelperLabel.Location = New-Object System.Drawing.Point(20, 116)
$script:HelperLabel.Size = New-Object System.Drawing.Size(480, 20)
$statusGroup.Controls.Add($script:HelperLabel)

$toggleHotkeyLabel = New-Object System.Windows.Forms.Label
$toggleHotkeyLabel.Text = 'Toggle hotkey'
$toggleHotkeyLabel.Location = New-Object System.Drawing.Point(20, 146)
$toggleHotkeyLabel.Size = New-Object System.Drawing.Size(95, 20)
$statusGroup.Controls.Add($toggleHotkeyLabel)

$script:ToggleHotkeyTextBox = New-Object System.Windows.Forms.TextBox
$script:ToggleHotkeyTextBox.Location = New-Object System.Drawing.Point(120, 143)
$script:ToggleHotkeyTextBox.Size = New-Object System.Drawing.Size(120, 25)
$script:ToggleHotkeyTextBox.Text = $script:Settings.Hotkey.Display
$statusGroup.Controls.Add($script:ToggleHotkeyTextBox)

$exitHotkeyLabel = New-Object System.Windows.Forms.Label
$exitHotkeyLabel.Text = 'Exit hotkey'
$exitHotkeyLabel.Location = New-Object System.Drawing.Point(260, 146)
$exitHotkeyLabel.Size = New-Object System.Drawing.Size(70, 20)
$statusGroup.Controls.Add($exitHotkeyLabel)

$script:ExitHotkeyTextBox = New-Object System.Windows.Forms.TextBox
$script:ExitHotkeyTextBox.Location = New-Object System.Drawing.Point(336, 143)
$script:ExitHotkeyTextBox.Size = New-Object System.Drawing.Size(145, 25)
$script:ExitHotkeyTextBox.Text = if ($script:Settings.ExitHotkey) { $script:Settings.ExitHotkey.Display } else { '' }
$statusGroup.Controls.Add($script:ExitHotkeyTextBox)

$saveSettingsButton = New-Object System.Windows.Forms.Button
$saveSettingsButton.Text = 'Save Settings'
$saveSettingsButton.Location = New-Object System.Drawing.Point(20, 178)
$saveSettingsButton.Size = New-Object System.Drawing.Size(130, 30)
$statusGroup.Controls.Add($saveSettingsButton)

$hideNowButton = New-Object System.Windows.Forms.Button
$hideNowButton.Text = 'Hide Now'
$hideNowButton.Location = New-Object System.Drawing.Point(164, 178)
$hideNowButton.Size = New-Object System.Drawing.Size(100, 30)
$statusGroup.Controls.Add($hideNowButton)

$restoreButton = New-Object System.Windows.Forms.Button
$restoreButton.Text = 'Restore'
$restoreButton.Location = New-Object System.Drawing.Point(276, 178)
$restoreButton.Size = New-Object System.Drawing.Size(100, 30)
$statusGroup.Controls.Add($restoreButton)

$toggleButton = New-Object System.Windows.Forms.Button
$toggleButton.Text = 'Toggle'
$toggleButton.Location = New-Object System.Drawing.Point(388, 178)
$toggleButton.Size = New-Object System.Drawing.Size(92, 30)
$statusGroup.Controls.Add($toggleButton)

$detailsGroup = New-Object System.Windows.Forms.GroupBox
$detailsGroup.Text = 'Window Details'
$detailsGroup.Location = New-Object System.Drawing.Point(346, 244)
$detailsGroup.Size = New-Object System.Drawing.Size(526, 328)
$detailsGroup.Anchor = 'Top,Bottom,Left,Right'
$script:MainForm.Controls.Add($detailsGroup)

$detailsTabs = New-Object System.Windows.Forms.TabControl
$detailsTabs.Location = New-Object System.Drawing.Point(12, 24)
$detailsTabs.Size = New-Object System.Drawing.Size(500, 292)
$detailsTabs.Anchor = 'Top,Bottom,Left,Right'
$detailsGroup.Controls.Add($detailsTabs)

$currentWindowsTab = New-Object System.Windows.Forms.TabPage
$currentWindowsTab.Text = 'Current Windows'
$currentWindowsTab.AutoScroll = $true
$detailsTabs.Controls.Add($currentWindowsTab)

$currentWindowsHintLabel = New-Object System.Windows.Forms.Label
$currentWindowsHintLabel.Text = 'Select a visible window and decide whether it should always hide or always stay visible.'
$currentWindowsHintLabel.Location = New-Object System.Drawing.Point(12, 10)
$currentWindowsHintLabel.Size = New-Object System.Drawing.Size(470, 32)
$currentWindowsTab.Controls.Add($currentWindowsHintLabel)

$script:CurrentWindowsListView = New-Object System.Windows.Forms.ListView
$script:CurrentWindowsListView.Location = New-Object System.Drawing.Point(12, 46)
$script:CurrentWindowsListView.Size = New-Object System.Drawing.Size(470, 118)
$script:CurrentWindowsListView.Anchor = 'Top,Left,Right'
$script:CurrentWindowsListView.View = 'Details'
$script:CurrentWindowsListView.FullRowSelect = $true
$script:CurrentWindowsListView.GridLines = $true
$script:CurrentWindowsListView.HideSelection = $false
$script:CurrentWindowsListView.Columns.Add('Program', 90) | Out-Null
$script:CurrentWindowsListView.Columns.Add('Behavior', 130) | Out-Null
$script:CurrentWindowsListView.Columns.Add('Window Title', 230) | Out-Null
$currentWindowsTab.Controls.Add($script:CurrentWindowsListView)

$ruleKeywordLabel = New-Object System.Windows.Forms.Label
$ruleKeywordLabel.Text = 'Title keyword'
$ruleKeywordLabel.Location = New-Object System.Drawing.Point(12, 172)
$ruleKeywordLabel.Size = New-Object System.Drawing.Size(120, 20)
$currentWindowsTab.Controls.Add($ruleKeywordLabel)

$script:RuleKeywordTextBox = New-Object System.Windows.Forms.TextBox
$script:RuleKeywordTextBox.Location = New-Object System.Drawing.Point(12, 194)
$script:RuleKeywordTextBox.Size = New-Object System.Drawing.Size(470, 25)
$script:RuleKeywordTextBox.Anchor = 'Top,Left,Right'
$currentWindowsTab.Controls.Add($script:RuleKeywordTextBox)

$hideSelectedWindowButton = New-Object System.Windows.Forms.Button
$hideSelectedWindowButton.Text = 'Always Hide'
$hideSelectedWindowButton.Location = New-Object System.Drawing.Point(12, 228)
$hideSelectedWindowButton.Size = New-Object System.Drawing.Size(110, 28)
$hideSelectedWindowButton.Anchor = 'Top,Left'
$currentWindowsTab.Controls.Add($hideSelectedWindowButton)
