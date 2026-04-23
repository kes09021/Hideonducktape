$keepSelectedWindowButton = New-Object System.Windows.Forms.Button
$keepSelectedWindowButton.Text = 'Keep Visible'
$keepSelectedWindowButton.Location = New-Object System.Drawing.Point(132, 228)
$keepSelectedWindowButton.Size = New-Object System.Drawing.Size(110, 28)
$keepSelectedWindowButton.Anchor = 'Top,Left'
$currentWindowsTab.Controls.Add($keepSelectedWindowButton)

$refreshCurrentWindowsButton = New-Object System.Windows.Forms.Button
$refreshCurrentWindowsButton.Text = 'Refresh Windows'
$refreshCurrentWindowsButton.Location = New-Object System.Drawing.Point(252, 228)
$refreshCurrentWindowsButton.Size = New-Object System.Drawing.Size(110, 28)
$refreshCurrentWindowsButton.Anchor = 'Top,Left'
$currentWindowsTab.Controls.Add($refreshCurrentWindowsButton)

$currentWindowsTipLabel = New-Object System.Windows.Forms.Label
$currentWindowsTipLabel.Text = 'Tip: shorten the title keyword if the full Chrome title changes too often.'
$currentWindowsTipLabel.Location = New-Object System.Drawing.Point(12, 262)
$currentWindowsTipLabel.Size = New-Object System.Drawing.Size(470, 18)
$currentWindowsTipLabel.Anchor = 'Top,Left,Right'
$currentWindowsTab.Controls.Add($currentWindowsTipLabel)

$hiddenWindowsTab = New-Object System.Windows.Forms.TabPage
$hiddenWindowsTab.Text = 'Hidden Windows'
$detailsTabs.Controls.Add($hiddenWindowsTab)

$script:HiddenListView = New-Object System.Windows.Forms.ListView
$script:HiddenListView.Location = New-Object System.Drawing.Point(12, 14)
$script:HiddenListView.Size = New-Object System.Drawing.Size(470, 230)
$script:HiddenListView.Anchor = 'Top,Bottom,Left,Right'
$script:HiddenListView.View = 'Details'
$script:HiddenListView.FullRowSelect = $true
$script:HiddenListView.GridLines = $true
$script:HiddenListView.Columns.Add('Program', 100) | Out-Null
$script:HiddenListView.Columns.Add('Window Title', 340) | Out-Null
$hiddenWindowsTab.Controls.Add($script:HiddenListView)

$hiddenHintLabel = New-Object System.Windows.Forms.Label
$hiddenHintLabel.Text = 'Minimize the app to the tray if you want the hotkey to keep running quietly.'
$hiddenHintLabel.Location = New-Object System.Drawing.Point(12, 252)
$hiddenHintLabel.Size = New-Object System.Drawing.Size(470, 20)
$hiddenHintLabel.Anchor = 'Bottom,Left,Right'
$hiddenWindowsTab.Controls.Add($hiddenHintLabel)

$savedRulesTab = New-Object System.Windows.Forms.TabPage
$savedRulesTab.Text = 'Saved Rules'
$detailsTabs.Controls.Add($savedRulesTab)

$savedRulesHintLabel = New-Object System.Windows.Forms.Label
$savedRulesHintLabel.Text = 'Rules are matched by program name plus a title keyword.'
$savedRulesHintLabel.Location = New-Object System.Drawing.Point(12, 10)
$savedRulesHintLabel.Size = New-Object System.Drawing.Size(470, 20)
$savedRulesTab.Controls.Add($savedRulesHintLabel)

$script:SavedRulesListView = New-Object System.Windows.Forms.ListView
$script:SavedRulesListView.Location = New-Object System.Drawing.Point(12, 38)
$script:SavedRulesListView.Size = New-Object System.Drawing.Size(470, 206)
$script:SavedRulesListView.Anchor = 'Top,Bottom,Left,Right'
$script:SavedRulesListView.View = 'Details'
$script:SavedRulesListView.FullRowSelect = $true
$script:SavedRulesListView.GridLines = $true
$script:SavedRulesListView.HideSelection = $false
$script:SavedRulesListView.Columns.Add('Program', 90) | Out-Null
$script:SavedRulesListView.Columns.Add('Action', 130) | Out-Null
$script:SavedRulesListView.Columns.Add('Title Keyword', 230) | Out-Null
$savedRulesTab.Controls.Add($script:SavedRulesListView)

$removeSavedRuleButton = New-Object System.Windows.Forms.Button
$removeSavedRuleButton.Text = 'Remove Selected Rule'
$removeSavedRuleButton.Location = New-Object System.Drawing.Point(12, 250)
$removeSavedRuleButton.Size = New-Object System.Drawing.Size(160, 28)
$removeSavedRuleButton.Anchor = 'Bottom,Left'
$savedRulesTab.Controls.Add($removeSavedRuleButton)

$refreshSavedRulesButton = New-Object System.Windows.Forms.Button
$refreshSavedRulesButton.Text = 'Refresh Rules'
$refreshSavedRulesButton.Location = New-Object System.Drawing.Point(182, 250)
$refreshSavedRulesButton.Size = New-Object System.Drawing.Size(110, 28)
$refreshSavedRulesButton.Anchor = 'Bottom,Left'
$savedRulesTab.Controls.Add($refreshSavedRulesButton)

$script:FooterLabel = New-Object System.Windows.Forms.Label
$script:FooterLabel.Location = New-Object System.Drawing.Point(16, 580)
$script:FooterLabel.Size = New-Object System.Drawing.Size(856, 20)
$script:FooterLabel.Anchor = 'Bottom,Left,Right'
$script:FooterLabel.Text = 'Ready.'
$script:MainForm.Controls.Add($script:FooterLabel)

$notifyMenu = New-Object System.Windows.Forms.ContextMenuStrip
$openMenuItem = $notifyMenu.Items.Add('Open Window Hider')
$toggleMenuItem = $notifyMenu.Items.Add('Toggle Hidden State')
$restoreAllMenuItem = $notifyMenu.Items.Add('Restore Hidden Windows')
$exitMenuItem = $notifyMenu.Items.Add('Exit')

$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:NotifyIcon.Text = 'Window Hider'
$script:TrayIcon = New-TrayAppIcon
$script:NotifyIcon.Icon = $script:TrayIcon
$script:NotifyIcon.Visible = $true
$script:NotifyIcon.ContextMenuStrip = $notifyMenu

$refreshTimer = New-Object System.Windows.Forms.Timer
$refreshTimer.Interval = 2000

$refreshAfterCheck = {
    if ($script:IsRefreshingTargetList) {
        return
    }

    if ($script:MainForm.IsHandleCreated) {
        $script:MainForm.BeginInvoke([System.Action]{
                Update-RuntimeTargets
                Update-UiStatus
            }) | Out-Null
        return
    }

    Update-RuntimeTargets
    Update-UiStatus
}

$script:TargetListBox.Add_ItemCheck($refreshAfterCheck)
$addProgramButton.Add_Click({ Add-ManualTarget })
$script:AddProgramTextBox.Add_KeyDown({
        param($sender, $eventArgs)
        if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Add-ManualTarget
            $eventArgs.SuppressKeyPress = $true
        }
    })

$script:CurrentWindowsListView.Add_SelectedIndexChanged({
        $window = Get-SelectedCurrentWindow
        if ($null -ne $window -and $null -ne $script:RuleKeywordTextBox) {
            $script:RuleKeywordTextBox.Text = [string]$window.Title
        }
    })

$hideSelectedWindowButton.Add_Click({ Add-WindowRuleFromSelection -Action 'hide' })
$keepSelectedWindowButton.Add_Click({ Add-WindowRuleFromSelection -Action 'keep' })
$refreshCurrentWindowsButton.Add_Click({
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage 'Current windows refreshed.'
    })

$removeSavedRuleButton.Add_Click({ Remove-SelectedSavedRule })
$refreshSavedRulesButton.Add_Click({
        Refresh-SavedRulesList
        Set-StatusMessage 'Saved rules refreshed.'
    })

$refreshButton.Add_Click({
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage 'Program list refreshed.'
    })

$clearButton.Add_Click({
        for ($index = 0; $index -lt $script:TargetListBox.Items.Count; $index++) {
            $script:TargetListBox.SetItemChecked($index, $false)
        }
        Update-RuntimeTargets
        Update-UiStatus
        Set-StatusMessage 'All checks cleared.'
    })

$saveSettingsButton.Add_Click({ Save-UiChanges })
$hideNowButton.Add_Click({
        $message = Hide-SelectedWindows
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage $message
    })

$restoreButton.Add_Click({
        $message = Restore-HiddenWindows
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage $message
    })

$toggleButton.Add_Click({
        $message = Toggle-SelectedWindows
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage $message
    })

$script:HotkeySink.add_HotkeyPressed({
        param($sender, $eventArgs)

        if ($eventArgs.HotkeyId -eq $script:ToggleHotkeyId) {
            $message = Toggle-SelectedWindows
            Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
            Set-StatusMessage $message
            return
        }

        if ($eventArgs.HotkeyId -eq $script:ExitHotkeyId) {
            $script:MainForm.Close()
        }
    })

$script:MainForm.Add_Shown({
        try {
            Register-Hotkeys -Target $script:HotkeySink -Settings $script:Settings
            Refresh-TargetList -CheckedTargets ($script:Settings.Targets | Sort-Object)
            Set-StatusMessage 'Window Hider is running. Use the hotkey or the buttons here.'
            $refreshTimer.Start()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Window Hider') | Out-Null
            $script:MainForm.Close()
        }
    })

$script:MainForm.Add_Resize({
        if ($script:MainForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            $script:MainForm.ShowInTaskbar = $false
            $script:MainForm.Hide()
            if (-not $script:TrayHintShown) {
                $script:NotifyIcon.ShowBalloonTip(1200, 'Window Hider', 'Still running in the tray. Double-click the tray icon to reopen.', [System.Windows.Forms.ToolTipIcon]::Info)
                $script:TrayHintShown = $true
            }
        }
    })

$script:NotifyIcon.Add_DoubleClick({ Show-MainForm })
$openMenuItem.Add_Click({ Show-MainForm })
$toggleMenuItem.Add_Click({
        $message = Toggle-SelectedWindows
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage $message
    })
$restoreAllMenuItem.Add_Click({
        $message = Restore-HiddenWindows
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
        Set-StatusMessage $message
    })
$exitMenuItem.Add_Click({
        $script:IsShuttingDown = $true
        $script:MainForm.Close()
    })

$refreshTimer.Add_Tick({
        Refresh-TargetList -CheckedTargets @(Get-CheckedTargetNames)
    })

$script:MainForm.Add_FormClosing({
        if (-not $script:IsShuttingDown) {
            $script:IsShuttingDown = $true
        }
    })

$script:MainForm.Add_FormClosed({
        $refreshTimer.Stop()
        try {
            Restore-HiddenWindows | Out-Null
        }
        catch {
        }
        Unregister-Hotkeys -Target $script:HotkeySink
        $script:HotkeySink.Dispose()
        $script:NotifyIcon.Visible = $false
        $script:NotifyIcon.Dispose()
        if ($null -ne $script:TrayIcon) {
            $script:TrayIcon.Dispose()
        }
        $refreshTimer.Dispose()
    })

if ($SmokeTest) {
    Refresh-TargetList -CheckedTargets ($script:Settings.Targets | Sort-Object)
    Write-Host 'UI smoke test OK'
    Write-Host "Loaded targets : $((@(Get-CheckedTargetNames)) -join ', ')"
    Write-Host "Catalog items  : $($script:TargetListBox.Items.Count)"
    Write-Host "Hidden windows : $($script:HiddenWindows.Count)"
    Unregister-Hotkeys -Target $script:HotkeySink
    $script:HotkeySink.Dispose()
    $script:NotifyIcon.Visible = $false
    $script:NotifyIcon.Dispose()
    if ($null -ne $script:TrayIcon) {
        $script:TrayIcon.Dispose()
    }
    $refreshTimer.Dispose()
    $script:MainForm.Dispose()
    exit 0
}

[System.Windows.Forms.Application]::Run($script:MainForm)
