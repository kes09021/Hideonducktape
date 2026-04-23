function Format-TargetEntry {
    param([pscustomobject]$Entry)

    if ($Entry.HiddenCount -gt 0 -and $Entry.VisibleCount -gt 0) {
        return "{0}  |  visible: {1}, hidden: {2}" -f $Entry.ExeName, $Entry.VisibleCount, $Entry.HiddenCount
    }

    if ($Entry.HiddenCount -gt 0) {
        return "{0}  |  hidden windows: {1}" -f $Entry.ExeName, $Entry.HiddenCount
    }

    if ($Entry.VisibleCount -gt 0) {
        return "{0}  |  visible windows: {1}" -f $Entry.ExeName, $Entry.VisibleCount
    }

    return "{0}  |  not running right now" -f $Entry.ExeName
}

function Get-CheckedTargetNames {
    $targets = [System.Collections.Generic.List[string]]::new()

    if ($null -eq $script:TargetListBox) {
        return [string[]]$targets.ToArray()
    }

    foreach ($index in $script:TargetListBox.CheckedIndices) {
        $entry = $script:TargetEntries[[int]$index]
        if ($null -ne $entry) {
            $targets.Add($entry.ExeName) | Out-Null
        }
    }

    return [string[]]$targets.ToArray()
}

function Refresh-TargetList {
    param([string[]]$CheckedTargets)

    $catalogData = Get-TargetCatalog -SelectedTargets $CheckedTargets
    $script:LatestInventory = @($catalogData.Inventory)
    $script:TargetEntries = @($catalogData.Catalog)

    $checkedSet = New-TargetSet -Names $CheckedTargets

    $script:IsRefreshingTargetList = $true
    $script:TargetListBox.BeginUpdate()

    try {
        $script:TargetListBox.Items.Clear()

        for ($index = 0; $index -lt $script:TargetEntries.Count; $index++) {
            $entry = $script:TargetEntries[$index]
            [void]$script:TargetListBox.Items.Add((Format-TargetEntry -Entry $entry))
            if ($checkedSet.Contains($entry.ExeName)) {
                $script:TargetListBox.SetItemChecked($index, $true)
            }
        }
    }

    finally {
        $script:TargetListBox.EndUpdate()
        $script:IsRefreshingTargetList = $false
    }

    Update-RuntimeTargets
    Refresh-CurrentWindowsList
    Refresh-SavedRulesList
    Update-UiStatus
}

function Update-RuntimeTargets {
    $script:RuntimeTargets = New-TargetSet -Names @(Get-CheckedTargetNames)
}

function Get-SelectedCurrentWindow {
    if ($null -eq $script:CurrentWindowsListView -or $script:CurrentWindowsListView.SelectedItems.Count -eq 0) {
        return $null
    }

    return $script:CurrentWindowsListView.SelectedItems[0].Tag
}

function Get-SelectedSavedRule {
    if ($null -eq $script:SavedRulesListView -or $script:SavedRulesListView.SelectedItems.Count -eq 0) {
        return $null
    }

    return $script:SavedRulesListView.SelectedItems[0].Tag
}

function Refresh-CurrentWindowsList {
    if ($null -eq $script:CurrentWindowsListView) {
        return
    }

    $selectedTargets = @(Get-CheckedTargetNames)
    $selectedSet = New-TargetSet -Names $selectedTargets
    $windows = @(
        Get-MatchingWindows -Targets $selectedSet -Inventory $script:LatestInventory |
            Sort-Object ProcessName, Title
    )

    $script:CurrentWindowsListView.BeginUpdate()

    try {
        $script:CurrentWindowsListView.Items.Clear()
        foreach ($window in $windows) {
            $behavior = Resolve-WindowBehavior -Window $window -Rules (Get-WindowRulesArray)
            $item = New-Object System.Windows.Forms.ListViewItem($window.ProcessName)
            [void]$item.SubItems.Add($behavior.Label)
            [void]$item.SubItems.Add([string]$window.Title)
            $item.Tag = $window
            [void]$script:CurrentWindowsListView.Items.Add($item)
        }
    }
    finally {
        $script:CurrentWindowsListView.EndUpdate()
    }
}

function Refresh-SavedRulesList {
    if ($null -eq $script:SavedRulesListView) {
        return
    }

    $rules = @(
        Get-SavedRulesForSelection |
            Sort-Object ProcessName, Action, TitleContains
    )

    $script:SavedRulesListView.BeginUpdate()

    try {
        $script:SavedRulesListView.Items.Clear()
        foreach ($rule in $rules) {
            $item = New-Object System.Windows.Forms.ListViewItem($rule.ProcessName)
            [void]$item.SubItems.Add((Get-WindowRuleActionLabel -Action $rule.Action))
            [void]$item.SubItems.Add($rule.TitleContains)
            $item.Tag = $rule
            [void]$script:SavedRulesListView.Items.Add($item)
        }
    }
    finally {
        $script:SavedRulesListView.EndUpdate()
    }
}

function Add-WindowRuleFromSelection {
    param(
        [ValidateSet('hide', 'keep')]
        [string]$Action
    )

    $window = Get-SelectedCurrentWindow
    if ($null -eq $window) {
        [System.Windows.Forms.MessageBox]::Show('Choose a window from the Current Windows tab first.', 'Window Hider') | Out-Null
        return
    }

    $keyword = if ($null -ne $script:RuleKeywordTextBox) { Normalize-OptionalText -Value $script:RuleKeywordTextBox.Text } else { '' }
    if ([string]::IsNullOrWhiteSpace($keyword)) {
        $keyword = [string]$window.Title
    }

    try {
        Set-WindowRule -ProcessName $window.ProcessName -TitleContains $keyword -Action $Action
        Refresh-CurrentWindowsList
        Refresh-SavedRulesList
        Update-UiStatus
        Set-StatusMessage ("Saved runtime rule for {0}. Click Save Settings to keep it." -f $window.ProcessName)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Window Hider') | Out-Null
    }
}

function Remove-SelectedSavedRule {
    $rule = Get-SelectedSavedRule
    if ($null -eq $rule) {
        [System.Windows.Forms.MessageBox]::Show('Choose a saved rule from the Saved Rules tab first.', 'Window Hider') | Out-Null
        return
    }

    Remove-WindowRule -ProcessName $rule.ProcessName -TitleContains $rule.TitleContains
    Refresh-CurrentWindowsList
    Refresh-SavedRulesList
    Update-UiStatus
    Set-StatusMessage 'Saved rule removed from runtime. Click Save Settings to keep it.'
}

function Hide-SelectedWindows {
    Sync-HiddenWindows

    if ($script:RuntimeTargets.Count -eq 0) {
        return 'Select at least one program first.'
    }

    $candidateWindows = @(Get-MatchingWindows -Targets $script:RuntimeTargets -Inventory $script:LatestInventory)
    if ($candidateWindows.Count -eq 0) {
        return 'No visible windows were found for the selected programs.'
    }

    $windowsToHide = @()
    foreach ($window in $candidateWindows) {
        $behavior = Resolve-WindowBehavior -Window $window -Rules (Get-WindowRulesArray)
        if ($behavior.ShouldHide) {
            $windowsToHide += $window
        }
    }

    if ($windowsToHide.Count -eq 0) {
        return 'No windows matched the current window rules. Add a hide rule or remove a keep-visible rule.'
    }

    $script:HiddenWindows.Clear()
    foreach ($window in $windowsToHide) {
        [void][Win32WindowTools]::ShowWindow($window.Handle, [Win32WindowTools]::SW_HIDE)
        $script:HiddenWindows.Add($window) | Out-Null
    }

    return "Hidden {0} window(s)." -f $windowsToHide.Count
}

function Restore-HiddenWindows {
    Sync-HiddenWindows

    $restored = 0
    foreach ($window in $script:HiddenWindows.ToArray()) {
        if ([Win32WindowTools]::IsWindow($window.Handle)) {
            [void][Win32WindowTools]::ShowWindow($window.Handle, [Win32WindowTools]::SW_SHOW)
            $restored++
        }
    }

    $script:HiddenWindows.Clear()

    if ($restored -eq 0) {
        return 'There are no hidden windows to restore.'
    }

    return "Restored {0} window(s)." -f $restored
}

function Toggle-SelectedWindows {
    Sync-HiddenWindows

    if ($script:HiddenWindows.Count -gt 0) {
        return Restore-HiddenWindows
    }

    return Hide-SelectedWindows
}

function Set-StatusMessage {
    param([string]$Message)

    $script:FooterLabel.Text = $Message
}

function Update-UiStatus {
    Sync-HiddenWindows

    $selectedTargets = @(Get-CheckedTargetNames)
    $selectedSet = New-TargetSet -Names $selectedTargets
    $visibleMatches = @(Get-MatchingWindows -Targets $selectedSet -Inventory $script:LatestInventory)
    $windowRuleCount = @(Get-WindowRulesArray).Count

    if ($selectedTargets.Count -eq 0) {
        $script:CurrentStateLabel.Text = 'Current status: choose target programs'
        $script:CurrentStateLabel.ForeColor = [System.Drawing.Color]::DarkGoldenrod
    }
    elseif ($script:HiddenWindows.Count -gt 0 -and $visibleMatches.Count -gt 0) {
        $script:CurrentStateLabel.Text = 'Current status: mixed (some hidden, some visible)'
        $script:CurrentStateLabel.ForeColor = [System.Drawing.Color]::DarkOrange
    }
    elseif ($script:HiddenWindows.Count -gt 0) {
        $script:CurrentStateLabel.Text = 'Current status: hidden'
        $script:CurrentStateLabel.ForeColor = [System.Drawing.Color]::Firebrick
    }
    elseif ($visibleMatches.Count -gt 0) {
        $script:CurrentStateLabel.Text = 'Current status: visible'
        $script:CurrentStateLabel.ForeColor = [System.Drawing.Color]::ForestGreen
    }
    else {
        $script:CurrentStateLabel.Text = 'Current status: no matching windows open'
        $script:CurrentStateLabel.ForeColor = [System.Drawing.Color]::SteelBlue
    }

    $exitText = if ($script:Settings.ExitHotkey) { $script:Settings.ExitHotkey.Display } else { 'not set' }
    $script:SummaryLabel.Text = "Selected apps: $($selectedTargets.Count)  |  Visible target windows: $($visibleMatches.Count)  |  Hidden windows: $($script:HiddenWindows.Count)  |  Window rules: $windowRuleCount"
    $script:HotkeyStatusLabel.Text = "Toggle: $($script:Settings.Hotkey.Display)  |  Exit: $exitText"
    $script:HelperLabel.Text = 'Tip: use Current Windows or Saved Rules to keep one Chrome window visible while hiding the others.'

    $script:HiddenListView.BeginUpdate()
    $script:HiddenListView.Items.Clear()
    foreach ($window in $script:HiddenWindows.ToArray()) {
        $item = New-Object System.Windows.Forms.ListViewItem($window.ProcessName)
        [void]$item.SubItems.Add([string]$window.Title)
        [void]$script:HiddenListView.Items.Add($item)
    }
    $script:HiddenListView.EndUpdate()
}

function Register-OneHotkey {
    param(
        [System.IntPtr]$Handle,
        [int]$Id,
        [pscustomobject]$Definition,
        [string]$Purpose
    )

    if (-not [Win32WindowTools]::RegisterHotKey($Handle, $Id, [uint32]$Definition.Modifiers, [uint32]$Definition.KeyCode)) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "Failed to register $Purpose hotkey '$($Definition.Display)' (Win32 error $errorCode). Choose a different shortcut."
    }
}

function Unregister-Hotkeys {
    param([System.Object]$Target)

    if ($null -eq $Target) {
        return
    }

    [void][Win32WindowTools]::UnregisterHotKey($Target.Handle, $script:ToggleHotkeyId)
    [void][Win32WindowTools]::UnregisterHotKey($Target.Handle, $script:ExitHotkeyId)
    $script:HotkeysRegistered = $false
}

function Register-Hotkeys {
    param(
        [System.Object]$Target,
        [pscustomobject]$Settings
    )

    Register-OneHotkey -Handle $Target.Handle -Id $script:ToggleHotkeyId -Definition $Settings.Hotkey -Purpose 'toggle'
    if ($Settings.ExitHotkey) {
        Register-OneHotkey -Handle $Target.Handle -Id $script:ExitHotkeyId -Definition $Settings.ExitHotkey -Purpose 'exit'
    }

    $script:HotkeysRegistered = $true
}

function Apply-Hotkeys {
    param(
        [System.Object]$Target,
        [pscustomobject]$NewSettings,
        [pscustomobject]$RollbackSettings
    )
