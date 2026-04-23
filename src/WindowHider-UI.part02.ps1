    $json = $configObject | ConvertTo-Json -Depth 4
    Set-Content -LiteralPath $Settings.ConfigPath -Value $json -Encoding UTF8
}

function Copy-WindowRules {
    param([object[]]$Rules)

    $copiedRules = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rule in @($Rules)) {
        $copiedRules.Add([pscustomobject]@{
                ProcessName   = $rule.ProcessName
                TitleContains = $rule.TitleContains
                Action        = $rule.Action
            }) | Out-Null
    }

    return $copiedRules
}

function Get-WindowRulesArray {
    if ($null -eq $script:Settings) {
        return @()
    }

    if ($script:Settings.PSObject.Properties.Name -notcontains 'WindowRules') {
        return @()
    }

    if ($null -eq $script:Settings.WindowRules) {
        return @()
    }

    $rules = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rule in $script:Settings.WindowRules) {
        if ($null -eq $rule) {
            continue
        }

        if ($rule.PSObject.Properties.Name -notcontains 'ProcessName') {
            continue
        }

        $rules.Add($rule) | Out-Null
    }

    return [object[]]$rules.ToArray()
}

function Initialize-SettingsCollections {
    param([pscustomobject]$Settings)

    if ($null -eq $Settings) {
        return
    }

    if ($Settings.PSObject.Properties.Name -notcontains 'WindowRules' -or $null -eq $Settings.WindowRules) {
        $Settings | Add-Member -NotePropertyName WindowRules -NotePropertyValue (New-Object 'System.Collections.Generic.List[object]') -Force
    }
}

function Get-WindowRuleActionLabel {
    param([string]$Action)

    switch ($Action) {
        'hide' { return 'Always hide' }
        'keep' { return 'Always keep visible' }
        default { return $Action }
    }
}

function Test-WindowRuleMatch {
    param(
        [pscustomobject]$Window,
        [pscustomobject]$Rule
    )

    return $Window.Title.IndexOf($Rule.TitleContains, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Resolve-WindowBehavior {
    param(
        [pscustomobject]$Window,
        [object]$Rules
    )

    $allRules = @($Rules)
    $processRules = @(
        $allRules |
            Where-Object {
                $null -ne $_ -and
                $_.PSObject.Properties.Name -contains 'ProcessName' -and
                $_.ProcessName -eq $Window.ProcessName
            }
    )

    foreach ($rule in @($processRules | Where-Object { $_.Action -eq 'keep' })) {
        if (Test-WindowRuleMatch -Window $Window -Rule $rule) {
            return [pscustomobject]@{
                ShouldHide = $false
                Label      = 'Always keep visible'
                Rule       = $rule
            }
        }
    }

    $hideRules = @($processRules | Where-Object { $_.Action -eq 'hide' })
    if ($hideRules.Count -gt 0) {
        foreach ($rule in $hideRules) {
            if (Test-WindowRuleMatch -Window $Window -Rule $rule) {
                return [pscustomobject]@{
                    ShouldHide = $true
                    Label      = 'Always hide'
                    Rule       = $rule
                }
            }
        }

        return [pscustomobject]@{
            ShouldHide = $false
            Label      = 'Visible by window rules'
            Rule       = $null
        }
    }

    return [pscustomobject]@{
        ShouldHide = $true
        Label      = 'Hide by app'
        Rule       = $null
    }
}

function Set-WindowRule {
    param(
        [string]$ProcessName,
        [string]$TitleContains,
        [ValidateSet('hide', 'keep')]
        [string]$Action
    )

    $normalizedProcess = Normalize-ExecutableName -Name $ProcessName
    $keyword = Normalize-OptionalText -Value $TitleContains

    if ([string]::IsNullOrWhiteSpace($normalizedProcess)) {
        throw 'A valid program name is required for the window rule.'
    }

    if ([string]::IsNullOrWhiteSpace($keyword)) {
        throw 'Enter a title keyword for the window rule.'
    }

    $updatedRules = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rule in (Get-WindowRulesArray)) {
        $isSameRule = (
            $rule.ProcessName -eq $normalizedProcess -and
            [string]::Equals([string]$rule.TitleContains, $keyword, [System.StringComparison]::OrdinalIgnoreCase)
        )

        if (-not $isSameRule) {
            $updatedRules.Add($rule) | Out-Null
        }
    }

    $updatedRules.Add([pscustomobject]@{
            ProcessName   = $normalizedProcess
            TitleContains = $keyword
            Action        = $Action
        }) | Out-Null

    $script:Settings.WindowRules = $updatedRules
}

function Remove-WindowRule {
    param(
        [string]$ProcessName,
        [string]$TitleContains
    )

    $updatedRules = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rule in (Get-WindowRulesArray)) {
        $isSameRule = (
            $rule.ProcessName -eq $ProcessName -and
            [string]::Equals([string]$rule.TitleContains, [string]$TitleContains, [System.StringComparison]::OrdinalIgnoreCase)
        )

        if (-not $isSameRule) {
            $updatedRules.Add($rule) | Out-Null
        }
    }

    $script:Settings.WindowRules = $updatedRules
}

function Get-SavedRulesForSelection {
    $selectedTargets = @(Get-CheckedTargetNames)
    $selectedSet = New-TargetSet -Names $selectedTargets

    if ($selectedSet.Count -eq 0) {
        return Get-WindowRulesArray
    }

    return @((Get-WindowRulesArray) | Where-Object { $selectedSet.Contains($_.ProcessName) })
}

function Get-VisibleWindowInventory {
    $shellWindow = [Win32WindowTools]::GetShellWindow()
    $results = [System.Collections.Generic.List[object]]::new()

    $callback = [Win32WindowTools+EnumWindowsProc]{
        param([IntPtr]$Handle, [IntPtr]$Unused)

        if ($Handle -eq $shellWindow) {
            return $true
        }

        if (-not [Win32WindowTools]::IsWindowVisible($Handle)) {
            return $true
        }

        if ([Win32WindowTools]::IsIconic($Handle)) {
            return $true
        }

        if ([Win32WindowTools]::GetAncestor($Handle, [Win32WindowTools]::GA_ROOTOWNER) -ne $Handle) {
            return $true
        }

        $title = [Win32WindowTools]::GetWindowTitle($Handle)
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }

        $processId = [uint32]0
        [void][Win32WindowTools]::GetWindowThreadProcessId($Handle, [ref]$processId)
        if ($processId -eq 0) {
            return $true
        }

        try {
            $process = [System.Diagnostics.Process]::GetProcessById([int]$processId)
            $exeName = Normalize-ExecutableName -Name $process.ProcessName
        }
        catch {
            return $true
        }

        $results.Add([pscustomobject]@{
                Handle      = $Handle
                ProcessName = $exeName
                ProcessId   = $processId
                Title       = $title
            }) | Out-Null

        return $true
    }

    [void][Win32WindowTools]::EnumWindows($callback, [IntPtr]::Zero)
    return [object[]]$results.ToArray()
}

function Sync-HiddenWindows {
    $active = [System.Collections.Generic.List[object]]::new()

    foreach ($window in $script:HiddenWindows.ToArray()) {
        if (-not [Win32WindowTools]::IsWindow($window.Handle)) {
            continue
        }

        if ([Win32WindowTools]::IsWindowVisible($window.Handle)) {
            continue
        }

        $active.Add($window) | Out-Null
    }

    $script:HiddenWindows.Clear()
    foreach ($window in $active) {
        $script:HiddenWindows.Add($window) | Out-Null
    }
}

function Get-MatchingWindows {
    param(
        [System.Collections.Generic.HashSet[string]]$Targets,
        [object[]]$Inventory
    )

    if ($Targets.Count -eq 0) {
        return @()
    }

    $source = if ($Inventory) { @($Inventory) } else { @(Get-VisibleWindowInventory) }
    return @($source | Where-Object { $Targets.Contains($_.ProcessName) })
}

function Get-TargetCatalog {
    param([string[]]$SelectedTargets)

    Sync-HiddenWindows

    $inventory = Get-VisibleWindowInventory
    $visibleGroups = @{}
    foreach ($window in $inventory) {
        if (-not $visibleGroups.ContainsKey($window.ProcessName)) {
            $visibleGroups[$window.ProcessName] = [System.Collections.Generic.List[object]]::new()
        }
        $visibleGroups[$window.ProcessName].Add($window) | Out-Null
    }

    $hiddenGroups = @{}
    foreach ($window in $script:HiddenWindows.ToArray()) {
        if (-not $hiddenGroups.ContainsKey($window.ProcessName)) {
            $hiddenGroups[$window.ProcessName] = [System.Collections.Generic.List[object]]::new()
        }
        $hiddenGroups[$window.ProcessName].Add($window) | Out-Null
    }

    $allNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $SelectedTargets) {
        $normalized = Normalize-ExecutableName -Name $name
        if ($normalized) {
            [void]$allNames.Add($normalized)
        }
    }

    foreach ($name in $visibleGroups.Keys) {
        [void]$allNames.Add($name)
    }

    foreach ($name in $hiddenGroups.Keys) {
        [void]$allNames.Add($name)
    }

    $catalog = foreach ($name in ($allNames | Sort-Object)) {
        $visibleCount = if ($visibleGroups.ContainsKey($name)) { $visibleGroups[$name].Count } else { 0 }
        $hiddenCount = if ($hiddenGroups.ContainsKey($name)) { $hiddenGroups[$name].Count } else { 0 }

        [pscustomobject]@{
            ExeName      = $name
            VisibleCount = $visibleCount
            HiddenCount  = $hiddenCount
        }
    }

    [pscustomobject]@{
        Inventory = $inventory
        Catalog   = @($catalog)
    }
}
