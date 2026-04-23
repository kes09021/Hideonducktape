[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

Add-Type -Language CSharp @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class Win32WindowTools
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public UIntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    public const int WM_HOTKEY = 0x0312;
    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
    public const uint MOD_ALT = 0x0001;
    public const uint MOD_CONTROL = 0x0002;
    public const uint MOD_SHIFT = 0x0004;
    public const uint MOD_WIN = 0x0008;
    public const uint MOD_NOREPEAT = 0x4000;
    public const uint GA_ROOTOWNER = 3;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetShellWindow();

    [DllImport("user32.dll")]
    public static extern int GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);

    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    public static string GetWindowTitle(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        if (length == 0)
        {
            return string.Empty;
        }

        StringBuilder builder = new StringBuilder(length + 1);
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }
}
"@

function Write-Log {
    param(
        [string]$Message,
        [switch]$Always
    )

    if ($Always -or $script:Settings.LogToConsole) {
        $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "[$stamp] $Message"
    }
}

function Normalize-ExecutableName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }

    $trimmed = [System.IO.Path]::GetFileName($Name.Trim())
    if (-not $trimmed.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = "$trimmed.exe"
    }

    return $trimmed.ToLowerInvariant()
}

function Parse-Hotkey {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        throw 'Hotkey cannot be empty.'
    }

    $tokens = @(
        $Text.Split('+', [System.StringSplitOptions]::RemoveEmptyEntries) |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
    )

    if ($tokens.Count -eq 0) {
        throw "Could not parse hotkey '$Text'."
    }

    $modifiers = [uint32]0
    $keyCode = $null

    foreach ($token in $tokens) {
        $normalizedToken = $token.ToLowerInvariant()

        if ($normalizedToken -in @('ctrl', 'control')) {
            $modifiers = $modifiers -bor [Win32WindowTools]::MOD_CONTROL
            continue
        }

        if ($normalizedToken -eq 'alt') {
            $modifiers = $modifiers -bor [Win32WindowTools]::MOD_ALT
            continue
        }

        if ($normalizedToken -eq 'shift') {
            $modifiers = $modifiers -bor [Win32WindowTools]::MOD_SHIFT
            continue
        }

        if ($normalizedToken -in @('win', 'windows')) {
            $modifiers = $modifiers -bor [Win32WindowTools]::MOD_WIN
            continue
        }

        if ($null -ne $keyCode) {
            throw "Only one non-modifier key is allowed in hotkey '$Text'."
        }

        try {
            $keyCode = [uint32][System.Enum]::Parse([System.Windows.Forms.Keys], $token, $true)
        }
        catch {
            throw "Unsupported hotkey token '$token' in '$Text'."
        }
    }

    if ($null -eq $keyCode) {
        throw "Hotkey '$Text' must include a non-modifier key such as H or F12."
    }

    [pscustomobject]@{
        Display   = $Text
        Modifiers = ($modifiers -bor [Win32WindowTools]::MOD_NOREPEAT)
        KeyCode   = $keyCode
    }
}

function Read-Settings {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path"
    }

    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json

    $targets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in @($config.targets)) {
        $normalized = Normalize-ExecutableName -Name ([string]$entry)
        if ($null -ne $normalized) {
            [void]$targets.Add($normalized)
        }
    }

    if ($targets.Count -eq 0) {
        throw 'Config must include at least one executable in the targets array.'
    }

    $logToConsole = $true
    if ($null -ne $config.logToConsole) {
        $logToConsole = [bool]$config.logToConsole
    }

    [pscustomobject]@{
        ConfigPath    = (Resolve-Path -LiteralPath $Path).Path
        Hotkey        = Parse-Hotkey -Text ([string]$config.hotkey)
        ExitHotkey    = if ($config.exitHotkey) { Parse-Hotkey -Text ([string]$config.exitHotkey) } else { $null }
        Targets       = $targets
        LogToConsole  = $logToConsole
    }
}

function Register-GlobalHotkey {
    param(
        [int]$Id,
        [pscustomobject]$Definition,
        [string]$Purpose
    )

    if (-not [Win32WindowTools]::RegisterHotKey([IntPtr]::Zero, $Id, [uint32]$Definition.Modifiers, [uint32]$Definition.KeyCode)) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "Failed to register $Purpose hotkey '$($Definition.Display)' (Win32 error $errorCode). Try a different shortcut."
    }
}

function Get-MatchingWindows {
    param([System.Collections.Generic.HashSet[string]]$Targets)

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

        if (-not $Targets.Contains($exeName)) {
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
    return $results
}

function Hide-TargetWindows {
    $windows = Get-MatchingWindows -Targets $script:Settings.Targets

    if ($windows.Count -eq 0) {
        Write-Log -Message 'No matching visible windows were found.'
        return
    }

    $script:HiddenWindows.Clear()
    foreach ($window in $windows) {
        [void][Win32WindowTools]::ShowWindow($window.Handle, [Win32WindowTools]::SW_HIDE)
        $script:HiddenWindows.Add($window) | Out-Null
    }

    $summary = $windows | ForEach-Object { "$($_.ProcessName): $($_.Title)" }
    Write-Log -Message ("Hidden {0} window(s): {1}" -f $windows.Count, ($summary -join '; '))
}

function Restore-TargetWindows {
    $restored = [System.Collections.Generic.List[string]]::new()

    foreach ($window in @($script:HiddenWindows)) {
        if ([Win32WindowTools]::IsWindow($window.Handle)) {
            [void][Win32WindowTools]::ShowWindow($window.Handle, [Win32WindowTools]::SW_SHOW)
            $restored.Add("$($window.ProcessName): $($window.Title)") | Out-Null
        }
    }

    $script:HiddenWindows.Clear()

    if ($restored.Count -eq 0) {
        Write-Log -Message 'There were no hidden windows left to restore.'
        return
    }

    Write-Log -Message ("Restored {0} window(s): {1}" -f $restored.Count, ($restored -join '; '))
}

function Toggle-TargetWindows {
    if ($script:HiddenWindows.Count -gt 0) {
        Restore-TargetWindows
        return
    }

    Hide-TargetWindows
}

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path -Path $scriptRoot -ChildPath 'window-hider.config.json'
}

$script:Settings = Read-Settings -Path $ConfigPath
$script:HiddenWindows = [System.Collections.Generic.List[object]]::new()

if ($ValidateOnly) {
    $targets = $script:Settings.Targets | Sort-Object
    Write-Host "Config OK"
    Write-Host "Toggle hotkey : $($script:Settings.Hotkey.Display)"
    if ($script:Settings.ExitHotkey) {
        Write-Host "Exit hotkey   : $($script:Settings.ExitHotkey.Display)"
    }
    else {
        Write-Host 'Exit hotkey   : not configured'
    }
    Write-Host "Targets       : $($targets -join ', ')"
    exit 0
}

$toggleHotkeyId = 1
$exitHotkeyId = 2

try {
    Register-GlobalHotkey -Id $toggleHotkeyId -Definition $script:Settings.Hotkey -Purpose 'toggle'

    if ($script:Settings.ExitHotkey) {
        Register-GlobalHotkey -Id $exitHotkeyId -Definition $script:Settings.ExitHotkey -Purpose 'exit'
    }

    Write-Log -Message "Watching for $($script:Settings.Hotkey.Display) to hide or restore target windows." -Always
    if ($script:Settings.ExitHotkey) {
        Write-Log -Message "Use $($script:Settings.ExitHotkey.Display) to stop the tool." -Always
    }

    $shouldExit = $false
    $message = New-Object Win32WindowTools+MSG
    while (-not $shouldExit) {
        $result = [Win32WindowTools]::GetMessage([ref]$message, [IntPtr]::Zero, 0, 0)
        if ($result -le 0) {
            break
        }

        if ($message.message -ne [Win32WindowTools]::WM_HOTKEY) {
            continue
        }

        $hotkeyId = [int]$message.wParam
        if ($hotkeyId -eq $toggleHotkeyId) {
            Toggle-TargetWindows
            continue
        }

        if ($hotkeyId -eq $exitHotkeyId) {
            Write-Log -Message 'Exit hotkey pressed. Shutting down.' -Always
            $shouldExit = $true
        }
    }
}
finally {
    [void][Win32WindowTools]::UnregisterHotKey([IntPtr]::Zero, $toggleHotkeyId)
    [void][Win32WindowTools]::UnregisterHotKey([IntPtr]::Zero, $exitHotkeyId)
}
