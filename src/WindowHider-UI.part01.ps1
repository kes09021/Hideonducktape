[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$ValidateOnly,
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Language CSharp -ReferencedAssemblies @('System.Windows.Forms', 'System.Drawing') @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

public static class Win32WindowTools
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
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

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);

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

public class HotkeyEventArgs : EventArgs
{
    public int HotkeyId { get; private set; }

    public HotkeyEventArgs(int hotkeyId)
    {
        HotkeyId = hotkeyId;
    }
}

public class HotkeyAwareForm : Form
{
    public event EventHandler<HotkeyEventArgs> HotkeyPressed;

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == Win32WindowTools.WM_HOTKEY)
        {
            var handler = HotkeyPressed;
            if (handler != null)
            {
                handler(this, new HotkeyEventArgs(message.WParam.ToInt32()));
            }
        }

        base.WndProc(ref message);
    }
}

public class HotkeyMessageWindow : NativeWindow, IDisposable
{
    public event EventHandler<HotkeyEventArgs> HotkeyPressed;

    public HotkeyMessageWindow()
    {
        CreateParams parameters = new CreateParams();
        parameters.Caption = "WindowHiderHotkeySink";
        this.CreateHandle(parameters);
    }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg == Win32WindowTools.WM_HOTKEY)
        {
            var handler = HotkeyPressed;
            if (handler != null)
            {
                handler(this, new HotkeyEventArgs(message.WParam.ToInt32()));
            }
        }

        base.WndProc(ref message);
    }

    public void Dispose()
    {
        if (this.Handle != IntPtr.Zero)
        {
            this.DestroyHandle();
        }
    }
}
"@

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

function Normalize-OptionalText {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return $Value.Trim()
}

function New-TargetSet {
    param([string[]]$Names)

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @($Names)) {
        $normalized = Normalize-ExecutableName -Name $name
        if ($null -ne $normalized) {
            [void]$set.Add($normalized)
        }
    }

    return $set
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

        if ($normalizedToken -in @('`', 'backquote', 'grave', 'oem3', '~', 'tilde')) {
            $keyCode = [uint32][System.Windows.Forms.Keys]::Oem3
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

    $logToConsole = $true
    if ($null -ne $config.logToConsole) {
        $logToConsole = [bool]$config.logToConsole
    }

    $windowRuleEntries = @()
    if ($config.PSObject.Properties.Name -contains 'windowRules') {
        $windowRuleEntries = @($config.windowRules)
    }

    $windowRules = New-Object 'System.Collections.Generic.List[object]'
    $seenRuleKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $windowRuleEntries) {
        $processName = Normalize-ExecutableName -Name ([string]$entry.processName)
        $titleContains = if ($null -ne $entry.titleContains) { ([string]$entry.titleContains).Trim() } else { '' }
        $action = if ($null -ne $entry.action) { ([string]$entry.action).Trim().ToLowerInvariant() } else { '' }

        if ([string]::IsNullOrWhiteSpace($processName) -or [string]::IsNullOrWhiteSpace($titleContains)) {
            continue
        }

        if ($action -notin @('hide', 'keep')) {
            continue
        }

        $ruleKey = '{0}|{1}' -f $processName.ToLowerInvariant(), $titleContains.ToLowerInvariant()
        if (-not $seenRuleKeys.Add($ruleKey)) {
            continue
        }

        $windowRules.Add([pscustomobject]@{
                ProcessName   = $processName
                TitleContains = $titleContains
                Action        = $action
            }) | Out-Null
    }

    [pscustomobject]@{
        ConfigPath   = (Resolve-Path -LiteralPath $Path).Path
        Hotkey       = Parse-Hotkey -Text ([string]$config.hotkey)
        ExitHotkey   = if ($config.exitHotkey) { Parse-Hotkey -Text ([string]$config.exitHotkey) } else { $null }
        Targets      = $targets
        WindowRules  = $windowRules
        LogToConsole = $logToConsole
    }
}

function Save-Settings {
    param([pscustomobject]$Settings)

    $configObject = [pscustomobject]@{
        hotkey       = $Settings.Hotkey.Display
        exitHotkey   = if ($Settings.ExitHotkey) { $Settings.ExitHotkey.Display } else { $null }
        targets      = @($Settings.Targets | Sort-Object)
        windowRules  = @(
            @($Settings.WindowRules | Sort-Object ProcessName, Action, TitleContains) |
                ForEach-Object {
                    [pscustomobject]@{
                        processName   = $_.ProcessName
                        titleContains = $_.TitleContains
                        action        = $_.Action
                    }
                }
        )
        logToConsole = [bool]$Settings.LogToConsole
    }
