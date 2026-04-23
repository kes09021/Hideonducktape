# Window Hider

Window Hider is a small Windows utility written in PowerShell that can hide and restore selected application windows with a global hotkey.

It includes:

- a desktop UI for choosing target programs
- per-window rules such as "keep this Chrome window visible"
- tray mode so the tool can keep running in the background
- a simple launcher for double-click use

## Files

- `WindowHider-UI.ps1`: main UI app
- `Hide-ConfiguredWindows.ps1`: original hotkey-only version
- `window-hider.config.json`: default configuration
- `RUN_WINDOW_HIDER.vbs`: double-click launcher
- `start-window-hider.cmd`: visible launcher for debugging

## Features

- hide or restore selected apps with a global hotkey
- choose target executables from the UI
- keep one window visible while hiding other windows from the same app
- minimize to tray and keep the hotkey active
- restore hidden windows from the tray menu

## Quick Start

1. Edit `window-hider.config.json` if you want a different hotkey or default targets.
2. Double-click `RUN_WINDOW_HIDER.vbs`.
3. Select the apps you want to control.
4. Use the hotkey to hide and restore windows.

## Default Hotkeys

- toggle: `Alt+\``
- exit: `Ctrl+Alt+Shift+Q`

## Per-Window Example

If two Chrome windows are open and you want one of them to stay visible:

1. Check `chrome.exe` in the target list.
2. Open the `Current Windows` tab.
3. Click the Chrome window you want to keep visible.
4. Adjust the `Title keyword` if needed.
5. Click `Keep Visible`.
6. Click `Save Settings`.

## Requirements

- Windows
- PowerShell 5.1 or later

## Notes

- This project does not require Python.
- The tray icon appears as a small `H`.
- Closing the app restores any windows hidden by the tool.
