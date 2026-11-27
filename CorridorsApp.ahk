#Requires AutoHotkey v2.0
#SingleInstance Force
#Include <WebViewToo>
InstallKeybdHook

TraySetIcon("Cross-street with buildings.png")

; 1. Unpack bundled HTML and WebView2Loader.dll into the user's writable %TEMP% folder
; (Avoids write-permission errors if the IT-approved folder is read-only)
WorkDir     := A_Temp "\CorridorsVisualizer"
DirCreate(WorkDir)
LibDir := A_MyDocuments "\AutoHotkey\Lib"

HtmlPath    := WorkDir "\Corridors_App.html"
DllPath     := WorkDir "\WebView2Loader.dll"
UserDataDir := WorkDir "\WebViewCache"

FileInstall("CoordCheck.html", HtmlPath, 1)
FileInstall(LibDir "\64bit\WebView2Loader.dll", DllPath, 1)

; 2. Initialize WebViewGui with cache directed to %TEMP%
MainGui := WebViewGui("+Resize +MinSize500x600", "Signal Corridors Visualizer", {DataDir: UserDataDir})
MainGui.OnEvent("Close", (*) => CleanupAndExit())

; 3. Configure Custom Tray Icon Menu
SetupTrayMenu()

; 4. Load HTML and show centered on primary monitor (safe for docking/undocking)
MainGui.Navigate(HtmlPath)
width := "w" A_ScreenWidth/2
height := "h" A_ScreenHeight * .9
MainGui.Show(width " " height "Center")
; MainGui.Show("w520 h850 Center")

; ==============================================================================
; CONTEXT-SENSITIVE HOTKEYS (Active ONLY when Corridors window is focused)
; ==============================================================================

; #HotIf WinActive(MainGui.Hwnd)
#HotIf WinActive("ahk_pid " ProcessExist())

; Increase Hours, Minutes, Seconds (Ctrl + Shift + H / M / S)
^+h::AdjustWebOffset("hrs", 1)
^+m::AdjustWebOffset("min", 1)
^+s::AdjustWebOffset("sec", 1)

; Decrease Hours, Minutes, Seconds (Ctrl + Win + H / M / S)
^#h::AdjustWebOffset("hrs", -1)
^#m::AdjustWebOffset("min", -1)
^#s::AdjustWebOffset("sec", -1)

; Reset Hours, Minutes, and Seconds to 0 (Ctrl + 0 or Ctrl + Numpad0)
^0::
^Numpad0::ResetWebOffset()

#HotIf ; Reset hotkey context

; ==============================================================================
; TRAY MENU & HELPER FUNCTIONS
; ==============================================================================

SetupTrayMenu() {
    Tray := A_TrayMenu
    Tray.Delete()

    Tray.Add("Show Corridors", (*) => MainGui.Show())
    Tray.Add("Center && Reset Size", (*) => MainGui.Show("w520 h850 Center"))
    Tray.Add()
    Tray.Add("Reset Offset to 0`tCtrl+0", (*) => ResetWebOffset())
    Tray.Add("Reload Visualizer", (*) => MainGui.Reload())
    Tray.Add()
    Tray.Add("Keyboard Shortcuts...", (*) => ShowHelpBox())
    Tray.Add("Exit", (*) => CleanupAndExit())

    Tray.Default := "Show Corridors"
    A_IconTip := "Corridors Signal Visualizer"
}

ShowHelpBox() {
    MsgBox(
        "Active Window Hotkeys:`n`n"
        "Increase Offset:`n"
        "  • Ctrl + Shift + H  (+1 Hour)`n"
        "  • Ctrl + Shift + M  (+1 Minute)`n"
        "  • Ctrl + Shift + S  (+1 Second)`n`n"
        "Decrease Offset:`n"
        "  • Ctrl + Win + H    (-1 Hour)`n"
        "  • Ctrl + Win + M    (-1 Minute)`n"
        "  • Ctrl + Win + S    (-1 Second)`n`n"
        "Reset:`n"
        "  • Ctrl + 0 / Numpad0  (Zero all offsets)",
        "Corridors Keyboard Shortcuts",
        "Iconi Owner" MainGui.Hwnd
    )
}

AdjustWebOffset(unit, delta) {

    idMap := Map("hrs", "clkAdjHr", "min", "clkAdjMin", "sec", "clkAdjSec")
    targetId := idMap[unit]

    ; Clean single-line JS (avoids any AHK continuation-section formatting quirks)
    js := "(() => { const el = document.getElementById('" targetId "'); if (!el) return 'NOT_FOUND'; const step = parseFloat(el.step) || 1; const min = parseFloat(el.min); const max = parseFloat(el.max); let next = (parseFloat(el.value) || 0) + (" delta " * step); if (!isNaN(min)) next = Math.max(min, next); if (!isNaN(max)) next = Math.min(max, next); el.value = Math.round(next * 10) / 10; if (typeof Timing === 'function') Timing(el.value); el.dispatchEvent(new Event('change', { bubbles: true })); return el.value; })()"

    MainGui.ExecuteScript(js)
}

ResetWebOffset() {

    js := "(() => { ['clkAdjHr', 'clkAdjMin', 'clkAdjSec'].forEach(id => { const el = document.getElementById(id); if (el) { el.value = 0; if (typeof Timing === 'function') Timing(0); el.dispatchEvent(new Event('change', { bubbles: true })); } }); })()"

    MainGui.ExecuteScript(js)
}


CleanupAndExit() {
    try FileDelete(HtmlPath)
    ExitApp()
}
