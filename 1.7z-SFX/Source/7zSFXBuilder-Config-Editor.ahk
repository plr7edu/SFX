#Requires AutoHotkey v2.0
#SingleInstance Force

; ========================================
; TOGGLE TOOLTIP FUNCTION
; ========================================
global tooltipWindows := Map()

ToggleTooltip(flagType) {
    global tooltipWindows, mainGui, darkTheme

    ; Close tooltip if already open
    if (tooltipWindows.Has(flagType) && WinExist("ahk_id " . tooltipWindows[flagType].Hwnd)) {
        tooltipWindows[flagType].Destroy()
        tooltipWindows.Delete(flagType)
        return
    }

    ; Create tooltip content based on type
    if (flagType = "GUIFlags") {
        title := "GUI Flags - Custom/Additional Flags"
        content := "How to use:`n`n• Enter numeric values separated by + (plus sign)`n• Example: 32+64+128`n• Example: 1+4+512`n`nAvailable options:`n`n1 - Show extraction % in TitleBar (right side)`n2 - Don't display extraction % in TitleBar`n4 - Show extraction % under ProgressBar`n32 - Display icon in extraction progress window`n64 - Use 'BeginPrompt' window to specify extraction`n128 - Use separate window to specify extraction path`n256 - Confirm abolition of installation/extraction`n512 - Don't display icon in TitleBar of all windows`n1024 - Display icon in extraction path window`n4096 - Change button names (OK/Cancel instead of Yes/No)`n8192 - Don't show ProgressBar on TaskBar (Win7+)`n16384 - Show '&' symbol in texts`n`nMultiple flags example:`nTo combine flags 32, 64, and 128, enter: 32+64+128"
    } else {
        title := "Misc Flags - Custom/Additional Flags"
        content := "How to use:`n`n• Enter numeric values separated by + (plus sign)`n• Example: 1+2`n• Single flag example: 1`n`nAvailable options:`n`n1 - Don't verify free disk space needed for extraction`n2 - Don't verify free RAM needed for extraction`n`nMultiple flags example:`nTo combine flags 1 and 2, enter: 1+2"
    }

    ; Create dark-themed tooltip window
    tooltipGui := Gui("+AlwaysOnTop -SysMenu +Owner" . mainGui.Hwnd, title)
    tooltipGui.BackColor := "0x" . darkTheme["bgColor"]

    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        DWMWA_USE_IMMERSIVE_DARK_MODE := 19
        if (VerCompare(A_OSVersion, "10.0.18985") >= 0) {
            DWMWA_USE_IMMERSIVE_DARK_MODE := 20
        }
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", tooltipGui.Hwnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", 1, "Int", 4)
    }

    tooltipGui.SetFont("s10 bold", "Segoe UI")
    tooltipGui.AddText("x20 y20 w460 c" . darkTheme["textColor"], title)

    tooltipGui.SetFont("s9 norm", "Segoe UI")
    tooltipGui.AddText("x20 y50 w460 c" . darkTheme["textColor"], content)

    tooltipGui.SetFont("s10", "Segoe UI")
    closeBtn := tooltipGui.AddButton("x200 y+20 w100 h32", "Close")
    closeBtn.OnEvent("Click", (*) => (tooltipGui.Destroy(), tooltipWindows.Delete(flagType)))

    ApplyDarkMode(tooltipGui)

    ; Position tooltip near the main window
    tooltipGui.Show("w500 h" . (flagType = "GUIFlags" ? "550" : "350"))

    ; Store reference
    tooltipWindows[flagType] := tooltipGui
}

; ========================================
; CONFIGURATION
; ========================================
global CONFIG_FILE := A_ScriptDir . "\SFXConfig.ini"
global ICON_FILE := A_ScriptDir . "\Icon\7zsfxbuilder_config_editor.ico"
global ICON_PATH := A_ScriptDir . "\Icon\gui-icons\"

; ========================================
; GLOBAL VARIABLES
; ========================================
global mainGui
global ddlConfig, editFuzzySearch
global btnPin, btnNew, btnReset, btnSave, btnDelete
global editConfigName, editInstallPath, editTitle, editGUIMode
global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags
global chkMisc_4, chkMisc_8, editCustomMiscFlags
global chkSelfDelete, editBeginPrompt
global edit7zArchive, editSFXName, editSFXIcon
global ddlUseDefMod, editUPXCommands, editCopyright
global editExecuteFile, editExecuteParams
global configList := []
global configData := Map()
global isPinned := true
global isNewMode := false
global currentSection := ""

; Dark Theme Colors
global darkTheme := Map(
    "bgColor", "1E1E1E",
    "cardBg", "2D2D2D",
    "textColor", "FFFFFF",
    "textSecondary", "B0B0B0",
    "buttonBg", "3A3A3C",
    "buttonText", "FFFFFF",
    "buttonHover", "4A4A4C",
    "editBg", "3C3C3C",
    "editText", "FFFFFF",
    "success", "0E7A0D",
    "danger", "C42B1C",
    "warning", "FFC107",
    "scrollBg", "2D2D2D",
    "scrollThumb", "555555",
    "scrollThumbHover", "777777"
)

; Dark Mode Globals
global IsDarkMode := True
global DarkColors := Map("Background", 0x1E1E1E, "Controls", 0x2D2D2D, "Font", 0xFFFFFF)
global TextBackgroundBrush := 0
global WindowProcNew := 0
global WindowProcOld := 0

; ========================================
; DARK MODE FUNCTIONS
; ========================================
InitDarkMode() {
    global TextBackgroundBrush, DarkColors
    if (!TextBackgroundBrush)
        TextBackgroundBrush := DllCall("gdi32\CreateSolidBrush", "UInt", DarkColors["Background"], "Ptr")
}

WindowProc(hwnd, uMsg, wParam, lParam) {
    critical
    global IsDarkMode, DarkColors, TextBackgroundBrush, WindowProcOld
    static WM_CTLCOLOREDIT := 0x0133
    static WM_CTLCOLORLISTBOX := 0x0134
    static WM_CTLCOLORBTN := 0x0135
    static WM_CTLCOLORSTATIC := 0x0138
    static DC_BRUSH := 18

    if (IsDarkMode) {
        switch uMsg {
            case WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX:
                DllCall("gdi32\SetTextColor", "Ptr", wParam, "UInt", DarkColors["Font"])
                DllCall("gdi32\SetBkColor", "Ptr", wParam, "UInt", DarkColors["Controls"])
                DllCall("gdi32\SetDCBrushColor", "Ptr", wParam, "UInt", DarkColors["Controls"], "UInt")
                return DllCall("gdi32\GetStockObject", "Int", DC_BRUSH, "Ptr")
            case WM_CTLCOLORBTN:
                DllCall("gdi32\SetDCBrushColor", "Ptr", wParam, "UInt", DarkColors["Background"], "UInt")
                return DllCall("gdi32\GetStockObject", "Int", DC_BRUSH, "Ptr")
            case WM_CTLCOLORSTATIC:
                DllCall("gdi32\SetTextColor", "Ptr", wParam, "UInt", DarkColors["Font"])
                DllCall("gdi32\SetBkColor", "Ptr", wParam, "UInt", DarkColors["Background"])
                return TextBackgroundBrush
        }
    }
    return DllCall("user32\CallWindowProc", "Ptr", WindowProcOld, "Ptr", hwnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam)
}

SetWindowAttribute(GuiObj, DarkMode := True) {
    global DarkColors, TextBackgroundBrush
    static PreferredAppMode := Map("Default", 0, "AllowDark", 1, "ForceDark", 2, "ForceLight", 3, "Max", 4)

    if (!TextBackgroundBrush)
        TextBackgroundBrush := DllCall("gdi32\CreateSolidBrush", "UInt", DarkColors["Background"], "Ptr")

    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        DWMWA_USE_IMMERSIVE_DARK_MODE := 19
        if (VerCompare(A_OSVersion, "10.0.18985") >= 0) {
            DWMWA_USE_IMMERSIVE_DARK_MODE := 20
        }
        uxtheme := DllCall("kernel32\GetModuleHandle", "Str", "uxtheme", "Ptr")
        SetPreferredAppMode := DllCall("kernel32\GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
        FlushMenuThemes := DllCall("kernel32\GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")

        switch DarkMode {
            case True:
                DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", GuiObj.hWnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", True, "Int", 4)
                DllCall(SetPreferredAppMode, "Int", PreferredAppMode["ForceDark"])
                DllCall(FlushMenuThemes)
                GuiObj.BackColor := DarkColors["Background"]
            default:
                DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", GuiObj.hWnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", False, "Int", 4)
                DllCall(SetPreferredAppMode, "Int", PreferredAppMode["Default"])
                DllCall(FlushMenuThemes)
                GuiObj.BackColor := "Default"
        }
    }
}

SetWindowTheme(GuiObj, DarkMode := True) {
    static GWL_WNDPROC := -4
    static GWL_STYLE := -16
    static ES_MULTILINE := 0x0004
    static GetWindowLong := A_PtrSize = 8 ? "GetWindowLongPtr" : "GetWindowLong"
    static SetWindowLong := A_PtrSize = 8 ? "SetWindowLongPtr" : "SetWindowLong"
    static Init := False
    global IsDarkMode, WindowProcNew, WindowProcOld

    IsDarkMode := DarkMode
    Mode_Explorer := (DarkMode ? "DarkMode_Explorer" : "Explorer")
    Mode_CFD := (DarkMode ? "DarkMode_CFD" : "CFD")

    for hWnd, GuiCtrlObj in GuiObj {
        switch GuiCtrlObj.Type {
            case "Button":
                DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_Explorer, "Ptr", 0)
            case "CheckBox":
                DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_Explorer, "Ptr", 0)
                ; Force checkbox text color to white in dark mode
                if (DarkMode)
                    GuiCtrlObj.SetFont("cFFFFFF")
            case "UpDown":
                DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_Explorer, "Ptr", 0)
            case "ComboBox", "DDL":
                DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_CFD, "Ptr", 0)
                cbInfo := Buffer(40 + (3 * A_PtrSize), 0)
                NumPut("UInt", 40 + (3 * A_PtrSize), cbInfo, 0)
                if DllCall("user32\GetComboBoxInfo", "Ptr", GuiCtrlObj.hWnd, "Ptr", cbInfo) {
                    hwndList := NumGet(cbInfo, 40 + (2 * A_PtrSize), "Ptr")
                    if (hwndList) {
                        DllCall("uxtheme\SetWindowTheme", "Ptr", hwndList, "Str", Mode_Explorer, "Ptr", 0)
                    }
                }
            case "Edit":
                if (DllCall("user32\" GetWindowLong, "Ptr", GuiCtrlObj.hWnd, "Int", GWL_STYLE) & ES_MULTILINE) {
                    DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_Explorer, "Ptr", 0)
                } else {
                    DllCall("uxtheme\SetWindowTheme", "Ptr", GuiCtrlObj.hWnd, "Str", Mode_CFD, "Ptr", 0)
                }
                GuiCtrlObj.SetFont("cFFFFFF")
            case "Text":
                GuiCtrlObj.Opt("cFFFFFF")
        }
    }

    if !(Init) {
        WindowProcNew := CallbackCreate(WindowProc)
        WindowProcOld := DllCall("user32\" SetWindowLong, "Ptr", GuiObj.Hwnd, "Int", GWL_WNDPROC, "Ptr", WindowProcNew, "Ptr")
        Init := True
    }
}

ApplyDarkMode(GuiObj) {
    InitDarkMode()
    SetWindowAttribute(GuiObj, True)
    SetWindowTheme(GuiObj, True)
    DllCall("RedrawWindow", "Ptr", GuiObj.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0285)
    for hWnd, GuiCtrlObj in GuiObj {
        DllCall("InvalidateRect", "Ptr", GuiCtrlObj.Hwnd, "Ptr", 0, "Int", true)
    }
}

; ========================================
; DARK SCROLLBAR STYLING
; ========================================
ApplyDarkScrollbar(hwnd) {
    global darkTheme

    ; Enable custom scrollbar drawing
    static SIF_RANGE := 0x1
    static SIF_PAGE := 0x2
    static SIF_POS := 0x4
    static SIF_TRACKPOS := 0x10
    static SIF_ALL := SIF_RANGE | SIF_PAGE | SIF_POS | SIF_TRACKPOS

    ; Set dark theme for scrollbar
    DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
}

; ========================================
; SET BUTTON ICON
; ========================================
SetButtonIconIL(btn, iconPath, size := 24, marginLeft := 4, marginTop := 0, marginRight := 6, marginBottom := 0) {
    static BCM_SETIMAGELIST := 0x1602
    static ILC_COLOR32 := 0x20
    static ILC_MASK := 0x1
    if (!IsObject(btn) || iconPath = "" || !FileExist(iconPath))
        return
    hIcon := LoadPicture(iconPath, "Icon w" size " h" size, &imgType)
    if (!hIcon)
        return
    himl := DllCall("Comctl32\ImageList_Create", "Int", size, "Int", size, "UInt", ILC_COLOR32|ILC_MASK, "Int", 1, "Int", 1, "Ptr")
    if (!himl) {
        DllCall("user32\DestroyIcon", "Ptr", hIcon)
        return
    }
    DllCall("Comctl32\ImageList_AddIcon", "Ptr", himl, "Ptr", hIcon)
    DllCall("user32\DestroyIcon", "Ptr", hIcon)
    buf := Buffer(A_PtrSize + 16 + 4, 0)
    NumPut("Ptr", himl, buf, 0)
    NumPut("Int", marginLeft, buf, A_PtrSize + 0)
    NumPut("Int", marginTop, buf, A_PtrSize + 4)
    NumPut("Int", marginRight, buf, A_PtrSize + 8)
    NumPut("Int", marginBottom, buf, A_PtrSize + 12)
    NumPut("UInt", 0, buf, A_PtrSize + 16)
    DllCall("user32\SendMessage", "Ptr", btn.hWnd, "UInt", BCM_SETIMAGELIST, "Ptr", 0, "Ptr", buf.Ptr)
    try {
        if (btn.HasProp("hImgList") && btn.hImgList)
            DllCall("Comctl32\ImageList_Destroy", "Ptr", btn.hImgList)
    }
    btn.hImgList := himl
}

; ========================================
; FUNCTION DECLARATIONS (Must be before CreateMainGui)
; ========================================

; ========================================
; TOGGLE PIN/UNPIN
; ========================================
TogglePin(*) {
    global mainGui, btnPin, isPinned, ICON_PATH

    if isPinned {
        mainGui.Opt("-AlwaysOnTop")
        btnPin.Text := "Unpinned"
        if FileExist(ICON_PATH . "unpin.ico")
            SetButtonIconIL(btnPin, ICON_PATH . "unpin.ico", 24)
        isPinned := false
    } else {
        mainGui.Opt("+AlwaysOnTop")
        btnPin.Text := "Pinned"
        if FileExist(ICON_PATH . "pin.ico")
            SetButtonIconIL(btnPin, ICON_PATH . "pin.ico", 24)
        isPinned := true
    }
}

; ========================================
; NEW CONFIG
; ========================================
NewConfig(*) {
    global ddlConfig, editFuzzySearch, isNewMode, currentSection

    isNewMode := true
    currentSection := ""

    ; Disable dropdown and search
    ddlConfig.Enabled := false
    editFuzzySearch.Enabled := false

    ; Clear all fields
    ClearAllFields()
}

; ========================================
; RESET GUI
; ========================================
ResetGui(*) {
    global ddlConfig, editFuzzySearch, isNewMode, currentSection

    isNewMode := false
    currentSection := ""

    ; Enable dropdown and search
    ddlConfig.Enabled := true
    editFuzzySearch.Enabled := true

    ; Clear dropdown selection
    ddlConfig.Choose(0)
    editFuzzySearch.Value := ""

    ; Clear all fields
    ClearAllFields()
}

; ========================================
; SAVE CONFIG
; ========================================
SaveConfig(*) {
    global CONFIG_FILE, configList, configData, isNewMode, currentSection
    global editConfigName, editInstallPath, ddlConfig, editFuzzySearch

    ; Validate required fields
    configName := Trim(editConfigName.Value)
    if (configName = "") {
        ShowDarkMessage("Validation Error", "Configuration Name is required!", "error")
        return
    }

    installPath := Trim(editInstallPath.Value)
    if (installPath = "") {
        ShowDarkMessage("Validation Error", "Install Path is required!", "error")
        return
    }

    ; Check if renaming existing config
    if (!isNewMode && currentSection != "" && currentSection != configName) {
        ; Delete old section
        try {
            IniDelete(CONFIG_FILE, currentSection)
        }
    }

    ; Write all configuration values
    try {
        IniWrite(installPath, CONFIG_FILE, configName, "InstallPath")

        ; GUI Flags
        guiFlags := BuildGUIFlags()
        IniWrite(guiFlags, CONFIG_FILE, configName, "GUIFlags")

        ; Title (optional)
        titleVal := Trim(editTitle.Value)
        if (titleVal != "")
            IniWrite('"' . titleVal . '"', CONFIG_FILE, configName, "Title")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "Title")

        ; GUI Mode (optional)
        guiModeVal := Trim(editGUIMode.Value)
        if (guiModeVal != "")
            IniWrite('"' . guiModeVal . '"', CONFIG_FILE, configName, "GUIMode")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "GUIMode")

        ; Misc Flags (optional)
        miscFlags := BuildMiscFlags()
        if (miscFlags != '""')
            IniWrite(miscFlags, CONFIG_FILE, configName, "MiscFlags")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "MiscFlags")

        ; Self Delete (optional)
        if (chkSelfDelete.Value)
            IniWrite('"1"', CONFIG_FILE, configName, "SelfDelete")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "SelfDelete")

        ; Begin Prompt (optional, convert newlines to \n)
        promptVal := Trim(editBeginPrompt.Value)
        if (promptVal != "") {
            promptVal := StrReplace(promptVal, "`r`n", "\n")
            promptVal := StrReplace(promptVal, "`n", "\n")
            IniWrite('"' . promptVal . '"', CONFIG_FILE, configName, "BeginPrompt")
        } else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "BeginPrompt")

        ; 7z Archive
        IniWrite(Trim(edit7zArchive.Value), CONFIG_FILE, configName, "7zSFXBuilder_7zArchive")

        ; SFX Name
        IniWrite(Trim(editSFXName.Value), CONFIG_FILE, configName, "7zSFXBuilder_SFXName")

        ; SFX Icon
        iconVal := Trim(editSFXIcon.Value)
        if (iconVal != "")
            IniWrite(iconVal, CONFIG_FILE, configName, "7zSFXBuilder_SFXIcon")

        ; Use Default Module
        IniWrite(ddlUseDefMod.Text, CONFIG_FILE, configName, "7zSFXBuilder_UseDefMod")

        ; UPX Commands
        IniWrite(Trim(editUPXCommands.Value), CONFIG_FILE, configName, "7zSFXBuilder_UPXCommands")

        ; Legal Copyright
        IniWrite(Trim(editCopyright.Value), CONFIG_FILE, configName, "7zSFXBuilder_Res_LegalCopyright")

        ; Execute File (optional)
        execFileVal := Trim(editExecuteFile.Value)
        if (execFileVal != "")
            IniWrite('"' . execFileVal . '"', CONFIG_FILE, configName, "ExecuteFile")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "ExecuteFile")

        ; Execute Parameters (optional)
        execParamsVal := Trim(editExecuteParams.Value)
        if (execParamsVal != "")
            IniWrite('"' . execParamsVal . '"', CONFIG_FILE, configName, "ExecuteParameters")
        else if (!isNewMode)
            IniDelete(CONFIG_FILE, configName, "ExecuteParameters")

        ShowDarkMessage("Success", "Configuration '" . configName . "' saved successfully!", "success")

        ; Reload configurations and update dropdown
        LoadConfigurations()
        ddlConfig.Delete()
        ddlConfig.Add(configList)

        ; Find and select the saved config
        for index, name in configList {
            if (name = configName) {
                ddlConfig.Choose(index)
                break
            }
        }

        ; Re-enable dropdown if was in new mode
        if (isNewMode) {
            ddlConfig.Enabled := true
            editFuzzySearch.Enabled := true
            isNewMode := false
        }

        currentSection := configName

    } catch Error as e {
        ShowDarkMessage("Error", "Failed to save configuration!`n`n" . e.Message, "error")
    }
}

; ========================================
; DELETE CONFIG
; ========================================
DeleteConfig(*) {
    global CONFIG_FILE, configList, currentSection, ddlConfig

    if (currentSection = "") {
        ShowDarkMessage("No Selection", "Please select a configuration to delete.", "warning")
        return
    }

    ; Show confirmation dialog
    if (!ShowConfirmDialog("Delete Configuration", "Are you sure you want to permanently delete the configuration:`n`n'" . currentSection . "'?`n`nThis action cannot be undone.")) {
        return
    }

    try {
        ; Delete the section
        IniDelete(CONFIG_FILE, currentSection)

        ShowDarkMessage("Success", "Configuration '" . currentSection . "' deleted successfully!", "success")

        ; Reload configurations
        LoadConfigurations()
        ddlConfig.Delete()
        ddlConfig.Add(configList)

        ; Reset GUI
        ResetGui()

    } catch Error as e {
        ShowDarkMessage("Error", "Failed to delete configuration!`n`n" . e.Message, "error")
    }
}

; ========================================
; ON CONFIG CHANGE (DROPDOWN)
; ========================================
OnConfigChange(*) {
    global ddlConfig, configData, currentSection

    selectedConfig := ddlConfig.Text
    if (selectedConfig = "" || !configData.Has(selectedConfig))
        return

    currentSection := selectedConfig
    LoadConfigToFields(selectedConfig)
}

; ========================================
; FUZZY SEARCH
; ========================================
OnFuzzySearch(*) {
    global editFuzzySearch, ddlConfig, configList

    searchText := StrLower(editFuzzySearch.Value)
    if (searchText = "")
        return

    for index, configName in configList {
        if (InStr(StrLower(configName), searchText)) {
            ddlConfig.Choose(index)
            OnConfigChange()
            return
        }
    }
}

; ========================================
; BROWSE FOLDER
; ========================================
BrowseFolder(editField) {
    global mainGui, isPinned

    if isPinned
        mainGui.Opt("-AlwaysOnTop")

    folderPath := DirSelect(, 3, "Select Installation Folder")

    if isPinned
        mainGui.Opt("+AlwaysOnTop")

    if (folderPath != "")
        editField.Value := folderPath
}

; ========================================
; BROWSE FILE
; ========================================
BrowseFile(editField, filterName, saveMode := False) {
    global mainGui, isPinned

    if isPinned
        mainGui.Opt("-AlwaysOnTop")

    mode := saveMode ? 16 : 1  ; 16 = prompt to create, 1 = file must exist
    filePath := FileSelect(mode, , "Select " . filterName, filterName)

    if isPinned
        mainGui.Opt("+AlwaysOnTop")

    if (filePath != "")
        editField.Value := filePath
}

; ========================================
; CLEAR ALL FIELDS
; ========================================
ClearAllFields() {
    global editConfigName, editInstallPath, editTitle, editGUIMode
    global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags
    global chkMisc_4, chkMisc_8, editCustomMiscFlags
    global chkSelfDelete, editBeginPrompt
    global edit7zArchive, editSFXName, editSFXIcon
    global ddlUseDefMod, editUPXCommands, editCopyright
    global editExecuteFile, editExecuteParams

    editConfigName.Value := ""
    editInstallPath.Value := ""
    editTitle.Value := ""
    editGUIMode.Value := ""

    chkGUI_8.Value := 0
    chkGUI_16.Value := 0
    chkGUI_2048.Value := 0
    editCustomGUIFlags.Value := ""

    chkMisc_4.Value := 0
    chkMisc_8.Value := 0
    editCustomMiscFlags.Value := ""

    chkSelfDelete.Value := 0
    editBeginPrompt.Value := ""

    edit7zArchive.Value := ""
    editSFXName.Value := ""
    editSFXIcon.Value := ""

    ddlUseDefMod.Choose(1)
    editUPXCommands.Value := "--best --all-methods"
    editCopyright.Value := "Copyright © 2026 PL-DEV"

    editExecuteFile.Value := ""
    editExecuteParams.Value := ""
}

; ========================================
; LOAD CONFIG TO FIELDS
; ========================================
LoadConfigToFields(sectionName) {
    global configData
    global editConfigName, editInstallPath, editTitle, editGUIMode
    global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags
    global chkMisc_4, chkMisc_8, editCustomMiscFlags
    global chkSelfDelete, editBeginPrompt
    global edit7zArchive, editSFXName, editSFXIcon
    global ddlUseDefMod, editUPXCommands, editCopyright
    global editExecuteFile, editExecuteParams

    if (!configData.Has(sectionName))
        return

    data := configData[sectionName]

    ; Configuration Name
    editConfigName.Value := sectionName

    ; Install Path
    editInstallPath.Value := GetValue(data, "InstallPath")

    ; Title
    editTitle.Value := GetValue(data, "Title")

    ; GUI Mode
    editGUIMode.Value := GetValue(data, "GUIMode")

    ; Parse GUI Flags
    ParseGUIFlags(GetValue(data, "GUIFlags"))

    ; Parse Misc Flags
    ParseMiscFlags(GetValue(data, "MiscFlags"))

    ; Self Delete
    chkSelfDelete.Value := (GetValue(data, "SelfDelete") = "1") ? 1 : 0

    ; Begin Prompt (convert \n to actual newlines)
    promptText := GetValue(data, "BeginPrompt")
    promptText := StrReplace(promptText, "\n", "`n")
    editBeginPrompt.Value := promptText

    ; 7z Archive
    edit7zArchive.Value := GetValue(data, "7zSFXBuilder_7zArchive")

    ; SFX Name
    editSFXName.Value := GetValue(data, "7zSFXBuilder_SFXName")

    ; SFX Icon
    editSFXIcon.Value := GetValue(data, "7zSFXBuilder_SFXIcon")

    ; Use Default Module
    modValue := GetValue(data, "7zSFXBuilder_UseDefMod")
    if (modValue != "") {
        switch modValue {
            case "7zsd_LZMA2_x64": ddlUseDefMod.Choose(1)
            case "7zsd_LZMA_x64": ddlUseDefMod.Choose(2)
            case "7zsd_LZMA2_x86": ddlUseDefMod.Choose(3)
            case "7zsd_LZMA_x86": ddlUseDefMod.Choose(4)
            default: ddlUseDefMod.Choose(1)
        }
    }

    ; UPX Commands
    upxValue := GetValue(data, "7zSFXBuilder_UPXCommands")
    editUPXCommands.Value := (upxValue != "") ? upxValue : "--best --all-methods"

    ; Legal Copyright
    copyrightValue := GetValue(data, "7zSFXBuilder_Res_LegalCopyright")
    editCopyright.Value := (copyrightValue != "") ? copyrightValue : "Copyright © 2026 PL-DEV"

    ; Execute File
    editExecuteFile.Value := GetValue(data, "ExecuteFile")

    ; Execute Parameters
    editExecuteParams.Value := GetValue(data, "ExecuteParameters")
}

; ========================================
; PARSE GUI FLAGS
; ========================================
ParseGUIFlags(flagString) {
    global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags

    ; Clear all
    chkGUI_8.Value := 0
    chkGUI_16.Value := 0
    chkGUI_2048.Value := 0
    editCustomGUIFlags.Value := ""

    ; Remove quotes
    flagString := Trim(flagString, '"')
    if (flagString = "")
        return

    ; Parse known flags
    if InStr(flagString, "8")
        chkGUI_8.Value := 1
    if InStr(flagString, "16")
        chkGUI_16.Value := 1
    if InStr(flagString, "2048")
        chkGUI_2048.Value := 1

    ; Detect custom flags
    customFlags := ""
    for flag in StrSplit(flagString, "+") {
        flag := Trim(flag)
        if (flag != "8" && flag != "16" && flag != "2048" && flag != "") {
            customFlags .= flag . "+"
        }
    }
    editCustomGUIFlags.Value := RTrim(customFlags, "+")
}

; ========================================
; PARSE MISC FLAGS
; ========================================
ParseMiscFlags(flagString) {
    global chkMisc_4, chkMisc_8, editCustomMiscFlags

    ; Clear all
    chkMisc_4.Value := 0
    chkMisc_8.Value := 0
    editCustomMiscFlags.Value := ""

    ; Remove quotes
    flagString := Trim(flagString, '"')
    if (flagString = "")
        return

    ; Parse known flags
    if InStr(flagString, "4")
        chkMisc_4.Value := 1
    if InStr(flagString, "8")
        chkMisc_8.Value := 1

    ; Detect custom flags
    customFlags := ""
    for flag in StrSplit(flagString, "+") {
        flag := Trim(flag)
        if (flag != "4" && flag != "8" && flag != "") {
            customFlags .= flag . "+"
        }
    }
    editCustomMiscFlags.Value := RTrim(customFlags, "+")
}

; ========================================
; BUILD GUI FLAGS STRING
; ========================================
BuildGUIFlags() {
    global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags

    result := ""
    if chkGUI_8.Value
        result .= "8+"
    if chkGUI_16.Value
        result .= "16+"
    if chkGUI_2048.Value
        result .= "2048+"

    ; Add custom flags
    customFlags := Trim(editCustomGUIFlags.Value)
    if (customFlags != "")
        result .= customFlags . "+"

    result := RTrim(result, "+")
    return (result != "") ? '"' . result . '"' : '""'
}

; ========================================
; BUILD MISC FLAGS STRING
; ========================================
BuildMiscFlags() {
    global chkMisc_4, chkMisc_8, editCustomMiscFlags

    result := ""
    if chkMisc_4.Value
        result .= "4+"
    if chkMisc_8.Value
        result .= "8+"

    ; Add custom flags
    customFlags := Trim(editCustomMiscFlags.Value)
    if (customFlags != "")
        result .= customFlags . "+"

    result := RTrim(result, "+")
    return (result != "") ? '"' . result . '"' : '""'
}

; ========================================
; GET VALUE HELPER
; ========================================
GetValue(dataMap, key) {
    if (dataMap.Has(key)) {
        value := dataMap[key]
        ; Remove quotes if present
        value := Trim(value, '"')
        return value
    }
    return ""
}

; ========================================
; DARK MESSAGE BOX
; ========================================
ShowDarkMessage(title, message, iconType := "info") {
    global mainGui, darkTheme

    dialogWidth := 450
    dialogHeight := 200

    msgGui := Gui("+AlwaysOnTop -SysMenu +Owner" . mainGui.Hwnd, title)
    msgGui.BackColor := "0x" . darkTheme["bgColor"]

    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        DWMWA_USE_IMMERSIVE_DARK_MODE := 19
        if (VerCompare(A_OSVersion, "10.0.18985") >= 0) {
            DWMWA_USE_IMMERSIVE_DARK_MODE := 20
        }
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", msgGui.Hwnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", 1, "Int", 4)
    }

    msgGui.SetFont("s10", "Segoe UI")

    iconSymbol := "ℹ"
    iconColor := darkTheme["textColor"]
    switch iconType {
        case "error":
            iconSymbol := "✖"
            iconColor := darkTheme["danger"]
        case "success":
            iconSymbol := "✓"
            iconColor := darkTheme["success"]
        case "warning":
            iconSymbol := "⚠"
            iconColor := darkTheme["warning"]
    }

    msgGui.SetFont("s32 bold")
    msgGui.AddText("x20 y20 w60 h60 Center c" . iconColor, iconSymbol)

    msgGui.SetFont("s11 bold", "Segoe UI")
    msgGui.AddText("x90 y25 w340 c" . darkTheme["textColor"], title)

    msgGui.SetFont("s9 norm", "Segoe UI")
    msgGui.AddText("x90 y55 w340 h90 c" . darkTheme["textColor"], message)

    msgGui.SetFont("s10", "Segoe UI")
    okBtn := msgGui.AddButton("x185 y155 w80 h32", "OK")
    okBtn.OnEvent("Click", (*) => msgGui.Destroy())

    ApplyDarkMode(msgGui)
    msgGui.Show("w" . dialogWidth . " h" . dialogHeight)
    WinWaitClose(msgGui.Hwnd)
}

; ========================================
; DARK CONFIRMATION DIALOG
; ========================================
ShowConfirmDialog(title, message) {
    global mainGui, darkTheme

    dialogWidth := 450
    dialogHeight := 200
    result := false

    confirmGui := Gui("+AlwaysOnTop -SysMenu +Owner" . mainGui.Hwnd, title)
    confirmGui.BackColor := "0x" . darkTheme["bgColor"]

    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        DWMWA_USE_IMMERSIVE_DARK_MODE := 19
        if (VerCompare(A_OSVersion, "10.0.18985") >= 0) {
            DWMWA_USE_IMMERSIVE_DARK_MODE := 20
        }
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", confirmGui.Hwnd, "Int", DWMWA_USE_IMMERSIVE_DARK_MODE, "Int*", 1, "Int", 4)
    }

    confirmGui.SetFont("s10", "Segoe UI")

    confirmGui.SetFont("s32 bold")
    confirmGui.AddText("x20 y20 w60 h60 Center c" . darkTheme["warning"], "⚠")

    confirmGui.SetFont("s11 bold", "Segoe UI")
    confirmGui.AddText("x90 y25 w340 c" . darkTheme["textColor"], title)

    confirmGui.SetFont("s9 norm", "Segoe UI")
    confirmGui.AddText("x90 y55 w340 h80 c" . darkTheme["textColor"], message)

    confirmGui.SetFont("s10", "Segoe UI")
    yesBtn := confirmGui.AddButton("x145 y150 w70 h35", "Yes")
    noBtn := confirmGui.AddButton("x235 y150 w70 h35", "No")

    yesBtn.OnEvent("Click", (*) => (result := true, confirmGui.Destroy()))
    noBtn.OnEvent("Click", (*) => (result := false, confirmGui.Destroy()))

    ApplyDarkMode(confirmGui)
    confirmGui.Show("w" . dialogWidth . " h" . dialogHeight)
    WinWaitClose(confirmGui.Hwnd)

    return result
}

; ========================================
; INITIALIZE
; ========================================
Initialize()

Initialize() {
    SetupDarkTrayMenu()
    if (!FileExist(CONFIG_FILE)) {
        MsgBox("Configuration file not found!`n`nExpected: " . CONFIG_FILE, "Error", "Icon!")
        ExitApp
    }
    LoadConfigurations()
    CreateMainGui()
}

; ========================================
; SETUP DARK TRAY MENU
; ========================================
SetupDarkTrayMenu() {
    if (FileExist(ICON_FILE))
        TraySetIcon(ICON_FILE)

    A_TrayMenu.Delete()
    A_TrayMenu.Add("Show", (*) => (mainGui.Show(), mainGui.Flash()))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Reload Script", (*) => Reload())
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    A_TrayMenu.Default := "Show"

    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        try {
            uxtheme := DllCall("GetModuleHandle", "Str", "uxtheme", "Ptr")
            SetPreferredAppMode := DllCall("GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
            FlushMenuThemes := DllCall("GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")
            if (SetPreferredAppMode)
                DllCall(SetPreferredAppMode, "Int", 2)
            if (FlushMenuThemes)
                DllCall(FlushMenuThemes)
        }
    }
}

; ========================================
; LOAD CONFIGURATIONS
; ========================================
LoadConfigurations() {
    global configList, configData, CONFIG_FILE
    configList := []
    configData := Map()

    sections := IniRead(CONFIG_FILE)
    if (sections = "")
        return

    sectionLines := StrSplit(sections, "`n", "`r")
    for index, sectionName in sectionLines {
        if (sectionName != "") {
            configList.Push(sectionName)
            configData[sectionName] := Map()

            sectionContent := IniRead(CONFIG_FILE, sectionName)
            lines := StrSplit(sectionContent, "`n", "`r")
            for line in lines {
                if (InStr(line, "=")) {
                    parts := StrSplit(line, "=", , 2)
                    key := Trim(parts[1])
                    value := Trim(parts[2])
                    configData[sectionName][key] := value
                }
            }
        }
    }
}

; ========================================
; CREATE MAIN GUI
; ========================================
CreateMainGui() {
    global mainGui, ddlConfig, editFuzzySearch, darkTheme
    global btnPin, btnNew, btnReset, btnSave, btnDelete
    global editConfigName, editInstallPath, editTitle, editGUIMode
    global chkGUI_8, chkGUI_16, chkGUI_2048, editCustomGUIFlags
    global chkMisc_4, chkMisc_8, editCustomMiscFlags
    global chkSelfDelete, editBeginPrompt
    global edit7zArchive, editSFXName, editSFXIcon
    global ddlUseDefMod, editUPXCommands, editCopyright
    global editExecuteFile, editExecuteParams
    global configList, ICON_PATH

    mainGui := Gui("+AlwaysOnTop +DPIScale", "All-in-One 7z SFX Builder Config Editor")
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.BackColor := "0x" . darkTheme["bgColor"]

    ; Set GUI icon
    if (FileExist(ICON_FILE)) {
        try {
            mainGui.Show("Hide")
            hIcon := LoadPicture(ICON_FILE, "Icon1 w32 h32", &imgType)
            if (hIcon) {
                SendMessage(0x80, 1, hIcon, , mainGui.Hwnd)
                SendMessage(0x80, 0, hIcon, , mainGui.Hwnd)
            }
        }
    }

    ; Top Buttons Row
    btnPin := mainGui.AddButton("x20 y20 w110 h35", "Pinned")
    if FileExist(ICON_PATH . "pin.ico")
        SetButtonIconIL(btnPin, ICON_PATH . "pin.ico", 24)
    btnPin.OnEvent("Click", TogglePin)

    btnNew := mainGui.AddButton("x145 y20 w110 h35", "New")
    if FileExist(ICON_PATH . "new.ico")
        SetButtonIconIL(btnNew, ICON_PATH . "new.ico", 24)
    btnNew.OnEvent("Click", NewConfig)

    btnReset := mainGui.AddButton("x270 y20 w110 h35", "Reset")
    if FileExist(ICON_PATH . "reset.ico")
        SetButtonIconIL(btnReset, ICON_PATH . "reset.ico", 24)
    btnReset.OnEvent("Click", ResetGui)

    btnSave := mainGui.AddButton("x395 y20 w110 h35", "Save")
    if FileExist(ICON_PATH . "save.ico")
        SetButtonIconIL(btnSave, ICON_PATH . "save.ico", 24)
    btnSave.OnEvent("Click", SaveConfig)

    btnDelete := mainGui.AddButton("x520 y20 w110 h35", "Delete")
    if FileExist(ICON_PATH . "delete.ico")
        SetButtonIconIL(btnDelete, ICON_PATH . "delete.ico", 24)
    btnDelete.OnEvent("Click", DeleteConfig)

    ; Dropdown and Search
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y75 c" . darkTheme["textSecondary"], "Select Configuration:")
    mainGui.SetFont("s14 norm")
    ddlConfig := mainGui.Add("DropDownList", "x20 y95 w610 R12", configList)
    ddlConfig.OnEvent("Change", OnConfigChange)

    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y140 c" . darkTheme["textSecondary"], "Quick Search (Fuzzy):")
    mainGui.SetFont("s12")
    editFuzzySearch := mainGui.AddEdit("x20 y160 w610 h35 -VScroll")
    editFuzzySearch.OnEvent("Change", OnFuzzySearch)

    ; Scrollable area
    mainGui.SetFont("s10 norm")

    ; Configuration Name
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y215 c" . darkTheme["textColor"], "Configuration Name:")
    mainGui.SetFont("s10 norm")
    editConfigName := mainGui.AddEdit("x20 y240 w610 h30 -VScroll")
    editConfigName.ToolTip := "The section name in SFXConfig.ini (e.g., '3_APT_Setup')"

    ; Install Path
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y285 c" . darkTheme["textColor"], "Install Path:")
    mainGui.SetFont("s10 norm")
    editInstallPath := mainGui.AddEdit("x20 y310 w550 h30 -VScroll")
    editInstallPath.ToolTip := "Extraction destination path`nExamples: C:\#PL-DEV, %APPDATA%\Nilesoft Shell\, %%S (current path)"
    btnBrowseInstall := mainGui.AddButton("x580 y310 w50 h30", "...")
    btnBrowseInstall.OnEvent("Click", (*) => BrowseFolder(editInstallPath))

    ; Title
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y355 c" . darkTheme["textColor"], "Title (Optional):")
    mainGui.SetFont("s10 norm")
    editTitle := mainGui.AddEdit("x20 y380 w610 h30 -VScroll")
    editTitle.ToolTip := "Custom title for the SFX extraction window"

    ; GUI Mode
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y425 c" . darkTheme["textColor"], "GUI Mode (Optional):")
    mainGui.SetFont("s10 norm")
    editGUIMode := mainGui.AddEdit("x20 y450 w610 h30 -VScroll")
    editGUIMode.ToolTip := "GUI display mode (usually '2' for silent mode)"

    ; GUI Flags GroupBox
    mainGui.SetFont("s11 bold")
    mainGui.AddGroupBox("x20 y495 w610 h180 c" . darkTheme["textColor"], "GUI Flags")
    mainGui.SetFont("s10 norm")

    ; Get checkbox dimensions for proper alignment
    SGW := SysGet(71)  ; SM_CXMENUCHECK
    SGH := SysGet(72)  ; SM_CYMENUCHECK

    chkGUI_8 := mainGui.AddCheckbox("x40 y520 h" . SGH . " w" . SGW)
    chkGUI_8.ToolTip := "Use Windows XP styles (scheme)`nApplies modern visual styles to the extraction dialog`nMost commonly used flag (appears in almost all configs)"
    mainGui.AddText("x+5 yp cFFFFFF h" . SGH, "8 - Use Windows XP styles (scheme)")

    chkGUI_16 := mainGui.AddCheckbox("x40 y550 h" . SGH . " w" . SGW)
    chkGUI_16.ToolTip := "Use bold font for extraction percentage`nMakes the progress percentage text bold and easier to read`nCommonly combined with flag 8"
    mainGui.AddText("x+5 yp cFFFFFF h" . SGH, "16 - Use bold font for extraction percentage")

    chkGUI_2048 := mainGui.AddCheckbox("x40 y580 h" . SGH . " w" . SGW)
    chkGUI_2048.ToolTip := "Display icon in 'BeginPrompt' window`nShows custom icon in password/warning prompt dialogs`nUsed primarily in encrypted SFX archives"
    mainGui.AddText("x+5 yp cFFFFFF h" . SGH, "2048 - Display icon in 'BeginPrompt' window")

    mainGui.AddText("x40 y615 cFFFFFF", "Custom/Additional Flags:")
    editCustomGUIFlags := mainGui.AddEdit("x40 y635 w510 h25 -VScroll")
    btnGUIFlagsInfo := mainGui.AddButton("x560 y635 w40 h25", "ℹ️")
    btnGUIFlagsInfo.OnEvent("Click", (*) => ToggleTooltip("GUIFlags"))

    ; Misc Flags GroupBox
    mainGui.SetFont("s11 bold")
    mainGui.AddGroupBox("x20 y690 w610 h140 c" . darkTheme["textColor"], "Misc Flags")
    mainGui.SetFont("s10 norm")

    chkMisc_4 := mainGui.AddCheckbox("x40 y715 h" . SGH . " w" . SGW)
    chkMisc_4.ToolTip := "Require administrator privileges`nForces the SFX to run with admin rights`nNeeded for installing to system directories (Program Files, etc.)"
    mainGui.AddText("x+5 yp cFFFFFF 0x200 h" . SGH, "4 - Require administrator privileges")

    chkMisc_8 := mainGui.AddCheckbox("x40 y745 h" . SGH . " w" . SGW)
    chkMisc_8.ToolTip := "Show password request window`nDisplays password prompt after BeginPrompt and ExtractPath dialogs`nUsed for encrypted archives"
    mainGui.AddText("x+5 yp cFFFFFF 0x200 h" . SGH, "8 - Show password request window")

    mainGui.AddText("x40 y775 cFFFFFF", "Custom/Additional Flags:")
    editCustomMiscFlags := mainGui.AddEdit("x40 y795 w510 h25 -VScroll")
    btnMiscFlagsInfo := mainGui.AddButton("x560 y795 w40 h25", "ℹ️")
    btnMiscFlagsInfo.OnEvent("Click", (*) => ToggleTooltip("MiscFlags"))

    ; Self Delete
    chkSelfDelete := mainGui.AddCheckbox("x20 y845 h" . SGH . " w" . SGW)
    chkSelfDelete.ToolTip := "Automatically delete the SFX executable after extraction completes"
    mainGui.AddText("x+5 yp cFFFFFF 0x200 h" . SGH, "Self Delete (Delete SFX after extraction)")

    ; Begin Prompt
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y880 c" . darkTheme["textColor"], "Begin Prompt (Optional):")
    mainGui.SetFont("s10 norm")
    editBeginPrompt := mainGui.AddEdit("x20 y905 w610 h80 Multi")
    editBeginPrompt.ToolTip := "Warning message shown before extraction`nUsed for encrypted archives or important notices`nUse \n for line breaks"

    ; 7z Archive
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1000 c" . darkTheme["textColor"], "7z Archive:")
    mainGui.SetFont("s10 norm")
    edit7zArchive := mainGui.AddEdit("x20 y1025 w550 h30 -VScroll")
    edit7zArchive.ToolTip := "Path to the .7z archive file to be converted to SFX"
    btnBrowse7z := mainGui.AddButton("x580 y1025 w50 h30", "...")
    btnBrowse7z.OnEvent("Click", (*) => BrowseFile(edit7zArchive, "7z Archive (*.7z)"))

    ; SFX Name
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1070 c" . darkTheme["textColor"], "SFX Name (Output .exe):")
    mainGui.SetFont("s10 norm")
    editSFXName := mainGui.AddEdit("x20 y1095 w550 h30 -VScroll")
    editSFXName.ToolTip := "Path and name for the output SFX executable file"
    btnBrowseSFX := mainGui.AddButton("x580 y1095 w50 h30", "...")
    btnBrowseSFX.OnEvent("Click", (*) => BrowseFile(editSFXName, "Executable (*.exe)", True))

    ; SFX Icon
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1140 c" . darkTheme["textColor"], "SFX Icon:")
    mainGui.SetFont("s10 norm")
    editSFXIcon := mainGui.AddEdit("x20 y1165 w550 h30 -VScroll")
    editSFXIcon.ToolTip := "Path to the .ico file for the SFX executable icon"
    btnBrowseIcon := mainGui.AddButton("x580 y1165 w50 h30", "...")
    btnBrowseIcon.OnEvent("Click", (*) => BrowseFile(editSFXIcon, "Icon Files (*.ico)"))

    ; Use Default Module
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1210 c" . darkTheme["textColor"], "Use Default Module:")
    mainGui.SetFont("s10 norm")
    ddlUseDefMod := mainGui.AddDropDownList("x20 y1235 w610", ["7zsd_LZMA2_x64", "7zsd_LZMA_x64", "7zsd_LZMA2_x86", "7zsd_LZMA_x86"])
    ddlUseDefMod.Choose(1)
    ddlUseDefMod.ToolTip := "7z SFX module to use (x64 for 64-bit, x86 for 32-bit)"

    ; UPX Commands
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1280 c" . darkTheme["textColor"], "UPX Commands:")
    mainGui.SetFont("s10 norm")
    editUPXCommands := mainGui.AddEdit("x20 y1305 w610 h30 -VScroll")
    editUPXCommands.Value := "--best --all-methods"
    editUPXCommands.ToolTip := "UPX compression parameters for the final SFX executable"

    ; Execute File
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1350 c" . darkTheme["textColor"], "Execute File (Optional):")
    mainGui.SetFont("s10 norm")
    editExecuteFile := mainGui.AddEdit("x20 y1375 w550 h30 -VScroll")
    editExecuteFile.ToolTip := "Program to run after extraction completes"
    btnBrowseExec := mainGui.AddButton("x580 y1375 w50 h30", "...")
    btnBrowseExec.OnEvent("Click", (*) => BrowseFile(editExecuteFile, "Executable (*.exe;*.bat;*.cmd;*.vbs)"))

    ; Execute Parameters
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1420 c" . darkTheme["textColor"], "Execute Parameters (Optional):")
    mainGui.SetFont("s10 norm")
    editExecuteParams := mainGui.AddEdit("x20 y1445 w610 h30 -VScroll")
    editExecuteParams.ToolTip := "Command-line parameters for the executed program"

    ; Legal Copyright
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y1490 c" . darkTheme["textColor"], "Legal Copyright:")
    mainGui.SetFont("s10 norm")
    editCopyright := mainGui.AddEdit("x20 y1515 w610 h30 -VScroll")
    editCopyright.Value := "Copyright © 2026 PL-DEV"
    editCopyright.ToolTip := "Copyright notice embedded in the SFX executable"

    ; Apply dark mode
    ApplyDarkMode(mainGui)

    ; Add scrollbar support and apply dark scrollbar
    DllCall("SetWindowLong", "Ptr", mainGui.Hwnd, "Int", -16, "Int", DllCall("GetWindowLong", "Ptr", mainGui.Hwnd, "Int", -16) | 0x00200000)  ; WS_VSCROLL

    ; Apply dark theme to scrollbar
    ApplyDarkScrollbar(mainGui.Hwnd)

    ; Calculate center position
    guiWidth := 650
    guiHeight := 800
    screenWidth := A_ScreenWidth
    screenHeight := A_ScreenHeight
    xPos := (screenWidth - guiWidth) // 2
    yPos := (screenHeight - guiHeight) // 2

    mainGui.Show("w" . guiWidth . " h" . guiHeight . " x" . xPos . " y" . yPos)

    ; Set up scrollbar
    SCROLLINFO := Buffer(28, 0)
    NumPut("UInt", 28, SCROLLINFO, 0)  ; cbSize
    NumPut("UInt", 0x17, SCROLLINFO, 4)  ; fMask (SIF_RANGE | SIF_PAGE | SIF_DISABLENOSCROLL)
    NumPut("Int", 0, SCROLLINFO, 8)  ; nMin
    NumPut("Int", 1570, SCROLLINFO, 12)  ; nMax (total content height)
    NumPut("UInt", 800, SCROLLINFO, 16)  ; nPage (visible height)
    DllCall("SetScrollInfo", "Ptr", mainGui.Hwnd, "Int", 1, "Ptr", SCROLLINFO, "Int", 1)  ; SB_VERT = 1

    ; Handle scroll events
    OnMessage(0x115, Gui_ScrollEvent)  ; WM_VSCROLL
    OnMessage(0x020A, Gui_MouseWheelEvent)  ; WM_MOUSEWHEEL - Mouse wheel support
}

; ========================================
; SCROLL EVENT HANDLER
; ========================================
Gui_ScrollEvent(wParam, lParam, msg, hwnd) {
    global mainGui
    if (hwnd != mainGui.Hwnd)
        return

    static SB_LINEUP := 0
    static SB_LINEDOWN := 1
    static SB_PAGEUP := 2
    static SB_PAGEDOWN := 3
    static SB_THUMBTRACK := 5
    static SB_THUMBPOSITION := 4
    static SB_TOP := 6
    static SB_BOTTOM := 7

    scrollCode := wParam & 0xFFFF

    ; Get current scroll position
    SCROLLINFO := Buffer(28, 0)
    NumPut("UInt", 28, SCROLLINFO, 0)
    NumPut("UInt", 0x17, SCROLLINFO, 4)  ; fMask
    DllCall("GetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", SCROLLINFO)

    currentPos := NumGet(SCROLLINFO, 20, "Int")  ; nPos
    newPos := currentPos

    switch scrollCode {
        case SB_LINEUP:
            newPos := currentPos - 30
        case SB_LINEDOWN:
            newPos := currentPos + 30
        case SB_PAGEUP:
            newPos := currentPos - 100
        case SB_PAGEDOWN:
            newPos := currentPos + 100
        case SB_THUMBTRACK, SB_THUMBPOSITION:
            newPos := (wParam >> 16) & 0xFFFF
        case SB_TOP:
            newPos := 0
        case SB_BOTTOM:
            newPos := 1570
    }

    ; Clamp position
    if (newPos < 0)
        newPos := 0
    if (newPos > 770)  ; Max scroll = nMax - nPage
        newPos := 770

    ; Update scroll position
    NumPut("Int", newPos, SCROLLINFO, 20)
    DllCall("SetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", SCROLLINFO, "Int", 1)

    ; Scroll the window content
    DllCall("ScrollWindow", "Ptr", hwnd, "Int", 0, "Int", currentPos - newPos, "Ptr", 0, "Ptr", 0)
}

; ========================================
; MOUSE WHEEL EVENT HANDLER
; ========================================
Gui_MouseWheelEvent(wParam, lParam, msg, hwnd) {
    global mainGui
    if (hwnd != mainGui.Hwnd)
        return

    ; Extract wheel delta (positive = scroll up, negative = scroll down)
    wheelDelta := (wParam >> 16) & 0xFFFF
    ; Convert to signed value
    if (wheelDelta > 0x7FFF)
        wheelDelta := wheelDelta - 0x10000

    ; Get current scroll position
    SCROLLINFO := Buffer(28, 0)
    NumPut("UInt", 28, SCROLLINFO, 0)
    NumPut("UInt", 0x17, SCROLLINFO, 4)  ; fMask
    DllCall("GetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", SCROLLINFO)

    currentPos := NumGet(SCROLLINFO, 20, "Int")  ; nPos

    ; Calculate new position (scroll 3 lines per wheel notch)
    scrollAmount := 90  ; 3 lines * 30 pixels per line
    newPos := currentPos - (wheelDelta > 0 ? scrollAmount : -scrollAmount)

    ; Clamp position
    if (newPos < 0)
        newPos := 0
    if (newPos > 770)  ; Max scroll = nMax - nPage
        newPos := 770

    ; Update scroll position
    NumPut("Int", newPos, SCROLLINFO, 20)
    DllCall("SetScrollInfo", "Ptr", hwnd, "Int", 1, "Ptr", SCROLLINFO, "Int", 1)

    ; Scroll the window content
    DllCall("ScrollWindow", "Ptr", hwnd, "Int", 0, "Int", currentPos - newPos, "Ptr", 0, "Ptr", 0)

    return 0  ; Prevent default handling
}