#Requires AutoHotkey v2.0
#SingleInstance Force

; ========================================
; CONFIGURATION
; ========================================
global LOG_FILE := A_ScriptDir . "\Logs\SFXHelper.log"
global CONFIG_FILE := A_ScriptDir . "\SFXConfig.ini"
global SFX_BUILDER_EXE := "C:\Program Files (x86)\7z SFX Builder\7z SFX Builder.exe"
global CONFIG_EDITOR_EXE := A_ScriptDir . "\7zSFXBuilder-Config-Editor.exe"  ; Add this line
global ICON_FILE := A_ScriptDir . "\Icon\7zsfxbuilder_healper.ico"
global ICON_PATH := A_ScriptDir . "\Icon\gui-icons\"

; ========================================
; GLOBAL VARIABLES
; ========================================
global mainGui
global dropZoneText, fileNameText, btnPrep, btnPin, btnAdd, btnReset, btnAdvance
global ddlConfig, editFuzzySearch
global droppedFilePath := ""
global selectedConfig := ""
global isPinned := true
global configList := []
global configData := Map()

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
    "dropZone", "2D2D2D",
    "dropZoneBorder", "4A4A4C"
)

; ========================================
; DARK MODE GLOBALS
; ========================================
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
    static WM_CTLCOLOREDIT    := 0x0133
    static WM_CTLCOLORLISTBOX := 0x0134
    static WM_CTLCOLORBTN     := 0x0135
    static WM_CTLCOLORSTATIC  := 0x0138
    static DC_BRUSH           := 18

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

    Mode_Explorer  := (DarkMode ? "DarkMode_Explorer"  : "Explorer")
    Mode_CFD       := (DarkMode ? "DarkMode_CFD"       : "CFD")
    Mode_ItemsView := (DarkMode ? "DarkMode_ItemsView" : "ItemsView")

    for hWnd, GuiCtrlObj in GuiObj {

        switch GuiCtrlObj.Type {

            case "Button", "CheckBox", "UpDown":
                DllCall("uxtheme\SetWindowTheme",
                    "Ptr", GuiCtrlObj.hWnd,
                    "Str", Mode_Explorer,
                    "Ptr", 0)

            case "ComboBox", "DDL":
           {
            ; Apply dark theme to main control
            DllCall("uxtheme\SetWindowTheme",
                "Ptr", GuiCtrlObj.hWnd,
                "Str", Mode_CFD,
                "Ptr", 0)

            ; --- FORCE DARK SCROLLBAR ON DROPDOWN LIST ---
            ; COMBOBOXINFO structure size: 40 + (3 * PtrSize)
            cbInfo := Buffer(40 + (3 * A_PtrSize), 0)
            NumPut("UInt", 40 + (3 * A_PtrSize), cbInfo, 0)

            if DllCall("user32\GetComboBoxInfo",
                "Ptr", GuiCtrlObj.hWnd,
                "Ptr", cbInfo)
            {
                ; hwndList is at offset: 40 + (2 * PtrSize)
                hwndList := NumGet(cbInfo, 40 + (2 * A_PtrSize), "Ptr")

                if (hwndList) {
                    DllCall("uxtheme\SetWindowTheme",
                        "Ptr", hwndList,
                        "Str", Mode_Explorer,
                        "Ptr", 0)
                }
            }
        }

            case "Edit":
                if (DllCall("user32\" GetWindowLong,
                    "Ptr", GuiCtrlObj.hWnd,
                    "Int", GWL_STYLE) & ES_MULTILINE)
                {
                    DllCall("uxtheme\SetWindowTheme",
                        "Ptr", GuiCtrlObj.hWnd,
                        "Str", Mode_Explorer,
                        "Ptr", 0)
                }
                else
                {
                    DllCall("uxtheme\SetWindowTheme",
                        "Ptr", GuiCtrlObj.hWnd,
                        "Str", Mode_CFD,
                        "Ptr", 0)
                }

                GuiCtrlObj.SetFont("cFFFFFF")

            case "Text":
                GuiCtrlObj.Opt("cFFFFFF")
        }
    }

    if !(Init) {
        WindowProcNew := CallbackCreate(WindowProc)
        WindowProcOld := DllCall("user32\" SetWindowLong,
            "Ptr", GuiObj.Hwnd,
            "Int", GWL_WNDPROC,
            "Ptr", WindowProcNew,
            "Ptr")
        Init := True
    }
}


ApplyDarkMode(GuiObj) {
    InitDarkMode()
    SetWindowAttribute(GuiObj, True)
    SetWindowTheme(GuiObj, True)

    ; Force redraw
    DllCall("RedrawWindow", "Ptr", GuiObj.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0285)
    for hWnd, GuiCtrlObj in GuiObj {
        DllCall("InvalidateRect", "Ptr", GuiCtrlObj.Hwnd, "Ptr", 0, "Int", true)
    }
}

; ========================================
; INITIALIZE
; ========================================
Initialize()

Initialize() {
    WriteLog("========================================")
    WriteLog("SFX Helper Tool - Session Started")
    WriteLog("Script Directory: " . A_ScriptDir)
    WriteLog("========================================")

    ; Set tray icon
    if (FileExist(ICON_FILE)) {
        TraySetIcon(ICON_FILE)
        WriteLog("Tray icon set: " . ICON_FILE)
    } else {
        WriteLog("WARNING: Icon file not found at: " . ICON_FILE)
    }

    ; Setup dark mode tray menu
    SetupDarkTrayMenu()

    ; Create Logs folder if not exists
    logsDir := A_ScriptDir . "\Logs"
    if (!DirExist(logsDir)) {
        try {
            DirCreate(logsDir)
            WriteLog("Created Logs directory: " . logsDir)
        } catch Error as e {
            MsgBox("Failed to create Logs directory: " . e.Message, "Error", "Icon!")
            ExitApp
        }
    }

    ; Check if config file exists
    if (!FileExist(CONFIG_FILE)) {
        WriteLog("ERROR: SFXConfig.ini not found at: " . CONFIG_FILE)
        MsgBox("Configuration file not found!`n`nExpected location:`n" . CONFIG_FILE . "`n`nPlease ensure SFXConfig.ini is in the same folder as this script.", "Configuration Missing", "Icon!")
        ExitApp
    }

    WriteLog("Configuration file found: " . CONFIG_FILE)

    ; Check if config editor exists
    if (!FileExist(CONFIG_EDITOR_EXE)) {
        WriteLog("WARNING: Config Editor not found at: " . CONFIG_EDITOR_EXE)
    } else {
        WriteLog("Config Editor found: " . CONFIG_EDITOR_EXE)
    }

    ; Load configurations
    LoadConfigurations()

    ; Create GUI
    CreateMainGui()
}

; ========================================
; LAUNCH CONFIG EDITOR
; ========================================
LaunchConfigEditor(*) {
    global CONFIG_EDITOR_EXE

    WriteLog("Edit button clicked - Launching Config Editor")

    if (!FileExist(CONFIG_EDITOR_EXE)) {
        WriteLog("ERROR: Config Editor not found at: " . CONFIG_EDITOR_EXE)
        ShowDarkMessage("Config Editor Not Found", "7zSFXBuilder-Config-Editor.exe not found at:`n" . CONFIG_EDITOR_EXE, "error")
        return
    }

    try {
        Run('"' . CONFIG_EDITOR_EXE . '"')
        WriteLog("Config Editor launched successfully")
    } catch Error as e {
        WriteLog("ERROR: Failed to launch Config Editor - " . e.Message)
        ShowDarkMessage("Error", "Failed to launch Config Editor!`n" . e.Message, "error")
    }
}

; ========================================
; DEBUG LOGGING FUNCTION
; ========================================
WriteLog(message) {
    global LOG_FILE
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    logEntry := "[" . timestamp . "] " . message . "`n"
    try {
        FileAppend(logEntry, LOG_FILE, "UTF-8")
    }
}

; ========================================
; SETUP DARK MODE TRAY MENU
; ========================================
SetupDarkTrayMenu() {
    WriteLog("Setting up dark mode tray menu")

    ; Remove default menu items
    A_TrayMenu.Delete()

    ; Add custom menu items
    A_TrayMenu.Add("Show", (*) => (mainGui.Show(), mainGui.Flash()))
    A_TrayMenu.Add("Open Logs Folder", (*) => Run(A_ScriptDir . "\Logs"))
    A_TrayMenu.Add("Open Config Folder", (*) => Run(A_ScriptDir . "\7zSFXBuilderConfig"))
    A_TrayMenu.Add()  ; Separator
    A_TrayMenu.Add("Reload Script", (*) => Reload())
    A_TrayMenu.Add("Exit", (*) => ExitApp())

    ; Set default action (double-click)
    A_TrayMenu.Default := "Show"

    ; Apply dark mode to tray menu (Windows 10/11)
    if (VerCompare(A_OSVersion, "10.0.17763") >= 0) {
        try {
            uxtheme := DllCall("GetModuleHandle", "Str", "uxtheme", "Ptr")
            SetPreferredAppMode := DllCall("GetProcAddress", "Ptr", uxtheme, "Ptr", 135, "Ptr")
            FlushMenuThemes := DllCall("GetProcAddress", "Ptr", uxtheme, "Ptr", 136, "Ptr")

            if (SetPreferredAppMode) {
                DllCall(SetPreferredAppMode, "Int", 2)  ; ForceDark = 2
            }
            if (FlushMenuThemes) {
                DllCall(FlushMenuThemes)
            }

            WriteLog("Dark mode applied to tray menu")
        } catch Error as e {
            WriteLog("WARNING: Could not apply dark mode to tray menu - " . e.Message)
        }
    }
}

; ========================================
; LOAD CONFIGURATIONS FROM INI
; ========================================
LoadConfigurations() {
    global configList, configData, CONFIG_FILE

    WriteLog("Loading configurations from INI file...")

    configList := []
    configData := Map()

    ; Read all sections from INI
    sections := IniRead(CONFIG_FILE)

    if (sections = "") {
        WriteLog("ERROR: No configurations found in INI file")
        MsgBox("No configurations found in SFXConfig.ini!", "Error", "Icon!")
        ExitApp
    }

    ; Parse sections (each line is a section name)
    sectionLines := StrSplit(sections, "`n", "`r")

    for index, sectionName in sectionLines {
        if (sectionName != "") {
            configList.Push(sectionName)

            ; Read all keys from this section
            sectionContent := IniRead(CONFIG_FILE, sectionName)
            configData[sectionName] := sectionContent

            WriteLog("Loaded config: " . sectionName)
        }
    }

    WriteLog("Total configurations loaded: " . configList.Length)
}

; ========================================
; SET BUTTON ICON FUNCTION
; ========================================
SetButtonIconIL(btn, iconPath, size := 24, marginLeft := 4, marginTop := 0, marginRight := 6, marginBottom := 0) {
    static BCM_SETIMAGELIST := 0x1602
    static ILC_COLOR32 := 0x20
    static ILC_MASK := 0x1
    if (!IsObject(btn) || iconPath = "") {
        return
    }
    if (!FileExist(iconPath)) {
        return
    }
    hIcon := LoadPicture(iconPath, "Icon w" size " h" size, &imgType)
    if (!hIcon) {
        return
    }
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
        if (btn.HasProp("hImgList") && btn.hImgList) {
            DllCall("Comctl32\ImageList_Destroy", "Ptr", btn.hImgList)
        }
    }
    btn.hImgList := himl
}

; ========================================
; CREATE MAIN GUI
; ========================================
CreateMainGui() {
    global mainGui, dropZoneText, fileNameText, btnPrep, darkTheme
    global btnPin, btnAdd, btnReset, btnAdvance
    global ddlConfig, editFuzzySearch, configList, ICON_PATH

    WriteLog("Creating main GUI")

    mainGui := Gui("+AlwaysOnTop +DPIScale", "7z SFX Helper Tool")
    mainGui.SetFont("s10", "Segoe UI")
    mainGui.BackColor := "0x" . darkTheme["bgColor"]

    ; Set GUI icon
    if (FileExist(ICON_FILE)) {
        try {
            mainGui.Show("Hide")  ; Show hidden first to set icon
            hIcon := LoadPicture(ICON_FILE, "Icon1 w32 h32", &imgType)
            if (hIcon) {
                SendMessage(0x80, 1, hIcon, , mainGui.Hwnd)  ; WM_SETICON, ICON_LARGE
                SendMessage(0x80, 0, hIcon, , mainGui.Hwnd)  ; WM_SETICON, ICON_SMALL
                WriteLog("GUI icon set successfully")
            }
        } catch Error as e {
            WriteLog("WARNING: Failed to set GUI icon - " . e.Message)
        }
    }

    ; Title
    mainGui.SetFont("s14 bold")
    mainGui.AddText("x20 y20 w660 Center c" . darkTheme["textColor"], "7z SFX HELPER TOOL")

    ; Four Buttons with Icons
    mainGui.SetFont("s10 norm")
    btnPin := mainGui.AddButton("x20 y60 w150 h35", "Pinned")
    if FileExist(ICON_PATH . "pin.ico") {
        SetButtonIconIL(btnPin, ICON_PATH . "pin.ico", 24)
    }
    btnPin.OnEvent("Click", TogglePin)

    btnAdd := mainGui.AddButton("x190 y60 w150 h35", "Add File")
    if FileExist(ICON_PATH . "add.ico") {
        SetButtonIconIL(btnAdd, ICON_PATH . "add.ico", 24)
    }
    btnAdd.OnEvent("Click", BrowseFile)

    btnReset := mainGui.AddButton("x360 y60 w150 h35", "Reset")
    if FileExist(ICON_PATH . "reset.ico") {
        SetButtonIconIL(btnReset, ICON_PATH . "reset.ico", 24)
    }
    btnReset.OnEvent("Click", ResetGui)

    btnAdvance := mainGui.AddButton("x530 y60 w150 h35", "Edit")
    if FileExist(ICON_PATH . "edit.ico") {
        SetButtonIconIL(btnAdvance, ICON_PATH . "edit.ico", 24)
    }
    btnAdvance.OnEvent("Click", LaunchConfigEditor)

    ; Drop Down List for Manual Configuration Selection
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y110 c" . darkTheme["textSecondary"], "Select Configuration:")
    mainGui.SetFont("s14 norm")
    ddlConfig := mainGui.Add("DropDownList", "w660 R12", ConfigList)
    ddlConfig.OnEvent("Change", OnConfigChange)

    ; Fuzzy Search Input (increased label size)
    mainGui.SetFont("s11 bold")
    mainGui.AddText("x20 y175 c" . darkTheme["textSecondary"], "Quick Search (Fuzzy):")
    mainGui.SetFont("s12")
    editFuzzySearch := mainGui.AddEdit("x20 y200 w660 h40 -VScroll")
    editFuzzySearch.OnEvent("Change", OnFuzzySearch)

    ; Drop Zone (adjusted Y position)
    mainGui.SetFont("s10 norm")
    dropZoneText := mainGui.AddText("x20 y255 w660 h200 Border Center 0x200 Background" . darkTheme["dropZone"], "`n`n`nDRAG & DROP .7z FILE HERE`n`n`n")
    dropZoneText.SetFont("s12 bold c" . darkTheme["textSecondary"])

    ; File Name Display
    mainGui.SetFont("s9")
    fileNameText := mainGui.AddText("x20 y465 w660 Center c" . darkTheme["textSecondary"], "No file selected")

    ; PREP Button (larger, bold text)
    mainGui.SetFont("s14 bold")
    btnPrep := mainGui.AddButton("x250 y505 w200 h50 Disabled", "PREP")
    btnPrep.OnEvent("Click", ProcessPrep)

    ; Register WM_DROPFILES message
    WriteLog("Registering WM_DROPFILES message handler")
    OnMessage(0x233, WM_DROPFILES)

    ; Enable drag and drop on the GUI window
    WriteLog("Enabling drag and drop on GUI window, HWND: " . mainGui.Hwnd)
    result := DllCall("shell32\DragAcceptFiles", "Ptr", mainGui.Hwnd, "Int", 1, "Int")
    WriteLog("DragAcceptFiles result: " . result)

    ; Apply dark mode
    WriteLog("Applying dark mode theme")
    ApplyDarkMode(mainGui)

    ; Calculate center position
    guiWidth := 700
    guiHeight := 585
    screenWidth := A_ScreenWidth
    screenHeight := A_ScreenHeight
    xPos := (screenWidth - guiWidth) // 2
    yPos := (screenHeight - guiHeight) // 2

    mainGui.Show("w" . guiWidth . " h" . guiHeight . " x" . xPos . " y" . yPos)

    WriteLog("GUI displayed successfully")
}

; ========================================
; PIN/UNPIN TOGGLE
; ========================================
TogglePin(*) {
    global mainGui, btnPin, isPinned, ICON_PATH

    if isPinned {
        mainGui.Opt("-AlwaysOnTop")
        btnPin.Text := "Unpinned"
        if FileExist(ICON_PATH . "unpin.ico") {
            SetButtonIconIL(btnPin, ICON_PATH . "unpin.ico", 24)
        }
        isPinned := false
        WriteLog("Window unpinned")
    } else {
        mainGui.Opt("+AlwaysOnTop")
        btnPin.Text := "Pinned"
        if FileExist(ICON_PATH . "pin.ico") {
            SetButtonIconIL(btnPin, ICON_PATH . "pin.ico", 24)
        }
        isPinned := true
        WriteLog("Window pinned")
    }
}

; ========================================
; BROWSE FILE BUTTON
; ========================================
BrowseFile(*) {
    global droppedFilePath, dropZoneText, fileNameText, btnPrep, mainGui, isPinned

    WriteLog("Browse file button clicked")

    ; Temporarily unpin if pinned
    if isPinned {
        mainGui.Opt("-AlwaysOnTop")
    }

    selectedFile := FileSelect(3, , "Select .7z Archive File", "7z Archive Files (*.7z)")

    ; Re-pin if was pinned
    if isPinned {
        mainGui.Opt("+AlwaysOnTop")
    }

    if (selectedFile != "") {
        WriteLog("File selected via browse: " . selectedFile)

        ; Check if it's a .7z file
        SplitPath(selectedFile, , , &fileExt)

        if (StrLower(fileExt) = "7z") {
            droppedFilePath := selectedFile
            SplitPath(selectedFile, &fileName)
            fileNameText.Value := "Selected: " . fileName
            fileNameText.SetFont("cFFFFFF")
            CheckEnablePrep()
            dropZoneText.Value := "`n`n`n✓ FILE LOADED`n`n`n"
            dropZoneText.SetFont("c0E7A0D")
            WriteLog("Valid .7z file loaded: " . fileName)
        } else {
            WriteLog("Invalid file extension selected: " . fileExt)
            ShowDarkMessage("Invalid File", "Please select only .7z archive files!", "error")
        }
    } else {
        WriteLog("File selection cancelled")
    }
}

; ========================================
; DRAG & DROP HANDLER
; ========================================
WM_DROPFILES(wParam, lParam, msg, hwnd) {
    global droppedFilePath, dropZoneText, fileNameText, btnPrep

    WriteLog("WM_DROPFILES message received, wParam: " . wParam . ", hwnd: " . hwnd)

    ; Get number of files dropped
    fileCount := DllCall("shell32\DragQueryFileW", "Ptr", wParam, "UInt", 0xFFFFFFFF, "Ptr", 0, "UInt", 0, "UInt")
    WriteLog("Number of files dropped: " . fileCount)

    if (fileCount > 0) {
        ; Get the file path
        bufferSize := 260
        filePathBuffer := Buffer(bufferSize * 2, 0)

        length := DllCall("shell32\DragQueryFileW", "Ptr", wParam, "UInt", 0, "Ptr", filePathBuffer.Ptr, "UInt", bufferSize, "UInt")

        if (length > 0) {
            filePath := StrGet(filePathBuffer, "UTF-16")
            WriteLog("File path retrieved: " . filePath)

            ; Get file extension
            SplitPath(filePath, &fileName, , &fileExt)
            WriteLog("File name: " . fileName . ", Extension: " . fileExt)

            ; Check if it's a .7z file
            if (StrLower(fileExt) = "7z") {
                droppedFilePath := filePath
                fileNameText.Value := "Selected: " . fileName
                fileNameText.SetFont("cFFFFFF")
                CheckEnablePrep()
                dropZoneText.Value := "`n`n`n✓ FILE LOADED`n`n`n"
                dropZoneText.SetFont("c0E7A0D")
                WriteLog("Valid .7z file accepted: " . fileName)
            } else {
                WriteLog("Invalid file type dropped: " . fileExt)
                ShowDarkMessage("Invalid File", "Please drop only .7z archive files!", "error")
            }
        } else {
            WriteLog("ERROR: DragQueryFileW returned 0 length")
        }
    }

    DllCall("shell32\DragFinish", "Ptr", wParam)
    WriteLog("DragFinish called")
}

; ========================================
; CONFIGURATION DROPDOWN CHANGE
; ========================================
OnConfigChange(*) {
    global ddlConfig, selectedConfig

    selectedConfig := ddlConfig.Text
    WriteLog("Configuration selected: " . selectedConfig)
    CheckEnablePrep()
}

; ========================================
; FUZZY SEARCH HANDLER
; ========================================
OnFuzzySearch(*) {
    global editFuzzySearch, ddlConfig, configList

    searchText := StrLower(editFuzzySearch.Value)

    if (searchText = "") {
        WriteLog("Fuzzy search cleared")
        return
    }

    WriteLog("Fuzzy search query: " . searchText)

    ; Search for matching configuration
    for index, configName in configList {
        if (InStr(StrLower(configName), searchText)) {
            ddlConfig.Choose(index)
            WriteLog("Fuzzy search matched: " . configName)
            OnConfigChange()
            return
        }
    }

    WriteLog("Fuzzy search: No match found for '" . searchText . "'")
}

; ========================================
; CHECK AND ENABLE PREP BUTTON
; ========================================
CheckEnablePrep() {
    global droppedFilePath, selectedConfig, btnPrep

    if (droppedFilePath != "" && selectedConfig != "") {
        btnPrep.Enabled := true
        WriteLog("PREP button enabled")
    } else {
        btnPrep.Enabled := false
        WriteLog("PREP button disabled - Missing: " . (droppedFilePath = "" ? "File" : "") . (selectedConfig = "" ? " Config" : ""))
    }
}

; ========================================
; PREP BUTTON HANDLER
; ========================================
ProcessPrep(*) {
    global droppedFilePath, selectedConfig, configData, CONFIG_FILE, SFX_BUILDER_EXE

    WriteLog("========================================")
    WriteLog("PREP button clicked")
    WriteLog("Processing file: " . droppedFilePath)
    WriteLog("Using configuration: " . selectedConfig)

    if (droppedFilePath = "") {
        WriteLog("ERROR: No file selected")
        ShowDarkMessage("No File", "Please drag and drop a .7z file first!", "warning")
        return
    }

    if (selectedConfig = "") {
        WriteLog("ERROR: No configuration selected")
        ShowDarkMessage("No Configuration", "Please select a configuration from the dropdown!", "warning")
        return
    }

    ; Create 7zSFXBuilderConfig folder
    configFolderPath := A_ScriptDir . "\7zSFXBuilderConfig"
    WriteLog("Creating config folder: " . configFolderPath)

    if (!DirExist(configFolderPath)) {
        try {
            DirCreate(configFolderPath)
            WriteLog("Config folder created successfully")
        } catch Error as e {
            WriteLog("ERROR: Failed to create config folder - " . e.Message)
            ShowDarkMessage("Error", "Failed to create config folder!`n" . e.Message, "error")
            return
        }
    } else {
        WriteLog("Config folder already exists")
    }

    ; Prepare config.txt path
    configFilePath := configFolderPath . "\config.txt"
    WriteLog("Config file path: " . configFilePath)

    ; Get the selected configuration content
    if (!configData.Has(selectedConfig)) {
        WriteLog("ERROR: Configuration data not found for: " . selectedConfig)
        ShowDarkMessage("Error", "Configuration data not found!", "error")
        return
    }

    configContent := configData[selectedConfig]
    WriteLog("Retrieved configuration content for: " . selectedConfig)

    ; Prepare archive and EXE paths
    SplitPath(droppedFilePath, , &fileDir, &fileExt, &fileNameNoExt)
    archivePath := droppedFilePath
    exePath := fileDir . "\" . fileNameNoExt . ".exe"

    WriteLog("Archive path: " . archivePath)
    WriteLog("EXE path: " . exePath)

    ; Replace template values in configuration
    configContent := StrReplace(configContent, '7zSFXBuilder_7zArchive=""', '7zSFXBuilder_7zArchive=' . archivePath)
    configContent := StrReplace(configContent, '7zSFXBuilder_SFXName=""', '7zSFXBuilder_SFXName=' . exePath)

    WriteLog("Template values replaced in configuration")

    ; Add the required header
    finalConfig := ";!@Install@!UTF-8!`n" . configContent . "`n;!@InstallEnd@!"

    ; Write config.txt file
    WriteLog("Writing config.txt file...")
    try {
        if (FileExist(configFilePath)) {
            FileDelete(configFilePath)
            WriteLog("Existing config.txt deleted")
        }

        FileAppend(finalConfig, configFilePath, "UTF-8")
        WriteLog("config.txt created successfully")
    } catch Error as e {
        WriteLog("ERROR: Failed to write config.txt - " . e.Message)
        ShowDarkMessage("Error", "Failed to write config.txt!`n" . e.Message, "error")
        return
    }

    ; Check if 7z SFX Builder exists
    if (!FileExist(SFX_BUILDER_EXE)) {
        WriteLog("ERROR: 7z SFX Builder not found at: " . SFX_BUILDER_EXE)
        ShowDarkMessage("7z SFX Builder Not Found", "7z SFX Builder not found at:`n" . SFX_BUILDER_EXE, "error")
        return
    }
    WriteLog("7z SFX Builder found")

    ; Launch 7z SFX Builder with config file
    WriteLog("Launching 7z SFX Builder with config file...")
    try {
        Run('"' . SFX_BUILDER_EXE . '" "' . configFilePath . '"')
        WriteLog("7z SFX Builder launched successfully")
    } catch Error as e {
        WriteLog("ERROR: Failed to launch 7z SFX Builder - " . e.Message)
        ShowDarkMessage("Error", "Failed to launch 7z SFX Builder!`n" . e.Message, "error")
        return
    }

    ; Success message
    WriteLog("Process completed successfully!")
    WriteLog("========================================")
    ShowDarkMessage("Success", "Process completed successfully!`n`nConfig file created at:`n" . configFilePath . "`n`n7z SFX Builder has been launched.", "success")
}

; ========================================
; RESET GUI
; ========================================
ResetGui(*) {
    global droppedFilePath, dropZoneText, fileNameText, btnPrep, darkTheme
    global ddlConfig, editFuzzySearch, selectedConfig

    WriteLog("Reset button clicked - Resetting GUI to initial state")

    droppedFilePath := ""
    selectedConfig := ""

    dropZoneText.Value := "`n`n`nDRAG & DROP .7z FILE HERE`n`n`n"
    dropZoneText.SetFont("c" . darkTheme["textSecondary"])

    fileNameText.Value := "No file selected"
    fileNameText.SetFont("c" . darkTheme["textSecondary"])

    ddlConfig.Choose(0)
    editFuzzySearch.Value := ""

    btnPrep.Enabled := false

    WriteLog("GUI reset completed")
}

; ========================================
; DARK MODE MESSAGE BOX
; ========================================
ShowDarkMessage(title, message, iconType := "info") {
    global mainGui, darkTheme

    WriteLog("Showing message - Title: " . title . ", Type: " . iconType)

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

    ; Icon symbol based on type
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

    ; Apply dark mode to message dialog
    ApplyDarkMode(msgGui)

    msgGui.Show("w" . dialogWidth . " h" . dialogHeight)
    WinWaitClose(msgGui.Hwnd)
}
