#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon

; ========================================
; CONFIGURATION - EDIT THESE VALUES
; ========================================
global PROGRAM_NAME := "Zip2exe"                                    ; Display name of the program
global PROGRAM_PATH := "C:\Program Files (x86)\NSIS\Bin\zip2exe.exe"      ; Full path to the executable
global PROGRAM_ARGS := ""                                         ; Command line arguments (leave empty if none)

; ========================================
; SCRIPT START - DO NOT EDIT BELOW
; ========================================

; Check if program exists
if (!FileExist(PROGRAM_PATH)) {
    MsgBox(
        PROGRAM_NAME . " not found!`n`n" .
        "Expected location:`n" . PROGRAM_PATH . "`n`n" .
        "Please verify " . PROGRAM_NAME . " is installed.",
        "Error - " . PROGRAM_NAME . " Not Found",
        "Icon! 48"
    )
    ExitApp
}

; Launch program
try {
    if (PROGRAM_ARGS != "") {
        Run('"' . PROGRAM_PATH . '" ' . PROGRAM_ARGS)
    } else {
        Run('"' . PROGRAM_PATH . '"')
    }
} catch Error as e {
    MsgBox(
        "Failed to launch " . PROGRAM_NAME . "!`n`n" .
        "Error: " . e.Message,
        "Error - Launch Failed",
        "Icon! 48"
    )
    ExitApp
}

; Exit script after launching
ExitApp