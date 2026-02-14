; Simple Self-Extractor NSIS Script with Progress Bar Only
; This template will be processed by the VBS automation script

!include "MUI2.nsh"

; Compression settings
SetCompressor lzma
SetCompressorDictSize 64

; General settings
Name "ZIPFILE_NAME_PLACEHOLDER"
OutFile "ZIPFILE_NAME_PLACEHOLDER.exe"
RequestExecutionLevel user

; Modern UI Settings - Only show progress bar
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "${NSISDIR}\Contrib\Graphics\Header\nsis.bmp"

; Auto-close settings - window closes immediately after installation
!define MUI_FINISHPAGE_NOAUTOCLOSE
!define MUI_INSTFILESPAGE_FINISHHEADER_TEXT "Extraction Complete"
!define MUI_INSTFILESPAGE_FINISHHEADER_SUBTEXT "Files have been extracted successfully."

; Only include the instfiles page (progress bar)
!insertmacro MUI_PAGE_INSTFILES

; Language
!insertmacro MUI_LANGUAGE "English"

; Installation section - extracts to current directory
Section "Extract Files" SecExtract
    ; Get the directory where the exe is located
    StrCpy $INSTDIR "$EXEDIR"
    
    ; Set output path to current directory
    SetOutPath "$INSTDIR"
    
    ; Extract all files from the zip
    File /r "ZIPFILE_PATH_PLACEHOLDER\*.*"
    
    ; Force window to close immediately after extraction
    SetAutoClose true
    
SectionEnd

; Function to ensure immediate closure
Function .onInstSuccess
    ; Exit immediately without waiting
    Quit
FunctionEnd