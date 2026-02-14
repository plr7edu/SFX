' VBS Script to Automatically Create NSIS Self-Extractor
' Place this script in the same directory as your .zip file and template.nsi

Dim fso, currentDir, zipFile, zipFileName, zipBaseName, tempDir
Dim nsiTemplate, nsiContent, nsiFile, shell, nsisPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Get current directory
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Find the first .zip file in current directory
zipFile = ""
zipFileName = ""
zipBaseName = ""

For Each file In fso.GetFolder(currentDir).Files
    If LCase(fso.GetExtensionName(file.Name)) = "zip" Then
        zipFile = file.Path
        zipFileName = file.Name
        zipBaseName = Left(zipFileName, Len(zipFileName) - 4) ' Remove .zip extension
        Exit For
    End If
Next

If zipFile = "" Then
    MsgBox "No .zip file found in current directory!", vbCritical, "Error"
    WScript.Quit
End If

' Create temporary extraction directory
tempDir = currentDir & "\temp_extract"
If fso.FolderExists(tempDir) Then
    fso.DeleteFolder tempDir, True
End If
fso.CreateFolder tempDir

' Extract zip file to temporary directory
Dim zipApp
Set zipApp = CreateObject("Shell.Application")
Set sourceFolder = zipApp.Namespace(zipFile)
Set destFolder = zipApp.Namespace(tempDir)

destFolder.CopyHere sourceFolder.Items, 4 + 16 ' 4 = No progress dialog, 16 = Yes to all

' Wait for extraction to complete
WScript.Sleep 2000

' Read the NSIS template
nsiTemplate = currentDir & "\Template.nsi"
If Not fso.FileExists(nsiTemplate) Then
    MsgBox "template.nsi not found! Please ensure the NSIS template file exists.", vbCritical, "Error"
    WScript.Quit
End If

' Read template content
Set file = fso.OpenTextFile(nsiTemplate, 1)
nsiContent = file.ReadAll
file.Close

' Replace placeholders
nsiContent = Replace(nsiContent, "ZIPFILE_NAME_PLACEHOLDER", zipBaseName)
nsiContent = Replace(nsiContent, "ZIPFILE_PATH_PLACEHOLDER", tempDir)
nsiContent = Replace(nsiContent, "INSTALLER_PATH_PLACEHOLDER", currentDir)

' Create the actual .nsi file
nsiFile = currentDir & "\" & zipBaseName & ".nsi"
Set file = fso.CreateTextFile(nsiFile, True)
file.Write nsiContent
file.Close

' Compile with NSIS
nsisPath = "C:\Program Files (x86)\NSIS\makensis.exe"
If Not fso.FileExists(nsisPath) Then
    nsisPath = "C:\Program Files\NSIS\makensis.exe"
    If Not fso.FileExists(nsisPath) Then
        MsgBox "NSIS not found! Please install NSIS or update the path.", vbCritical, "Error"
        WScript.Quit
    End If
End If

' Run NSIS compiler
Dim compileCmd
compileCmd = """" & nsisPath & """ """ & nsiFile & """"

Dim exec
Set exec = shell.Exec(compileCmd)

' Wait for compilation to complete and capture output
Dim output, errorOutput
output = ""
errorOutput = ""

Do While exec.Status = 0
    WScript.Sleep 100
    If Not exec.StdOut.AtEndOfStream Then
        output = output & exec.StdOut.ReadAll
    End If
    If Not exec.StdErr.AtEndOfStream Then
        errorOutput = errorOutput & exec.StdErr.ReadAll
    End If
Loop

' Read any remaining output
If Not exec.StdOut.AtEndOfStream Then
    output = output & exec.StdOut.ReadAll
End If
If Not exec.StdErr.AtEndOfStream Then
    errorOutput = errorOutput & exec.StdErr.ReadAll
End If

If exec.ExitCode = 0 Then
    MsgBox "Self-extractor created successfully: " & zipBaseName & ".exe", vbInformation, "Success"
    ' Clean up temporary files
    fso.DeleteFolder tempDir, True
    fso.DeleteFile nsiFile
Else
    ' Show detailed error information
    Dim errorMsg
    errorMsg = "NSIS Compilation Failed!" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "Exit Code: " & exec.ExitCode & vbCrLf & vbCrLf
    
    If errorOutput <> "" Then
        errorMsg = errorMsg & "Error Output:" & vbCrLf & errorOutput & vbCrLf & vbCrLf
    End If
    
    If output <> "" Then
        errorMsg = errorMsg & "Standard Output:" & vbCrLf & output
    End If
    
    errorMsg = errorMsg & vbCrLf & "NSI File: " & nsiFile
    errorMsg = errorMsg & vbCrLf & "You can manually check the .nsi file for errors."
    
    MsgBox errorMsg, vbCritical, "Compilation Error Details"
    
    ' Don't clean up files so user can inspect them
    MsgBox "Temporary files left for inspection:" & vbCrLf & "- " & nsiFile & vbCrLf & "- " & tempDir, vbInformation, "Debug Info"
End If