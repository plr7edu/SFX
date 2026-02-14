# SFX Tools Collection

A streamlined toolkit for creating self-extracting archives with 7z, WinRAR, and NSIS with automated workflows.

## 📁 Structure

```
1.7z-SFX/              # 7z SFX Create All Tools with automated workflows
2.WinRAR-SFX/          # WinRAR SFX Create All Tools with automated workflows
3.NSIS-SFX/            # NSIS SFX Create All Tools with automated workflows
Icons/                 # Icon resources
```

## 🛠️ Tools

### 7z SFX Builder 

**Quick Method:**

1. Run `7zSFXBuilder-Helper.exe`
2. Drag `.7z` file into window
3. Select config → Click PREP (Load from `SFXConfig.ini`)
4. "7z SFX Builder" Open with Preload selected config. 

**Config Editor:**

- Run `7zSFXBuilder-Config-Editor.exe`
- Create/edit configurations
- Save to `SFXConfig.ini`

### WinRAR SFX

1. Place batch file in folder with **one subfolder**
2. Run `Create Current Path Folder Silent Extract WinRAR SFX.bat`
3. Gets `FolderName.exe` (silent extract)

### NSIS SFX

1. Place `.zip` + scripts in same folder
2. Run `Create-NSIS-Self-Extract-Simple-Installer.vbs`
3. Gets `filename.exe` installer



## 📌 Requirements

- **[7z](https://www.7-zip.org/)**

- [**WinRAR**](https://www.win-rar.com/start.html?&L=0)

- [**NSIS**](https://nsis.sourceforge.io/Main_Page)

- [**AutoHotkey v2.0**](https://www.autohotkey.com/)  

- [**AutoHotkey v2.1 Alpha 18**](https://www.autohotkey.com/boards/viewtopic.php?t=118744#p615486) 

- [**7z SFX Builder**](https://sourceforge.net/projects/s-zipsfxbuilder/) 

- [**7z SFX Constructor**](https://sourceforge.net/projects/sfxconstructor/)

  



