; DeskGrid NSIS Installer Script
; Simple installer without MUI

Name "DeskGrid"
OutFile "DeskGrid-Setup.exe"
Unicode True
InstallDir "$PROGRAMFILES64\DeskGrid"
RequestExecutionLevel admin

; Metadata
VIProductVersion "1.0.0.0"
VIAddVersionKey "ProductName" "DeskGrid"
VIAddVersionKey "FileDescription" "DeskGrid Installer"
VIAddVersionKey "FileVersion" "1.0.0"

; Installer Section
Section "DeskGrid" SecMain
    SetOutPath "$INSTDIR"
    
    ; Copy all published files
    File /r "publish\*.*"
    
    ; Create Start Menu shortcuts
    CreateDirectory "$SMPROGRAMS\DeskGrid"
    CreateShortcut "$SMPROGRAMS\DeskGrid\DeskGrid.lnk" "$INSTDIR\DeskGrid.exe"
    CreateShortcut "$SMPROGRAMS\DeskGrid\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
    
    ; Create Desktop shortcut
    CreateShortcut "$DESKTOP\DeskGrid.lnk" "$INSTDIR\DeskGrid.exe"
    
    ; Write registry keys
    WriteRegStr HKLM "Software\DeskGrid" "InstallPath" "$INSTDIR"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "DisplayName" "DeskGrid"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "DisplayIcon" "$INSTDIR\DeskGrid.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "Publisher" "DeskGrid"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "DisplayVersion" "1.0.0"
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "NoModify" 1
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid" "NoRepair" 1
    
    ; Add to startup
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "DeskGrid" "$INSTDIR\DeskGrid.exe"
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; Uninstaller Section
Section "Uninstall"
    ; Kill DeskGrid if it's running
    nsExec::ExecToLog 'taskkill /F /IM DeskGrid.exe'
    ; Wait a moment for process to terminate
    Sleep 1000
    
    ; Remove files
    RMDir /r "$INSTDIR"
    
    ; Remove shortcuts
    Delete "$SMPROGRAMS\DeskGrid\DeskGrid.lnk"
    Delete "$SMPROGRAMS\DeskGrid\Uninstall.lnk"
    RMDir "$SMPROGRAMS\DeskGrid"
    Delete "$DESKTOP\DeskGrid.lnk"
    
    ; Remove registry keys
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskGrid"
    DeleteRegKey HKLM "Software\DeskGrid"
    DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "DeskGrid"
SectionEnd
