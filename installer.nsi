; =========================================================
; AR Studio Pro - Complete NSIS Installer & Uninstaller Script
; =========================================================

!define APP_NAME "AR Studio Pro"
!define COMP_NAME "AR Studio"
!define VERSION "26.009.21"
!define EXEC_NAME "MAIN.exe"
!define ICON_NAME "logo.ico"

; Installer Configuration
Name "${APP_NAME}"
OutFile "AR_Studio_Pro_Setup.exe"
InstallDir "$PROGRAMFILES64\${APP_NAME}"
InstallDirRegKey HKLM "Software\${APP_NAME}" "Install_Dir"
RequestExecutionLevel admin

; UI Configuration (Modern UI 2)
!include "MUI2.nsh"

!define MUI_ICON "${ICON_NAME}"
!define MUI_UNICON "${ICON_NAME}"
!define MUI_ABORTWARNING

; Installer Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Uninstaller Pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Language
!insertmacro MUI_LANGUAGE "English"

; =========================================================
; Installation Section
; =========================================================
Section "MainSection" SEC01
    SetOutPath "$INSTDIR"
    SetOverwrite on

    ; Single EXE file (--onefile mode ke liye)
    File "dist\MAIN.exe"
    
    ; Icon file ko instdir me copy karein taaki uninstaller/shortcuts ise use kar sakein
    File "${ICON_NAME}"

    ; Registry Keys
    WriteRegStr HKLM "Software\${APP_NAME}" "Install_Dir" "$INSTDIR"
    
    ; Control Panel Uninstaller Settings
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" '"$INSTDIR\uninstall.exe"'
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\${ICON_NAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${COMP_NAME}"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${VERSION}"

    ; Write Uninstaller
    WriteUninstaller "$INSTDIR\uninstall.exe"

    ; Desktop Shortcut
    CreateShortCut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${EXEC_NAME}" "" "$INSTDIR\${ICON_NAME}" 0

    ; Start Menu Shortcuts
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortCut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${EXEC_NAME}" "" "$INSTDIR\${ICON_NAME}" 0
    CreateShortCut "$SMPROGRAMS\${APP_NAME}\Uninstall ${APP_NAME}.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
SectionEnd

; =========================================================
; Uninstallation Section
; =========================================================
Section "Uninstall"
    ; Files aur Folder delete karein
    Delete "$INSTDIR\${EXEC_NAME}"
    Delete "$INSTDIR\${ICON_NAME}"
    Delete "$INSTDIR\uninstall.exe"
    RMDir /r "$INSTDIR"

    ; Shortcuts remove karein
    Delete "$DESKTOP\${APP_NAME}.lnk"
    RMDir /r "$SMPROGRAMS\${APP_NAME}"

    ; Registry Keys delete karein
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
    DeleteRegKey HKLM "Software\${APP_NAME}"
SectionEnd