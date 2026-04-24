!ifndef MUI_LICENSEPAGE_CHECKBOX
  !define MUI_LICENSEPAGE_CHECKBOX
!endif

!ifndef MUI_LICENSEPAGE_CHECKBOX_TEXT
  !define MUI_LICENSEPAGE_CHECKBOX_TEXT "Ich akzeptiere das Lizenzabkommen und den Haftungsausschluss."
!endif

!ifndef BUILD_UNINSTALLER
!macro customPageAfterChangeDir
  Page custom ExistingInstallNoticePage
!macroend

Function ExistingInstallNoticePage
  ReadRegStr $0 HKCU "Software\${APP_GUID}" InstallLocation
  StrCmp $0 "" 0 showExistingInstallNotice
  ReadRegStr $0 HKLM "Software\${APP_GUID}" InstallLocation
  StrCmp $0 "" 0 showExistingInstallNotice
  Goto doneExistingInstallNotice

  Goto showExistingInstallNotice

  showExistingInstallNotice:
    MessageBox MB_OK|MB_ICONINFORMATION "Es wurde bereits eine aeltere Version von Repetierer gefunden.$\r$\n$\r$\nDie Installation aktualisiert die Anwendung. Deine Excel-Datei, Einstellungen und Backup-Daten koennen weiterverwendet werden."

  doneExistingInstallNotice:
  Abort
FunctionEnd
!endif
