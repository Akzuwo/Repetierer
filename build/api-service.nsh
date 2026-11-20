; Start only the headless service after installation, never the desktop UI.
!macro customInstall
  Exec '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" --background-api --api-register-startup'
!macroend
!macro customUnInstall
  Delete "$SMSTARTUP\Repetierer API.vbs"
!macroend
