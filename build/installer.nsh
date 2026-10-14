; BTS Office — extra Start menu shortcuts for each app
!macro customInstall
  CreateDirectory "$SMPROGRAMS\BTS Office"
  CreateShortCut "$SMPROGRAMS\BTS Office\BTS Docs.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=docs" "$INSTDIR\resources\file-docs.ico" 0
  CreateShortCut "$SMPROGRAMS\BTS Office\BTS Sheets.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=sheets" "$INSTDIR\resources\file-sheets.ico" 0
  CreateShortCut "$SMPROGRAMS\BTS Office\BTS Slides.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=slides" "$INSTDIR\resources\file-slides.ico" 0
  CreateShortCut "$SMPROGRAMS\BTS Office\BTS PDF.lnk" "$INSTDIR\${APP_EXECUTABLE_FILENAME}" "--app=pdf" "$INSTDIR\resources\file-pdf.ico" 0
!macroend
!macro customUnInstall
  Delete "$SMPROGRAMS\BTS Office\BTS Docs.lnk"
  Delete "$SMPROGRAMS\BTS Office\BTS Sheets.lnk"
  Delete "$SMPROGRAMS\BTS Office\BTS Slides.lnk"
  Delete "$SMPROGRAMS\BTS Office\BTS PDF.lnk"
!macroend
