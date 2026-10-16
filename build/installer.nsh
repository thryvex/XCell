; XCell: ลงทะเบียนเป็นตัวเลือก "เปิดด้วย" ของ .xlsx และ .csv เท่านั้น (SPEC 6.2)
; ไม่เขียนค่าเริ่มต้นของ Software\Classes\.xlsx / .csv จึงไม่แย่งเป็นโปรแกรมเริ่มต้น
; (fileAssociations ของ electron-builder บน NSIS เขียนค่าเริ่มต้นนั้น จึงไม่ใช้บน Windows)

!macro xcellRegisterOpenWith EXT DESC
  WriteRegStr SHELL_CONTEXT "Software\Classes\XCell.${EXT}" "" "${DESC}"
  WriteRegStr SHELL_CONTEXT "Software\Classes\XCell.${EXT}\DefaultIcon" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}",0'
  WriteRegStr SHELL_CONTEXT "Software\Classes\XCell.${EXT}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "XCell.${EXT}" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".${EXT}" ""
!macroend

!macro xcellUnregisterOpenWith EXT
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.${EXT}\OpenWithProgids" "XCell.${EXT}"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\XCell.${EXT}"
!macroend

!macro customInstall
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "XCell"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  !insertmacro xcellRegisterOpenWith "xlsx" "Excel Worksheet (XCell)"
  !insertmacro xcellRegisterOpenWith "csv" "CSV Document (XCell)"
  ; แจ้ง Explorer ให้อ่านการลงทะเบียนใหม่ (SHCNE_ASSOCCHANGED)
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro customUnInstall
  !insertmacro xcellUnregisterOpenWith "xlsx"
  !insertmacro xcellUnregisterOpenWith "csv"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\Applications\${APP_EXECUTABLE_FILENAME}"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
