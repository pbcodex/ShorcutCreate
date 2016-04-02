Enumeration Window
  #MainForm
EndEnumeration

If OpenWindow(#MainForm, 0, 0, 300, 400, "Demo", #PB_Window_SystemMenu|#PB_Window_ScreenCentered)      
  Repeat : Until WaitWindowEvent(10) = #PB_Event_CloseWindow
EndIf
; IDE Options = PureBasic 5.42 LTS (Windows - x86)
; CursorPosition = 6
; EnableUnicode
; EnableXP
; UseIcon = demo.ico
; Executable = Demo.exe