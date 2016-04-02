EnableExplicit

Declare createShellLink(Application.s, LinkFileName.s, arg.s, desc.s, dir.s, icon.s, index)

Global Application.s, LinkFileName.s, LinkName.s

;Nom de l'executable
Application = GetCurrentDirectory() + "Demo.exe"

;Nom du lien à afficher sur le bureau
LinkFileName = GetHomeDirectory() + "desktop\" + "Demo link.lnk"

;Titre du lien 
LinkName = "Open Demo"

;Création du lien sur le bureau
If createShellLink(Application, LinkFileName, "", LinkName, "", Application, 0) = 0
  MessageRequester("Information", "Le racourci est crée sur le bureau", #PB_MessageRequester_Ok)
EndIf

;By Rashad (Forum anglophone)
Procedure createShellLink(Application.s, LinkFileName.s, arg.s, desc.s, dir.s, icon.s, index)
  ;Application - path to the exe that is linked to, lnk - link name, dir - working
  ;directory, icon - path to the icon file, index - icon index in iconfile
  Protected hRes.l, mem.s, ppf.IPersistFile
  CompilerIf #PB_Compiler_Unicode
    Protected psl.IShellLinkW
  CompilerElse
    Protected psl.IShellLinkA
  CompilerEndIf
  
  ;make shure COM is active
  CoInitialize_(0)
  hRes = CoCreateInstance_(?CLSID_ShellLink, 0, 1, ?IID_IShellLink, @psl)
  
  If hRes = 0
    psl\SetPath(Application)
    psl\SetArguments(arg)
    psl\SetDescription(desc)
    psl\SetWorkingDirectory(dir)
    psl\SetIconLocation(icon, index)
    
    ;query IShellLink for the IPersistFile interface for saving the
    ;link in persistent storage
    hRes = psl\QueryInterface(?IID_IPersistFile, @ppf)
    
    If hRes = 0
      ;CompilerIf #PB_Compiler_Unicode
      ;save the link
      hRes = ppf\Save(LinkFileName, #True)
      ppf\Release()
    EndIf
    psl\Release()
  EndIf
  
  ;shut down COM
  CoUninitialize_()
  
  DataSection
    CLSID_ShellLink:
    Data.l $00021401
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46
    IID_IShellLink:
    CompilerIf #PB_Compiler_Unicode
      Data.l $000214F9
    CompilerElse
      Data.l $000214EE
    CompilerEndIf
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46
    IID_IPersistFile:
    Data.l $0000010b
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46
  EndDataSection
  ProcedureReturn hRes
EndProcedure
; IDE Options = PureBasic 5.42 LTS (Windows - x86)
; CursorPosition = 5
; Folding = -
; EnableUnicode
; EnableXP
; Executable = Demo.exe