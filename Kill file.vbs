option explicit

dltFile

Sub dltFile()
Dim wshshell, path, fso
Set wshshell = CreateObject("wscript.shell")
Set fso = CreateObject("Scripting.FileSystemObject")
path = wshshell.ExpandEnvironmentStrings("%USERPROFILE%") & "\My Documents\"
'msgbox  path & "Insight Software\Macro Express\Macro Log\SMT_SP_v3.vbs"
If fso.FileExists(path & "\Insight Software\Macro Express\Macro Logs\SMT_SP_v3.vbs") Then
    fso.deletefile(path & "\Insight Software\Macro Express\Macro Logs\SMT_SP_v3.vbs")
End If

End Sub