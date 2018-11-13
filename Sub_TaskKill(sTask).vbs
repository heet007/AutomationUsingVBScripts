Sub TaskKill(sTask)
   Dim oShell, sSystemRoot

   Set oShell = CreateObject("WScript.Shell")
   sSystemRoot = oShell.ExpandEnvironmentStrings("%SystemRoot%")
   oShell.Run sSystemRoot&"\System32\taskkill.exe /F /IM "& Chr(34) &sTask& Chr(34),0,True

   Set sSystemRoot = Nothing
   Set oShell = Nothing
End Sub
