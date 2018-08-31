'============================Function : SetFirewall=============================
' SetFirewall(Name,EXEPath,Protocol,Dir,Enable,Action,Profile) 
' Parameter Description: 
' --  Name=Name of the rule 
' --  EXEPath=”Resolved Path of the executable” |NOTE:- Resolve the path up to exe using environment variables Or using Session.Property("Directory Property")
' --  Protocol =tcp or udp
' --  Dir=Inbound(in) or outbound(out) rule
' --  Enable=yes or no
' --  Action=allow or block or custom
' --  Profile=Private and/or public and/or domain (To add rule in more than one profile use “,” E.g.: profile=private, domain )
'
'Note: Rule can’t be added for both the protocols at one time, to do so use separate command with protocol value replaced. Same applies for “dir” and “action” tags.

'=====================================================================
'Version Info:
'  *   Version 1.0.0.0  Created July 14,2017  by Hitesh Patel

'=====================================================================

Function SetFirewall(Name,Exepath,Protocol,Dir,Enable,Action,Profile)
On Error Resume Next
Set objshell = CreateObject("Wscript.Shell")
Set oEnv = objshell.Environment("PROCESS")
objshell.Run "cmd.exe /c " & "netsh advfirewall firewall add rule name=" & Chr(34) & Name & Chr(34) & " dir=" & Dir & " action=" & action & " program=" & Chr(34) & EXEPath & Chr(34) & " enable=" & enable & " profile=" & Profile & " protocol=" & Protocol,0, True
Set objshell = Nothing
End Function
