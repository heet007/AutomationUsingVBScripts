'*****************************************************************************
'* Script to fetch the Properties from MSI(database) File  -   Version 1.0.1.0
'* Author: Hitesh Patel
'* Purpose: For just 1 or 2 MSIs, we usually do this task by using orca or installshield , However for large packages such as AutoCAD or Microsoft or SAP etc., where there are more then 10-15 MSIs, fetching the properties manually becomes quite a lengthy and exhaustive process. This script is created to help you in such situation.
'* HOW TO USE: Pass MSI path and a property 
'* TIP: In case need to featch multiple values , Use array of desired properties to featch. for example see below Script
'*****************************************************************************
Function GetMSIPropValue(msi,Prop) 
 dbstr = "Select `Value` From Property WHERE `Property`='" & prop & Chr(39)
 GetMSIPropValue=""
 If msi = "" Or prop="" Then Exit Function End If
 Dim FS, TS, WI, DB2, View, Rec,arr2,dbstr,a
	 Set WI = CreateObject("WindowsInstaller.Installer")
	 Set DB2 = WI.OpenDatabase(msi,0)
 If Err.number Then Exit Function End If
 	Set View = DB2.OpenView(dbstr)
 	View.Execute
 	Set Rec = View.Fetch
 If Not Rec Is Nothing Then
  GetMSIPropValue=Rec.StringData(1)
 End If
' WI=nothing
' DB2=nothing
 End Function 
'------------------------------------------------------
On Error Resume Next
Dim vbtoolkit,Pcode, Ucode,PropArr,fso,OutputFile,strPath,folder,FileCount,MsiPath,textstr,arrytext,arritem,arrylen,a,objShell,objFolder 

PropArr = Array("ProductName","ProductVersion","ProductCode","UpgradeCode")    		'******* This array can be Modified 

sn = WScript.ScriptName '        Script Name
fn = WScript.ScriptFullName '    Fully Qualified Script Name
CurrentDir = Replace(fn, "\" & sn, "\") ' With No Quotes 
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )
    ' Return the path of the selected folder
    If IsObject( objfolder ) Then strPath= objFolder.Self.Path


If strPath = vbNull Then
    WScript.Echo "Cancelled"
Else
	Set OutputFile = fso.CreateTextFile(strPath & "\OutPutFile.csv", True) 		'******* The File Path can be Modified 
End If

arrytext=""
	For arritem=0 To (uBound (PropArr))
		arrytext =  arrytext & PropArr(arritem)
		arrytext = arrytext & Chr(44)
	Next

OutputFile.WriteLine(arrytext)
Set folder=fso.GetFolder(strPath)
Set FileCount=folder.Files

	For Each file In FileCount

		If (File.Type="Windows Installer Package" or File.Type="Windows Installer-pakket") Then
			MsiPath=  file.Path
			textstr=""
			For Each i In PropArr
			Pcode = GetMSIPropValue(MsiPath,i)
			textstr = textstr & Pcode
			textstr = textstr & Chr(44)
			Next 
			OutputFile.WriteLine(textstr)
		Else
			OutputFile.WriteLine("InvalidFile  " & file.Path )
		End If 

	Next

