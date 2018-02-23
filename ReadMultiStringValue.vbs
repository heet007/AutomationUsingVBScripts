'---------------------------------------------------------------------
'Name:   ReadMultiStringValue(RegHive,RegKey,RegValue)
'Programmer: Hitesh Patel
'Purpose: Read the REG_MULTI_SZ value from given registry.
'Example: ReadMultiStringValue("HKEY_LOCAL_MACHINE","SYSTEM\CurrentControdlSet\Control\Session Manager","PendingFileRenameOperations")
'Notes:  This Function returns an ARRAY of values if read is successful OR returns ZERO (Boolean value) if key is not present or  read is unsuccessful.
'Return Code Processing :  Catch the functions return code using IsArray() method
'-------------------------------------------------------------------
Function ReadMultiStringValue(RegHive,RegKey,RegValue)
	Select Case RegHive
	    Case "HKEY_CLASSES_ROOT"
	       Const HKEY_CLASSES_ROOT = &H80000000
	    Case "HKEY_CURRENT_USER"
	        Const HKEY_CURRENT_USER = &H80000001
	    Case "HKEY_LOCAL_MACHINE"
	        Const HKEY_LOCAL_MACHINE= &H80000002
	    Case "HKEY_USERS"
	        Const HKEY_USERS = &H80000003
	    Case Else
	End Select
'....................................	
	strComputer = "."
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"&_ 
	    strComputer & "\root\default:StdRegProv")
	strKeyPath = RegKey 
	strValueName =  RegValue 
	Return = objReg.GetMultiStringValue(HKEY_LOCAL_MACHINE,strKeyPath,_
	    strValueName,arrValues)
	    MsgBox "call to reg " & Return
'-----------------------
	If (Return = 0) And (Err.Number = 0) Then   
		ReadMultiStringValue = arrValues
	Else
	    ReadMultiStringValue = 0   
	End If
	END Function
'---------------------------------------------------------------------
