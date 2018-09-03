
'******************************************************************************
'* Function to modify the XML file at runtime
'* Author: Hitesh Patel - 18:21 31-08-2018
'* Parameters:
'*     1.	NodePath – ‘Full path to Node JOINED element’ of target XML, see example below for more info
'*     2.	Attr – Attribute of the XML which need to be modified
'*     3.	Data_static – data which is common among children’s of particular Node
'*     4.	Data_array – array of data, put ‘null’ if Data_Static is only needed/sufficient for your requirements.
'*     5.	FilePath – Target XML path
'* Example :
'*     1.	ModifyXML "/root/settings/sourceDir","value",Replace(currDir & "TC11\TC\Tc11.2.0a_win64","\", "\\"),"null",FilePath
'*     2.	dataArray=Array ("TC11/TC/Tc11.2.0a_win64","TC11/TC/Tc11.4.0_patch_1_wntx64/wntx64")
'*      	ModifyXML "/root/sourceLocations/coreLocations/directory","value",Replace(currDir,"\", "/"), dataArray,FilePath
'******************************************************************************
Function ModifyXML(NodePath,Attr,Data_static,Data_array,FilePath)
On Error Resume Next
Dim oXML,itr,ndFnd
Set oXML = CreateObject("Microsoft.XMLDOM")
oXML.async = False
oXML.load FilePath& "silent.xml"

		If 0= oXML.parseError Then
           	If IsArray(Data_array) Then
       			Set ndFnd = oXML.selectnodes(NodePath)
       			For item=0 to UBound(Data_array)
       				ndFnd(item).setAttribute Attr, Data_static&Data_array(item)
       			Next
       		Else 
    	 			Set ndFnd = oXML.selectsinglenode(NodePath)
					ndFnd.setAttribute Attr,Data_static				
			End If
End If
strResult = oXML.save(FilePath& "silent.xml")
End Function 
