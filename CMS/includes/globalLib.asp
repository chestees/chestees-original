<%
Dim Conn

cServerName = Request.ServerVariables("Server_Name")
cSiteName = "chestees"
cDBName = "chestees"
cFriendlySiteName = "Chestees.com"
cURL = "chestees.com"
cColor = "#E2E6EF"

cAdminUserID = cInt(Request.Cookies(cSiteName)("AdminAuthorized"))

strConn = "Driver={SQL Server};Server=198.71.226.2;Database=damptshirts;Uid=jdiehl;Pwd=redrock5%;"

cPath = "/"
cAdminPath = "/"

'Response.Write("C: " & cAdminUserID)
'Response.Write("<br>Index: " & InStr(Request.ServerVariables("script_name"),"/login.asp"))
'Response.Write("<br>Root: " & InStr(Request.ServerVariables("script_name"),"/"))
'Response.Write("<br>Approval: " & InStr(Request.ServerVariables("script_name"),"/approval.asp"))
'Response.Write("<br>Script: " & Request.ServerVariables("script_name"))
'_____________________________________________________________________________________________
'Cookie Check
'_____________________________________________________________________________________________
If InStr(Request.ServerVariables("script_name"),"login.asp") < 1 AND InStr(Request.ServerVariables("script_name"),"approval.asp") < 1 Then
	If Not IsNumeric(cAdminUserID) or cAdminUserID < 1 Then
		Response.Redirect ("/login.asp")
	End If	
End If

'_________________________________________________________________________________________________
'Place this subroutine to open db connection
Private Sub OpenDB()	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionString = strConn
	Conn.Open
End Sub
'_________________________________________________________________________________________________
'place this subroutine to close db connection
Private Sub CloseDB()
	Conn.Close
	Set Conn = Nothing
End Sub

'_____________________________________________________________________________________________
'Functions
'_____________________________________________________________________________________________
Function Stripper(inFileName)
	dim sOut,strorigFileName,arrSpecialChar,intCounter
	arrSpecialChar  = Array("%20","%"," ","#","+","(",")","&","$","@","!","*","<",">","?","/","|","\",",","'",":","...","ó",".")
	strorigFileName = replace(inFileName," ","-")
	intCounter = 0

	Do Until intCounter = 24
		sOut = replace(strorigFileName,arrSpecialChar(intCounter),"")
	  	intCounter = intCounter + 1
	  	strorigFileName = lcase(sOut)
	 Loop
	 'StripSpecialChar = response.write(strorigFileName)
	 Stripper = strorigFileName
End Function

Function SQLEncode(strValue)
	strValue = Replace(strValue,"'","''")
	
	If strValue = "" then
		SQLEncode = "NULL"
	Else 
		SQLEncode = "'" & strValue & "'"
	End If	
End Function

Function SQLDateEncode(strValue)
	strValue = Replace(strValue,"'","''")
	
	If strValue = "" then
		SQLDateEncode = "NULL"
	Else 
		SQLDateEncode = "'" & strValue & "'"
	End If
End Function

Function SQLNumEncode(strValue)
	If strValue = "" OR Not IsNumeric(strValue) then
		SQLNumEncode = "NULL"
	Else 
		SQLNumEncode = strValue
	End If
End Function

Function CheckBoxValue(i_value)
	If Lcase(i_value) = "on" then
		CheckBoxValue = 1
	Else
		CheckBoxValue = 0
	End If
End Function
%>