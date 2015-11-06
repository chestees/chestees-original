<!--#include virtual="/includes/globalLib.asp"-->
<%
Call OpenDB()

'Set Upload Path
strDirPath = Server.MapPath ("index.asp")
strDirPath = Replace(lcase(strDirPath), "admin\pictures\index.asp","uploads\pictures\")

'_____________________________________________________________________________________________
'Get Variables

intPictureID = cInt(Request("PictureID"))
btnSubmit = Request("Submit")

If intPictureID > 0 AND btnSubmit <> "" Then

	'Delete the old file
	SQL = "SELECT Filename, Thumbnail FROM tblPicture WHERE PictureID = " & intPictureID
		Set RS = Conn.Execute(SQL)
		
		If Not RS.EOF Then
		
			strFilename = RS("Filename")			
			strFilename = strDirPath & strFilename
			strThumbnail = RS("Thumbnail")			
			strThumbnail = strDirPath & strThumbnail
			
			Set fs=Server.CreateObject("Scripting.FileSystemObject") 
			if fs.FileExists(strFilename) then
			  fs.DeleteFile strFilename,True
			end if
			if fs.FileExists(strThumbnail) then
			  fs.DeleteFile strThumbnail,True
			end if
			set fs = Nothing
			
		End If
		
	RS.Close
	Set RS = Nothing
	'END Delete the old file
	
	SQL = "DELETE FROM tblPicture " & _
		"WHERE PictureID = " & intPictureID
		Conn.Execute(SQL)
	
	Call CloseDB()
	
	Response.Redirect("index.asp")
		
End If
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0">
<table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
  	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="<%=cAdminPath%>images/header.jpg"></td>
		</tr>
		<tr>
		  <td width="153" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  	<td align="center">
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="151" class="NavTitle">Navigation</td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
<!--#include virtual="/incNav.asp" -->
				</td>
			  </tr>
			</table>
		  </td>
		  <td width="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
		  <td width="645" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0" width="645">
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="38"></td>
					  <td width="621" class="PageTitle" align="right">SUBMITTED PICTURES: DELETE</td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="1"></td>
					  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
<form action="delete.asp" method="post">
<input type="hidden" name="PictureID" value="<%=intPictureID%>">
<%
OpenDB()

'_____________________________________________________________________________________________
'CREATE THE CART RECORDSET
SQL = "SELECT Filename, FullName, SubmitDate, Description FROM tblPicture WHERE PictureID = " & intPictureID
	Set RS = Conn.Execute(SQL)
	
	strFilename = RS("Filename")
	strFullName = RS("FullName")
	dtSubmitDate = RS("SubmitDate")
	strDescription = RS("Description")

Call CloseDB()
%>
					<tr bgcolor="<%=cColor%>">
					  <td style="border-left:1px dashed #A13846; border-top:1px dashed #A13846; width:110px;">
					  <img src="<%=cPath%>uploads/pictures/thumbnail_<%=intPictureID%>.jpg"></b><br>
					    <span class="Black11px"><%=strStyle%></span></td>
					  <td style="border-top:1px dashed #A13846; vertical-align:top;"><b><%If strFullName <> "" Then response.Write(strFullName) else response.Write("Name not given")%></b><br><i><%=dtSubmitDate%></i><br>
						<%If strDescription <> "" Then response.Write(strDescription)%></td>
					</tr>
					<tr>
					  <td colspan="3" style="border-top:1px dashed #A13846;"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
					<tr> 
					  <td colspan="3" style="text-align:right; padding-top:10px;"><input type="Submit" value="DELETE" name="Submit"></td>
					</tr>
</form>
				  </table>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
	<td bgcolor="#A13846" width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#A13846" height="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="2"></td>
  </tr>
</table>
<!--#include virtual="/incFooter.asp" -->

</body>
</html>