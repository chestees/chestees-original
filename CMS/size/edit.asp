<!--#include virtual="/includes/globalLib.asp"-->
<%
Call OpenDB()

'_____________________________________________________________________________________________
'Get Variables

intProductSizeID = cInt(Request("ProductSizeID"))
btnSubmit = Request("Submit")

'_____________________________________________________________________________________________
'ADD Record

If intProductSizeID = 0 AND btnSubmit <> "" Then
	
	strSize = Request("Size")
		
	SQL = "INSERT INTO tblProductSize (" & _
		"Size " & _
		") VALUES (" & _
		SQLEncode(strSize) & ")"
		Conn.Execute(SQL)
	
	Call CloseDB()
	
	Response.Redirect "index.asp"

'_____________________________________________________________________________________________
'EDIT Record

ElseIf intProductSizeID > 0 AND btnSubmit <> "" Then
	
	response.Write("HELLO")
	strSize = Request("Size")

	SQL = "UPDATE tblProductSize SET " & _
		"Size = " & SQLEncode(strSize) & " " & _
		"WHERE ProductSizeID = " & intProductSizeID
		Conn.Execute(SQL)

	Call CloseDB()
	
	Response.Redirect "index.asp"

'_____________________________________________________________________________________________
'VIEW Record

ElseIf intProductSizeID > 0 AND btnSubmit = "" Then

	SQL = "SELECT Size " & _
		"FROM tblProductSize " & _
		"WHERE ProductSizeID = " & intProductSizeID
		Set	RS = Conn.Execute(SQL)
		
		strSize = RS("Size")
	
		RS.Close
		Set RS = Nothing
		
		Call CloseDB()

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
					  <td width="621" class="PageTitle" align="right">
<%If intProductSizeID = 0 Then%>
					  	PRODUCTS :: SIZES :: ADD
<%Else%>
					  	PRODUCTS :: SIZES :: EDIT
<%End If%>
					  </td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="8"></td>
					  <td width="2"><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="20"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="2"></td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" align="center" bgcolor="#E4DAC4">~ <a class="Nav" href="index.asp">Back to listing</a> ~</td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="2"></td>
					  <td><img src="<%=cAdminPath%>images/filler.gif" width="2" height="1"></td>
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
<form action="edit.asp" method="post">
<input type="hidden" name="ProductSizeID" value="<%=intProductSizeID%>">
					<tr>
					  <td><span class="PinkB12px">Size</span><br>
				      <input name="Size" type="text" class="Text_300" value="<%=strSize%>"></td>
					</tr>
					<tr>
					  <td colspan="2"><hr color="#CAB689"></td>
					</tr>
					<tr>
					  <td colspan="2" align="right"><input type="submit" name="Submit" class="Submit" value="Submit"></td>
					</tr>
					<tr>
					  <td colspan="2" height="10"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
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