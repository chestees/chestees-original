<!--#include virtual="/includes/globalLib.asp"-->
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
					  <td width="621" class="PageTitle" align="right">STYLES LISTING</td>
					  <td width="20" bgcolor="#A13846"><img src="<%=cAdminPath%>images/filler.gif" width="20" height="1"></td>
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
					  <td colspan="2" align="center" bgcolor="#E4DAC4">~ <a class="Nav" href="edit.asp">Add record</a> ~</td>
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
<%
OpenDB()

SQL = "SELECT ProductStyleID, Style FROM tblProductStyle ORDER BY Style"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
	
	intProductStyleID = RS("ProductStyleID")
	strStyle = RS("Style")
	
	If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
					<tr bgcolor="<%=cColor%>">
					  <td><a class="Record" href="edit.asp?ProductStyleID=<%=intProductStyleID%>"><%=strStyle%></a></td>
					  <td width="30" align="right"><img src="<%=cAdminPath%>images/filler.gif" border="0" width="30" height="1"></td>
					</tr>
<%
	RS.MoveNext
	Loop
	
	RS.Close
	Set RS = nothing
	
Else
%>
					<tr bgcolor="#E2E6EF">
					  <td align="center">No Records</td>
					</tr>
<%
End If

Call CloseDB()
%>
					<tr>
					  <td colspan="2"><img src="<%=cAdminPath%>images/filler.gif" width="1" height="10"></td>
					</tr>
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