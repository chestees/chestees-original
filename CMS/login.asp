<!--#include virtual="/includes/globalLib.asp"-->
<!--#include virtual="/includes/adovbs.inc" -->
<%

OpenDB()

btnSubmit = Request("Submit")
strUsername = Request(trim("Username"))
strPassword = Request(trim("Userpass"))

If btnSubmit <> "" Then
	
	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	cmd.CommandText = "usp_AdminLogin"
	cmd.CommandType = adCmdStoredProc
	
	cmd.Parameters.Append cmd.CreateParameter("Username",adVarChar,adParamInput,50)
	cmd.Parameters("Username") = strUsername
	cmd.Parameters.Append cmd.CreateParameter("Userpass",adVarChar,adParamInput,50)
	cmd.Parameters("Userpass") = strPassword
	cmd.Parameters.Append cmd.CreateParameter("AdminUserID",adInteger,adParamOutput)
	cmd.Parameters("AdminUserID") = cAdminUserID
	
	cmd.Execute
	
	cAdminUserID = cmd.Parameters("AdminUserID")
	'response.Write("VALID: " & cAdminUserID)
	
	If cAdminUserID > 0 Then
	
		Response.Cookies(cSiteName)("AdminAuthorized") = cAdminUserID
		
		Set cmd = nothing
		
		CloseDB()
		
		Response.Redirect "/shopping/"
	Else 
		' Displayed when a failed login occurs.
		strNotice="<strong><font color=""#FF0000"" face=""Arial"" size=""3"">You are not authorized</font></strong>"		
	End If	
	
	Set cmd = nothing
	
	CloseDB()
	
End If
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0" onLoad="Form_Login.Username.focus();">
<table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
  	<td bgcolor="#A13846" width="2"><img src="/images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="images/header.jpg"></td>
		</tr>
		<tr>
		  <td width="153" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0">
			  <tr>
			  	<td align="center">
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="/images/filler.gif" width="2" height="38"></td>
					  <td width="151" class="NavTitle">Navigation</td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
				  <table width="143" border="0" cellspacing="0" cellpadding="0" align="center">
				  	<tr>
					  <td width="13"><img src="images/nav_Bullet.gif"></td>
					  <td><a href="http://www.chestees.com">chestees.com</a></td>
					</tr>
					<tr>
					  <td colspan="2"><img src="/images/filler.gif" width="2" height="15"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
		  </td>
		  <td width="2" bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
		  <td width="645" valign="top">
		  	<table border="0" cellspacing="0" cellpadding="0" width="645">
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0">
				  	<tr>
				  	  <td width="2"><img src="/images/filler.gif" width="2" height="38"></td>
					  <td width="623" class="PageTitle" align="right">Login</td>
					  <td width="20" bgcolor="#A13846"><img src="/images/filler.gif" width="20" height="1"></td>
					  <td width="2"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="/images/filler.gif" width="1" height="20"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="0" width="230" align="center">
<form action="login.asp" method="post" name="Form_Login">
					<tr>
					  <td colspan="4">Please login...</td>
					</tr>
					<tr>
					  <td colspan="4"><img src="/images/filler.gif" width="1" height="10"></td>
					</tr>
					<tr>
					  <td colspan="4" bgcolor="#A13846"><img src="/images/filler.gif" width="1" height="2"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td colspan="2"><img src="/images/filler.gif" width="1" height="10"></td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td>&nbsp;Username</a></td>
					  <td align="right"><input name="Username" type="text" class="Text_150">&nbsp;</td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td colspan="2"><img src="/images/filler.gif" width="1" height="10"></td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td>&nbsp;Password</a></td>
					  <td align="right"><input name="Userpass" type="password" class="Text_150">&nbsp;</td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td colspan="2"><img src="/images/filler.gif" width="1" height="10"></td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td colspan="2" align="right"><input name="Submit" type="submit" class="Submit" value="Submit">&nbsp;</td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					  <td colspan="2"><img src="/images/filler.gif" width="1" height="10"></td>
					  <td bgcolor="#A13846"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4" bgcolor="#A13846"><img src="/images/filler.gif" width="1" height="2"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="/images/filler.gif" width="1" height="25"></td>
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
	<td bgcolor="#A13846" width="2"><img src="/images/filler.gif" width="2" height="1"></td>
  </tr>
  <tr>
  	<td colspan="3" bgcolor="#A13846" height="2"><img src="/images/filler.gif" width="2" height="2"></td>
  </tr>
</table>
<!--#include file="incFooter.asp" -->

</body>
</html>