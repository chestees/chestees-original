<!--#include virtual="/includes/globalLib.asp"-->
<%
intCommentID = Request("CommentID")
blnApprove = Request("Approve")
strRedirect = Request("Redirect")

OpenDB()

If blnApprove = 1 Then
	
	SQL = "UPDATE tblComment SET Active = 1 WHERE CommentID = " & intCommentID
		Conn.Execute(SQL)
	strMessage = "APPROVED"

Else

	SQL = "DELETE FROM tblComment WHERE CommentID = " & intCommentID
		Conn.Execute(SQL)
	strMessage = "NOT APPROVED"
		
End If
		
If strRedirect <> "" Then
	Response.Redirect(strRedirect)
End If
CloseDB()
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="/css/stylesheet.css">
<link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>

<body leftmargin="0" topmargin="0" marginWidth="0" marginHeight="0">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
  	<td align="center"><%=strMessage%></td>
  </tr>
</table>

</body>
</html>