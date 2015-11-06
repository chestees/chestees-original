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
					  <td width="621" class="PageTitle" align="right">COMMENTS - UNAPPROVED</td>
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
				  	<tr>
					  <td style="text-align:right; border-bottom:1px dotted #a13b3b;" colspan="2">- <a href="index.asp">Approved</a> -</td>
					</tr>
<%
OpenDB()

SQL = "SELECT ProductID, Product, Image_Main FROM tblProduct WHERE Active = 1 AND Private = 0 ORDER BY ProductID DESC"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
	
	intProductID = RS("ProductID")
	strProduct = RS("Product")
	strImage_Main = RS("Image_Main")
%>
				  <tr>
					<td style="width:253px; vertical-align:top;"><img src="<%=cPath%>uploads/products/<%=strImage_Main%>" border="0"></td>
					<td valign="top">
					  <table style="width:100%;" border="0" cellspacing="0" cellpadding="4">
						  <tr bgcolor="#666666">
						  	<td><b>Members</b></td>
						  </tr>
<%
	'_____________________________________________________________________________________________
	'CREATE THE COMMENTS RECORDSET
	SQL = "SELECT C.Email, P.Comment, P.Rating, P.DatePosted FROM tblComment P INNER JOIN tblCustomer C ON P.CustomerID = C.CustomerID WHERE P.ProductID = " & intProductID & " AND P.Active = 0 ORDER BY P.DatePosted DESC"
		Set rsComments = Conn.Execute(SQL)
	
		If Not rsComments.EOF Then

			Do While Not rsComments.EOF
				
				strEmail = rsComments("Email")
				strComment = rsComments("Comment")
				intRating = cInt(rsComments("Rating"))
				dtDatePosted = rsComments("DatePosted")
	
			If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
						  <tr bgcolor="<%=cColor%>">
						  	<td><span class="PinkB12px">Rated: <%=intRating%> out of 5</span><br><b><%=dtDatePosted%></b><br><%=strEmail%><br>
							<div style="border:1px dotted #A13846; padding:15px;"><%=strComment%></div></td>
						  </tr>						
<%
			rsComments.MoveNext
			Loop

		Else
			If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
						  <tr bgcolor="<%=cColor%>">
						  	<td><span class="PinkB12px">0 comments from members.</span></td>
						  </tr>
<%
		End If

		rsComments.Close
		Set rsComments = Nothing
%>
						  <tr bgcolor="#666666">
						  	<td><b>Visitors</b></td>
						  </tr>
<%
	'_____________________________________________________________________________________________
	'CREATE THE COMMENTS RECORDSET
	SQL = "SELECT CommentID, Comment, Rating, DatePosted FROM tblComment WHERE ProductID = " & intProductID & " AND CustomerID = 0 AND Active = 0 ORDER BY DatePosted DESC"
		Set rsComments = Conn.Execute(SQL)
	
		If Not rsComments.EOF Then

			Do While Not rsComments.EOF
				
				intCommentID = rsComments("CommentID")
				strComment = rsComments("Comment")
				intRating = cInt(rsComments("Rating"))
				dtDatePosted = rsComments("DatePosted")
	
			If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
						  <tr bgcolor="<%=cColor%>">
						  	<td><span class="PinkB12px">Rated: <%=intRating%> out of 5</span><br><b><%=dtDatePosted%></b><br>
							<div style="border:1px dotted #A13846; padding:15px;"><a class="Delete" href="approval.asp?CommentID=<%=intCommentID%>&Approve=0&Redirect=unapproved.asp">DELETE</a><br><%=strComment%></div></td>
						  </tr>						
<%
			rsComments.MoveNext
			Loop
			
		Else
			If cColor = "#E2E6EF" Then cColor = "#EEF2FC" Else cColor = "#E2E6EF" End If
%>
						  <tr bgcolor="<%=cColor%>">
						  	<td><span class="PinkB12px">0 comments from anonymous.</span></td>
						  </tr>
<%
		End If

		rsComments.Close
		Set rsComments = Nothing
%>
						</table>
					  </td>
					</tr>
					<tr bgcolor="#FFFFFF">
					  <td colspan="2"><hr size="1"></td>
					</tr>
<%
	RS.MoveNext
	Loop

End If
	
RS.Close
Set RS = nothing
	
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