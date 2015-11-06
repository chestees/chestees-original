<!--#include virtual="/includes/globalLib.asp" -->
<!--#include virtual="/includes/adovbs.inc" -->
<%
Response.Buffer = False

strComment = server.HTMLEncode(trim(Request("Comment")))
strEmail = trim(Request("Email"))
dtDate = Now()

If strComment <> "" Then

	'_____________________________________________________________________________________________
	'OPEN DATABASE CONNECTION
	Call OpenDB()

	Set cmd = Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = Conn
	cmd.CommandText = "usp_ContactForm"
	cmd.CommandType = adCmdStoredProc
	
	cmd.Parameters.Append cmd.CreateParameter("Email",adVarChar,adParamInput,75)
	cmd.Parameters("Email") = strEmail
	cmd.Parameters.Append cmd.CreateParameter("Comment",adVarChar,adParamInput,-1)
	cmd.Parameters("Comment") = strComment
	cmd.Parameters.Append cmd.CreateParameter("SubmitDate",adDBTimeStamp,adParamInput)
	cmd.Parameters("SubmitDate") = dtDate

	cmd.Execute
	set cmd = nothing
		
	Call Email_Contact()
	
	Call CloseDB()
	
	Response.Write("<b>Thanks for the comment/question.</b> We will respond as soon as possible.<br><br>Make sure you become a friend on <a target='_blank' href='http://www.facebook.com/chestees'>Facebook</a> as well")

Else 
	Response.Write("Error")
End If
%>