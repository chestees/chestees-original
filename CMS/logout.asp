<!--#include virtual="/includes/globalLib.asp" -->
<%
Response.Cookies(cSiteName)("AdminAuthorized") = 0
response.redirect("/")
%>