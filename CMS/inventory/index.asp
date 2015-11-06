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
  	<td bgcolor="#A13846" width="2"><img src="/images/filler.gif" width="2" height="1"></td>
	<td bgcolor="#EEF2FC">
	  <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
	  	<tr>
		  <td colspan="3"><img src="/images/header.jpg"></td>
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
<!--#include virtual="/incNav.asp" -->
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
					  <td width="621" class="PageTitle" align="right">
					  	PRODUCT SALES SUMMARY
					  </td>
					  <td width="20" bgcolor="#A13846"><img src="/images/filler.gif" width="20" height="8"></td>
					  <td width="2"><img src="/images/filler.gif" width="2" height="1"></td>
					</tr>
					<tr>
					  <td colspan="4"><img src="/images/filler.gif" width="1" height="20"></td>
					</tr>
				  </table>
				</td>
			  </tr>
<%
Call OpenDB()

'_____________________________________________________________________________________________
'CREATE THE PRODUCTS RECORDSET
SQL = "SELECT ProductID, Product, Image_Index FROM tblProduct WHERE Active = 1 AND CategoryID = 1 ORDER BY DisplayOrder"
	Set rsProducts = Conn.Execute(SQL)
	
	Do While Not rsProducts.EOF
		
		intProductID = rsProducts("ProductID")
		strProduct = rsProducts("Product")
		strImage = rsProducts("Image_Index")
%>
              <tr>
			  	<td>
				  <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
<%
'_____________________________________________________________________________________________
'CREATE THE SALES RECORDSET
SQL = "SELECT A.Size, B.Style, C.Quantity from " & _
 	"((tblProduct D inner join tblProductDetail C on D.ProductID = C.ProductID) " & _
	"INNER JOIN tblProductSize A on A.ProductSizeID = C.ProductSizeID) " & _
	"INNER JOIN tblProductStyle B on B.ProductStyleID = C.ProductStyleID " & _
	"where D.ProductID = " & intProductID & " " & _
	"order by B.Style"
	Set rsCart = Conn.Execute(SQL)
	
If Not rsCart.EOF Then
%>
            <tr bgcolor="<%=myColor%>">
            	<td>
              	<div>
                	<div style="float:left"><img src="http://www.chestees.com/uploads/products/<%=strImage%>"></div>
                  <div style="margin-left:190px;">
                  	<div style="float:left; font-weight:700; width:150px;">Color</div>
                    <div style="float:left; font-weight:700; width:75px;">Size</div>
                    <div style="float:left; font-weight:700; width: 100px;">Quantity</div>
                  </div>

<%
	myColor = "#E2E2E2"
	curTotalPrice = 0
	intTotalQuantity = 0
	
	Do While Not rsCart.EOF
	
		strSize = rsCart("Size")
		strStyle = rsCart("Style")
		intQTY = rsCart("Quantity")

		If myColor = "#E2E2E2" Then myColor = "#FFFFFF" Else myColor = "#E2E2E2"
%>

                  <div style="margin-left:190px;">
                  	<div style="float:left; width:150px;"><%=strStyle%></div>
                    <div style="float:left; width:75px;"><%=strSize%></div>
                    <div style="float:left; width:100px; text-align:center;"><%=intQTY%></div>
                  </div>
<%
	rsCart.MoveNext
	Loop
%>

<%
Else
%>
            <tr>
              <td style="padding:10px; font-size:16px; font-weight:bold; border:1px dashed #E2E2E2; text-align:right;">NULL</td>
            </tr>
<%
End If
%>
						                </div>
              </td>
            </tr>
				  </table>
				</td>
			  </tr>
<%
	rsProducts.MoveNext
	Loop
%>
			  <tr>
          <td colspan="3">&nbsp;</td>
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
<!--#include virtual="/incFooter.asp" -->

</body>
</html>