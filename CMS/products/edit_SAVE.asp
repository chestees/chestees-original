<!--#include virtual="/includes/globalLib.asp"-->
<%
Call OpenDB()

'Set Upload Path
strDirPath = "../../uploads/products/"

'_____________________________________________________________________________________________
'Get Variables

intProductID = cInt(Request("ProductID"))
btnSubmit = Request("Submit")

'_____________________________________________________________________________________________
'ADD Record

If intProductID = 0 AND btnSubmit <> "" Then
	
	Set Upload = Server.CreateObject("Persits.Upload.1")	
	Upload.SaveVirtual(strDirPath)
	
	Set blnActive = Upload.Form("Active")
	Set strProduct = Upload.Form("Product")
	Set strDescription = Upload.Form("Description")
		
	SQL = "INSERT INTO tblProduct (" & _
		"Active, Product, Description " & _
		") VALUES (" & _
		CheckBoxValue(blnActive) & ", " & _
		SQLEncode(strProduct) & ", " & _
		SQLEncode(strDescription) & ")"
		Conn.Execute(SQL)
	
	'Gets the ID of the record just added
	SQL = "SELECT Max(ProductID) AS MaxID FROM tblProduct"
		Set rsMaxID = Conn.Execute(SQL)
		intProductID = rsMaxID("MaxID")
	
	rsMaxID.Close
	Set rsMaxID = Nothing

	If Upload.Files.Count > 0 Then

		'MAIN IMAGE
		Set objFilename1 = Upload.Files("Filename1")
		
		If Not objFilename1 Is Nothing Then
			
			strTemp = objFilename1.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & strExt
			DBfilename = "product_" & intProductID & strExt
				
			objFilename1.CopyVirtual NewFilePath
				
			objFilename1.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Main = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 1
		Set objFilename3 = Upload.Files("Filename3")
		
		If Not objFilename3 Is Nothing Then
			
			strTemp = objFilename3.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_1" & strExt
			DBfilename = "product_" & intProductID & "_1" & strExt
				
			objFilename3.CopyVirtual NewFilePath
				
			objFilename3.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt1 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 2
		Set objFilename4 = Upload.Files("Filename4")
		
		If Not objFilename4 Is Nothing Then
			
			strTemp = objFilename4.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_2" & strExt
			DBfilename = "product_" & intProductID & "_2" & strExt
				
			objFilename4.CopyVirtual NewFilePath
				
			objFilename4.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt2 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 1
		Set objFilename5 = Upload.Files("Filename5")
		
		If Not objFilename5 Is Nothing Then
			
			strTemp = objFilename5.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_3" & strExt
			DBfilename = "product_" & intProductID & "_3" & strExt
				
			objFilename5.CopyVirtual NewFilePath
				
			objFilename5.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt3 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'RENDERED IMAGE
		Set objFilename2 = Upload.Files("Filename2")
		
		If Not objFilename2 Is Nothing Then
			
			strTemp = objFilename2.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_Render_" & intProductID & strExt
			DBfilename = "product_Render_" & intProductID & strExt
				
			objFilename2.CopyVirtual NewFilePath
				
			objFilename2.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Render = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
				
	End If
	
	'UPDATE THE RELATIONAL SIZE TABLE
	SQL = "DELETE * FROM relProductToProductSize WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
		
	SQL = "SELECT ProductSizeID FROM tblProductSize"
		Set	rsSize = Conn.Execute(SQL)
		
	If Not rsSize.EOF Then
		Do While Not rsSize.EOF
			
			intProductSizeID = rsSize("ProductSizeID")
			Set intProductSizeID_Form = Upload.Form("ProductSizeID_" & intProductSizeID)
			'response.Write(intProductSizeID & "<br>")
			'response.Write(intProductSizeID_Form & "<br>")
			
			If intProductSizeID_Form = "ON" Then
			
				SQL = "INSERT INTO relProductToProductSize (" & _
					"ProductSizeID, ProductID " & _
					") VALUES (" & _
					SQLNumEncode(intProductSizeID) & ", " & _
					SQLNumEncode(intProductID) & ")"
					Conn.Execute(SQL)
					'response.Write(SQL & "<br>")
		
			End If			
			
		rsSize.MoveNext
		Loop
		
	End If
	
	rsSize.Close
	Set rsSize = Nothing
	
	'UPDATE THE RELATIONAL STYLE TABLE
	SQL = "DELETE * FROM relProductToProductStyle WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
		
	SQL = "SELECT ProductStyleID FROM tblProductStyle"
		Set	rsStyle = Conn.Execute(SQL)
		
	If Not rsStyle.EOF Then
		Do While Not rsStyle.EOF
			
			intProductStyleID = rsStyle("ProductStyleID")
			Set intProductStyleID_Form = Upload.Form("ProductStyleID_" & intProductStyleID)
			'response.Write(intProductStyleID & "<br>")
			'response.Write(intProductStyleID_Form & "<br>")
			'response.Flush()
			If intProductStyleID_Form = "ON" Then
			
				curPrice = Upload.Form("Price_" & intProductStyleID)
				
				SQL = "INSERT INTO relProductToProductStyle (" & _
					"ProductStyleID, ProductID, Price " & _
					") VALUES (" & _
					SQLNumEncode(intProductStyleID) & ", " & _
					SQLNumEncode(intProductID) & ", " & _
					SQLNumEncode(curPrice) & ")"
					Conn.Execute(SQL)
					'response.Write(SQL & "<br>")
		
			End If			
			
		rsStyle.MoveNext
		Loop
		
	End If
	
	rsStyle.Close
	Set rsStyle = Nothing
		
	Set Upload = Nothing
	
	Call CloseDB()
	
	Response.Redirect "index.asp"

'_____________________________________________________________________________________________
'EDIT Record

ElseIf intProductID > 0 AND btnSubmit <> "" Then

	Set Upload = Server.CreateObject("Persits.Upload.1")	
	Upload.SaveVirtual(strDirPath)	
	
	Set blnActive = Upload.Form("Active")
	Set strProduct = Upload.Form("Product")
	Set strDescription = Upload.Form("Description")

	SQL = "UPDATE tblProduct SET " & _
		"Active = " & CheckBoxValue(blnActive) & ", " & _
		"Product = " & SQLEncode(strProduct) & ", " & _
		"Description = " & SQLEncode(strDescription) & " " & _
		"WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
	
	If Upload.Files.Count > 0 Then

		'MAIN IMAGE
		Set objFilename1 = Upload.Files("Filename1")
		
		If Not objFilename1 Is Nothing Then
			
			strTemp = objFilename1.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & strExt
			DBfilename = "product_" & intProductID & strExt
				
			objFilename1.CopyVirtual NewFilePath
				
			objFilename1.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Main = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 1
		Set objFilename3 = Upload.Files("Filename3")
		
		If Not objFilename3 Is Nothing Then
			
			strTemp = objFilename3.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_1" & strExt
			DBfilename = "product_" & intProductID & "_1" & strExt
				
			objFilename3.CopyVirtual NewFilePath
				
			objFilename3.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt1 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 2
		Set objFilename4 = Upload.Files("Filename4")
		
		If Not objFilename4 Is Nothing Then
			
			strTemp = objFilename4.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_2" & strExt
			DBfilename = "product_" & intProductID & "_2" & strExt
				
			objFilename4.CopyVirtual NewFilePath
				
			objFilename4.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt2 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'ALT IMAGE 1
		Set objFilename5 = Upload.Files("Filename5")
		
		If Not objFilename5 Is Nothing Then
			
			strTemp = objFilename5.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_" & intProductID & "_3" & strExt
			DBfilename = "product_" & intProductID & "_3" & strExt
				
			objFilename5.CopyVirtual NewFilePath
				
			objFilename5.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Alt3 = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
		
		'RENDERED IMAGE
		Set objFilename2 = Upload.Files("Filename2")
		
		If Not objFilename2 Is Nothing Then
			
			strTemp = objFilename2.ExtractFileName
			strExt = Mid(strTemp, inStrRev(strTemp,"."))			
				
			NewFilePath = strDirPath & "product_Render_" & intProductID & strExt
			DBfilename = "product_Render_" & intProductID & strExt
				
			objFilename2.CopyVirtual NewFilePath
				
			objFilename2.Delete
				
			SQL = "UPDATE tblProduct SET " & _
				"Image_Render = " & SQLEncode(DBfilename) & " " & _
				"WHERE ProductID = " & intProductID
				Conn.Execute(SQL)
		
		End If
				
	End If
	
	'UPDATE THE RELATIONAL SIZE TABLE
	SQL = "DELETE * FROM relProductToProductSize WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
		
	SQL = "SELECT ProductSizeID FROM tblProductSize"
		Set	rsSize = Conn.Execute(SQL)
		
	If Not rsSize.EOF Then
		Do While Not rsSize.EOF
			
			intProductSizeID = rsSize("ProductSizeID")
			Set intProductSizeID_Form = Upload.Form("ProductSizeID_" & intProductSizeID)
			'response.Write(intProductSizeID & "<br>")
			'response.Write(intProductSizeID_Form & "<br>")
			
			If intProductSizeID_Form = "ON" Then
			
				SQL = "INSERT INTO relProductToProductSize (" & _
					"ProductSizeID, ProductID " & _
					") VALUES (" & _
					SQLNumEncode(intProductSizeID) & ", " & _
					SQLNumEncode(intProductID) & ")"
					Conn.Execute(SQL)
					'response.Write(SQL & "<br>")
		
			End If			
			
		rsSize.MoveNext
		Loop
		
	End If
	
	rsSize.Close
	Set rsSize = Nothing
	
	'UPDATE THE RELATIONAL STYLE TABLE
	SQL = "DELETE * FROM relProductToProductStyle WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
		
	SQL = "SELECT ProductStyleID FROM tblProductStyle"
		Set	rsStyle = Conn.Execute(SQL)
		
	If Not rsStyle.EOF Then
		Do While Not rsStyle.EOF
			
			intProductStyleID = rsStyle("ProductStyleID")
			Set intProductStyleID_Form = Upload.Form("ProductStyleID_" & intProductStyleID)
			'response.Write(intProductStyleID & "<br>")
			'response.Write(intProductStyleID_Form & "<br>")
			'response.Flush()
			If intProductStyleID_Form = "ON" Then
			
				curPrice = Upload.Form("Price_" & intProductStyleID)
				
				SQL = "INSERT INTO relProductToProductStyle (" & _
					"ProductStyleID, ProductID, Price " & _
					") VALUES (" & _
					SQLNumEncode(intProductStyleID) & ", " & _
					SQLNumEncode(intProductID) & ", " & _
					SQLNumEncode(curPrice) & ")"
					Conn.Execute(SQL)
					'response.Write(SQL & "<br>")
		
			End If			
			
		rsStyle.MoveNext
		Loop
		
	End If
	
	rsStyle.Close
	Set rsStyle = Nothing
		
	Set Upload = Nothing
	
	Call CloseDB()
	
	Response.Redirect "index.asp"

'_____________________________________________________________________________________________
'VIEW Record

ElseIf intProductID > 0 AND btnSubmit = "" Then

	SQL = "SELECT Active, Product, Description, Image_Main, Image_Alt1, Image_Alt2, Image_Alt3, Image_Render " & _
		"FROM tblProduct " & _
		"WHERE ProductID = " & intProductID
		Set	RS = Conn.Execute(SQL)
		
		blnActive = RS("Active")
		strProduct = RS("Product")
		strDescription = RS("Description")
		'curPrice = RS("Price")
		strFilename1 = RS("Image_Main")
		strFilename3 = RS("Image_Alt1")
		strFilename4 = RS("Image_Alt2")
		strFilename5 = RS("Image_Alt3")
		strFilename2 = RS("Image_Render")
	
		RS.Close
		Set RS = Nothing
		
		Call CloseDB()

End If
%>
<html>
<head>
<title><%=cFriendlySiteName%> | Administration</title>
<link rel="stylesheet" href="<%=cAdminPath%>css/stylesheet.css">
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
<%If intProductID = 0 Then%>
					  	PRODUCTS :: ADD
<%Else%>
					  	PRODUCTS :: EDIT
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
<form action="edit.asp?ProductID=<%=intProductID%>&Submit=True" method="post" ENCTYPE="multipart/form-data">
					<tr>
					  <td><span class="PinkB12px">Active?</span> <input name="Active" type="checkbox" value="ON" <%If blnActive Then Response.Write("checked")%>></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Product name</span><br>
				      <input name="Product" type="text" class="Text_300" value="<%=strProduct%>"></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Description</span><br>
				      <textarea name="Description" class="Textarea_400"><%=strDescription%></textarea></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Main image</span><br>
				      <input type="file" name="Filename1" value="<%=strFilename1%>"><br><font size="1" color="#CC0000">Image must be 243 x 246 pixels.</font>
                      <%If strFilename1 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="<%=cPath%>uploads/products/<%=strFilename1%>"><font size="1"><%=strFilename1%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Alt image 1</span><br>
				      <input type="file" name="Filename3" value="<%=strFilename3%>"><br><font size="1" color="#CC0000">Image must be 243 x 246 pixels.</font>
                      <%If strFilename3 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="<%=cPath%>uploads/products/<%=strFilename3%>"><font size="1"><%=strFilename3%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Alt image 2</span><br>
				      <input type="file" name="Filename4" value="<%=strFilename4%>"><br><font size="1" color="#CC0000">Image must be 243 x 246 pixels.</font>
                      <%If strFilename4 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="<%=cPath%>uploads/products/<%=strFilename4%>"><font size="1"><%=strFilename4%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Alt image 3</span><br>
				      <input type="file" name="Filename5" value="<%=strFilename5%>"><br><font size="1" color="#CC0000">Image must be 243 x 246 pixels.</font>
                      <%If strFilename5 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="<%=cPath%>uploads/products/<%=strFilename5%>"><font size="1"><%=strFilename5%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Rendered image</span><br>
				      <input type="file" name="Filename2" value="<%=strFilename2%>"><br><font size="1" color="#CC0000">Image must be 441 x 357 pixels.</font>
                      <%If strFilename2 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="<%=cPath%>uploads/products/<%=strFilename2%>"><font size="1"><%=strFilename2%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Sizes available</span><br>
<%
OpenDB()

SQL = "SELECT ProductSizeID, Size FROM tblProductSize"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
		
		intProductSizeID = RS("ProductSizeID")
		strSize = RS("Size")
		
		If intProductID > 0 Then
		
			SQL = "SELECT ProductSizeID FROM relProductToProductSize WHERE ProductID = " & intProductID & " AND ProductSizeID = " & intProductSizeID
				Set RS2 = Conn.Execute(SQL)
				
				If Not RS2.EOF Then
					intProductSizeID_Compare = RS2("ProductSizeID")
				End If
				
				RS2.Close
				Set RS2 = Nothing
				
		End If
%>
					  	 <input name="ProductSizeID_<%=intProductSizeID%>" type="checkbox" value="ON" <%If intProductSizeID = intProductSizeID_Compare Then Response.Write("checked")%>> <%=strSize%><br>
<%
	RS.MoveNext
	Loop
	
	RS.Close
	Set RS = nothing
	
End If

CloseDB()
%>
				      </td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Styles available</span><br>
<%
OpenDB()

SQL = "SELECT ProductStyleID, Style FROM tblProductStyle"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
		
		intProductStyleID = RS("ProductStyleID")
		strStyle = RS("Style")
		curPrice = ""
		
		If intProductID > 0 Then
		
			SQL = "SELECT ProductStyleID, Price FROM relProductToProductStyle WHERE ProductID = " & intProductID & " AND ProductStyleID = " & intProductStyleID
				Set RS2 = Conn.Execute(SQL)
				
				If Not RS2.EOF Then
					intProductStyleID_Compare = RS2("ProductStyleID")
					curPrice = RS2("Price")
				End If
				
				RS2.Close
				Set RS2 = Nothing
				
		End If
%>
					  	 <input name="ProductStyleID_<%=intProductStyleID%>" type="checkbox" value="ON" <%If intProductStyleID = intProductStyleID_Compare Then Response.Write("checked")%>> <%=strStyle%>
						 <input name="Price_<%=intProductStyleID%>" value="<%=curPrice%>" type="text" class="Text_50">
						 <br>
<%
	RS.MoveNext
	Loop
	
	RS.Close
	Set RS = nothing
	
End If

CloseDB()
%>
				      </td>
					</tr>
					<tr>
					  <td colspan="2"><hr color="#CAB689"></td>
					</tr>
					<tr>
					  <td colspan="2" align="right"><input type="submit" class="Submit" value="Submit"></td>
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