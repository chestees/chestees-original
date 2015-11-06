<!--#include virtual="/includes/globalLib.asp"-->
<%
Call OpenDB()

'Set Upload Path
'strDirPath = "\\fs1-n02\stor2wc1dfw1\407499\407510\www.chestees.com\web\content\uploads\products\"
strDirPath = "\\fs1-n02\stor2wc1dfw1\407499\407510\www.chestees.com\web\content\uploads\products\"

'_____________________________________________________________________________________________
'Get Variables

intProductID = cInt(Request("ProductID"))
btnSubmit = Request("Submit")

'_____________________________________________________________________________________________
'ADD Record

If intProductID = 0 AND btnSubmit <> "" Then
	
	Set Upload = Server.CreateObject("Persits.Upload.1")	
	Upload.Save(strDirPath)
	
	Set blnActive = Upload.Form("Active")
	Set blnPrivate = Upload.Form("Private")
	Set blnSpecial = Upload.Form("Special")
	Set strProduct = Upload.Form("Product")
	Set intCategoryID = Upload.Form("CategoryID")
	Set strDescription = Upload.Form("Description")
	Set strLinkDesc = Upload.Form("LinkDesc")
	Set strDigg = Upload.Form("Digg")
	Set strStumbleTitle = Upload.Form("StumbleTitle")

			
	SQL = "INSERT INTO tblProduct (" & _
		"Active, Private, Special, Product, CategoryID, Description, LinkDesc, Digg, StumbleTitle " & _
		") VALUES (" & _
		CheckBoxValue(blnActive) & ", " & _
		CheckBoxValue(blnPrivate) & ", " & _
		CheckBoxValue(blnSpecial) & ", " & _
		SQLEncode(strProduct) & ", " & _
		SQLNumEncode(intCategoryID) & ", " & _
		SQLEncode(strDescription) & ", " & _
		SQLEncode(strLinkDesc) & ", " & _
		SQLEncode(strDigg) & ", " & _
		SQLEncode(strStumbleTitle) & ")"
		Conn.Execute(SQL)
	
	'Gets the ID of the record just added
	SQL = "SELECT Max(ProductID) AS MaxID FROM tblProduct"
		Set rsMaxID = Conn.Execute(SQL)
		intProductID = rsMaxID("MaxID")
	
	rsMaxID.Close
	Set rsMaxID = Nothing

	If Upload.Files.Count > 0 Then

		'INDEX PAGE IMAGE 180 x 180
		Set objFilenameMain = Upload.Files("FilenameMain")
		
		If Not objFilenameMain Is Nothing Then
			
			strExt = objFilenameMain.Ext
			strNewFileName = Stripper(strProduct) & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilenameMain = strNewFileName
				
			objFilenameMain.Copy NewFilePath
				
			objFilenameMain.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_Index = " & SQLEncode(DBfilenameMain) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 1
		Set objFilename1 = Upload.Files("Filename1")
		
		If Not objFilename1 Is Nothing Then
			
			strExt = objFilename1.Ext
			strNewFileName = Stripper(strProduct) & "-1" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename1 = strNewFileName
				
			objFilename1.Copy NewFilePath
				
			objFilename1.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_1 = " & SQLEncode(DBfilename1) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 2
		Set objFilename2 = Upload.Files("Filename2")
		
		If Not objFilename2 Is Nothing Then
			
			strExt = objFilename2.Ext
			strNewFileName = Stripper(strProduct) & "-2" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename2 = strNewFileName
				
			objFilename2.Copy NewFilePath
				
			objFilename2.Delete
			
			SQL = "UPDATE tblProduct SET " & _
			"Image_2 = " & SQLEncode(DBfilename2) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 3
		Set objFilename3 = Upload.Files("Filename3")
		
		If Not objFilename3 Is Nothing Then
			
			strExt = objFilename3.Ext
			strNewFileName = Stripper(strProduct) & "-3" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename3 = strNewFileName
				
			objFilename3.Copy NewFilePath
				
			objFilename3.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_3 = " & SQLEncode(DBfilename3) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 4
		Set objFilename4 = Upload.Files("Filename4")
		
		If Not objFilename4 Is Nothing Then
			
			strExt = objFilename4.Ext
			strNewFileName = Stripper(strProduct) & "-4" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename4 = strNewFileName
				
			objFilename4.Copy NewFilePath
				
			objFilename4.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_4 = " & SQLEncode(DBfilename4) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
				
	End If
	
	'UPDATE THE PRODUCT DETAIL TABLE		
	SQL = "SELECT ProductStyleID FROM tblProductStyle"
		Set	rsStyle = Conn.Execute(SQL)
		
	If Not rsStyle.EOF Then
		Do While Not rsStyle.EOF
			
			intProductStyleID = rsStyle("ProductStyleID")
			Set intProductStyleID_Form = Upload.Form("ProductStyleID_" & intProductStyleID)

			If intProductStyleID_Form = "ON" Then
			
				curPrice = Upload.Form("Price_" & intProductStyleID)
				
				SQL = "SELECT ProductSizeID FROM tblProductSize"
					Set	rsSize = Conn.Execute(SQL)
					
				If Not rsSize.EOF Then
					Do While Not rsSize.EOF
						
						intProductSizeID = rsSize("ProductSizeID")
						Set intProductSize_Qty = Upload.Form("ProductSizeID_" & intProductSizeID & "_" & intProductStyleID)

						If intProductSize_Qty > 0 Then
						
							SQL = "INSERT INTO tblProductDetail (" & _
								"ProductStyleID, ProductSizeID, ProductID, Price, Quantity " & _
								") VALUES (" & _
								SQLNumEncode(intProductStyleID) & ", " & _
								SQLNumEncode(intProductSizeID) & ", " & _
								SQLNumEncode(intProductID) & ", " & _
								SQLNumEncode(curPrice) & ", " & _
								SQLNumEncode(intProductSize_Qty) & ")"
								Conn.Execute(SQL)

						End If
					
					rsSize.moveNext
					loop
				
				End If						
							
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
	Upload.Save(strDirPath)	
	
	Set blnActive = Upload.Form("Active")
	Set blnPrivate = Upload.Form("Private")
	Set blnSpecial = Upload.Form("Special")
	Set strProduct = Upload.Form("Product")
	Set intCategoryID = Upload.Form("CategoryID")
	Set strDescription = Upload.Form("Description")
	Set strLinkDesc = Upload.Form("LinkDesc")
	Set strDigg = Upload.Form("Digg")
	Set strStumbleTitle = Upload.Form("StumbleTitle")

	SQL = "UPDATE tblProduct SET " & _
		"Active = " & CheckBoxValue(blnActive) & ", " & _
		"Private = " & CheckBoxValue(blnPrivate) & ", " & _
		"Special = " & CheckBoxValue(blnSpecial) & ", " & _
		"Product = " & SQLEncode(strProduct) & ", " & _
		"CategoryID = " & SQLNumEncode(intCategoryID) & ", " & _
		"Description = " & SQLEncode(strDescription) & ", " & _
		"LinkDesc = " & SQLEncode(strLinkDesc) & ", " & _
		"Digg = " & SQLEncode(strDigg) & ", " & _
		"StumbleTitle = " & SQLEncode(strStumbleTitle) & " " & _
		"WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
	
	If Upload.Files.Count > 0 Then

		'INDEX PAGE IMAGE 180 x 180
		Set objFilenameMain = Upload.Files("FilenameMain")
		
		If Not objFilenameMain Is Nothing Then
			
			strExt = objFilenameMain.Ext
			strNewFileName = Stripper(strProduct) & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilenameMain = strNewFileName
				
			objFilenameMain.Copy NewFilePath
				
			objFilenameMain.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_Index = " & SQLEncode(DBfilenameMain) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 1
		Set objFilename1 = Upload.Files("Filename1")
		
		If Not objFilename1 Is Nothing Then
			
			strExt = objFilename1.Ext
			strNewFileName = Stripper(strProduct) & "-1" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename1 = strNewFileName
				
			objFilename1.Copy NewFilePath
				
			objFilename1.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_1 = " & SQLEncode(DBfilename1) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 2
		Set objFilename2 = Upload.Files("Filename2")
		
		If Not objFilename2 Is Nothing Then
			
			strExt = objFilename2.Ext
			strNewFileName = Stripper(strProduct) & "-2" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename2 = strNewFileName
				
			objFilename2.Copy NewFilePath
				
			objFilename2.Delete
			
			SQL = "UPDATE tblProduct SET " & _
			"Image_2 = " & SQLEncode(DBfilename2) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 3
		Set objFilename3 = Upload.Files("Filename3")
		
		If Not objFilename3 Is Nothing Then
			
			strExt = objFilename3.Ext
			strNewFileName = Stripper(strProduct) & "-3" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename3 = strNewFileName
				
			objFilename3.Copy NewFilePath
				
			objFilename3.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_3 = " & SQLEncode(DBfilename3) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
		
		'IMAGE 4
		Set objFilename4 = Upload.Files("Filename4")
		
		If Not objFilename4 Is Nothing Then
			
			strExt = objFilename4.Ext
			strNewFileName = Stripper(strProduct) & "-4" & strExt
			
			NewFilePath = strDirPath & strNewFileName
			DBfilename4 = strNewFileName
				
			objFilename4.Copy NewFilePath
				
			objFilename4.Delete
		
			SQL = "UPDATE tblProduct SET " & _
			"Image_4 = " & SQLEncode(DBfilename4) & " " & _
			"WHERE ProductID = " & intProductID
			Conn.Execute(SQL)
		End If
				
	End If
	
	'UPDATE THE PRODUCT DETAIL TABLE
	SQL = "DELETE FROM tblProductDetail WHERE ProductID = " & intProductID
		Conn.Execute(SQL)
		
	SQL = "SELECT ProductStyleID FROM tblProductStyle"
		Set	rsStyle = Conn.Execute(SQL)
		
	If Not rsStyle.EOF Then
		Do While Not rsStyle.EOF
			
			intProductStyleID = rsStyle("ProductStyleID")
			Set intProductStyleID_Form = Upload.Form("ProductStyleID_" & intProductStyleID)

			If intProductStyleID_Form = "ON" Then
			
				curPrice = Upload.Form("Price_" & intProductStyleID)
				
				SQL = "SELECT ProductSizeID FROM tblProductSize"
					Set	rsSize = Conn.Execute(SQL)
					
				If Not rsSize.EOF Then
					Do While Not rsSize.EOF
						
						intProductSizeID = rsSize("ProductSizeID")
						Set intProductSize_Qty = Upload.Form("ProductSizeID_" & intProductSizeID & "_" & intProductStyleID)
						'Response.Write(intProductStyleID & "<br>")
						'Response.Write(intProductSize_Qty & "<br>")
						'Response.Write(intProductSizeID & "<br><br>")
						'Response.Flush()
						
						If intProductSize_Qty > 0 Then
						
							SQL = "INSERT INTO tblProductDetail (" & _
								"ProductStyleID, ProductSizeID, ProductID, Price, Quantity " & _
								") VALUES (" & _
								SQLNumEncode(intProductStyleID) & ", " & _
								SQLNumEncode(intProductSizeID) & ", " & _
								SQLNumEncode(intProductID) & ", " & _
								SQLNumEncode(curPrice) & ", " & _
								SQLNumEncode(intProductSize_Qty) & ")"
								Conn.Execute(SQL)
								'Response.Write(SQL & "<br>")
								'Response.Flush()

						End If
					
					rsSize.moveNext
					loop
				
				End If						
							
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

	SQL = "SELECT Active, Private, Special, Product, CategoryID, Description, Image_Index, Image_1, Image_2, Image_3, Image_4, LinkDesc, Digg, StumbleTitle " & _
		"FROM tblProduct " & _
		"WHERE ProductID = " & intProductID
		Set	RS = Conn.Execute(SQL)
		
		blnActive = RS("Active")
		blnPrivate = RS("Private")
		blnSpecial = RS("Special")
		strProduct = RS("Product")
		intSelCategoryID = RS("CategoryID")
		strDescription = RS("Description")
		strFilenameMain = RS("Image_Index")
		strFilename1 = RS("Image_1")
		strFilename2 = RS("Image_2")
		strFilename3 = RS("Image_3")
		strFilename4 = RS("Image_4")
		strLinkDesc = RS("LinkDesc")
		strDigg = RS("Digg")
		strStumbleTitle = RS("StumbleTitle")
	
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
					  <td><span class="PinkB12px">Private?</span> <input name="Private" type="checkbox" value="ON" <%If blnPrivate Then Response.Write("checked")%>></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Special?</span> <input name="Special" type="checkbox" value="ON" <%If blnSpecial Then Response.Write("checked")%>></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Product name</span><br>
				      <input name="Product" type="text" class="Text_300" value="<%=strProduct%>"></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Category</span><br>
				      <select class="Text_300" size="1" name="CategoryID">
					  <option value="0">_____________________</option>
<%
OpenDB()
SQL = "SELECT CategoryID, Category FROM tblCategory"
	Set	rsCat = Conn.Execute(SQL)
	
	Do While Not rsCat.EOF
	
		intCategoryID = rsCat("CategoryID")
		strCategory = rsCat("Category")
%>
					  <option value="<%=intCategoryID%>"<%If intCategoryID = cInt(intSelCategoryID) Then Response.Write(" SELECTED")%>><%=strCategory%></option>
<%
	rsCat.MoveNext
	Loop
	
rsCat.Close
Set rsCat = Nothing
CloseDB()
%>
					  </select>
					  </td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Description</span><br>
				      <textarea name="Description" class="Textarea_400"><%=strDescription%></textarea></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Digg Link</span><br>
				      <input name="Digg" type="text" class="Text_300" style="Width:500px;" value="<%=strDigg%>"></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">StumbleTitle</span><br>
				      <textarea name="StumbleTitle" class="Textarea_400" style="Width:500px;"><%=strStumbleTitle%></textarea></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Link Description</span><br>
				      <textarea name="LinkDesc" class="Textarea_400" style="Width:500px;"><%=strLinkDesc%></textarea></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Main image</span><br>
				      <input type="file" name="FilenameMain" value="<%=strFilenameMain%>">
                      <%If strFilenameMain <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="http://www.chestees.com/uploads/products/<%=strFilenameMain%>"><font size="1"><%=strFilenameMain%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Image 1</span><br>
				      <input type="file" name="Filename1" value="<%=strFilename1%>">
                      <%If strFilename1 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="http://www.chestees.com/uploads/products/<%=strFilename1%>"><font size="1"><%=strFilename1%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Image 2</span><br>
				      <input type="file" name="Filename2" value="<%=strFilename2%>">
                      <%If strFilename2 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="http://www.chestees.com/uploads/products/<%=strFilename2%>"><font size="1"><%=strFilename2%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Image 3</span><br>
				      <input type="file" name="Filename3" value="<%=strFilename3%>">
                      <%If strFilename3 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="http://www.chestees.com/uploads/products/<%=strFilename3%>"><font size="1"><%=strFilename3%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
                    <tr>
					  <td><span class="PinkB12px">Image 4</span><br>
				      <input type="file" name="Filename4" value="<%=strFilename4%>">
                      <%If strFilename4 <> "" Then%>
                        <br><font size="1" color="#CC0000">Existing File: <a target="_blank" href="http://www.chestees.com/uploads/products/<%=strFilename4%>"><font size="1"><%=strFilename4%></font></a></font>
						  <%Else%>
                        <br><font size="1" color="#CC0000">No file exists</font>
                      <%End If%></td>
					</tr>
					<tr>
					  <td><span class="PinkB12px">Styles and sizes available</span></td>
					</tr>
					<tr>
					  <td>
<%
OpenDB()

SQL = "SELECT ProductStyleID, Style FROM tblProductStyle ORDER BY Style"
	Set	RS = Conn.Execute(SQL)
	
If Not RS.EOF Then
	
	Do While Not RS.EOF
		
		intProductStyleID = RS("ProductStyleID")
		strStyle = RS("Style")
		curPrice = ""
		
		If intProductID > 0 Then
		
			SQL = "SELECT ProductStyleID, Price FROM tblProductDetail WHERE ProductID = " & intProductID & " AND ProductStyleID = " & intProductStyleID
				Set RS2 = Conn.Execute(SQL)
				
				If Not RS2.EOF Then
					intProductStyleID_Compare = RS2("ProductStyleID")
					curPrice = RS2("Price")
				End If
				
				RS2.Close
				Set RS2 = Nothing
				
				If intProductStyleID = intProductStyleID_Compare Then 
					blnMatch = true
					blnHighlight = "Highlight"
				Else
					blnMatch = false
					blnHighlight = ""
				End If
		End If
%>
					  	 <table cellpadding="0" cellspacing="0" border="0" class="<%=blnHighlight%>" width="591">
						   <tr>
						   	<td width="20"><input name="ProductStyleID_<%=intProductStyleID%>" type="checkbox" value="ON" <%If blnMatch Then Response.Write("checked")%>></td>
							<td><b><%=strStyle%></b></td>
							<td align="right"><input name="Price_<%=intProductStyleID%>" value="<%=curPrice%>" type="text" class="Text_50"></td>
						  </tr>
						  <tr>
						  	<td>&nbsp;</td>
							<td colspan="2">
							  <table cellpadding="0" cellspacing="0" border="0" width="599">
							  	<tr>
								  <td>
<%
SQL = "SELECT ProductSizeID, Size FROM tblProductSize"
	Set	rsSize = Conn.Execute(SQL)
	
If Not rsSize.EOF Then
	
	Do While Not rsSize.EOF
		
		intProductSizeID = rsSize("ProductSizeID")
		strSize = rsSize("Size")
		
		If intProductID > 0 Then
		
			SQL = "SELECT ProductSizeID, Quantity FROM tblProductDetail WHERE ProductID = " & intProductID & " AND ProductSizeID = " & intProductSizeID & " AND ProductStyleID = " & intProductStyleID
				Set rsSize2 = Conn.Execute(SQL)
				
				If Not rsSize2.EOF Then
					'intProductSizeID_Compare = rsSize2("ProductSizeID")
					intQuantity = rsSize2("Quantity")
				Else 
					intQuantity = 0
				End If
				
				rsSize2.Close
				Set rsSize2 = Nothing
		
		Else
			intQuantity = 0
		End If
%>
								  <%=strSize%>&nbsp;
								  <input name="ProductSizeID_<%=intProductSizeID%>_<%=intProductStyleID%>" value="<%=intQuantity%>" type="text" class="Text_20">
<%
	rsSize.MoveNext
	Loop
	
	rsSize.Close
	Set rsSize = nothing
	
End If
%>
								  </td>
								</tr>
							  </table>
							</td>
						  </tr>
						</table>
						 
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