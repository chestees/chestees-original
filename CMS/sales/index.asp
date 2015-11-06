<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp"-->
<%
'_____________________________________________________________________________________________
'Get Variables
btnSubmit = Request("Submit")
dtYear = year(now)
i = 0

Call OpenDB()
Dim myFixedArray(5) 'Fixed size array
myFixedArray(0) = "DateOrdered >= '2008-01-01' AND DateOrdered <= '2008-12-31 11:59:59 PM'"
myFixedArray(1) = "DateOrdered >= '2009-01-01' AND DateOrdered <= '2009-12-31 11:59:59 PM'"
myFixedArray(2) = "DateOrdered >= '2010-01-01' AND DateOrdered <= '2010-12-31 11:59:59 PM'"
myFixedArray(3) = "DateOrdered >= '2011-01-01' AND DateOrdered <= '2011-12-31 11:59:59 PM'"
myFixedArray(4) = "DateOrdered >= '2012-01-01' AND DateOrdered <= '2012-12-31 11:59:59 PM'"
myFixedArray(5) = "DateOrdered >= '2013-01-01' AND DateOrdered <= '2013-12-31 11:59:59 PM'"
%>
<script type="text/javascript" src="https://www.google.com/jsapi"></script>
<script type="text/javascript">
  google.load('visualization', '1', { packages: ['corechart'] });
</script>
<script type="text/javascript">
function drawVisualization() {
  // Create and populate the data table.
  var data = google.visualization.arrayToDataTable([
    ['Year', 'Purchase', 'Shipping', 'Discount', 'Total'],

<%
For Each item In myFixedArray
	
	curPurchaseAmount_2 = 0
	cShippingCost_2 = 0
	curDiscountAmount_2 = 0
	curTotalAmount_2 = 0
	'_____________________________________________________________________________________________
	'CREATE THE PRODUCTS RECORDSET
	SQL = "SELECT PurchaseAmount, ShippingCost, DiscountAmount, TotalAmount "
		SQL = SQL & " FROM  tblOrder "
		SQL = SQL & "WHERE " & item & " ORDER BY DateOrdered DESC"
		Set rsCart = Conn.Execute(SQL)

		Do While Not rsCart.EOF
			curPurchaseAmount = rsCart("PurchaseAmount")
			cShippingCost = rsCart("ShippingCost")
			curDiscountAmount = rsCart("DiscountAmount")
			curTotalAmount = rsCart("TotalAmount")
			
			curPurchaseAmount_2 = curPurchaseAmount_2 + curPurchaseAmount
			cShippingCost_2 = cShippingCost_2 + cShippingCost
			curDiscountAmount_2 = curDiscountAmount_2 + curDiscountAmount
			curTotalAmount_2 = curTotalAmount_2 + curTotalAmount
			
		rsCart.MoveNext
		Loop
		Select Case i
			Case 0
				thisMonth = "2008"
			Case 1
				thisMonth = "2009"
			Case 2
				thisMonth = "2010"
			Case 3
				thisMonth = "2011"
			Case 4
				thisMonth = "2012"
			Case 5
				thisMonth = "2013"
		End Select
%>
  ['<%=thisMonth%>',  <%=curPurchaseAmount_2%>, <%=cShippingCost_2%>, <%=curDiscountAmount_2%>, <%=curTotalAmount_2%>],
<%
	i = i + 1
	curTotal_PurchaseAmount = curTotal_PurchaseAmount + curPurchaseAmount_2
			curTotal_ShippingAmount = curTotal_ShippingAmount + cShippingCost_2
			curTotal_DiscountAmount = curTotal_DiscountAmount + curDiscountAmount_2
			curTotal_TotalAmount = curTotal_TotalAmount + curTotalAmount_2
Next
%>
  ]);

  // Create and draw the visualization.
  new google.visualization.ColumnChart(document.getElementById('visualization')).
    draw(data,
      {title:"Yearly Sales",
      hAxis: {title: "Year"},
      width: 591, height: 500, backgroundColor: 'eef2fc'}
    );
}
google.setOnLoadCallback(drawVisualization);
</script>
<html>
<head>
  <title>
    <%=cFriendlySiteName%>
    | Administration</title>
  <link rel="stylesheet" href="/css/stylesheet.css">
  <link rel="shortcut icon" href="/images/favicon.ico" type="image/x-icon">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="804" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td bgcolor="#A13846" width="2">
        <img src="/images/filler.gif" width="2" height="1">
      </td>
      <td bgcolor="#EEF2FC">
        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td colspan="3">
              <img src="/images/header.jpg">
            </td>
          </tr>
          <tr>
            <td width="153" valign="top">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td align="center">
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="38">
                        </td>
                        <td width="151" class="NavTitle">
                          Navigation
                        </td>
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
            <td width="2" bgcolor="#A13846">
              <img src="/images/filler.gif" width="2" height="1">
            </td>
            <td width="645" valign="top">
              <table border="0" cellspacing="0" cellpadding="0" width="645">
                <tr>
                  <td>
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="38">
                        </td>
                        <td width="621" class="PageTitle" align="right">
                          SALES SUMMARY
                        </td>
                        <td width="20" bgcolor="#A13846">
                          <img src="/images/filler.gif" width="20" height="8">
                        </td>
                        <td width="2">
                          <img src="/images/filler.gif" width="2" height="1">
                        </td>
                      </tr>
                      <tr>
                        <td colspan="4">
                          <img src="/images/filler.gif" width="1" height="20">
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>
                    <table border="0" cellspacing="0" cellpadding="4" width="599" align="center">
                      <tr>
                        <td>
                          <div id="visualization"></div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
      <td bgcolor="#A13846" width="2">
        <img src="/images/filler.gif" width="2" height="1">
      </td>
    </tr>
    <tr>
      <td colspan="3" bgcolor="#A13846" height="2">
        <img src="/images/filler.gif" width="2" height="2">
      </td>
    </tr>
  </table>
  <!--#include virtual="/incFooter.asp" -->
</body>
</html>
