<%@ Language=VBScript %>
<%
If Request.ServerVariables("CONTENT_LENGTH") <> 0 Then

Dim strAmount, strTransDet, strMerchRef
	Set PxObj = Server.CreateObject("PxAccess.PxAccessCtrl")

	PxObj.AmountInput = Request.Form("amount")
	PxObj.TxnType = "Purchase"
	PxObj.CurrencyInput = "NZD"
	PxObj.UserId = "Obtain this from DPS"
	PxObj.MerchantReference = Request.Form("merchantreference")
	PxObj.TxnData1 = Request.Form("txndata1")
	PxObj.TxnData2 = Request.Form("txndata2")
	PxObj.TxnData3 = Request.Form("txndata3")
	PxObj.UrlSuccess = Request.Form("urlsuccess") 
	PxObj.UrlFail = Request.Form("urlfail") 
	PxObj.EmailAddress = Request.Form("email")
	PxObj.DoGenerateRequest
	
	StrRequest = PxObj.Request  
	Set PxObj = Nothing

	Response.Redirect StrRequest
End If		
%>
<html>
<head>
<title>Test PXAccess</title>
</head>
<body>
<form action="pxaccess.asp" method="post" id="form1" name="form1">
 <table border=1>
  <tr>
   <td>Amount</td>
   <td><input type="text" name="amount" value="1.23"></td>
  </tr>
  <tr>
   <td>MerchantReference</td>
   <td><input size=100 type="text" maxlength="64" name="merchantreference" value="merchant reference(appears on txn reports)"></td>
  </tr>
  <tr>
   <td>TxnData1</td>
   <td><input size=100 type="text" name="txndata1" value="optional txndata1"></td>
  </tr>
  <tr>
   <td>TxnData2</td>
   <td><input size=100 type="text" name="txndata2" value="optional txndata2"></td>
  </tr>
  <tr>
   <td>TxnData3</td>
   <td><input size=100 type="text" name="txndata3" value="optional txndata3"></td>
  </tr>
  <tr>
   <td>UrlSuccess</td>
   <td><input size=100 type="text" name="urlsuccess" value="http://www.yourwebsite.co.nz/success.asp"></td>
  </tr>
  <tr>
   <td>UrlFail</td><td><input size=100 type="text" name="urlfail" value="http://www.yourwebsite.co.nz/failed.asp"></td>
  </tr>
  <tr>
   <td>Email Address</td>
   <td><input size=100 type="text" name="email" value="test@test.co.nz"></td>
  </tr>										
  <tr>
   <td>&nbsp;</td>
   <td><input type="Submit" name="Submit" value="Submit"></td>
  </tr>
</table>
</form>






<br>
<br>


<form action="pxaccess.asp" method="post" id="form1" name="form1">
 <table border=1>
  <tr>
   <td>Amount</td>
   <td><input type="text" name="amount" value="1.23"></td>
  </tr>
  <tr>
   <td>MerchantReference</td>
   <td><input size=100 type="text" maxlength="64" name="merchantreference" value="merchant reference(appears on txn reports)"></td>
  </tr>
  <tr>
   <td>TxnData1</td>
   <td><input size=100 type="text" name="txndata1" value="optional txndata1"></td>
  </tr>
  <tr>
   <td>TxnData2</td>
   <td><input size=100 type="text" name="txndata2" value="optional txndata2"></td>
  </tr>
  <tr>
   <td>TxnData3</td>
   <td><input size=100 type="text" name="txndata3" value="optional txndata3"></td>
  </tr>
  <tr>
   <td>UrlSuccess</td>
   <td><input size=100 type="text" name="urlsuccess" value="http://www.yourwebsite.co.nz/success.asp"></td>
  </tr>
  <tr>
   <td>UrlFail</td><td><input size=100 type="text" name="urlfail" value="http://www.yourwebsite.co.nz/failed.asp"></td>
  </tr>
  <tr>
   <td>Email Address</td>
   <td><input size=100 type="text" name="email" value="test@test.co.nz"></td>
  </tr>										
  <tr>
   <td>&nbsp;</td>
   <td><input type="Submit" name="Submit" value="Submit"></td>
  </tr>
</table>
</form>



</body>
</html>

