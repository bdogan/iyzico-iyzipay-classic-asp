<!--#include file="iyzico/iyzico.asp" -->
<%

	Dim pBinNumber, pPrice, InstallmentResponse
	pBinNumber = Request.Form("bin_number")
	pPrice = Request.Form("price")
	If (NOT IsEmpty(pBinNumber)) Then
		Set Iyzico = New oIyzico
		Iyzico.ConversationId = "123456"
		Set Options = Iyzico.CreateOptions("62A2FzHv838Yjt7wDddKOmcxMijGBQYj", "1wIl3hHs8zt4wgkIpZDECSAWQDVEiDO8", "https://api.iyzipay.com")
		
		Set InstallmentRequest = Iyzico.CreateRequest("InstallmentInfo", Options)
		
		InstallmentRequest.BinNumber = pBinNumber
		InstallmentRequest.Price = pPrice
		
		Set InstallmentResponse = Iyzico.CreateResponse(InstallmentRequest)
		
		Set InstallmentRequest = Nothing
		Set Options = Nothing
		'Set Iyzico = Nothing
	Else
		pBinNumber = "454671"
		pPrice = "1"
	End If
%>
<!DOCTYPE html>
<html>
  <head>
    <meta charset="iso-8859-9">
    <title>Test Iyzico</title>
  </head>
  <body>
    <h3>Iyzico Installment Info</h3>
	<form method="POST">
		<input name="bin_number" placeholder="bin number" value="<%=pBinNumber%>" />
		<input name="price" placeholder="price" value="<%=pPrice%>" />
		<button type="submit">Send</button>
	</form>
	<%
		If (NOT IsEmpty(InstallmentResponse)) Then
			Iyzico.pr(InstallmentResponse)
		End If
	%>
  </body>
</html>
