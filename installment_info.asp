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
	Else
		pBinNumber = "454671"
		pPrice = "1"
	End If
%>
<!DOCTYPE html>
<html>
	<head>
		<meta charset="ISO-8859-9">
		<title>Iyzico Api - ASP</title>
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" crossorigin="anonymous">
	</head>
	<body>
		<div class="container">
			<h3>Iyzico Api</h3>
			<hr size="1" />
			<ol class="breadcrumb">
				<li><a href="default.asp">Main Page</a></li>
				<li class="active">InstallmentInfo</li>
			</ol>
			<hr size="1" />
			
			<form method="POST">
				<p><input name="bin_number" class="form-control" placeholder="bin number" value="<%=pBinNumber%>" /></p>
				<p><input name="price" placeholder="price" class="form-control" value="<%=pPrice%>" /></p>
				<button class="btn btn-default" type="submit">Send</button>
			</form>
			<%
				If (NOT IsEmpty(InstallmentResponse)) Then
					Response.Write "<hr size=""1"" />"
					Response.Write "<h4>Request</h4>"
					Iyzico.pr(Iyzico.LastRequest)
					Response.Write "<h4>Response</h4>"
					Iyzico.pr(InstallmentResponse)
				End If
			%>
		</div>
	</body>
</html>
