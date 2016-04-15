<!--#include file="iyzico/iyzico.asp" -->
<%
	Dim pPaymentId, pConversationData, pPaymentConversationId	
	pPaymentId = Request.Form("paymentId")
	pConversationData = Request.Form("conversationData")
	pPaymentConversationId = Request.Form("conversationId")
	pSendAuth = Request.Form("sendAuth")
	
	Set Iyzico = New oIyzico
	If (pSendAuth = "ok") Then
		
		Iyzico.ConversationId = "123456"
		Set Options = Iyzico.CreateOptions("62A2FzHv838Yjt7wDddKOmcxMijGBQYj", "1wIl3hHs8zt4wgkIpZDECSAWQDVEiDO8", "https://api.iyzipay.com")
		
		Set ThreeDsPaymentAuthRequest = Iyzico.CreateRequest("ThreeDsPaymentAuth", Options)
		
		ThreeDsPaymentAuthRequest.PaymentId = pPaymentId
		ThreeDsPaymentAuthRequest.ConversationData = pConversationData
		
		Set ThreeDsPaymentAuthResponse = Iyzico.CreateResponse(ThreeDsPaymentAuthRequest)
		
		Set ThreeDsPaymentAuthRequest = Nothing
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
				<li><a href="threeds_payment.asp">3DSecure Payment</a></li>
				<li class="active">3DSecure Callback</li>
			</ol>
			<hr size="1" />
			
			<h4>3DSecureCallback Post Response</h4>
			<%
				Dim Item, fieldName, fieldValue
				Dim a, b, c, d

				Set d = Server.CreateObject("Scripting.Dictionary")

				For Each Item In Request.Form
					fieldName = Item
					fieldValue = Request.Form(Item)

					d.Add fieldName, fieldValue
				Next
				
				Iyzico.pr(d)
				
				Set d = Nothing
			%>
			
			<hr size="1" />
			<form method="POST">
				<p><input name="paymentId" class="form-control" placeholder="payment id" value="<%=pPaymentId%>" /></p>
				<p><input name="conversationId" class="form-control" placeholder="conversation id" value="<%=pPaymentConversationId%>" /></p>
				<p><input name="conversationData" class="form-control" placeholder="conversation data" value="<%=pConversationData%>" /></p>
				<input type="hidden" name="sendAuth" value="ok"/>
				<button class="btn btn-default" type="submit">Send</button>
			</form>
			<%
				If (NOT IsEmpty(ThreeDsPaymentAuthResponse)) Then
					Response.Write "<hr size=""1"" />"
					Response.Write "<h4>Request</h4>"
					Iyzico.pr(Iyzico.LastRequest)
					Response.Write "<h4>Response</h4>"
					Iyzico.pr(ThreeDsPaymentAuthResponse)
				End If
			%>
		</div>
	</body>
</html>
