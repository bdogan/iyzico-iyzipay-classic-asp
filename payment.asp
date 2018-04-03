<!--#include file="iyzico/iyzico.asp" -->
<%

	Dim pBinNumber, pPrice, PaymentResponse
	
	pPrice = Request.Form("price")
	pPaidPrice = Request.Form("paid_price")
	pCurr = Request.Form("curr")
	If (IsEmpty(pCurr)) Then pCurr = "TRY"
	
	pCardHolder = Request.Form("card_holder")
	pCardNumber = Request.Form("card_number")	
	pExpireYear = Request.Form("expire_year")	
	pExpireMonth = Request.Form("expire_month")	
	pCvc = Request.Form("cvc")	
	pRegisterCard = Request.Form("register_card")	
	pInstallment = Request.Form("installment")	
	
	If (NOT IsEmpty(pCardHolder)) Then
		Set Iyzico = New oIyzico
		Iyzico.ConversationId = "123456"
		Set Options = Iyzico.CreateOptions("62A2FzHv838Yjt7wDddKOmcxMijGBQYj", "1wIl3hHs8zt4wgkIpZDECSAWQDVEiDO8", "https://api.iyzipay.com")
		
		Set PaymentRequest = Iyzico.CreateRequest("Payment", Options)
			
		PaymentRequest.Price = pPrice
		PaymentRequest.PaidPrice = pPaidPrice
		PaymentRequest.Curr = pCurr
		PaymentRequest.Installment = pInstallment
		PaymentRequest.BasketId = "B67832"
		PaymentRequest.PaymentChannel = IyzicoPaymentChannelWeb
		PaymentRequest.PaymentGroup = IyzicoPaymentGroupProduct
		
		Set PaymentCard = Iyzico.CreateModel("PaymentCard")
		PaymentCard.CardHolderName = pCardHolder
		PaymentCard.CardNumber = pCardNumber
		PaymentCard.ExpireYear = pExpireYear
		PaymentCard.ExpireMonth = pExpireMonth
		PaymentCard.Cvc = pCvc
		PaymentCard.RegisterCard = pRegisterCard
		PaymentRequest.SetPaymentCard(PaymentCard)
		
		Set Buyer = Iyzico.CreateModel("Buyer")
		Buyer.Id = "B0111"
		Buyer.Name = "John"
		Buyer.Surname = "Due"
		Buyer.IdentityNumber = "74300864791"
		Buyer.Email = "john@foo.com"
		Buyer.GsmNumber = "1234567890"
		Buyer.RegistrationDate = "2015-10-05 12:43:35"
		Buyer.LastLoginDate = "2013-04-21 15:12:09"
		Buyer.RegistrationAddress = "fýstýkçýþahap Nidakule Goztepe Merdivenköy Mah. Bora Sok. No:1"
		Buyer.City = "Istanbul"
		Buyer.Country = "Turkey"
		Buyer.Ip = "85.34.78.112"
		PaymentRequest.SetBuyer(Buyer)
		
		Set BillingAddress = Iyzico.CreateModel("Address")
		BillingAddress.Address = "Nidakule Göztepe Merdivenköy Mah. Bora Sok. No:1"
		BillingAddress.ZipCode = "34742"
		BillingAddress.ContactName = "Jane Doe"
		BillingAddress.City = "Istanbul"
		BillingAddress.Country = "Turkey"
		PaymentRequest.SetBillingAddress(BillingAddress)
		
		Dim BasketItems(1)
		Set BasketItems(0) = Iyzico.CreateModel("BasketItem")
		BasketItems(0).Id = "BI101"
		BasketItems(0).Name = "Binocular"
		BasketItems(0).Category1 = "Collectibles"
		BasketItems(0).Category2 = "Accessories"
		BasketItems(0).ItemType = IyzicoBasketItemTypePhysical
		BasketItems(0).Price = 0.5
		Set BasketItems(1) = Iyzico.CreateModel("BasketItem")
		BasketItems(1).Id = "BI102"
		BasketItems(1).Name = "Binocular2"
		BasketItems(1).Category1 = "Collectibles"
		BasketItems(1).Category2 = "Accessories"
		BasketItems(1).ItemType = IyzicoBasketItemTypePhysical
		BasketItems(1).Price = 0.5
		PaymentRequest.SetBasketItems(BasketItems)
		
		Set ShippingAddress = Iyzico.CreateModel("Address")
		ShippingAddress.Address = "Nidakule Göztepe Merdivenköy Mah. Bora Sok. No:1"
		ShippingAddress.ZipCode = "34742"
		ShippingAddress.ContactName = "Jane Doe"
		ShippingAddress.City = "Istanbul"
		ShippingAddress.Country = "Turkey"
		PaymentRequest.SetShippingAddress(ShippingAddress)
		
		Set PaymentResponse = Iyzico.CreateResponse(PaymentRequest)
		
		Set PaymentRequest = Nothing
		Set Options = Nothing
	Else
		pRegisterCard = "0"
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
				<li class="active">Payment</li>
			</ol>
			<hr size="1" />
			
			<form method="POST">
				<p><input name="price" class="form-control" placeholder="price" value="<%=pPrice%>" /></p>
				<p><input name="paid_price" class="form-control" placeholder="paid price" value="<%=pPaidPrice%>" /></p>
				<p><input name="curr" class="form-control" placeholder="curr" value="<%=pCurr%>" /></p>
				<p><input name="installment" class="form-control" placeholder="installment" value="<%=pInstallment%>" /></p>
				<hr size="1" />
				<p><input name="card_holder" class="form-control" placeholder="card holder" value="<%=pCardHolder%>" /></p>
				<p><input name="card_number" class="form-control" placeholder="card number" value="<%=pCardNumber%>" /></p>
				<p><input name="expire_year" class="form-control" placeholder="expire year" value="<%=pExpireYear%>" /></p>
				<p><input name="expire_month" class="form-control" placeholder="expire month" value="<%=pExpireMonth%>" /></p>
				<p><input name="cvc" class="form-control" placeholder="cvc" value="<%=pCvc%>" /></p>
				<p><input name="register_card" class="form-control" placeholder="register card" value="<%=pRegisterCard%>" /></p>
				<button class="btn btn-default" type="submit">Send</button>
			</form>
			<%
				If (NOT IsEmpty(PaymentResponse)) Then
					Response.Write "<hr size=""1"" />"
					Response.Write "<h4>Request</h4>"
					Iyzico.pr(Iyzico.LastRequest)
					Response.Write "<h4>Response</h4>"
					Iyzico.pr(PaymentResponse)
				End If
			%>
		</div>
	</body>
</html>
