<%

Class rThreeDsPayment

	Public Path
	Public Method
	Private F
	Public Sub Class_Initialize
		Path = "/payment/iyzipos/initialize3ds/ecom"
		Method = "POST"
		Set F = New oIzyicoRequestFormatter
	End Sub 
	
	Private pPrice
	Public Property Get Price
		Price = F.FormatPrice(pPrice)
	End Property
	Public Property Let Price(pVal)
		pPrice = pVal
	End Property
	
	Private pCallbackUrl
	Public Property Get CallbackUrl
		CallbackUrl = pCallbackUrl
	End Property
	Public Property Let CallbackUrl(pVal)
		pCallbackUrl = pVal
	End Property
    
	Private pPaidPrice
	Public Property Get PaidPrice
		PaidPrice = F.FormatPrice(pPaidPrice)
	End Property
	Public Property Let PaidPrice(pVal)
		pPaidPrice = pVal
	End Property
	
	Private pInstallment
	Public Property Get Installment
		If (IsEmpty(pInstallment)) Then pInstallment = 1
		Installment = pInstallment
	End Property
	Public Property Let Installment(pVal)
		pInstallment = pVal
	End Property
	
	Private pPaymentChannel
	Public Property Get PaymentChannel
		PaymentChannel = pPaymentChannel
	End Property
	Public Property Let PaymentChannel(pVal)
		pPaymentChannel = pVal
	End Property
	
	Private pBasketId
	Public Property Get BasketId
		BasketId = pBasketId
	End Property
	Public Property Let BasketId(pVal)
		pBasketId = pVal
	End Property
	
	Private pPaymentGroup
	Public Property Get PaymentGroup
		PaymentGroup = pPaymentGroup
	End Property
	Public Property Let PaymentGroup(pVal)
		pPaymentGroup = pVal
	End Property
	
	Private pPaymentCard
	Public Sub SetPaymentCard(vPaymentCard)
		Set pPaymentCard = vPaymentCard
	End Sub
	Public Property Get PaymentCard
		If (IsEmpty(pPaymentCard)) Then Exit Property
		Set PaymentCard = pPaymentCard
	End Property

	Private pBuyer
	Public Sub SetBuyer(vBuyer)
		Set pBuyer = vBuyer
	End Sub
	Public Property Get Buyer
		If (IsEmpty(pBuyer)) Then Exit Property
		Set Buyer = pBuyer
	End Property
	
	Private pShippingAddress
	Public Sub SetShippingAddress(vShippingAddress)
		Set pShippingAddress = vShippingAddress
	End Sub
	Public Property Get ShippingAddress
		If (IsEmpty(pShippingAddress)) Then Exit Property
		Set ShippingAddress = pShippingAddress
	End Property
	
	Private pBillingAddress
	Public Sub SetBillingAddress(vBillingAddress)
		Set pBillingAddress = vBillingAddress
	End Sub
	Public Property Get BillingAddress
		If (IsEmpty(pBillingAddress)) Then Exit Property
		Set BillingAddress = pBillingAddress
	End Property
	
	Private pBasketItems
	Public Sub SetBasketItems(vBasketItems)
		pBasketItems = vBasketItems
	End Sub
	Public Property Get BasketItems
		BasketItems = pBasketItems
	End Property
	
	Private pPaymentSource
	Public Property Get PaymentSource
		PaymentSource = pPaymentSource
	End Property
	Public Property Let PaymentSource(pVal)
		pPaymentSource = pVal
	End Property

	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		If (NOT IsEmpty(Price)) Then Data.Add "price", Price
		If (NOT IsEmpty(PaidPrice)) Then Data.Add "paidPrice", PaidPrice
		If (NOT IsEmpty(Installment)) Then Data.Add "installment", Installment
		If (NOT IsEmpty(PaymentChannel)) Then Data.Add "paymentChannel", PaymentChannel
		If (NOT IsEmpty(BasketId)) Then Data.Add "basketId", BasketId
		If (NOT IsEmpty(PaymentGroup)) Then Data.Add "paymentGroup", PaymentGroup
		
		If (NOT IsEmpty(PaymentCard)) Then Data.Add "paymentCard", PaymentCard.Data
		If (NOT IsEmpty(Buyer)) Then Data.Add "buyer", Buyer.Data
		If (NOT IsEmpty(ShippingAddress)) Then Data.Add "shippingAddress", ShippingAddress.Data
		If (NOT IsEmpty(BillingAddress)) Then Data.Add "billingAddress", BillingAddress.Data	
		If (NOT IsEmpty(BasketItems)) Then
			Dim pElements(), Cursor : Cursor = 0 : ReDim pElements(UBOUND(BasketItems))
			Dim pBasketItem
			For Each pBasketItem In BasketItems
				Set pElements(Cursor) = pBasketItem.Data
				Cursor = Cursor + 1
			Next
			Data.Add "basketItems", pElements
		End If
		If (NOT IsEmpty(CallbackUrl)) Then Data.Add "callbackUrl", CallbackUrl
	End Property

End Class

%>