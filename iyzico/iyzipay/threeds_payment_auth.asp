<%

Class rThreeDsPaymentAuth

	Public Path
	Public Method
	Public Sub Class_Initialize
		Path = "/payment/iyzipos/auth3ds/ecom"
		Method = "POST"	
	End Sub 
	
	Private pPaymentId
	Public Property Get PaymentId
		PaymentId = pPaymentId
	End Property
	Public Property Let PaymentId(pVal)
		pPaymentId = pVal
	End Property

	Private pConversationData
	Public Property Get ConversationData
		ConversationData = pConversationData
	End Property
	Public Property Let ConversationData(pVal)
		pConversationData = pVal
	End Property
	
	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		If (NOT IsEmpty(PaymentId)) Then Data.Add "paymentId", PaymentId
		If (NOT IsEmpty(PaymentConversationId)) Then Data.Add "paymentConversationId", PaymentConversationId
		If (NOT IsEmpty(ConversationData)) Then Data.Add "conversationData", ConversationData
	End Property

End Class

%>