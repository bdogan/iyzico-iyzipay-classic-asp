<%

Class rPaymentCard
	
	Private pCardHolderName
	Public Property Get CardHolderName
		CardHolderName = pCardHolderName
	End Property
	Public Property Let CardHolderName(pVal)
		pCardHolderName = pVal
	End Property
	
	Private pCardNumber
	Public Property Get CardNumber
		CardNumber = pCardNumber
	End Property
	Public Property Let CardNumber(pVal)
		pCardNumber = pVal
	End Property
	
	Private pExpireMonth
	Public Property Get ExpireMonth
		ExpireMonth = pExpireMonth
	End Property
	Public Property Let ExpireMonth(pVal)
		pExpireMonth = pVal
	End Property
	
	Private pExpireYear
	Public Property Get ExpireYear
		ExpireYear = pExpireYear
	End Property
	Public Property Let ExpireYear(pVal)
		pExpireYear = pVal
	End Property
	
	Private pCvc
	Public Property Get Cvc
		Cvc = pCvc
	End Property
	Public Property Let Cvc(pVal)
		pCvc = pVal
	End Property
	
	Private pRegisterCard
	Public Property Get RegisterCard
		If (IsEmpty(pRegisterCard)) Then pRegisterCard = "0"
		RegisterCard = pRegisterCard
	End Property
	Public Property Let RegisterCard(pVal)
		pRegisterCard = pVal
	End Property
	
	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		If (NOT IsEmpty(CardHolderName)) Then Data.Add "cardHolderName", CardHolderName
		If (NOT IsEmpty(CardNumber)) Then Data.Add "cardNumber", CardNumber
		If (NOT IsEmpty(ExpireYear)) Then Data.Add "expireYear", ExpireYear
		If (NOT IsEmpty(ExpireMonth)) Then Data.Add "expireMonth", ExpireMonth
		If (NOT IsEmpty(Cvc)) Then Data.Add "cvc", Cvc
		If (NOT IsEmpty(RegisterCard)) Then Data.Add "registerCard", RegisterCard
	End Property

End Class

%>