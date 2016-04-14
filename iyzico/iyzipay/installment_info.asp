<%

Class rInstallmentInfo

	Public Path
	Public Method
	Private F
	Public Sub Class_Initialize
		Path = "/payment/iyzipos/installment"
		Method = "POST"
		Set F = New oIzyicoRequestFormatter
	End Sub 

	Private pBinNumber
	Public Property Get BinNumber
		BinNumber = pBinNumber
	End Property
	Public Property Let BinNumber(pVal)
		pBinNumber = pVal
	End Property
	
	Public pPrice
	Public Property Get Price
		Price = F.FormatPrice(pPrice)
	End Property
	Public Property Let Price(pVal)
		pPrice = pVal
	End Property
	
	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		Data.Add "binNumber", BinNumber
		If (NOT IsEmpty(Price)) Then Data.Add "price", Price
	End Property

End Class

%>