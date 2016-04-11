<%

Class rInstallmentInfo

	Public Path
	Public Method
	Public Sub Class_Initialize
		Path = "/payment/iyzipos/installment"
		Method = "POST"
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
		Price = pPrice
	End Property
	Public Property Let Price(pVal)
		pPrice = pVal
	End Property
	
	Public Property Get Hash
		Hash = "binNumber=" & BinNumber & ",price=" & Price
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		Data.Add "binNumber", BinNumber
		Data.Add "price", Price
	End Property

End Class

%>