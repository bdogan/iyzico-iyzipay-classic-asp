<%

Class mAddress

	Private pContactName
	Public Property Get ContactName
		ContactName = pContactName
	End Property
	Public Property Let ContactName(pVal)
		pContactName = pVal
	End Property
	
	Private pCity
	Public Property Get City
		City = pCity
	End Property
	Public Property Let City(pVal)
		pCity = pVal
	End Property
	
	Private pCountry
	Public Property Get Country
		Country = pCountry
	End Property
	Public Property Let Country(pVal)
		pCountry = pVal
	End Property
	
	Private pAddress
	Public Property Get Address
		Address = pAddress
	End Property
	Public Property Let Address(pVal)
		pAddress = pVal
	End Property
	
	Private pZipCode
	Public Property Get ZipCode
		ZipCode = pZipCode
	End Property
	Public Property Let ZipCode(pVal)
		pZipCode = pVal
	End Property
	
	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		If (NOT IsEmpty(Address)) Then Data.Add "address", Address
		If (NOT IsEmpty(ZipCode)) Then Data.Add "zipCode", ZipCode
		If (NOT IsEmpty(ContactName)) Then Data.Add "contactName", ContactName
		If (NOT IsEmpty(City)) Then Data.Add "city", City
		If (NOT IsEmpty(Country)) Then Data.Add "country", Country
	End Property

End Class

%>