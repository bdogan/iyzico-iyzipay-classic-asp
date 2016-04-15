<%

Class mBuyer

	Private pId
	Public Property Get Id
		Id = pId
	End Property
	Public Property Let Id(pVal)
		pId = pVal
	End Property
	
	Private pName
	Public Property Get Name
		Name = pName
	End Property
	Public Property Let Name(pVal)
		pName = pVal
	End Property
	
	Private pSurname
	Public Property Get Surname
		Surname = pSurname
	End Property
	Public Property Let Surname(pVal)
		pSurname = pVal
	End Property
	
	Private pGsmNumber
	Public Property Get GsmNumber
		GsmNumber = pGsmNumber
	End Property
	Public Property Let GsmNumber(pVal)
		pGsmNumber = pVal
	End Property
	
	Private pEmail
	Public Property Get Email
		Email = pEmail
	End Property
	Public Property Let Email(pVal)
		pEmail = pVal
	End Property
	
	Private pIdentityNumber
	Public Property Get IdentityNumber
		IdentityNumber = pIdentityNumber
	End Property
	Public Property Let IdentityNumber(pVal)
		pIdentityNumber = pVal
	End Property
	
	Private pLastLoginDate
	Public Property Get LastLoginDate
		LastLoginDate = pLastLoginDate
	End Property
	Public Property Let LastLoginDate(pVal)
		pLastLoginDate = pVal
	End Property
	
	Private pRegistrationDate
	Public Property Get RegistrationDate
		RegistrationDate = pRegistrationDate
	End Property
	Public Property Let RegistrationDate(pVal)
		pRegistrationDate = pVal
	End Property
	
	Private pRegistrationAddress
	Public Property Get RegistrationAddress
		RegistrationAddress = pRegistrationAddress
	End Property
	Public Property Let RegistrationAddress(pVal)
		pRegistrationAddress = pVal
	End Property
	
	Private pIp
	Public Property Get Ip
		Ip = pIp
	End Property
	Public Property Let Ip(pVal)
		pIp = pVal
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
	
	Public Property Get Hash
		Hash = Iyzico.GenerateHashFromData(Data)
	End Property
	
	Public Property Get Data
		Set Data = Server.CreateObject("Scripting.Dictionary")
		If (NOT IsEmpty(Id)) Then Data.Add "id", Id
		If (NOT IsEmpty(Name)) Then Data.Add "name", Name
		If (NOT IsEmpty(Surname)) Then Data.Add "surname", Surname
		If (NOT IsEmpty(IdentityNumber)) Then Data.Add "identityNumber", IdentityNumber
		If (NOT IsEmpty(Email)) Then Data.Add "email", Email
		If (NOT IsEmpty(GsmNumber)) Then Data.Add "gsmNumber", GsmNumber
		If (NOT IsEmpty(RegistrationDate)) Then Data.Add "registrationDate", RegistrationDate
		If (NOT IsEmpty(LastLoginDate)) Then Data.Add "lastLoginDate", LastLoginDate
		If (NOT IsEmpty(RegistrationAddress)) Then Data.Add "registrationAddress", RegistrationAddress
		If (NOT IsEmpty(City)) Then Data.Add "city", City
		If (NOT IsEmpty(Country)) Then Data.Add "country", Country
		If (NOT IsEmpty(Ip)) Then Data.Add "ip", Ip
	End Property

End Class

%>