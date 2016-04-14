<%

Class rBasketItem
	
	Private F
	Public Sub Class_Initialize
		Set F = New oIzyicoRequestFormatter
	End Sub 
	
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
	
	Private pCategory1
	Public Property Get Category1
		Category1 = pCategory1
	End Property
	Public Property Let Category1(pVal)
		pCategory1 = pVal
	End Property
	
	Private pCategory2
	Public Property Get Category2
		Category2 = pCategory2
	End Property
	Public Property Let Category2(pVal)
		pCategory2 = pVal
	End Property
	
	Private pItemType
	Public Property Get ItemType
		ItemType = pItemType
	End Property
	Public Property Let ItemType(pVal)
		pItemType = pVal
	End Property
	
	Private pPrice
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
		If (NOT IsEmpty(Id)) Then Data.Add "id", Id
		If (NOT IsEmpty(Price)) Then Data.Add "price", Price
		If (NOT IsEmpty(Name)) Then Data.Add "name", Name
		If (NOT IsEmpty(Category1)) Then Data.Add "category1", Category1
		If (NOT IsEmpty(Category2)) Then Data.Add "category2", Category2
		If (NOT IsEmpty(ItemType)) Then Data.Add "itemType", ItemType
	End Property

End Class

%>