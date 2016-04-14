<!--#include file="iyzipay/installment_info.asp" -->
<!--#include file="lib/aspJSON1.17.asp" -->
<%

Class oIyzico
	
	Public Sub pr(ByVal pObj)
		Response.Flush
		Response.Write "<pre>"
		Response.Write DebugStr(pObj)
		Response.Write "</pre>"
		Response.Flush
	End Sub	
	
	Public Property Get IsType(ByRef pVal, ByRef pTypeName)
		IsType = (TypeName(pVal) = pTypeName)
	End Property
	
	Public Property Get IsString(ByVal pStr)
		IsString = (TypeName(pStr) = "String")
	End Property
	
	Public Property Get IsArray(ByVal pArr)
		IsArray = (TypeName(pArr) = "Variant()")
	End Property
		
	Public Property Get IsDictionary(ByVal pDict)
		IsDictionary = (TypeName(pDict) = "Dictionary")
	End Property
	
	Public Property Get IsAspJson(ByVal pJson)
		IsAspJson = (TypeName(pJson) = "aspJSON")
	End Property

	Public Property Get IsBoolean(ByVal pReq)
		IsBoolean = (TypeName(pReq) = "Boolean")
	End Property
	
	Public Property Get DebugStr(ByVal pObj)
		Dim dKey, Cursor, resultArr, arrObj
		If (IsNull(pObj)) Then pObj = "" : Exit Property
		If (IsArray(pObj)) Then
			Cursor = 0 : ReDim resultArr(ArrayCount(pObj))
			For Each arrObj In pObj
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & Cursor & "] => " & Replace(DebugStr(arrObj), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = "Array (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsDictionary(pObj)) Then
			Cursor = 0 : ReDim resultArr(pObj.Count)
			For Each dKey In pObj.Keys
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & dKey & "] => " & Replace(DebugStr(pObj(dKey)), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = TypeName(pObj) & " (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsObject(pObj)) Then
			On Error Resume Next
			DebugStr = Cstr(pObj)
			If (CheckError) Then DebugStr = "[Object] " & TypeName(pObj)
			Err.Clear
		Else
			DebugStr = Server.HTMLEncode(pObj)
		End If
	End Property
	
	Private pLocale
	Public Property Get Locale
		If (IsEmpty(pLocale)) Then pLocale = "tr"
		Locale = pLocale
	End Property
	Public Property Let Locale(pVal)
		pLocale = pVal
	End Property
	
	Private pConversationId
	Public Property Get ConversationId
		If (IsEmpty(pConversationId)) Then pConversationId = RandomString(4, 4)
		ConversationId = pConversationId
	End Property
	Public Property Let ConversationId(pVal)
		pConversationId = pVal
	End Property
	
	Private pOptions
	
	' Create iyzipay request
	Public Property Get CreateRequest(pName, vOptions)
		Set pOptions = vOptions
		Set CreateRequest = Eval("New r" & pName)
	End Property
	
	' Create iyzipay options for general request
	Public Property Get CreateOptions(pApiKey, pSecretKey, pBaseUrl)
		Set CreateOptions = New oIyzicoOptions
		CreateOptions.ApiKey = pApiKey
		CreateOptions.SecretKey = pSecretKey
		CreateOptions.BaseUrl = pBaseUrl
	End Property
	
	Public Property Get GenerateHashFromData(pData)
		Dim pKey, pHashValues(), Cursor : Cursor = 0 : ReDim pHashValues(pData.Count - 1)
		For Each pKey In pData.Keys
			pHashValues(Cursor) = pKey & "=" & pData.Item(pKey)
			Cursor = Cursor + 1
		Next
		GenerateHashFromData = Join(pHashValues, ",")
	End Property
	
	Private Property Get RequestHeaders(pRequest)
		Set RequestHeaders = Server.CreateObject("Scripting.Dictionary")
		RequestHeaders.Add "Accept", "application/json"
		RequestHeaders.Add "Content-Type", "application/json"
		
		Dim pRand : pRand = RandomString(7, 7)
		
		RequestHeaders.Add "Authorization", AuthString(pRequest, pRand)
		RequestHeaders.Add "x-iyzi-rnd", pRand
	End Property
	
	Private Property Get AuthString(pRequest, pRand)
		AuthString = "IYZWS " & pOptions.ApiKey & ":" & base64_sha1(pOptions.ApiKey & pRand & pOptions.SecretKey & GenerateRequestHash(pRequest))
	End Property
	
	Private Function base64_sha1(pVal)
		pr(pVal)
		Set oCrypt = Server.CreateObject("System.Security.Cryptography.SHA1Managed")
		base64_sha1 = Base64Encode(oCrypt.ComputeHash_2(ToBytes(pVal, "ascii")), "ascii")
		Set oCrypt = Nothing
	End Function
	
	Private Property Get ToBytes(pStr, pEncoding)
		Dim objStrm : Set objStrm = CreateObject("ADODB.Stream")
		objStrm.Open
		objStrm.Type = 2
		objStrm.CharSet = pEncoding
		objStrm.WriteText pStr
		objStrm.Position = 0
		objStrm.Type = 1
		ToBytes = objStrm.Read
		Set objStrm = Nothing
	End Property
	
	Private Function Base64Encode(pStr, pEncoding)
		Dim objBase64 : Set objBase64 = CreateObject("System.Security.Cryptography.ToBase64Transform")
		Dim i_size : i_size = objBase64.InputBlockSize
		Dim o_size : o_size = objBase64.OutputBlockSize
		Dim n_block
		
		Dim objStrm : Set objStrm = CreateObject("ADODB.Stream")
		objStrm.Open
		objStrm.Type = 1
		
		Dim bytes : bytes = pStr
		
		If (LenB(bytes) Mod i_size = 0) Then n_block = LenB(bytes) / i_size Else n_block = LenB(bytes) \ i_size + 1
		
		Dim i
		For i = 0 To n_block - 1
			Dim b_len : If LenB(bytes) < (i + 1) * i_size Then b_len = LenB(bytes) - i * i_size Else b_len = i_size
			Dim data : data = objBase64.TransformFinalBlock((bytes), i * i_size, b_len)
			objStrm.Write data
		Next
		
		objStrm.Position = 0
		objStrm.Type = 2
		objStrm.CharSet = pEncoding
		Base64Encode = objStrm.ReadText
		Set objStrm = Nothing
		Set objBase64 = Nothing
	End Function

	Private Property Get GenerateRequestHash(pRequest)
		GenerateRequestHash = "[locale=" & Locale & ",conversationId=" & ConversationId
		Dim innerHash : innerHash = pRequest.Hash
		If (Len(innerHash) > 0) Then GenerateRequestHash = GenerateRequestHash & "," & innerHash
		GenerateRequestHash = GenerateRequestHash & "]"
	End Property
	
	Private Property Get RequestData(pRequest)
		Set RequestData = Server.CreateObject("Scripting.Dictionary")
		RequestData.Add "locale", Locale
		RequestData.Add "conversationId", ConversationId
		Dim InnerRequestData, InnerKey : Set InnerRequestData = pRequest.Data
		For Each InnerKey In InnerRequestData.Keys
			RequestData.Add	InnerKey, InnerRequestData.Item(InnerKey)
		Next
	End Property
	
	Public Property Get CreateResponse(pRequest)
		Set CreateResponse = GetResponse(pRequest.Method, pOptions.BaseUrl & pRequest.Path, RequestData(pRequest), RequestHeaders(pRequest))
	End Property
	
	Public Property Get GetResponse(pMethod, pUrl, pData, pHeader)
		If (IsEmpty(pMethod)) Then pMethod = "GET"
		If (IsEmpty(pHeader)) Then Set pHeader = Server.CreateObject("Scripting.Dictionary")
		
		Dim pRequestBody : pRequestBody = ""
		If (pMethod = "POST" OR pMethod = "PUT") Then
			If (TypeName(pData) = "Dictionary") Then 
				Dim oJSON : Set oJSON = New aspJSON
				Set oJSON.data = pData
				pRequestBody = oJSON.JSONoutput()
				Set oJSON = Nothing
			Else
				pRequestBody = pData
			End If
		Else
			If (TypeName(pData) = "Dictionary") Then 
				Dim qs : qs = ""
				If (pData.Count > 0) Then
					Dim pKeyValues(), pKey, Cursor : ReDim pKeyValues(pData.Count - 1) : Cursor = 0
					For Each pKey In pData.Keys
						pKeyValues(Cursor) = "pKey=" & Server.URLEncode(pData.Item(pKey))
						Cursor = Cursor + 1
					Next
					qs = Join(pKeyValues, "&")
				End If
				If (InStr("?", pUrl) > 0) Then pUrl = pUrl & "&" & qs Else pUrl = pUrl & "?" & qs
			End If
		End If
		Dim objXML : Set objXML = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
		objXML.setTimeouts 100000, 100000, 200000, 200000
		objXML.Open pMethod, pUrl, False
		If (pHeader.Count > 0) Then
			Dim pHeaderKey
			For Each pHeaderKey In pHeader.Keys
				objXML.setRequestHeader pHeaderKey, pHeader.Item(pHeaderKey)
			Next
		End If
		objXML.Send pRequestBody
		Set GetResponse = ParseResponse(objXML.ResponseText)
		Set objXML = Nothing
	End Property
	
	Private Property Get ParseResponse(pResponseText)
		On Error Resume Next
		Dim oJSON : Set oJSON = New aspJSON
		oJSON.loadJSON(pResponseText)
		Set ParseResponse = oJSON.data
		Set oJSON = Nothing
		If (Err.Number <> 0) Then Set ParseResponse = Server.CreateObject("Scripting.Dictionary") : Err.Clear
	End Property
	
	Private Property Get RandomString(stringCount, integerCount)
		Randomize()
		Dim CharacterSetArray
		CharacterSetArray = Array(_
			Array(stringCount, "abcdefghijklmnopqrstuvwxyz"), _
			Array(integerCount, "0123456789") _
		)
		dim i
		dim j
		dim Count
		dim Chars
		dim Index
		dim Temp
		for i = 0 to UBound(CharacterSetArray)
			Count = CharacterSetArray(i)(0)
			Chars = CharacterSetArray(i)(1)
			for j = 1 to Count
				Index = Int(Rnd() * Len(Chars)) + 1
				Temp = Temp & Mid(Chars, Index, 1)
			next
		next
		dim TempCopy
		do until Len(Temp) = 0
			Index = Int(Rnd() * Len(Temp)) + 1
			TempCopy = TempCopy & Mid(Temp, Index, 1)
			Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)
		loop
		RandomString = TempCopy
	End Property
	
End Class

Class oIyzicoOptions
	
	Private pApiKey
	Public Property Get ApiKey
		ApiKey = pApiKey
	End Property
	Public Property Let ApiKey(pVal)
		pApiKey = pVal
	End Property
	
	Public pSecretKey
	Public Property Get SecretKey
		SecretKey = pSecretKey
	End Property
	Public Property Let SecretKey(pVal)
		pSecretKey = pVal
	End Property	
	
	Public pBaseUrl
	Public Property Get BaseUrl
		BaseUrl = pBaseUrl
	End Property
	Public Property Let BaseUrl(pVal)
		pBaseUrl = pVal
	End Property
	
End Class

Class oIzyicoRequestFormatter

	Private pRegEx
	Private Property Get RegEx
		If (IsEmpty(pRegEx)) Then Set pRegEx = New RegExp
		Set RegEx = pRegEx
	End Property
	
	Private Property Get RegexReplace(ByRef pRegEx, ByRef pHaystack, ByRef pNeedle)
		With RegEx
			.Pattern = pRegEx
			.IgnoreCase = True
			.Global = True
			.MultiLine = True
		End With
		RegexReplace = regEx.Replace(pHaystack, pNeedle)
	End Property

	Public Property Get FormatPrice(ByVal pPrice)
		If (IsEmpty(pPrice)) Then FormatPrice = Empty : Exit Property
		pPrice = Cstr(pPrice)
		If (InStr(pPrice, ".") > 0 AND InStr(pPrice, ",") > 0) Then
			If (InStr(pPrice, ",") > InStr(pPrice, ".")) Then
				pPrice = Replace(pPrice, ".", "")
				pPrice = Replace(pPrice, ",", ".")
			Else
				pPrice = Replace(pPrice, ",", "")
			End If
		Else
			If (InStr(pPrice, ",") > 0) Then
				pPrice = Replace(pPrice, ",", ".")
			End If
		End If
		If (NOT IsNumeric(pPrice)) Then FormatPrice = Empty : Exit Property
		
		If (InStr(pPrice, ".") = 0) Then 
			pPrice = pPrice & ".0"
		Else
			pPrice = RegexReplace("0+$", pPrice, "")
			pPrice = RegexReplace("(.*)\.$", pPrice, "$1.0")
		End If
		FormatPrice = pPrice
	End Property

End Class

%>