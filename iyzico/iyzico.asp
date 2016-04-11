<!--#include file="iyzipay/installment_info.asp" -->
<!--#include file="lib/aspJSON1.17.asp" -->
<!--#include file="lib/base64encoder.asp" -->
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
	
	Public Property Get IsAspList(ByVal pDict)
		IsAspList = (TypeName(pDict) = "AspList")
	End Property
			
	Public Property Get IsAspJson(ByVal pJson)
		IsAspJson = (TypeName(pJson) = "aspJSON")
	End Property
	
	Public Property Get IsRequestCollection(ByVal pReq)
		IsRequestCollection = (TypeName(pReq) = "IRequestDictionary")
	End Property

	Public Property Get IsApplicationObj(ByVal pReq)
		IsApplicationObj = (TypeName(pReq) = "IApplicationObject")
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
		ElseIf (IsDictionary(pObj) OR IsAspList(pObj)) Then
			Cursor = 0 : ReDim resultArr(pObj.Count)
			For Each dKey In pObj.Keys
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & dKey & "] => " & Replace(DebugStr(pObj(dKey)), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = TypeName(pObj) & " (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsAspJson(pObj)) Then
			Cursor = 0 : ReDim resultArr(pObj.data.Count)
			For Each dKey In pObj.data.Keys
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & dKey & "] => " & Replace(DebugStr(pObj.data(dKey)), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = "aspJSON (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsRequestCollection(pObj)) Then
			Cursor = 0 : ReDim resultArr(pObj.Count)
			For Each dKey In pObj
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & dKey & "] => " & Replace(DebugStr(pObj(dKey)), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = "IRequestDictionary (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsType(pObj, "Files")) Then
			Cursor = 0 : ReDim resultArr(pObj.Count)
			For Each dKey In pObj
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & Cursor & "] => " & Replace(DebugStr(dKey), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = "Files (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
		ElseIf (IsApplicationObj(pObj)) Then
			Cursor = 0 : ReDim resultArr(pObj.Contents.Count)
			For Each dKey In pObj.Contents
				resultArr(Cursor) = resultArr(Cursor) & VBTAB & "[" & dKey & "] => " & Replace(DebugStr(pObj(dKey)), VBCRLF, VBCRLF & VBTAB)
				Cursor = Cursor + 1
			Next
			DebugStr = "IRequestDictionary (" & VBCRLF & Join(resultArr, VBCRLF) & ")"
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
	
	Private Property Get RequestHeaders(pRequest)
		Set RequestHeaders = Server.CreateObject("Scripting.Dictionary")
		RequestHeaders.Add "Accept", "application/json"
		RequestHeaders.Add "Content-Type", "application/json"
		
		Dim pRand : pRand = RandomString(7, 7)
		
		RequestHeaders.Add "Authorization", AuthString(pRequest, pRand)
		RequestHeaders.Add "x-iyzi-rnd", pRand
	End Property
	
	Private Property Get AuthString(pRequest, pRand)
		AuthString = "IYZWS " & pOptions.ApiKey & ":" & b64_sha1(pOptions.ApiKey & pRand & pOptions.SecretKey & GenerateRequestHash(pRequest))
	End Property
	
	Private Function b64_sha1(pVal)
		Set oEncoding = CreateObject("System.Text.UTF8Encoding")
		Set oCrypt = Server.CreateObject("System.Security.Cryptography.SHA1Managed")
		Dim aBytes : aBytes = oEncoding.GetBytes_4(pVal)
		Dim aBinResult : aBinResult = oCrypt.ComputeHash_2((aBytes))
		b64_sha1 = base64_encode(BinaryToString(aBinResult)) & "="
		Set oEncoding = Nothing
		Set oCrypt = Nothing
	End Function
	
	Private Property Get BinaryToString(Binary)
	  'Antonin Foller, http://www.motobit.com
	  'Optimized version of a simple BinaryToString algorithm.
	  
	  Dim cl1, cl2, cl3, pl1, pl2, pl3
	  Dim L
	  cl1 = 1
	  cl2 = 1
	  cl3 = 1
	  L = LenB(Binary)
	  
	  Do While cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		If cl3>300 Then
		  pl2 = pl2 & pl3
		  pl3 = ""
		  cl3 = 1
		  cl2 = cl2 + 1
		  If cl2>200 Then
			pl1 = pl1 & pl2
			pl2 = ""
			cl2 = 1
		  End If
		End If
	  Loop
	  BinaryToString = pl1 & pl2 & pl3
	End Property

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

Class izBase64Encoder
	
	Public Property Get FromString(sText)
		On Error Resume Next
		Dim oXML, oNode
		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")
		Set oNode = oXML.CreateElement("base64")
		oNode.dataType = "bin.base64"
		oNode.nodeTypedValue = Stream_StringToBinary(sText)
		FromString = oNode.text
		Set oNode = Nothing
		Set oXML = Nothing
		If (NOT Utils.CheckError) Then Exit Property
		Err.Clear
		FromString = Empty
	End Property

	Public Property Get ToString(ByVal vCode)
		On Error Resume Next
		Dim oXML, oNode
		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")
		Set oNode = oXML.CreateElement("base64")
		oNode.dataType = "bin.base64"
		oNode.text = vCode
		ToString = Stream_BinaryToString(oNode.nodeTypedValue)
		Set oNode = Nothing
		Set oXML = Nothing
		If (NOT Utils.CheckError) Then Exit Property
		Err.Clear
		ToString = Empty
	End Property

	'Stream_StringToBinary Function
	'2003 Antonin Foller, http://www.motobit.com
	'Text - string parameter To convert To binary data
	Function Stream_StringToBinary(Text)
		Const adTypeText = 2
		Const adTypeBinary = 1

		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = adTypeText
		BinaryStream.CharSet = "ISO-8859-9"
		BinaryStream.Open
		BinaryStream.WriteText Text
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeBinary
		BinaryStream.Position = 0
		Stream_StringToBinary = BinaryStream.Read
		Set BinaryStream = Nothing
	End Function

	'Stream_BinaryToString Function
	'2003 Antonin Foller, http://www.motobit.com
	'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
	Private Function Stream_BinaryToString(Binary)
		Const adTypeText = 2
		Const adTypeBinary = 1
		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = adTypeBinary
		BinaryStream.Open
		BinaryStream.Write Binary
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeText
		BinaryStream.CharSet = "ISO-8859-9"
		Stream_BinaryToString = BinaryStream.ReadText
		Set BinaryStream = Nothing
	End Function

End Class

%>

<script language="javascript" type="text/javascript" runat="server">
var Base64={_keyStr:"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",encode:function(e){var t="";var n,r,i,s,o,u,a;var f=0;e=Base64._utf8_encode(e);while(f<e.length){n=e.charCodeAt(f++);r=e.charCodeAt(f++);i=e.charCodeAt(f++);s=n>>2;o=(n&3)<<4|r>>4;u=(r&15)<<2|i>>6;a=i&63;if(isNaN(r)){u=a=64}else if(isNaN(i)){a=64}t=t+this._keyStr.charAt(s)+this._keyStr.charAt(o)+this._keyStr.charAt(u)+this._keyStr.charAt(a)}return t},decode:function(e){var t="";var n,r,i;var s,o,u,a;var f=0;e=e.replace(/[^A-Za-z0-9+/=]/g,"");while(f<e.length){s=this._keyStr.indexOf(e.charAt(f++));o=this._keyStr.indexOf(e.charAt(f++));u=this._keyStr.indexOf(e.charAt(f++));a=this._keyStr.indexOf(e.charAt(f++));n=s<<2|o>>4;r=(o&15)<<4|u>>2;i=(u&3)<<6|a;t=t+String.fromCharCode(n);if(u!=64){t=t+String.fromCharCode(r)}if(a!=64){t=t+String.fromCharCode(i)}}t=Base64._utf8_decode(t);return t},_utf8_encode:function(e){e=e.replace(/rn/g,"n");var t="";for(var n=0;n<e.length;n++){var r=e.charCodeAt(n);if(r<128){t+=String.fromCharCode(r)}else if(r>127&&r<2048){t+=String.fromCharCode(r>>6|192);t+=String.fromCharCode(r&63|128)}else{t+=String.fromCharCode(r>>12|224);t+=String.fromCharCode(r>>6&63|128);t+=String.fromCharCode(r&63|128)}}return t},_utf8_decode:function(e){var t="";var n=0;var r=c1=c2=0;while(n<e.length){r=e.charCodeAt(n);if(r<128){t+=String.fromCharCode(r);n++}else if(r>191&&r<224){c2=e.charCodeAt(n+1);t+=String.fromCharCode((r&31)<<6|c2&63);n+=2}else{c2=e.charCodeAt(n+1);c3=e.charCodeAt(n+2);t+=String.fromCharCode((r&15)<<12|(c2&63)<<6|c3&63);n+=3}}return t}}
</script>