<%
'实现整合系统数据与BBSXP数据的同步

Function APICheckUserName()
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey
	Set ApiObject = New BBSXPApi
		ApiObject.AppendNode "action","checkname",0,False
		ApiObject.AppendNode "username",UserName,1,False
		SafeKey = LCase(Mid(Md5(ApiObject.XmlNode("username")&SiteConfig("APISafeKey")),9,16))
		ApiObject.AppendNode "syskey",SafeKey,0,False
		ApiObject.SendPost
		If ApiObject.CheckStatus = "1" Then
			Response.write("<img src='images/check_error.gif' align=absmiddle />&nbsp;"&ApiObject.Message&"")
			Set ApiObject = Nothing
			Response.End()
		End If
	Set ApiObject = Nothing
End Function


Function APIAddUser()
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey,ApiPassword
	Set ApiObject = New BBSXPApi
	ApiObject.AppendNode "action","reguser",0,False
	ApiObject.AppendNode "username",UserName,1,False
	ApiObject.AppendNode "password",UserPassword,0,False
	ApiObject.AppendNode "email",UserEmail,1,False
	if PasswordAnswer<>empty then
		ApiObject.AppendNode "question",PasswordQuestion,1,False
		ApiObject.AppendNode "answer",PasswordAnswer,1,False
	end if
	
	SafeKey = LCase(Mid(Md5(ApiObject.XmlNode("username")&SiteConfig("APISafeKey")),9,16))
	ApiObject.AppendNode "syskey",SafeKey,0,False
	ApiObject.SendPost
	If ApiObject.CheckStatus = "1" Then
		Message = Message&ApiObject.Message
	Else
		ApiPassword = LCase(Mid(Md5(UserPassword),9,16))
		ApiSaveCookie = ApiObject.APISetCookie(SafeKey,UserName,ApiPassword,"3")
	End If
	Set ApiObject = Nothing
End Function


Function APIUpdateUser(UserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer)
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey
	Set ApiObject = New BBSXPApi
		ApiObject.AppendNode "action","update",0,False
		ApiObject.AppendNode "username",UserName,1,False
		SafeKey = LCase(Mid(Md5(ApiObject.XmlNode("username")&SiteConfig("APISafeKey")),9,16))
		ApiObject.AppendNode "syskey",SafeKey,0,False
		ApiObject.AppendNode "password",UserPassword,1,False
		If ""&UserEmail&""<>"" Then ApiObject.AppendNode "email",UserEmail,1,False
		If ""&PasswordAnswer&""<>"" Then
			ApiObject.AppendNode "question",PasswordQuestion,1,False
			ApiObject.AppendNode "answer",PasswordAnswer,1,False
		End If
		ApiObject.SendPost
		If ApiObject.CheckStatus = "1" Then
			APIUpdateUser = ApiObject.Message
		End If
	Set ApiObject = Nothing
End Function


Function APIDeleteUser()
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey
	Set ApiObject = New BBSXPApi
		ApiObject.AppendNode "action","delete",0,False
		ApiObject.AppendNode "username",UserName,1,False
		SafeKey = LCase(Mid(Md5(ApiObject.XmlNode("username")&SiteConfig("APISafeKey")),9,16))
		ApiObject.AppendNode "syskey",SafeKey,0,False
		ApiObject.SendPost
		If ApiObject.CheckStatus = "1" Then
			APIDeleteUser = ApiObject.Message
		End If
	Set ApiObject = Nothing
End Function


Function APIUserLogin()
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey,ApiPassword
	SafeKey = LCase(Mid(Md5(UserName&SiteConfig("APISafeKey")),9,16))
	ApiPassword = LCase(Mid(Md5(UserPassword),9,16))
	Set ApiObject = New BBSXPApi
		ApiSaveCookie = ApiObject.APISetCookie(SafeKey,UserName,ApiPassword,"3")
	Set ApiObject = Nothing
	Response.Write ApiSaveCookie
	Response.Flush
End Function


Function APIUserLoginOut()
	If SiteConfig("APIEnable")=0 Then Response.End()
	Dim ApiObject,SafeKey
	SafeKey = LCase(Mid(Md5(TempUserName&SiteConfig("APISafeKey")),9,16))
	Set ApiObject = New BBSXPApi
		ApiSaveCookie = ApiObject.APISetCookie(SafeKey,TempUserName,"","")
	Set ApiObject = Nothing
	Response.Write ApiSaveCookie
	Response.Flush
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Class BBSXPApi
	Public AppID,CheckStatus,GetData,GetAppid
	Private XmlDoc,XmlHttp,MessageCode,ArrUrls,SafeKey,XmlFilePath

	Public Property Get Message()
		Message = MessageCode
	End Property
	Public Property Get XmlNode(Byval Str)
		If XmlDoc.documentElement.selectSingleNode(Str) is Nothing Then
			XmlNode = "Null"
		Else
			XmlNode = XmlDoc.documentElement.selectSingleNode(Str).text
		End If
	End Property

	'--------------------------------------------------

	Private Sub Class_Initialize()
		GetAppid = ""
		AppID = "BBSXP2008"
		ArrUrls = Split(Trim(SiteConfig("APIUrlList")),"|")
		CheckStatus = "1"
		SafeKey = SiteConfig("APISafeKey")
		MessageCode = ""
		XmlFilePath = Server.MapPath("Xml/API_User.xml")
		Set XmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
		Set GetData = Server.Createobject("Scripting.Dictionary")
		XmlDoc.ASYNC = False
		LoadXmlData()
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(XmlDoc) Then Set XmlDoc = Nothing
		If IsObject(GetData) Then Set GetData = Nothing
	End Sub

	Public Sub LoadXmlData()
		If Not XmlDoc.Load(XmlFilePath) Then
			XmlDoc.LoadXml "<?xml version=""1.0"" encoding="""&BBSxpCharset&"""?><root/>"
		End If
		AppendNode "appID",AppID,1,False
	End Sub
	
	'--------------------------------------------------
	Public Sub AppendNode(Byval NodeName,Byval NodeText,Byval NodeType ,Byval blnEncode)
		Dim ChildNode,CreateCDATASection
		NodeName = Lcase(NodeName)
		If XmlDoc.documentElement.selectSingleNode(NodeName) is nothing Then
			Set ChildNode = XmlDoc.documentElement.appendChild(XmlDoc.createNode(1,NodeName,""))
		Else
			Set ChildNode = XmlDoc.documentElement.selectSingleNode(NodeName)
		End If
		If blnEncode = True Then
			NodeText = AnsiToUnicode(NodeText)
		End If
		If NodeType = 1 Then
			ChildNode.Text = ""
			Set CreateCDATASection = XmlDoc.createCDATASection(Replace(NodeText,"]]>","]]&gt;"))
			ChildNode.appendChild(createCDATASection)
		Else
			ChildNode.Text = NodeText
		End If
	End Sub

	'--------------------------------------------------
	Public Sub SendPost()
		Dim i,GetXmlDoc,LoadAppid
		Set Xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set GetXmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
		For i = 0 to Ubound(ArrUrls)
			XmlHttp.Open "POST", Trim(ArrUrls(i)), false
			XmlHttp.SetRequestHeader "content-type", "text/xml"
			XmlHttp.Send XmlDoc
			If GetXmlDoc.load(XmlHttp.responseXML) Then
				LoadAppid = Lcase(GetXmlDoc.documentElement.selectSingleNode("appid").Text)
				GetData.add LoadAppid,GetXmlDoc
				CheckStatus = GetXmlDoc.documentElement.selectSingleNode("status").Text
				MessageCode = MessageCode & LoadAppid & "：" & GetXmlDoc.documentElement.selectSingleNode("body/message").Text
				If CheckStatus = "1" Then '当发生错误时退出
					Exit For
				End If
			Else
				CheckStatus = "1"
				MessageCode = "HTTP错误 "&XmlHttp.Status&"！"
				Exit For
			End If
		Next
		Set GetXmlDoc = Nothing
		Set XmlHttp = Nothing
	End Sub

	'--------------------------------------------------
	'写COOKIE调用 
	'参数
	'API_SafeKey 密钥，API_UserName 用户名，API_Password 加密的用户密码 ，API_CookieExpired 保存COOKIE时间
	'--------------------------------------------------
	Public Function APISetCookie(Byval API_SafeKey,Byval API_UserName,Byval API_Password,Byval API_CookieExpired)
		Dim i,TempStr
		TempStr = ""
		For i = 0 to Ubound(ArrUrls)
			TempStr = TempStr & vbNewLine & "<script language=""JavaScript"" src="""&Trim(ArrUrls(i))&"?syskey="&Server.URLEncode(API_SafeKey)&"&username="&Server.URLEncode(API_UserName)&"&password="&Server.URLEncode(API_Password)&"&savecookie="&Server.URLEncode(API_CookieExpired)&"""></script>"
		Next
		APISetCookie = TempStr
	End Function

	'--------------------------------------------------
	Private Function AnsiToUnicode(ByVal str)
		Dim i, j, c, i1, i2, u, fs, f, p
		AnsiToUnicode = ""
		p = ""
		For i = 1 To Len(str)
			c = Mid(str, i, 1)
			j = AscW(c)
			If j < 0 Then
				j = j + 65536
			End If
			If j >= 0 And j <= 128 Then
				If p = "c" Then
					AnsiToUnicode = " " & AnsiToUnicode
					p = "e"
				End If
				AnsiToUnicode = AnsiToUnicode & c
			Else
				If p = "e" Then
					AnsiToUnicode = AnsiToUnicode & " "
					p = "c"
				End If
				AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
			End If
		Next
	End Function
End Class
%>