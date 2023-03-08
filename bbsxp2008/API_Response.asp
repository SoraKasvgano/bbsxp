<!--#include file="Conn.asp"-->
<%
dim XmlDoc,Message,UserName,Menu,Appid,CheckStatus
Message=""
CheckStatus = 1

'定义全局变量
Dim UserID,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer,UserInfoList,IsMD5,Expired

Set XmlDoc=Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
XmlDoc.async = False


If SiteConfig("APIEnable")=0 Then
	CheckStatus = 1
	Message = "BBSXP系统整合接口未开启！"
	ResponseData()
	Set XmlDoc=nothing
ElseIf Request.QueryString<>"" Then
	APIResponseCookies()
Else
	If Not XmlDoc.Load(Request) Then
		CheckStatus = 1
		Message = "数据非法，操作中止！"
		Appid = "未知"
	Else
		If CheckSafeKey() Then
			Select Case Menu
				Case "checkname"
					APICheckUser()
				Case "reguser"
					APIAddUser()
				Case "login"
					APILogin()
				Case "logout"
					APILoginOut()
				Case "update"
					APIEditUser()
				Case "delete"
					APIDelUser()
				Case else
			End Select
		End If
	End If
	ResponseData()
	Set XmlDoc=nothing
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ResponseData()
	If Menu <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>BBSXP</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "BBSXP"
	XmlDoc.documentElement.selectSingleNode("status").text = CheckStatus
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Message,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet=BBSxpCharset
	Response.Write "<?xml version=""1.0"" encoding="""&BBSxpCharset&"""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CheckSafeKey()
	CheckSafeKey = False
	Dim SafeKey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username") is Nothing Then
		CheckStatus = 1
		Message = Message & "<li>非法请求。"
		Exit Function
	End If

	UserName = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("username").text))
	SafeKey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Menu = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = XmlDoc.documentElement.selectSingleNode("appid").text
	
	If SiteConfig("APISafeKey") = "APITEST" or SafeKey = "SafeKey" Then
		CheckStatus = 1
		Message = Message & "默认非法请求。"
		Exit Function
	End If
	
	Dim NewSafeKey
	NewSafeKey = LCase(Mid(Md5(UserName&SiteConfig("APISafeKey")),9,16))
	If SafeKey=NewSafeKey Then
		CheckSafeKey = True
	Else
		CheckStatus = 1
		Message = Message & "安全密钥验证失败，请与管理员联系。"
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub APICheckUser()
	Dim i
	
	Message=CheckUser(UserName)
	
	If Message<>"" Then
		CheckStatus = 1
		Exit Sub
	End If
	
	
	Sql="Select * From ["&TablePrefix&"Users] Where UserName='"&UserName&"'"
	Set Rs = Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		Message = "您填写的用户名已经被注册。"
		CheckStatus = 1
		Exit Sub
	Else
		CheckStatus = 0
		Message = "验证通过。"
	End If
	Rs.Close
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APIAddUser()
	Dim i
	UserPassword = HTMLEncode(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	PasswordQuestion = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("question").text))
	PasswordAnswer = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("answer").text))

	if AddUser=False then
		CheckStatus = 1
		Exit Sub
	end if

	if SiteConfig("AccountActivation")=0 or SiteConfig("AccountActivation")=2 then
		ResponseCookies "UserID",UserID,"999"
		ResponseCookies "UserPassword",UserPassword,"999"
	end if
	
	Message = "注册成功。"
	CheckStatus = 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APIEditUser()
	UserPassword = HTMLEncode(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	PasswordQuestion = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("question").text))
	PasswordAnswer = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("answer").text))
		
	if ModifyUserPassword(UserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer)=False Then
		CheckStatus = 1
	else
		CheckStatus = 0
		Messenge = "基本资料修改成功。"
	end if
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APIDelUser()
	Rs.open "Select UserName from ["&TablePrefix&"Users] where UserName='"&UserName&"'",Conn,1,3
	if Rs.eof then
		CheckStatus = 1
		Message="操作不成功，用户资料不存在！"
	else
		Rs.delete
		Rs.update
		CheckStatus = 0
		Message="用户资料删除成功！"
	end if
	Rs.Close
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APILogin()
	Dim SaveCookie
	UserPassword = HTMLEncode(XmlDoc.documentElement.selectSingleNode("password").text)
	Expired = XmlDoc.documentElement.selectSingleNode("savecookie").text
	If UserName="" or UserPassword="" Then
		CheckStatus = 1
		Message = Message & "请输入用户名或密码。"
		Exit Sub
	End If

	If Expired="" or Not IsNumeric(Expired) Then Expired = 0

	If UserLogin=True then
		CheckStatus = 0
		Message="已经成功登录!"
	else
		CheckStatus = 1
	end if
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APILoginOut()
		Execute("Delete from ["&TablePrefix&"UserOnline] where sessionid='"&session.sessionid&"'")
		CleanCookies()
		Message="已经成功退出!"
End Sub

Sub APIResponseCookies()
	Dim SafeKey,NewSafeKey
	SafeKey = Request.QueryString("syskey")
	UserName = HTMLEncode(Request.QueryString("username"))
	UserName=unescape(UserName)		'修正中文编码问题,否则中文用户验证不通过。
	UserPassword = Request.QueryString("password")
	Expired = Request.QueryString("savecookie")
	IsMD5=True
	
	If UserName="" or SafeKey="" Then Exit Sub
	
	NewSafeKey = LCase(Mid(Md5(UserName&SiteConfig("APISafeKey")),9,16))
	If SafeKey<>NewSafeKey Then Exit Sub
	
	If Expired="" or Not IsNumeric(Expired) Then Expired = 0
	
	
	If UserPassword = "" Then	'用户退出
		UserLoginOut
	Else						'用户登陆
		UserLogin
	End If
End Sub
%>