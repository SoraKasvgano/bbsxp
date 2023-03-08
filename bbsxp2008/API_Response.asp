<!--#include file="Conn.asp"-->
<%
dim XmlDoc,Message,UserName,Menu,Appid,CheckStatus
Message=""
CheckStatus = 1

'����ȫ�ֱ���
Dim UserID,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer,UserInfoList,IsMD5,Expired

Set XmlDoc=Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
XmlDoc.async = False


If SiteConfig("APIEnable")=0 Then
	CheckStatus = 1
	Message = "BBSXPϵͳ���Ͻӿ�δ������"
	ResponseData()
	Set XmlDoc=nothing
ElseIf Request.QueryString<>"" Then
	APIResponseCookies()
Else
	If Not XmlDoc.Load(Request) Then
		CheckStatus = 1
		Message = "���ݷǷ���������ֹ��"
		Appid = "δ֪"
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
		Message = Message & "<li>�Ƿ�����"
		Exit Function
	End If

	UserName = HTMLEncode(Trim(XmlDoc.documentElement.selectSingleNode("username").text))
	SafeKey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Menu = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = XmlDoc.documentElement.selectSingleNode("appid").text
	
	If SiteConfig("APISafeKey") = "APITEST" or SafeKey = "SafeKey" Then
		CheckStatus = 1
		Message = Message & "Ĭ�ϷǷ�����"
		Exit Function
	End If
	
	Dim NewSafeKey
	NewSafeKey = LCase(Mid(Md5(UserName&SiteConfig("APISafeKey")),9,16))
	If SafeKey=NewSafeKey Then
		CheckSafeKey = True
	Else
		CheckStatus = 1
		Message = Message & "��ȫ��Կ��֤ʧ�ܣ��������Ա��ϵ��"
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
		Message = "����д���û����Ѿ���ע�ᡣ"
		CheckStatus = 1
		Exit Sub
	Else
		CheckStatus = 0
		Message = "��֤ͨ����"
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
	
	Message = "ע��ɹ���"
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
		Messenge = "���������޸ĳɹ���"
	end if
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APIDelUser()
	Rs.open "Select UserName from ["&TablePrefix&"Users] where UserName='"&UserName&"'",Conn,1,3
	if Rs.eof then
		CheckStatus = 1
		Message="�������ɹ����û����ϲ����ڣ�"
	else
		Rs.delete
		Rs.update
		CheckStatus = 0
		Message="�û�����ɾ���ɹ���"
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
		Message = Message & "�������û��������롣"
		Exit Sub
	End If

	If Expired="" or Not IsNumeric(Expired) Then Expired = 0

	If UserLogin=True then
		CheckStatus = 0
		Message="�Ѿ��ɹ���¼!"
	else
		CheckStatus = 1
	end if
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub APILoginOut()
		Execute("Delete from ["&TablePrefix&"UserOnline] where sessionid='"&session.sessionid&"'")
		CleanCookies()
		Message="�Ѿ��ɹ��˳�!"
End Sub

Sub APIResponseCookies()
	Dim SafeKey,NewSafeKey
	SafeKey = Request.QueryString("syskey")
	UserName = HTMLEncode(Request.QueryString("username"))
	UserName=unescape(UserName)		'�������ı�������,���������û���֤��ͨ����
	UserPassword = Request.QueryString("password")
	Expired = Request.QueryString("savecookie")
	IsMD5=True
	
	If UserName="" or SafeKey="" Then Exit Sub
	
	NewSafeKey = LCase(Mid(Md5(UserName&SiteConfig("APISafeKey")),9,16))
	If SafeKey<>NewSafeKey Then Exit Sub
	
	If Expired="" or Not IsNumeric(Expired) Then Expired = 0
	
	
	If UserPassword = "" Then	'�û��˳�
		UserLoginOut
	Else						'�û���½
		UserLogin
	End If
End Sub
%>