<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->
<%

if Request("menu")="OUT" then
	HtmlTop
	if ReturnUrl="" and Http_Referer<>"" then ReturnUrl=Http_Referer
	if ReturnUrl="" then ReturnUrl="Default.asp"
	
	if Request_Method <> "POST" then error("<li>�ύ��ʽ����</li><li>������ʹ�õ���"&Request_Method&"�ύ��ʽ��</li>")
	TempUserName=CookieUserName
	UserLoginOut
		
		'--------------------  API Start  --------------------
		If SiteConfig("APIEnable")=1 Then
			Dim ApiSaveCookie
			APIUserLoginOut()
		End If
		'--------------------  API End  ----------------------
	TempUserName=""
	succeed "�Ѿ��ɹ��˳�",ReturnUrl
	
elseif Request_Method = "POST" then
	ReturnUrl=Request.Form("ReturnUrl")
	UserName=HTMLEncode(Request.Form("UserName"))
	UserPassword=Trim(Request.Form("UserPassword"))
	IsMD5=False

	if SiteConfig("EnableAntiSpamTextGenerateForLogin")=1 then
		if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then
			AlertForModal("��֤�����")
		else
			Session("VerifyCode")=""
		end if
	end if
	if UserName=empty then AlertForModal("�û���û������")
	if UserPassword=empty then AlertForModal("����û������")
	if instr(UserName,"&")>0 and instr(UserName,";")>0 then AlertForModal("����ϵͳ�����������û����к��������ַ�����ʹ��Email��ʽ��¼")
	
	Invisible=Request("Invisible")

	if Request("IsSave")="1" then
		Expired=9999
	else
		Expired=0
	end if


	Dim Message
	if UserLogin=True then
		'--------------------  API Start  --------------------
		If SiteConfig("APIEnable")=1 Then
			APIUserLogin()
		End If
		'--------------------  API End  ----------------------
	else
		AlertForModal(Message)
	end if
%>
	<script language="javascript" type="text/javascript">
		<%if ReturnUrl<>"" then%>
		window.location.href = "<%=ReturnUrl%>";
		<%else%>
		if(top == self){window.location.href = "Default.asp";}else{parent.window.location.reload();}
		<%end if%>
	</script>
<%
else


Response.Clear()
%>
<title>�û���¼</title>
<script language="javascript" type="text/javascript">if(top == self)window.location.href = "Default.asp";</script>
<script src="Utility/global.js" type="text/javascript"></script>
<style type="text/css">body,table{font-size:9pt;}</style>

<form action="Login.asp" method="POST" name="form">
	<table cellspacing="1" cellpadding="2" width="100%">
		<tr>
			<td width="30%" align="right">�û����ƣ�</td>
			<td>
				<input type="text" name="UserName" value="�û���/�����ʼ���ַ" onblur="if (this.value==''){ this.value='�û���/�����ʼ���ַ';this.style.color='#999';}else{this.style.color='';}" onfocus="if (this.value=='�û���/�����ʼ���ַ') {this.value='';this.style.color='';}" title="�û���/�����ʼ���ַ" style="width:150px; color:#999" />&nbsp; <a href="javascript:parent.window.location.href='CreateUser.asp';">û��ע��?</a>
			</td>
		</tr>
		<tr>
			<td width="30%" align="right">�û����룺</td>
			<td>
				<input type="password" name="UserPassword" style="width:150px" />&nbsp; <a href="javascript:parent.window.location.href='RecoverPassword.asp';">�һ�����?</a>
			</td>
		</tr>
		<%if SiteConfig("EnableAntiSpamTextGenerateForLogin")=1 then%>
		<tr>
			<td width="30%" align="right">�� ֤ �룺</td>
			<td>
				<input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span>
			</td>
		</tr>
		<%end if%>
		<tr>
			<td align="right" width="30%">��¼��ʽ��</td>
			<td>
				<input type="checkbox" value="1" name="IsSave" id="IsSave" /><label for="IsSave">�Զ���¼</label>
				<input type="checkbox" value="1" name="Invisible" id="Invisible" /><label for="Invisible">�����¼</label>
			</td>
		</tr>
		<tr>
			<td align="center" colspan="2">
				<input type="submit" value=" ��¼ " />��<input type="button" onclick="javascript:parent.BBSXP_Modal.Close();" value=" ȡ �� " />
			</td>
		</tr>
	</table>
</form>
<%
end if
%>