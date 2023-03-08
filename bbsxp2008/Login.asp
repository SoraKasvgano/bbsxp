<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->
<%

if Request("menu")="OUT" then
	HtmlTop
	if ReturnUrl="" and Http_Referer<>"" then ReturnUrl=Http_Referer
	if ReturnUrl="" then ReturnUrl="Default.asp"
	
	if Request_Method <> "POST" then error("<li>提交方式错误！</li><li>您本次使用的是"&Request_Method&"提交方式！</li>")
	TempUserName=CookieUserName
	UserLoginOut
		
		'--------------------  API Start  --------------------
		If SiteConfig("APIEnable")=1 Then
			Dim ApiSaveCookie
			APIUserLoginOut()
		End If
		'--------------------  API End  ----------------------
	TempUserName=""
	succeed "已经成功退出",ReturnUrl
	
elseif Request_Method = "POST" then
	ReturnUrl=Request.Form("ReturnUrl")
	UserName=HTMLEncode(Request.Form("UserName"))
	UserPassword=Trim(Request.Form("UserPassword"))
	IsMD5=False

	if SiteConfig("EnableAntiSpamTextGenerateForLogin")=1 then
		if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then
			AlertForModal("验证码错误！")
		else
			Session("VerifyCode")=""
		end if
	end if
	if UserName=empty then AlertForModal("用户名没有输入")
	if UserPassword=empty then AlertForModal("密码没有输入")
	if instr(UserName,"&")>0 and instr(UserName,";")>0 then AlertForModal("由于系统升级，您的用户名中含有特殊字符，请使用Email方式登录")
	
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
<title>用户登录</title>
<script language="javascript" type="text/javascript">if(top == self)window.location.href = "Default.asp";</script>
<script src="Utility/global.js" type="text/javascript"></script>
<style type="text/css">body,table{font-size:9pt;}</style>

<form action="Login.asp" method="POST" name="form">
	<table cellspacing="1" cellpadding="2" width="100%">
		<tr>
			<td width="30%" align="right">用户名称：</td>
			<td>
				<input type="text" name="UserName" value="用户名/电子邮件地址" onblur="if (this.value==''){ this.value='用户名/电子邮件地址';this.style.color='#999';}else{this.style.color='';}" onfocus="if (this.value=='用户名/电子邮件地址') {this.value='';this.style.color='';}" title="用户名/电子邮件地址" style="width:150px; color:#999" />&nbsp; <a href="javascript:parent.window.location.href='CreateUser.asp';">没有注册?</a>
			</td>
		</tr>
		<tr>
			<td width="30%" align="right">用户密码：</td>
			<td>
				<input type="password" name="UserPassword" style="width:150px" />&nbsp; <a href="javascript:parent.window.location.href='RecoverPassword.asp';">找回密码?</a>
			</td>
		</tr>
		<%if SiteConfig("EnableAntiSpamTextGenerateForLogin")=1 then%>
		<tr>
			<td width="30%" align="right">验 证 码：</td>
			<td>
				<input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span> <span id="CheckVerifyCode" style="color:red"></span>
			</td>
		</tr>
		<%end if%>
		<tr>
			<td align="right" width="30%">登录方式：</td>
			<td>
				<input type="checkbox" value="1" name="IsSave" id="IsSave" /><label for="IsSave">自动登录</label>
				<input type="checkbox" value="1" name="Invisible" id="Invisible" /><label for="Invisible">隐身登录</label>
			</td>
		</tr>
		<tr>
			<td align="center" colspan="2">
				<input type="submit" value=" 登录 " />　<input type="button" onclick="javascript:parent.BBSXP_Modal.Close();" value=" 取 消 " />
			</td>
		</tr>
	</table>
</form>
<%
end if
%>