<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->

<%
HtmlTop


UserEmail=HTMLEncode(Request("UserEmail"))
ActivationKey=HTMLEncode(Request("ActivationKey"))
ReferrerName=HTMLEncode(Request("ReferrerName"))


if SiteConfig("AllowNewUserRegistration")= 0 then error("本论坛暂时不允许新用户注册！")

if SiteConfig("AccountActivation")=2 then

	if ActivationKey=empty then error("本论坛需要邀请码才能注册！")
	if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and Email='"&UserEmail&"'").eof then error("邀请码错误！")

end if




if Request.Form("menu")="AddUserName" then

	UserName=HTMLEncode(Request.Form("UserName"))
	UserPassword=Trim(Request.Form("UserPassword"))
	RetypePassword=Trim(Request.Form("RetypePassword"))
	PasswordQuestion=HTMLEncode(Request.Form("PasswordQuestion"))
	PasswordQuestion2=HTMLEncode(Request.Form("PasswordQuestion2"))
	PasswordAnswer=Request("PasswordAnswer")


	if ""&PasswordQuestion2&""<>"" then PasswordQuestion=PasswordQuestion2
	if UserName="" then Message=Message&"<li>您的用户名没有填写</li>"
	if instr(Request.QueryString("UserName"),UserName)=0 then Message=Message&"<li>用户名中不能含有URL所不能传送的特殊符号</li>"
	if PasswordQuestion="" then Message=Message&"<li>密码提示问题没有填写</li>"
	if PasswordAnswer="" then Message=Message&"<li>密码提示问题答案没有填写</li>"

	if SiteConfig("EnableAntiSpamTextGenerateForRegister")=1 then
		if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then
			Message=Message&"<li>验证码错误！</li>"
		else
			Session("VerifyCode")=""
		end if
	end if

	if SiteConfig("AccountActivation")=1 then 
		Randomize
		UserPassword=int(rnd*999999)+1
	else
		if Len(UserPassword)<6 then Message=Message&"<li>密码必须至少包含 6 个字符</li>"
		if UserPassword<>RetypePassword then Message=Message&"<li>您 2 次输入的密码不相同</li>"
	end if
	TempUserPassword=UserPassword

	if Message<>"" then error(""&Message&"")


	'--------------------  API Start  --------------------
	If SiteConfig("APIEnable")=1 Then
		Dim ApiSaveCookie
		APIAddUser()
	End If
	if Message<>"" then error(""&Message&"")
	'--------------------  API End  ----------------------


	Dim UserID
	if AddUser=False then error(""&Message&"")
	
	
	if SiteConfig("AccountActivation")=0 or SiteConfig("AccountActivation")=2 then
		ResponseCookies "UserID",UserID,"999"
		ResponseCookies "UserPassword",UserPassword,"999"
		'--------------------  API Start  --------------------
		If SiteConfig("APIEnable")=1 Then
			Response.Write ApiSaveCookie
			Response.Flush
		End If
		'--------------------  API End  ----------------------
	end if

	if SiteConfig("AccountActivation")=2 then Execute("Delete from ["&TablePrefix&"UserActivation] where Email='"&UserEmail&"'")

	LoadingEmailXml("NewUserAccountCreated")
	MailBody=Replace(MailBody,"[UserName]",UserName)
	MailBody=Replace(MailBody,"[password]",TempUserPassword)
	SendMail UserEmail,MailSubject,MailBody

	Message=Message&"<li>注册新用户资料成功</li><li><a href=Default.asp><a href=EditProfile.asp>填写个人详细资料</a></a></li>"
	succeed Message,"Default.asp"
end if


if Request("menu")="" and CreateUserAgreement<>"" then
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 注册协议</div>
<br /><center>
<textarea name="textarea" rows="18" readonly="readonly" style="width:100%"><%=CreateUserAgreement%></textarea><br /><br />
<input type="submit" value=" 同 意 " onclick="window.location.href='CreateUser.asp?menu=WriteProfile';" />
<input type="submit" value=" 不同意 " onclick="history.back()" /></center>
<%
else
%>
<script type="text/javascript" src="Utility/pswdplc.js"></script>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 填写注册资料</div>
<form method="post" name="form" action="?UserName=" onsubmit="this.action=this.action+this.UserName.value">
<input type="hidden" name="menu" value="AddUserName" />
<input type="hidden" name="ReferrerName" value="<%=ReferrerName%>" />
<input type="hidden" name="ActivationKey" value="<%=ActivationKey%>" />
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" valign="middle" align="Left">&nbsp;注册用户资料</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="30%"><b>用户名：</b></td>
	  <td align="Left" width="77%"><input type="text" name="UserName" size="40" onblur="CheckUserName(this.value)" /> <span id="CheckUserName" style="color:#FF0000"></span></td>
	</tr>
<%	if SiteConfig("AccountActivation")<>1 then%>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>密码：</b><br />
	  密码必须至少包含 6 个字符</td>
	  <td align="Left" valign="middle" width="77%"><input type="password" name="UserPassword" size="40" maxlength="16" onkeyup=EvalPwd(this.value); onchange=EvalPwd(this.value); onblur="CheckPassword(this.value)" /> <span id="CheckPassword" style="color:#FF0000"></span></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>密码强度：</b></td>
		<td align="Left" valign="middle" width="77%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">弱</td>
					<td id=iMedium align="center">中</td>
					<td id=iStrong align="center">强</td>
				</tr>
			</table>		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>重新键入密码：</b><br />请与您的密码保持一致</td>
		<td align="Left" valign="middle" width="77%"><input type="password" name="RetypePassword" size="40" onblur="CheckRetypePassword(this.value)" /> <span id="CheckRetypePassword" style="color:#FF0000"></span></td>
	</tr>
<%	end if%>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>您的Email地址：</b><br /><%if SiteConfig("AccountActivation")=1 then%><font color="FF0000">密码将通过Email发送</font><%end if%></td>
	 	<td align="Left" valign="middle" width="77%">
		<input type="text" name="UserEmail" size="40" onblur="CheckMail(this.value)" <%if SiteConfig("AccountActivation")=2 then%> value="<%=UserEmail%>" readonly<%end if%> /> <span id="CheckMail" style="color:#FF0000"></span></td>
	</tr>

<%	if SiteConfig("EnableAntiSpamTextGenerateForRegister")=1 then%>
	<tr class="CommonListCell">
		<td align="right"><b>验证码：</b></td>
		<td><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span> <span id="CheckVerifyCode" style="color:red"></span></td>
	</tr>
<%	end if%>
	<tr class="CommonListCell">
		<td align="right" valign="middle"><b>密码提示问题：</b><br>如果您忘记了密码，系统会向您询问机密答案</td>
		<td>
		<select name="PasswordQuestion" onchange="$('PasswordQuestion2').style.display = this.options[this.selectedIndex].value=='Define' ? '':'none';">
			<option value="" selected="selected">选择一个</option>
			<option value="母亲的出生地点">母亲的出生地点</option>
			<option value="儿童时期最好的朋友">儿童时期最好的朋友</option>
			<option value="第一个宠物的名字">第一个宠物的名字</option>
			<option value="最喜欢的老师">最喜欢的老师</option>
			<option value="最喜欢的历史人物">最喜欢的历史人物</option>
			<option value="祖父的职业">祖父的职业</option>
			<option value="Define">自定义</option>
		</select>　<span id=PasswordQuestion2 style="display:none"><input type=text name=PasswordQuestion2 size=20 /></span>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>机密答案：</b></td>
		<td><input type="text" name="PasswordAnswer" size="40" /></td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td valign="middle" colspan="2"><input type="submit" value=" 注 册 " /></td>
	</tr>
</table>
</form>
<script language="JavaScript" type="text/javascript">
function CheckUserName(UserName) {
	if(UserName.length > <%=SiteConfig("UserNameMaxLength")%> || UserName.length <<%=SiteConfig("UserNameMinLength")%>) {
		ShowCheckResult("CheckUserName","您输入的用户名长度应该在 <%=SiteConfig("UserNameMinLength")%>-<%=SiteConfig("UserNameMaxLength")%> 范围内","error");
		return;
	}
	

	Ajax_CallBack(false,"CheckUserName","Loading.asp?menu=CheckUserName&UserNameLength="+UserName.length+"&UserName=" + escape(UserName));
}

function CheckPassword(UserPassword) {
	if (UserPassword.length < 6){
		ShowCheckResult("CheckPassword", "密码必须至少包含 6 个字符","error");
		return;
	}

	ShowCheckResult("CheckPassword", "","right");

}
function CheckRetypePassword(RetypePassword) {
	if (RetypePassword != document.form.UserPassword.value){
		ShowCheckResult("CheckRetypePassword", "您 2 次输入的密码不相同","error");
		return;
	}
	
	if (RetypePassword != ''){
		ShowCheckResult("CheckRetypePassword", "","right");
	}
}
function CheckMail(Mail) {
	if(Mail.indexOf("@") == -1 || Mail.indexOf(".") == -1) {
		ShowCheckResult("CheckMail", "您没有输入Email或输入有误","error");
		return;
	}
	
	Ajax_CallBack(false,"CheckMail","Loading.asp?menu=CheckMail&Mail=" + escape(Mail));
}
</script>
<%
end if

HtmlBottom
%>