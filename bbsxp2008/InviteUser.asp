<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")
if SiteConfig("AccountActivation")<>2 then error("系统未开启 邀请码注册 选项！")
if SiteConfig("SelectMailMode")="" then error("系统未开启 邮件发送 功能！")

if Request("menu")="InvitingOk" then
	if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("验证码错误！")
	UserEmail=HTMLEncode(Request("UserEmail"))
	if UserEmail="" then error("请输入受邀人的Email地址！")
	if instr(UserEmail,"@")=0 then error("受邀人的Email地址填写错误")
	
	If not Execute("Select UserID From ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"'" ).eof Then Error("<li>"&UserEmail&" 已经在本论坛注册了")

	
	Randomize
	ActivationKey=int(rnd*9999999999)+1

	LoadingEmailXml("InviteNewUser")
	MailBody=Replace(MailBody,"[UserName]",CookieUserName)
	MailBody=Replace(MailBody,"[SiteName]",SiteConfig("SiteName"))
	MailBody=Replace(MailBody,"[InviteURL]","<a target=_blank href="&SiteConfig("SiteUrl")&"/CreateUser.asp?menu=WriteProfile&ActivationKey="&ActivationKey&"&ReferrerName="&CookieUserName&"&UserEmail="&UserEmail&">"&SiteConfig("SiteUrl")&"/CreateUser.asp?menu=WriteProfile&ActivationKey="&ActivationKey&"&ReferrerName="&CookieUserName&"&UserEmail="&UserEmail&"</a>")
	SendMail UserEmail,MailSubject,MailBody
	
	Execute("insert into ["&TablePrefix&"UserActivation] (ActivationKey,Email) values ('"&ActivationKey&"','"&UserEmail&"')")
	Session("VerifyCode")=""
	succeed "邀请码发送成功！","Default.asp"
else
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 发送邀请码</div>
<table cellspacing=1 cellpadding=5 width=80% class=CommonListArea>
	<form action="?menu=InvitingOk" method="POST" name=form>
		<tr class=CommonListTitle>
			<td align="center" colspan="2">发送邀请码</td>
		</tr>
		<tr class="CommonListCell">
			<td width="30%" align="right">验证码：</td>
			<td>
				<input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span> <span id="CheckVerifyCode" style="color:red"></span>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">受邀人电子邮箱：</td>
			<td>
				<input type="text" name="UserEmail" size=30 />
			</td>
		</tr>
		<tr class="CommonListCell">
			<td valign="top" align="center" colspan="2">
				<input type="submit" value=" 发送 " /> <input type="button" onclick="javascript:history.back()" value=" 取消 " />
			</td>
		</tr>
	</form>
	</table>
<%
end if
HtmlBottom
%>