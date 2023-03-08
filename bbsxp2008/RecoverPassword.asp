<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->
<%
HtmlTop

UserName=HTMLEncode(Request("UserName"))
UserEmail=HTMLEncode(Request("UserEmail"))
ActivationKey=HTMLEncode(Request("ActivationKey"))

select case Request("menu")
	case ""
		default
	case "SelectMethod"
		if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("验证码错误！")
		if ""&UserName&""="" then error("请输入用户名！")
		if Execute("Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'").eof then error("请输入您要找回密码的用户名！")
		SelectMethod
	case "MailRecover"
		if SiteConfig("SelectMailMode")="" then error("系统未开启 邮件 功能！")
		
		if UserEmail="" then error("请输入Email地址！")
		if UserName<>"" then UserNameSql="and UserName='"&UserName&"'"
		sql="Select * from ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"' "&UserNameSql&""
		Rs.Open sql,Conn,1
		if Rs.eof then error("论坛中找不到相关的资料")
		UserEmail=Rs("UserEmail")
		UserName=Rs("UserName")
		Rs.close

		Randomize
		ActivationKey=int(rnd*9999999999)+1
		
		LoadingEmailXml("RecoverPassword")
		MailBody=Replace(MailBody,"[UserName]",UserName)
		MailBody=Replace(MailBody,"[RecoverPasswordURL]","<a target=_blank href="&SiteConfig("SiteUrl")&"/RecoverPassword.asp?menu=MailRecoverok&username="&UserName&"&ActivationKey="&ActivationKey&">"&SiteConfig("SiteUrl")&"/RecoverPassword.asp?menu=MailRecoverok&username="&UserName&"&ActivationKey="&ActivationKey&"</a>")
		MailBody=Replace(MailBody,"[IPAddress]",REMOTE_ADDR)
		SendMail UserEmail,MailSubject,MailBody
		
		Execute("insert into ["&TablePrefix&"UserActivation] (ActivationKey,UserName) values ('"&ActivationKey&"','"&UserName&"')")
		Session("VerifyCode")=""
		log(""&UserName&"申请找回密码，Email:"&UserEmail&"")
		succeed "请到邮箱中取回密码","Login.asp"
		
	case "setNewPassword"
		UserPassword=Trim(Request("UserPassword"))
		UserPassword2=Trim(Request("UserPassword2"))
		
		if UserPassword<>UserPassword2 then error"<li>您的新密码和确认新密码不同"
		if Len(UserPassword)<6 then error"<li>新密码必须至少包含 6 个字符"

		if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and UserName='"&UserName&"'").eof then error("找回密码的信息已过期！请重新提交找回！")
		Execute("Delete from ["&TablePrefix&"UserActivation] where UserName='"&UserName&"'")
		
		'--------------------  API Start  --------------------
		Message=""
		If SiteConfig("APIEnable")=1 Then APIUpdateUser UserName,UserPassword,"","",""
		if Message<>"" then error(""&Message&"")
		'--------------------  API End  ----------------------
		
		ModifyUserPassword UserName,UserPassword,"","",""

		Message=Message&"<li>新密码设置成功</li><li><a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">请返回登录</a></li>"
		succeed Message,"Login.asp"

	case "QuestionRecover"
		UserPassword=Trim(Request("UserPassword"))
		UserPassword2=Trim(Request("UserPassword2"))
		
		if UserName="" then error "没有输入用户名！"
		if UserPassword<>UserPassword2 then error"<li>您的新密码和确认新密码不同"
		if Len(UserPassword)<6 then error"<li>新密码必须至少包含 6 个字符"
		
		Set Rs=Execute("Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'")
			if Not Rs.eof then
				if ""&Rs("PasswordAnswer")&""="" then Alert("您的密码问题答案为空，不能通过此方式取回密码")
				if md5(""&Request("PasswordAnswer")&"")<>Rs("PasswordAnswer") then Alert("答案错误")
				Rs.close
					
				'--------------------  API Start  --------------------
				Message=""
				If SiteConfig("APIEnable")=1 Then APIUpdateUser UserName,UserPassword,""
				if Message<>"" then error(""&Message&"")
				'--------------------  API End  ----------------------
		
				ModifyUserPassword UserName,UserPassword,"","",""
		
		
				Message=Message&"<li>新密码设置成功</li><li><a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">请返回登录</a></li>"
				succeed Message,"Login.asp"

			else
				Rs.close
				error "您输入的用户不存在"
			End if
		Rs.close
		
		
	case "MailRecoverok"
		if SiteConfig("SelectMailMode")="" then error("系统未开启 邮件 功能！")
		if UserName="" then error "URL不完整，没有用户名！"
		if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and UserName='"&UserName&"'").eof then error("找回密码的信息已过期！请重新提交找回！")

		SetNewPassword
end select

Sub default
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 找回密码</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea align="center">
<form action="RecoverPassword.asp?menu=SelectMethod" method="POST">
	<tr class=CommonListTitle>
		<td align="center">取回密码</td>
	</tr>
	<tr class="CommonListCell">
    	<td>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td colspan="2" style="padding-left:50px;">第一步：请输入你的用户名：</td>
		</tr>
		<tr>
			<td align="right" width="30%"><b>用户名：</b></td>
			<td><input name="UserName" size="30" /></td>
		</tr>
		<tr>
			<td align="right" width="30%"><b>验证码：</b></td>
			<td>
            <input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span> <span id="CheckVerifyCode" style="color:red"></span>
            </td>
		</tr>
		<tr>
			<td align="center" colspan="2"> <input type="submit" value=" 下一步 " /> </td>
		</tr>
		</table>
        </td>
    </tr>
</form>
</table>



<br /><center><a href="javascript:history.back()">BACK </a><br />
<%
End Sub

Sub SelectMethod
%>
<script language="javascript" type="text/javascript" src="Utility/pswdplc.js"></script>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 找回密码</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea align="center">
<form action="RecoverPassword.asp" method="POST">
<input type=hidden name=UserName value="<%=UserName%>" />
	<tr class=CommonListTitle>
		<td align="center">取回密码</td>
	</tr>
	<tr class="CommonListCell">
    	<td>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td colspan="2" style="padding-left:50px;">第二步：请选择一项来重新设置您的密码：</td>
		</tr>
        
		<tr>
			<td colspan="2" style="padding-left:100px;">
            	<input type="radio" value="QuestionRecover" name="menu" id="QuestionRecover" onclick="SelectOption('Question')" checked="checked" />&nbsp;<label for="QuestionRecover">使用 密码问题&提示答案 取回密码</label>
            </td>
		</tr>
		<tr>
			<td colspan="2" style="display:" id="Question">
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
			<tr>
				<td width="30%" align="right"><b>密码提示问题：</b></td>
				<td><%=Execute("Select PasswordQuestion from ["&TablePrefix&"Users] where UserName='"&UserName&"'")(0)%></td>
			</tr>
			<tr>
				<td width="30%" align="right"><b>提示问题答案：</b></td>
				<td><input size="15" name="PasswordAnswer" /></td>
			</tr>
			<tr>
	    		<td align="right" width="30%"><b>新密码：</b></td>
				<td><input type="password" name="UserPassword" size="40" maxLength="16" onkeyup="EvalPwd(this.value);" onchange="EvalPwd(this.value);" /></td>
			</tr>
			<tr>
	    		<td align="right" width="30%"><b>密码强度：</b></td>
	    		<td>
					<table border="0" width="250" cellspacing="1" cellpadding="2">
						<tr bgcolor="#f1f1f1">
							<td id=iWeak align="center">弱</td>
							<td id=iMedium align="center">中</td>
							<td id=iStrong align="center">强</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align="right" width="30%"><b>重新键入新密码：</b><br />请与您的新密码保持一致</td>
				<td valign="middle"> <input type="password" name="UserPassword2" size="40" maxlength="16" /></td>
			</tr>
			<tr>
				<td align="center" colspan="2"> <input type="submit" value=" 确定 " />　<input type="button" onclick="javascript:history.back()" value=" 取消 " /> </td>
			</tr>
			</table>
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            </td>
		</tr>
<%
if SiteConfig("SelectMailMode")<>"" then
%>
		<tr>
			<td colspan="2" style="padding-left:100px;">
            	<input value="MailRecover" name="menu" type="radio" id="MailRecover" onclick="SelectOption('Email')" />&nbsp;<label for="MailRecover">使用 Email 取回密码</label>
            </td>
		</tr>
		<tr>
			<td colspan="2" style="display:none" id="Email">
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
            <tr>
				<td align="right"><b>电子邮件地址：</b></td>
				<td><input name="UserEmail" size="30" /></td>
			</tr>
			<tr>
				<td align="center" colspan="2"> <input type="submit" value=" 确定 " />　<input type="button" onclick="javascript:history.back()" value=" 取消 " /> </td>
			</tr>
			</table>
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            </td>
		</tr>
<%end if%>
		</table>
        </td>
    </tr>
</form>
</table>

<script language="javascript" type="text/javascript">
function SelectOption(ID){
	$("Question").style.display = "none";
	try{
		$("Email").style.display = "none";
	}catch(e){}
	$(ID).style.display = "";
}
</script>


<br /><center><a href="javascript:history.back()">BACK </a><br />

<%
End Sub
Sub SetNewPassword
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 找回密码</div>
<script language="javascript" type="text/javascript" src="Utility/pswdplc.js"></script>
<form method="POST" name="form" action="?Menu=setNewPassword" onsubmit="return VerifyInput();">
<input type=hidden name=UserName value="<%=UserName%>" />
<input type=hidden name=ActivationKey value="<%=ActivationKey%>" />
<table width="100%" border="0" cellspacing="1" cellpadding="5" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center" colspan=2>设置 <%=UserName%> 的新密码</td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b> 新密码：</b></td>
		<td align="Left" width="55%"> <input type="password" name="UserPassword" size="40" maxLength="16" onkeyup="EvalPwd(this.value);" onchange="EvalPwd(this.value);" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b>密码强度：</b></td>
	    <td align="Left" width="55%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">弱</td>
					<td id=iMedium align="center">中</td>
					<td id=iStrong align="center">强</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="45%"><b>重新键入新密码：</b><br />请与您的新密码保持一致</td>
		<td align="Left" valign="middle" width="55%"> <input type="password" name="UserPassword2" size="40" maxlength="16" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="2"> <input type="submit" value=" 确 定 " /></td>
	</tr>
</table>
</form>
<script language="JavaScript">
function VerifyInput()
{
	if (document.form.UserPassword.value.length < 6)
	{
		alert("您的新密码长度必须大于5！");
		document.form.UserPassword.focus();
		return false;
	}
	if (document.form.UserPassword.value != document.form.UserPassword2.value)
	{
		alert("您二次键入的新密码不同！");
		document.form.UserPassword.focus();
		return false;
	}
	return true;
}
</script>
<%
End Sub



HtmlBottom
%>