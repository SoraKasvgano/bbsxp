<!-- #include file="Setup.asp" -->
<%
HtmlHeadTitle="联系我们"
HtmlTop

if Request_Method <> "POST" then
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 联系我们</div>
<form method="Post" name="form"  onSubmit="return chkInput(this);">
<table cellspacing="1" cellpadding="5" width="700" class="CommonListArea" align="center">
	<tr class=CommonListTitle>
		<td colspan="2" align="center">发送邮件给论坛管理员</td>
	</tr>
	<tr class="CommonListCell">
		<td>
        
        <div style="width:640px" align="left">
			<fieldset>
				<legend>个人信息</legend>
				<table cellpadding="0" cellspacing="3" border="0">
				<tr>
					<td>您的姓名 :<br /><input type="text" name="UserName" value="<%=CookieUserName%>" size="50" /></td>
				</tr>
				<tr>
					<td>Email 地址 :<br /><input type="text" name="UserEmail" value="<%=CookieUserEmail%>" size="50" /></td>
				</tr>
				<%if CookieUserName=empty then%><tr>
					<td>验证码：<br /><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span></td>
				</tr><%end if%>
				</table>
			</fieldset>
       		<br>
           		<fieldset>

				<legend>邮件信息</legend>
				<table cellpadding="0" cellspacing="3" border="0">
				<tr>
					<td>
						主题 :<br />
						<input type="text" name="MailSubject" value="" size="50" />
					</td>
				</tr>
				<tr>
					<td>
						内容 :<br />
						<textarea name="MailBody" rows="10" cols="50" wrap="virtual" style="width:540px"></textarea>
					</td>
				</tr>
				</table>
			</fieldset>
	</div>
    	</td>
	</tr>
    
    
	<tr class="CommonListCell">
		<td align=center>
        <%if SiteConfig("SelectMailMode")="" then%>
        <input type="button" value=" 发送 " accesskey="s"/ onClick="ToMail()">
	<%else%>        
         <input type="submit" value=" 发送 " accesskey="s"/>
	<%end if%>         
         <input type="reset" class="button" name="reset" value="重置表单" accesskey="r" /></td>
	</tr>
    
</table>
</form>
<script language="javascript" type="text/javascript">
function chkInput(form){
	if (form.UserName.value==''){
		alert('请输入您的姓名');
		return false;
	}
	if (form.UserEmail.value==''){
		alert('请输入Email 地址');
		return false;
	}
	if (form.MailSubject.value==''){
		alert('请输入主题');
		return false;
	}
	if (form.MailBody.value==''){
		alert('请输入内容');
		return false;
	}

	return true;
}
<%if SiteConfig("SelectMailMode")="" then%>
function ToMail(){
    document.form.action = "mailto:<%=SiteConfig("WebMasterEmail")%>?subject=" + form.MailSubject.value+"&body="+form.MailBody.value;
    document.form.submit();
}
<%end if%>
</script>



<%
HtmlBottom
end if

if CookieUserName=empty then
	if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then
		error("<li>验证码错误！</li>")
	else
		Session("VerifyCode")=""
	end if
end if


UserName=HTMLEncode(Request.Form("UserName"))
UserEmail=HTMLEncode(Request.Form("UserEmail"))
MailSubject=HTMLEncode(Request.Form("MailSubject"))
MailBody=HTMLEncode(Request.Form("MailBody"))

MailBody="发件人：<A href='mailto:"&UserEmail&"'>"&UserName&"</a><br /><br />"&MailBody


	SendMail SiteConfig("WebMasterEmail"),MailSubject,MailBody
	Message="<li>邮件发送成功<li><a href=Default.asp>返回首页</a>"
	succeed Message,"SendMessage.asp"

%>