<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->

<%
HtmlTop


UserEmail=HTMLEncode(Request("UserEmail"))
ActivationKey=HTMLEncode(Request("ActivationKey"))
ReferrerName=HTMLEncode(Request("ReferrerName"))


if SiteConfig("AllowNewUserRegistration")= 0 then error("����̳��ʱ���������û�ע�ᣡ")

if SiteConfig("AccountActivation")=2 then

	if ActivationKey=empty then error("����̳��Ҫ���������ע�ᣡ")
	if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and Email='"&UserEmail&"'").eof then error("���������")

end if




if Request.Form("menu")="AddUserName" then

	UserName=HTMLEncode(Request.Form("UserName"))
	UserPassword=Trim(Request.Form("UserPassword"))
	RetypePassword=Trim(Request.Form("RetypePassword"))
	PasswordQuestion=HTMLEncode(Request.Form("PasswordQuestion"))
	PasswordQuestion2=HTMLEncode(Request.Form("PasswordQuestion2"))
	PasswordAnswer=Request("PasswordAnswer")


	if ""&PasswordQuestion2&""<>"" then PasswordQuestion=PasswordQuestion2
	if UserName="" then Message=Message&"<li>�����û���û����д</li>"
	if instr(Request.QueryString("UserName"),UserName)=0 then Message=Message&"<li>�û����в��ܺ���URL�����ܴ��͵��������</li>"
	if PasswordQuestion="" then Message=Message&"<li>������ʾ����û����д</li>"
	if PasswordAnswer="" then Message=Message&"<li>������ʾ�����û����д</li>"

	if SiteConfig("EnableAntiSpamTextGenerateForRegister")=1 then
		if Request.Form("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then
			Message=Message&"<li>��֤�����</li>"
		else
			Session("VerifyCode")=""
		end if
	end if

	if SiteConfig("AccountActivation")=1 then 
		Randomize
		UserPassword=int(rnd*999999)+1
	else
		if Len(UserPassword)<6 then Message=Message&"<li>����������ٰ��� 6 ���ַ�</li>"
		if UserPassword<>RetypePassword then Message=Message&"<li>�� 2 ����������벻��ͬ</li>"
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

	Message=Message&"<li>ע�����û����ϳɹ�</li><li><a href=Default.asp><a href=EditProfile.asp>��д������ϸ����</a></a></li>"
	succeed Message,"Default.asp"
end if


if Request("menu")="" and CreateUserAgreement<>"" then
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ע��Э��</div>
<br /><center>
<textarea name="textarea" rows="18" readonly="readonly" style="width:100%"><%=CreateUserAgreement%></textarea><br /><br />
<input type="submit" value=" ͬ �� " onclick="window.location.href='CreateUser.asp?menu=WriteProfile';" />
<input type="submit" value=" ��ͬ�� " onclick="history.back()" /></center>
<%
else
%>
<script type="text/javascript" src="Utility/pswdplc.js"></script>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ��дע������</div>
<form method="post" name="form" action="?UserName=" onsubmit="this.action=this.action+this.UserName.value">
<input type="hidden" name="menu" value="AddUserName" />
<input type="hidden" name="ReferrerName" value="<%=ReferrerName%>" />
<input type="hidden" name="ActivationKey" value="<%=ActivationKey%>" />
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" valign="middle" align="Left">&nbsp;ע���û�����</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="30%"><b>�û�����</b></td>
	  <td align="Left" width="77%"><input type="text" name="UserName" size="40" onblur="CheckUserName(this.value)" /> <span id="CheckUserName" style="color:#FF0000"></span></td>
	</tr>
<%	if SiteConfig("AccountActivation")<>1 then%>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>���룺</b><br />
	  ����������ٰ��� 6 ���ַ�</td>
	  <td align="Left" valign="middle" width="77%"><input type="password" name="UserPassword" size="40" maxlength="16" onkeyup=EvalPwd(this.value); onchange=EvalPwd(this.value); onblur="CheckPassword(this.value)" /> <span id="CheckPassword" style="color:#FF0000"></span></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>����ǿ�ȣ�</b></td>
		<td align="Left" valign="middle" width="77%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">��</td>
					<td id=iMedium align="center">��</td>
					<td id=iStrong align="center">ǿ</td>
				</tr>
			</table>		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>���¼������룺</b><br />�����������뱣��һ��</td>
		<td align="Left" valign="middle" width="77%"><input type="password" name="RetypePassword" size="40" onblur="CheckRetypePassword(this.value)" /> <span id="CheckRetypePassword" style="color:#FF0000"></span></td>
	</tr>
<%	end if%>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>����Email��ַ��</b><br /><%if SiteConfig("AccountActivation")=1 then%><font color="FF0000">���뽫ͨ��Email����</font><%end if%></td>
	 	<td align="Left" valign="middle" width="77%">
		<input type="text" name="UserEmail" size="40" onblur="CheckMail(this.value)" <%if SiteConfig("AccountActivation")=2 then%> value="<%=UserEmail%>" readonly<%end if%> /> <span id="CheckMail" style="color:#FF0000"></span></td>
	</tr>

<%	if SiteConfig("EnableAntiSpamTextGenerateForRegister")=1 then%>
	<tr class="CommonListCell">
		<td align="right"><b>��֤�룺</b></td>
		<td><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span></td>
	</tr>
<%	end if%>
	<tr class="CommonListCell">
		<td align="right" valign="middle"><b>������ʾ���⣺</b><br>��������������룬ϵͳ������ѯ�ʻ��ܴ�</td>
		<td>
		<select name="PasswordQuestion" onchange="$('PasswordQuestion2').style.display = this.options[this.selectedIndex].value=='Define' ? '':'none';">
			<option value="" selected="selected">ѡ��һ��</option>
			<option value="ĸ�׵ĳ����ص�">ĸ�׵ĳ����ص�</option>
			<option value="��ͯʱ����õ�����">��ͯʱ����õ�����</option>
			<option value="��һ�����������">��һ�����������</option>
			<option value="��ϲ������ʦ">��ϲ������ʦ</option>
			<option value="��ϲ������ʷ����">��ϲ������ʷ����</option>
			<option value="�游��ְҵ">�游��ְҵ</option>
			<option value="Define">�Զ���</option>
		</select>��<span id=PasswordQuestion2 style="display:none"><input type=text name=PasswordQuestion2 size=20 /></span>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>���ܴ𰸣�</b></td>
		<td><input type="text" name="PasswordAnswer" size="40" /></td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td valign="middle" colspan="2"><input type="submit" value=" ע �� " /></td>
	</tr>
</table>
</form>
<script language="JavaScript" type="text/javascript">
function CheckUserName(UserName) {
	if(UserName.length > <%=SiteConfig("UserNameMaxLength")%> || UserName.length <<%=SiteConfig("UserNameMinLength")%>) {
		ShowCheckResult("CheckUserName","��������û�������Ӧ���� <%=SiteConfig("UserNameMinLength")%>-<%=SiteConfig("UserNameMaxLength")%> ��Χ��","error");
		return;
	}
	

	Ajax_CallBack(false,"CheckUserName","Loading.asp?menu=CheckUserName&UserNameLength="+UserName.length+"&UserName=" + escape(UserName));
}

function CheckPassword(UserPassword) {
	if (UserPassword.length < 6){
		ShowCheckResult("CheckPassword", "����������ٰ��� 6 ���ַ�","error");
		return;
	}

	ShowCheckResult("CheckPassword", "","right");

}
function CheckRetypePassword(RetypePassword) {
	if (RetypePassword != document.form.UserPassword.value){
		ShowCheckResult("CheckRetypePassword", "�� 2 ����������벻��ͬ","error");
		return;
	}
	
	if (RetypePassword != ''){
		ShowCheckResult("CheckRetypePassword", "","right");
	}
}
function CheckMail(Mail) {
	if(Mail.indexOf("@") == -1 || Mail.indexOf(".") == -1) {
		ShowCheckResult("CheckMail", "��û������Email����������","error");
		return;
	}
	
	Ajax_CallBack(false,"CheckMail","Loading.asp?menu=CheckMail&Mail=" + escape(Mail));
}
</script>
<%
end if

HtmlBottom
%>