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
		if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("��֤�����")
		if ""&UserName&""="" then error("�������û�����")
		if Execute("Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'").eof then error("��������Ҫ�һ�������û�����")
		SelectMethod
	case "MailRecover"
		if SiteConfig("SelectMailMode")="" then error("ϵͳδ���� �ʼ� ���ܣ�")
		
		if UserEmail="" then error("������Email��ַ��")
		if UserName<>"" then UserNameSql="and UserName='"&UserName&"'"
		sql="Select * from ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"' "&UserNameSql&""
		Rs.Open sql,Conn,1
		if Rs.eof then error("��̳���Ҳ�����ص�����")
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
		log(""&UserName&"�����һ����룬Email:"&UserEmail&"")
		succeed "�뵽������ȡ������","Login.asp"
		
	case "setNewPassword"
		UserPassword=Trim(Request("UserPassword"))
		UserPassword2=Trim(Request("UserPassword2"))
		
		if UserPassword<>UserPassword2 then error"<li>�����������ȷ�������벻ͬ"
		if Len(UserPassword)<6 then error"<li>������������ٰ��� 6 ���ַ�"

		if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and UserName='"&UserName&"'").eof then error("�һ��������Ϣ�ѹ��ڣ��������ύ�һأ�")
		Execute("Delete from ["&TablePrefix&"UserActivation] where UserName='"&UserName&"'")
		
		'--------------------  API Start  --------------------
		Message=""
		If SiteConfig("APIEnable")=1 Then APIUpdateUser UserName,UserPassword,"","",""
		if Message<>"" then error(""&Message&"")
		'--------------------  API End  ----------------------
		
		ModifyUserPassword UserName,UserPassword,"","",""

		Message=Message&"<li>���������óɹ�</li><li><a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">�뷵�ص�¼</a></li>"
		succeed Message,"Login.asp"

	case "QuestionRecover"
		UserPassword=Trim(Request("UserPassword"))
		UserPassword2=Trim(Request("UserPassword2"))
		
		if UserName="" then error "û�������û�����"
		if UserPassword<>UserPassword2 then error"<li>�����������ȷ�������벻ͬ"
		if Len(UserPassword)<6 then error"<li>������������ٰ��� 6 ���ַ�"
		
		Set Rs=Execute("Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'")
			if Not Rs.eof then
				if ""&Rs("PasswordAnswer")&""="" then Alert("�������������Ϊ�գ�����ͨ���˷�ʽȡ������")
				if md5(""&Request("PasswordAnswer")&"")<>Rs("PasswordAnswer") then Alert("�𰸴���")
				Rs.close
					
				'--------------------  API Start  --------------------
				Message=""
				If SiteConfig("APIEnable")=1 Then APIUpdateUser UserName,UserPassword,""
				if Message<>"" then error(""&Message&"")
				'--------------------  API End  ----------------------
		
				ModifyUserPassword UserName,UserPassword,"","",""
		
		
				Message=Message&"<li>���������óɹ�</li><li><a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">�뷵�ص�¼</a></li>"
				succeed Message,"Login.asp"

			else
				Rs.close
				error "��������û�������"
			End if
		Rs.close
		
		
	case "MailRecoverok"
		if SiteConfig("SelectMailMode")="" then error("ϵͳδ���� �ʼ� ���ܣ�")
		if UserName="" then error "URL��������û���û�����"
		if Execute("Select UserName from ["&TablePrefix&"UserActivation] where ActivationKey='"&ActivationKey&"' and UserName='"&UserName&"'").eof then error("�һ��������Ϣ�ѹ��ڣ��������ύ�һأ�")

		SetNewPassword
end select

Sub default
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �һ�����</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea align="center">
<form action="RecoverPassword.asp?menu=SelectMethod" method="POST">
	<tr class=CommonListTitle>
		<td align="center">ȡ������</td>
	</tr>
	<tr class="CommonListCell">
    	<td>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td colspan="2" style="padding-left:50px;">��һ��������������û�����</td>
		</tr>
		<tr>
			<td align="right" width="30%"><b>�û�����</b></td>
			<td><input name="UserName" size="30" /></td>
		</tr>
		<tr>
			<td align="right" width="30%"><b>��֤�룺</b></td>
			<td>
            <input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span>
            </td>
		</tr>
		<tr>
			<td align="center" colspan="2"> <input type="submit" value=" ��һ�� " /> </td>
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
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �һ�����</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea align="center">
<form action="RecoverPassword.asp" method="POST">
<input type=hidden name=UserName value="<%=UserName%>" />
	<tr class=CommonListTitle>
		<td align="center">ȡ������</td>
	</tr>
	<tr class="CommonListCell">
    	<td>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td colspan="2" style="padding-left:50px;">�ڶ�������ѡ��һ�������������������룺</td>
		</tr>
        
		<tr>
			<td colspan="2" style="padding-left:100px;">
            	<input type="radio" value="QuestionRecover" name="menu" id="QuestionRecover" onclick="SelectOption('Question')" checked="checked" />&nbsp;<label for="QuestionRecover">ʹ�� ��������&��ʾ�� ȡ������</label>
            </td>
		</tr>
		<tr>
			<td colspan="2" style="display:" id="Question">
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
			<tr>
				<td width="30%" align="right"><b>������ʾ���⣺</b></td>
				<td><%=Execute("Select PasswordQuestion from ["&TablePrefix&"Users] where UserName='"&UserName&"'")(0)%></td>
			</tr>
			<tr>
				<td width="30%" align="right"><b>��ʾ����𰸣�</b></td>
				<td><input size="15" name="PasswordAnswer" /></td>
			</tr>
			<tr>
	    		<td align="right" width="30%"><b>�����룺</b></td>
				<td><input type="password" name="UserPassword" size="40" maxLength="16" onkeyup="EvalPwd(this.value);" onchange="EvalPwd(this.value);" /></td>
			</tr>
			<tr>
	    		<td align="right" width="30%"><b>����ǿ�ȣ�</b></td>
	    		<td>
					<table border="0" width="250" cellspacing="1" cellpadding="2">
						<tr bgcolor="#f1f1f1">
							<td id=iWeak align="center">��</td>
							<td id=iMedium align="center">��</td>
							<td id=iStrong align="center">ǿ</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align="right" width="30%"><b>���¼��������룺</b><br />�������������뱣��һ��</td>
				<td valign="middle"> <input type="password" name="UserPassword2" size="40" maxlength="16" /></td>
			</tr>
			<tr>
				<td align="center" colspan="2"> <input type="submit" value=" ȷ�� " />��<input type="button" onclick="javascript:history.back()" value=" ȡ�� " /> </td>
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
            	<input value="MailRecover" name="menu" type="radio" id="MailRecover" onclick="SelectOption('Email')" />&nbsp;<label for="MailRecover">ʹ�� Email ȡ������</label>
            </td>
		</tr>
		<tr>
			<td colspan="2" style="display:none" id="Email">
            <hr width="80%" align="center" noshade="noshade" color="#CCCCCC" />
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
            <tr>
				<td align="right"><b>�����ʼ���ַ��</b></td>
				<td><input name="UserEmail" size="30" /></td>
			</tr>
			<tr>
				<td align="center" colspan="2"> <input type="submit" value=" ȷ�� " />��<input type="button" onclick="javascript:history.back()" value=" ȡ�� " /> </td>
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
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �һ�����</div>
<script language="javascript" type="text/javascript" src="Utility/pswdplc.js"></script>
<form method="POST" name="form" action="?Menu=setNewPassword" onsubmit="return VerifyInput();">
<input type=hidden name=UserName value="<%=UserName%>" />
<input type=hidden name=ActivationKey value="<%=ActivationKey%>" />
<table width="100%" border="0" cellspacing="1" cellpadding="5" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center" colspan=2>���� <%=UserName%> ��������</td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b> �����룺</b></td>
		<td align="Left" width="55%"> <input type="password" name="UserPassword" size="40" maxLength="16" onkeyup="EvalPwd(this.value);" onchange="EvalPwd(this.value);" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b>����ǿ�ȣ�</b></td>
	    <td align="Left" width="55%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">��</td>
					<td id=iMedium align="center">��</td>
					<td id=iStrong align="center">ǿ</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="45%"><b>���¼��������룺</b><br />�������������뱣��һ��</td>
		<td align="Left" valign="middle" width="55%"> <input type="password" name="UserPassword2" size="40" maxlength="16" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="2"> <input type="submit" value=" ȷ �� " /></td>
	</tr>
</table>
</form>
<script language="JavaScript">
function VerifyInput()
{
	if (document.form.UserPassword.value.length < 6)
	{
		alert("���������볤�ȱ������5��");
		document.form.UserPassword.focus();
		return false;
	}
	if (document.form.UserPassword.value != document.form.UserPassword2.value)
	{
		alert("�����μ���������벻ͬ��");
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