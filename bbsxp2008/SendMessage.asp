<!-- #include file="Setup.asp" -->
<%
HtmlHeadTitle="��ϵ����"
HtmlTop

if Request_Method <> "POST" then
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ��ϵ����</div>
<form method="Post" name="form"  onSubmit="return chkInput(this);">
<table cellspacing="1" cellpadding="5" width="700" class="CommonListArea" align="center">
	<tr class=CommonListTitle>
		<td colspan="2" align="center">�����ʼ�����̳����Ա</td>
	</tr>
	<tr class="CommonListCell">
		<td>
        
        <div style="width:640px" align="left">
			<fieldset>
				<legend>������Ϣ</legend>
				<table cellpadding="0" cellspacing="3" border="0">
				<tr>
					<td>�������� :<br /><input type="text" name="UserName" value="<%=CookieUserName%>" size="50" /></td>
				</tr>
				<tr>
					<td>Email ��ַ :<br /><input type="text" name="UserEmail" value="<%=CookieUserEmail%>" size="50" /></td>
				</tr>
				<%if CookieUserName=empty then%><tr>
					<td>��֤�룺<br /><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span></td>
				</tr><%end if%>
				</table>
			</fieldset>
       		<br>
           		<fieldset>

				<legend>�ʼ���Ϣ</legend>
				<table cellpadding="0" cellspacing="3" border="0">
				<tr>
					<td>
						���� :<br />
						<input type="text" name="MailSubject" value="" size="50" />
					</td>
				</tr>
				<tr>
					<td>
						���� :<br />
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
        <input type="button" value=" ���� " accesskey="s"/ onClick="ToMail()">
	<%else%>        
         <input type="submit" value=" ���� " accesskey="s"/>
	<%end if%>         
         <input type="reset" class="button" name="reset" value="���ñ�" accesskey="r" /></td>
	</tr>
    
</table>
</form>
<script language="javascript" type="text/javascript">
function chkInput(form){
	if (form.UserName.value==''){
		alert('��������������');
		return false;
	}
	if (form.UserEmail.value==''){
		alert('������Email ��ַ');
		return false;
	}
	if (form.MailSubject.value==''){
		alert('����������');
		return false;
	}
	if (form.MailBody.value==''){
		alert('����������');
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
		error("<li>��֤�����</li>")
	else
		Session("VerifyCode")=""
	end if
end if


UserName=HTMLEncode(Request.Form("UserName"))
UserEmail=HTMLEncode(Request.Form("UserEmail"))
MailSubject=HTMLEncode(Request.Form("MailSubject"))
MailBody=HTMLEncode(Request.Form("MailBody"))

MailBody="�����ˣ�<A href='mailto:"&UserEmail&"'>"&UserName&"</a><br /><br />"&MailBody


	SendMail SiteConfig("WebMasterEmail"),MailSubject,MailBody
	Message="<li>�ʼ����ͳɹ�<li><a href=Default.asp>������ҳ</a>"
	succeed Message,"SendMessage.asp"

%>