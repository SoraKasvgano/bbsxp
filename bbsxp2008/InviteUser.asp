<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
if SiteConfig("AccountActivation")<>2 then error("ϵͳδ���� ������ע�� ѡ�")
if SiteConfig("SelectMailMode")="" then error("ϵͳδ���� �ʼ����� ���ܣ�")

if Request("menu")="InvitingOk" then
	if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("��֤�����")
	UserEmail=HTMLEncode(Request("UserEmail"))
	if UserEmail="" then error("�����������˵�Email��ַ��")
	if instr(UserEmail,"@")=0 then error("�����˵�Email��ַ��д����")
	
	If not Execute("Select UserID From ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"'" ).eof Then Error("<li>"&UserEmail&" �Ѿ��ڱ���̳ע����")

	
	Randomize
	ActivationKey=int(rnd*9999999999)+1

	LoadingEmailXml("InviteNewUser")
	MailBody=Replace(MailBody,"[UserName]",CookieUserName)
	MailBody=Replace(MailBody,"[SiteName]",SiteConfig("SiteName"))
	MailBody=Replace(MailBody,"[InviteURL]","<a target=_blank href="&SiteConfig("SiteUrl")&"/CreateUser.asp?menu=WriteProfile&ActivationKey="&ActivationKey&"&ReferrerName="&CookieUserName&"&UserEmail="&UserEmail&">"&SiteConfig("SiteUrl")&"/CreateUser.asp?menu=WriteProfile&ActivationKey="&ActivationKey&"&ReferrerName="&CookieUserName&"&UserEmail="&UserEmail&"</a>")
	SendMail UserEmail,MailSubject,MailBody
	
	Execute("insert into ["&TablePrefix&"UserActivation] (ActivationKey,Email) values ('"&ActivationKey&"','"&UserEmail&"')")
	Session("VerifyCode")=""
	succeed "�����뷢�ͳɹ���","Default.asp"
else
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ����������</div>
<table cellspacing=1 cellpadding=5 width=80% class=CommonListArea>
	<form action="?menu=InvitingOk" method="POST" name=form>
		<tr class=CommonListTitle>
			<td align="center" colspan="2">����������</td>
		</tr>
		<tr class="CommonListCell">
			<td width="30%" align="right">��֤�룺</td>
			<td>
				<input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">�����˵������䣺</td>
			<td>
				<input type="text" name="UserEmail" size=30 />
			</td>
		</tr>
		<tr class="CommonListCell">
			<td valign="top" align="center" colspan="2">
				<input type="submit" value=" ���� " /> <input type="button" onclick="javascript:history.back()" value=" ȡ�� " />
			</td>
		</tr>
	</form>
	</table>
<%
end if
HtmlBottom
%>