<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" --><%

if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

if membercode < 5 then error("<li>��û��Ȩ�޽����̨")

if Request("menu")="out" then
session("pass")=""
Log("�˳���̨����")
response.redirect ""&Request.ServerVariables("script_name")&"/../"
end if

%>
<title><%=SiteSettings("SiteName")%>��̨���� - Powered By BBSXP</title>
<%
select case Request("menu")
case ""
index
case "Left"
Left
case "top"
top2
case "Login"
Login

case "pass"
pass



end select


sub top2
%><body topmargin="0" rightmargin="0" Leftmargin="0"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
	<tr>
		<td width="100%" class="a1" height="20" align="center"><b>
		<font face="��Բ">������ BBS ϵͳ -- BBSXP</font> <font color="ffff66">��ȫ</font>
		<font color="ff0033">����</font> <font color="33ff33">���� </font>
		<font color="0000ff">�ɿ� </font><font color="000000">������</font></b> </td>
	</tr>
</table>
<%
end sub




sub pass

session("pass")=md5(""&Request("pass")&"")
if SiteSettings("AdminPassword")<>session("pass") then error2("���������������")


Log("��¼��̨����")






%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>ϵͳ��Ϣ</b></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">������������<%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="50%">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">��������������ƣ�<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="50%">ACCESS ���ݿ�·����<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">FSO �ı���д��<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
		<td width="50%">SA-FileUp �ļ��ϴ���<%If Not IsObjInstalled("SoftArtisans.FileUp") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">JMail ���֧�֣�<%If Not IsObjInstalled("JMail.Message") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
		<td width="50%">JPEG ���֧�֣�<%If Not IsObjInstalled("Persits.Jpeg") Then%><font color="Persits.Jpegred"><b>��</b></font><%else%><b>��</b><%end if%></td>
	</tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>�����ݷ�ʽ</b></td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">���ٲ����û�</td>
		<td width="80%" height="23">
		<form method="POST" action="Admin_User.asp?menu=User2">
			<input size="13" name="UserName"> <input type="submit" value="���̲���">
		</td>
		</form>	</tr>
	<tr class="a4">
		<td width="20%" height="23">��ݹ�������</td>
		<td width="80%" height="23"><a href="Admin_bbs.asp?menu=classs">
		<font color="000000">������̳����</font></a> |
		<a href="Admin_bbs.asp?menu=bbsManage"><font color="000000">������̳����</font></a> 
		| <a href="Admin_bbs.asp?menu=upSiteSettingsok"><font color="000000">������̳����</font></a></td>
	</tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>��Ȩ��Ϣ</b></td>
	</tr>
	<tr class="a4">
		<td width="20%">ע����ַ</td>
		<td width="80%"><%=Request.ServerVariables("server_name")%></td>
</tr>
	<tr class="a4">
		<td width="20%">��Ȩʱ��</td>
		<td width="80%"><span id="Licence">?</span></td>
</tr>
	<tr class="a4">
		<td width="20%">ʹ������</td>
		<td width="80%"><span id="Class">?</span></td>
</tr>
	</table>
<br>

<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr class="a4">
		<td class="a1" colspan="2" height="25" align="center"><b>����BBSXP��̳</b></td>
	</tr>
	<tr class="a4">
		<td width="20%">��Ʒ����</td>
		<td width="80%">BBSXP������</td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">��Ʒ����</td>
		<td width="80%">YUZI�������޹�˾</td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23" valign="top">��ϵ����</td>
		<td width="80%">�绰��0595-22205408<br>
		���棺0595-22205409<br>
		��ַ��<a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a><br>
		��ַ���й�����ʡȪ����Ȫ��·ũ�д���25¥I��</td>
	</tr>
	</table>
<script src="http://Licence.yuzi.net/Licence.asp?BBSxpWeb=<%=Request.ServerVariables("server_name")%>"></script>

<%
Conn.execute("Delete from [BBSXP_Log] where DateCreated<"&SqlNowString&"-7")

htmlend
end sub




sub Login





%> <br>
<br>
<br>
<form action="Admin.asp" method="POST">
	<input type="hidden" value="pass" name="menu">
	<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class="a2">
		<tr>
			<td width="328" class="a1" height="25">
			<div align="center">
				��¼��̳����</div>
			</td>
		</tr>
		<tr>
			<td height="19" width="328" valign="top" class="a3">
			<div align="center">
				�������������: <input size="15" name="pass" type="password"><br>
				<input type="submit" value=" ��¼ ">��<input type="reset" value=" ȡ�� ">
			</div>
			</td>
		</tr>
	</table>
</form>
<%
htmlend
end sub

sub Left

%>
<style>BODY { MARGIN: 0px;}</style>
<table class="a2" cellSpacing="1" cellPadding="3" width="145" align="center" border="0">
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">ϵͳ����</td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=variable">
						<font color="000000">��������</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=SiteStat">
						<font color="000000">ͳ������</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=Showbanner">
						<font color="000000">�������</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=agreement">
						<font color="000000">ע���û�Э��</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			����<span>����</span></td>
		</tr>
		<tr>
			<td class="a3" id="TableTitleLink" height="25" align="center">
						<a target="main" href="Admin_Affiche.asp?menu=Affichelist">
						<font color="000000">�����������</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>�û�����</span></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=Showall">
						<font color="000000">ע���û�����</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=UserRank">
						<font color="000000">�û��ȼ�����</font></a></td>
		</tr>		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=Activation">
						<font color="000000">�����û�����</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>��̳����</span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=classs">
						<font color="000000">������̳����</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=ApplyManage">
						<font color="000000">����������̳</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=bbsManage">
						<font color="000000">������̳����</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=activation&checkbox=1&type=Recycle"><font color="000000">����վ</font></a> <a target="main" href="Admin_bbs.asp?menu=activation&checkbox=1&type=Censorship"><font color="000000">�����</font></a></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=upSiteSettings">
						<font color="000000">������̳����</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>��������</span> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=sendMail">
						<font color="000000">Ⱥ���ʼ�</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Message">
						<font color="000000">��ѶϢ����</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Link">
						<font color="000000">�������ӹ���</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Consortia">
						<font color="000000">�����������</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>�˵�����</span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_menu.asp">
						<font color="000000">��̳�˵�����</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>���ݿ���� </span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=bak">
						<font color="000000">ACCESS���ݿ�</font></a> </td>
		</tr>		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=statroom">
						<font color="000000">ͳ��ռ�ÿռ�</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=files">
						<font color="000000">�����ϴ��ļ�</font></a></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=PostAttachment">
						<font color="000000">�������ݿ⸽��</font></a> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=Posts">
						<font color="000000">�������ݱ�</font></a> </td>
		</tr>

		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>��������</span> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=circumstance">
						<font color="000000">������������</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=discreteness">
						<font color="000000">���֧�����</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=Log">
						<font color="000000">ϵͳ������¼</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
						 <a target="_top" href="?menu=out">�˳�����</a></td>
		</tr>
	</table><%
end sub
sub index


%>
<style type="text/css">.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}</style>
<script>
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>
<body scroll="no" topmargin="0" rightmargin="0" Leftmargin="0">

<table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
	<tr>
		<td align="middle" id="frmTitle" nowrap valign="center" name="frmTitle">
		<iframe frameborder="0" id="carnoc" name="carnoc" src="?menu=Left" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 170px; Z-INDEX: 2">
		</iframe></td>
		<td style="WIDTH: 100%">
		<iframe frameborder="0" id="main" name="top" scrolling="no" scrolling="yes" src="?menu=top" style="HEIGHT: 4%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe>
		<iframe frameborder="0" id="main" name="main" scrolling="yes" src="?menu=Login" style="HEIGHT: 96%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
		</iframe></td>
	</tr>
</table>
</center>

</html>
<%
end sub
CloseDatabase
%>