<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then  error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
if CookieUserRoleID <> 1 then Alert("��û��Ȩ�޽����̨")

HtmlHeadTitle="��̨����"
HtmlHead

if Request("menu")="out" then
	session.Abandon()
	Log("�˳���̨����")
	response.redirect ""&Request.ServerVariables("script_name")&"/../"
end if


select case Request("menu")
	case "AdminLeft"
		AdminLeft
	case "pass"
		pass
	case "Header"
		AdminHeader
	case "Login"
		Login
	case "main"
		main
	case else
		if session("pass")="" then response.redirect("?menu=Login")
		Default
end select

Sub pass
	if Request("pass")="" then Alert("���������Ա���룡")

	if Len(CookieUserPassword)=32 then
		session("pass")=MD5(""&Request("pass")&"")
	else
		session("pass")=SHA1(""&Request("pass")&"")
	end if

	if CookieUserPassword <> session("pass") then session("pass")="":Alert("����Ա��������뷵�����µ�¼��")
	Response.Write("<script language='JavaScript'>top.location.href='?';</script>")
End Sub

Sub main

%>
<body><div class="CommonListCell"  style='PADDING-TOP:15px; PADDING-RIGHT:5px; PADDING-LEFT:5px; PADDING-BOTTOM:15px;text-align:center;'>

<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" align="center">ϵͳ��Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">������������<%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="50%">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">�������������<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="50%">����������ϵͳ��<%=Request.ServerVariables("OS")%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">FSO ���֧�֣�<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
		<td width="50%">AspUpload ���֧�֣�<%If Not IsObjInstalled("Persits.Upload") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">AspEmail ���֧�֣�<%If Not IsObjInstalled("Persits.MailSender") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
		<td width="50%">AspJpeg ���֧�֣�<%If Not IsObjInstalled("Persits.Jpeg") Then%><font color="red"><b>��</b></font><%else%><b>��</b><%end if%></td>
	</tr>
</table>


<br />
<form method="POST" action="Admin_Setup.asp?menu=AdvancedSettingsUp">
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center">����Ա���</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center">
			<textarea name="AdminNotes" rows="9" style="width:95%"><%=AdminNotes%></textarea>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center"><input type="submit" value=" �� �� "></td>
	</tr>
</table>
</form>





<br />
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="3" align="center">�����ݷ�ʽ</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">�û�����</td>
		<td width="60%">��ǰ�� <b><font color=red><%=Execute("Select count(UserID) from ["&TablePrefix&"Users] where UserAccountStatus=0")(0)%></font></b> λ��ע���û��ȴ����</td>
		<td width="20%"><a href="Admin_User.asp?menu=SearchUser">�����û�</a></td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">��̳����</td>
		<td width="60%">��ǰ�� <b><font color=red><%=Execute("Select count(ForumID) from ["&TablePrefix&"Forums]")(0)%></font></b> ��̳�� <b><font color=red><%=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1")(0)%></font></b> ����������</td>
		<td width="20%"><a href="Admin_Forum.asp?menu=ForumAdd">�������̳</a></td>
	</tr>
</table>
<br />
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" align="center">����BBSXP��̳</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">��Ʒ����</td>
		<td width="80%">BBSXP Studio</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">��Ʒ��Ӫ</td>
		<td width="80%">Yuzi Corporation</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%" valign="top">��ϵ����</td>
		<td width="80%">
		Email��<a href="mailto:support@bbsxp.com">Support@bbsxp.com</a><br>
		WEB��<a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a>
	</tr>
</table>
<%



	Execute("Delete from ["&TablePrefix&"EventLog] where DateDiff("&SqlChar&"ww"&SqlChar&",EventDate,"&SqlNowString&") > 1")
	Execute("Delete from ["&TablePrefix&"Statistics] where DateDiff("&SqlChar&"yyyy"&SqlChar&",DateCreated,"&SqlNowString&") > 3")
	Execute("Delete from ["&TablePrefix&"UserActivation] where DateDiff("&SqlChar&"ww"&SqlChar&",DateCreated,"&SqlNowString&") > 1")

	AdminFooter
End Sub

Sub Login
%>
<br /><br /><br /><br /><br /><br /><br /><br />
<table width="400" border="0" cellspacing="1" cellpadding="5" align="center" class=CommonListArea>
<form action="?" method="POST">
<input type="hidden" value="pass" name="menu">
	<tr class=CommonListTitle>
		<td align="center">��¼��̳����</td>
	</tr>
	<tr class="CommonListCell">
	  <td>
	  
	  
	  <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="50%"><a href="Default.asp" target=_top><img src="Images/Logo.gif" border="0"/></a></td>
          <td width="50%"><a href="Default.asp" target=_top><%=SiteConfig("SiteName")%></a><br /><%=ForumsProgram%><br /><%=ForumsBuild%> ����������</td>
        </tr>
      </table>
      
      </td>
    </tr>
	<tr class="CommonListCell">
	  <td align="center">�� �� Ա��<input size="25" name="AdminUserName" type="text" value="<%=CookieUserName%>" Readonly></td>
    </tr>
	<tr class="CommonListCell">
		<td align="center">�������룺<input size="25" name="pass" type="Password"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center"><input type="submit" value=" ��¼ ">��<input type="button" onClick="javascript:history.back()" value=" ȡ�� "></td>
	</tr>
</form>	
</table>

<%
End Sub

Sub AdminLeft
%>
<style>BODY { MARGIN: 0px;}</style>
<table class=CommonListArea cellSpacing="1" cellPadding="5" width="145" align="center" border="0">
		<tr class=CommonListTitle>
			<td align="center">�衡��</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Setup.asp?menu=variable">��������</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Setup.asp?menu=AdvancedSettings">�߼�����</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">�ɡ�Ա</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=SearchUser">�����û�</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp">��������û�</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=AllRoles">�������н�ɫ</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=UserRank">�����û��ȼ�</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">������</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Reputation.asp">������������</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">�ۡ�̳</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ForumAdd">�½���̳</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ManageGroups">�� ̳ ��</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ApplyManage">�ۡ���̳</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=Tags">�ꡡ��ǩ</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=TreeView">����������̳</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=bbsManage">������̳����</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">������</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=BatchSendMail">Ⱥ���ʼ�</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=Message">��ѶϢ����</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=UpStatic">���¼�����</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=Link">�������ӹ���</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=AD">���������</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">�ˡ���</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showThemes">��̳������</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showMenu">��̳�˵�����</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showEmoticon">��̳�������</a></td>
		</tr>
		<tr class=CommonListTitle><td align="center">������</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=bak">ACCESS���ݿ�</a> </td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=statroom">ͳ��ռ�ÿռ�</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=PostAttachment">��������</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">ϵ��ͳ</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=circumstance">������������</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=discreteness">���֧�����</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=Statistics">��̳ͳ�Ƽ�¼</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=EventLog">��̳������¼</a></td>
		</tr>
</table><%
End Sub

Sub AdminHeader
	Log("��¼��̨����")
%>
	<body><table border=0 width=100% cellspacing=0 cellpadding=3><tr class=CommonListTitle><td><div id="BBSxpWelcome" style='Padding-left:15px;float:left'><%=ForumsProgram%>(<%=ForumsBuild%>)</div><div style='Text-align:right;Padding-right:15px;float:right'><a href='Default.asp' target='_blank'>��̳��ҳ</a>��<a href='?menu=main' target='main'>������ҳ</a>��<a target='_top' href='?menu=out'>�˳�����</a></div></td></tr></table>
<%
	response.flush()
	set XMLHTTP=server.createobject("MSXML2.ServerXMLHTTP")
	XMLHTTP.open "POST","http://www.bbsxp.com/Licence/",false
	XMLHTTP.Send "<root><ServerName>"&Request.ServerVariables("server_name")&"</ServerName><SiteName>"&SiteConfig("SiteName")&"</SiteName><SiteUrl>"&SiteConfig("SiteUrl")&"</SiteUrl><WebMasterEmail>"&SiteConfig("WebMasterEmail")&"</WebMasterEmail><Program>"&ForumsProgram&"</Program><Build>"&ForumsBuild&"</Build><Charset>"&BBSxpCharset&"</Charset><IPAddress>"&Request.ServerVariables("REMOTE_ADDR")&"</IPAddress><UserName>"&CookieUserName&"</UserName><HttpUserAgent>"&Request.ServerVariables("Http_User_Agent")&"</HttpUserAgent></root>"
	If XMLHTTP.ReadyState = 4 and XMLHTTP.Status = 200 Then getHTTPPage=XMLHTTP.responseText
	set XMLHTTP=nothing
%>
<script language="javascript">
	document.getElementById("BBSxpWelcome").innerHTML = "<%=getHTTPPage%>";
</script>
<%
End Sub

Sub Default
%><frameset cols="170,*" frameborder="no" ><frame src="?menu=AdminLeft" name="leftFrame" scrolling="yes" style="overflow: auto;"><frameset rows="20,*" framespacing="0"><frame src="?menu=Header" scrolling="no" marginheight="0"><frame src="?menu=main" name="main" marginheight="0"></frameset></frameset><%
End Sub

%>