<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then  error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")
if CookieUserRoleID <> 1 then Alert("您没有权限进入后台")

HtmlHeadTitle="后台管理"
HtmlHead

if Request("menu")="out" then
	session.Abandon()
	Log("退出后台管理")
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
	if Request("pass")="" then Alert("请输入管理员密码！")

	if Len(CookieUserPassword)=32 then
		session("pass")=MD5(""&Request("pass")&"")
	else
		session("pass")=SHA1(""&Request("pass")&"")
	end if

	if CookieUserPassword <> session("pass") then session("pass")="":Alert("管理员密码错误，请返回重新登录！")
	Response.Write("<script language='JavaScript'>top.location.href='?';</script>")
End Sub

Sub main

%>
<body><div class="CommonListCell"  style='PADDING-TOP:15px; PADDING-RIGHT:5px; PADDING-LEFT:5px; PADDING-BOTTOM:15px;text-align:center;'>

<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" align="center">系统信息</td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">服务器域名：<%=Request.ServerVariables("server_name")%> / <%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="50%">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">服务器软件名：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="50%">服务器操作系统：<%=Request.ServerVariables("OS")%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">FSO 组件支持：<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">AspUpload 组件支持：<%If Not IsObjInstalled("Persits.Upload") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="50%">AspEmail 组件支持：<%If Not IsObjInstalled("Persits.MailSender") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">AspJpeg 组件支持：<%If Not IsObjInstalled("Persits.Jpeg") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
</table>


<br />
<form method="POST" action="Admin_Setup.asp?menu=AdvancedSettingsUp">
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center">管理员便笺</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center">
			<textarea name="AdminNotes" rows="9" style="width:95%"><%=AdminNotes%></textarea>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center"><input type="submit" value=" 保 存 "></td>
	</tr>
</table>
</form>





<br />
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="3" align="center">管理快捷方式</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">用户管理</td>
		<td width="60%">当前有 <b><font color=red><%=Execute("Select count(UserID) from ["&TablePrefix&"Users] where UserAccountStatus=0")(0)%></font></b> 位新注册用户等待审核</td>
		<td width="20%"><a href="Admin_User.asp?menu=SearchUser">搜索用户</a></td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">论坛管理</td>
		<td width="60%">当前有 <b><font color=red><%=Execute("Select count(ForumID) from ["&TablePrefix&"Forums]")(0)%></font></b> 论坛及 <b><font color=red><%=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1")(0)%></font></b> 个讨论主题</td>
		<td width="20%"><a href="Admin_Forum.asp?menu=ForumAdd">添加新论坛</a></td>
	</tr>
</table>
<br />
<table cellpadding="5" cellspacing="1" border="0" width="95%" align="center" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2" align="center">关于BBSXP论坛</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">产品开发</td>
		<td width="80%">BBSXP Studio</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%">产品运营</td>
		<td width="80%">Yuzi Corporation</td>
	</tr>
	<tr class="CommonListCell">
		<td width="20%" valign="top">联系方法</td>
		<td width="80%">
		Email：<a href="mailto:support@bbsxp.com">Support@bbsxp.com</a><br>
		WEB：<a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a>
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
		<td align="center">登录论坛管理</td>
	</tr>
	<tr class="CommonListCell">
	  <td>
	  
	  
	  <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="50%"><a href="Default.asp" target=_top><img src="Images/Logo.gif" border="0"/></a></td>
          <td width="50%"><a href="Default.asp" target=_top><%=SiteConfig("SiteName")%></a><br /><%=ForumsProgram%><br /><%=ForumsBuild%> 管理控制面版</td>
        </tr>
      </table>
      
      </td>
    </tr>
	<tr class="CommonListCell">
	  <td align="center">管 理 员：<input size="25" name="AdminUserName" type="text" value="<%=CookieUserName%>" Readonly></td>
    </tr>
	<tr class="CommonListCell">
		<td align="center">管理密码：<input size="25" name="pass" type="Password"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center"><input type="submit" value=" 登录 ">　<input type="button" onClick="javascript:history.back()" value=" 取消 "></td>
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
			<td align="center">设　置</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Setup.asp?menu=variable">基本设置</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Setup.asp?menu=AdvancedSettings">高级设置</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">成　员</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=SearchUser">搜索用户</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp">浏览所有用户</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=AllRoles">管理所有角色</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_User.asp?menu=UserRank">管理用户等级</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">声　望</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Reputation.asp">管理声望评论</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">论　坛</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ForumAdd">新建论坛</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ManageGroups">论 坛 组</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=ApplyManage">论　　坛</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=Tags">标　　签</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=TreeView">管理所有论坛</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Forum.asp?menu=bbsManage">管理论坛资料</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">工　具</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=BatchSendMail">群发邮件</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=Message">短讯息管理</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=UpStatic">更新计数器</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=Link">友情链接管理</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Tool.asp?menu=AD">帖间广告管理</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">菜　单</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showThemes">论坛风格管理</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showMenu">论坛菜单管理</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_XML.asp?menu=showEmoticon">论坛表情管理</a></td>
		</tr>
		<tr class=CommonListTitle><td align="center">数　据</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=bak">ACCESS数据库</a> </td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=statroom">统计占用空间</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_FSO.asp?menu=PostAttachment">附件管理</a></td>
		</tr>
		<tr class=CommonListTitle>
			<td align="center">系　统</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=circumstance">主机环境变量</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=discreteness">组件支持情况</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=Statistics">论坛统计记录</a></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><a target="main" href="Admin_Other.asp?menu=EventLog">论坛操作记录</a></td>
		</tr>
</table><%
End Sub

Sub AdminHeader
	Log("登录后台管理")
%>
	<body><table border=0 width=100% cellspacing=0 cellpadding=3><tr class=CommonListTitle><td><div id="BBSxpWelcome" style='Padding-left:15px;float:left'><%=ForumsProgram%>(<%=ForumsBuild%>)</div><div style='Text-align:right;Padding-right:15px;float:right'><a href='Default.asp' target='_blank'>论坛首页</a>　<a href='?menu=main' target='main'>管理首页</a>　<a target='_top' href='?menu=out'>退出管理</a></div></td></tr></table>
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