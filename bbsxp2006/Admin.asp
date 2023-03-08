<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" --><%

if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

if membercode < 5 then error("<li>您没有权限进入后台")

if Request("menu")="out" then
session("pass")=""
Log("退出后台管理")
response.redirect ""&Request.ServerVariables("script_name")&"/../"
end if

%>
<title><%=SiteSettings("SiteName")%>后台管理 - Powered By BBSXP</title>
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
		<font face="幼圆">第六代 BBS 系统 -- BBSXP</font> <font color="ffff66">安全</font>
		<font color="ff0033">快速</font> <font color="33ff33">方便 </font>
		<font color="0000ff">可靠 </font><font color="000000">可升级</font></b> </td>
	</tr>
</table>
<%
end sub




sub pass

session("pass")=md5(""&Request("pass")&"")
if SiteSettings("AdminPassword")<>session("pass") then error2("社区管理密码错误")


Log("登录后台管理")






%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>系统信息</b></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">服务器域名：<%=Request.ServerVariables("server_name")%> 
		/ <%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="50%">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">服务器软件的名称：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="50%">ACCESS 数据库路径：<a target="_blank" href="<%=datapath%><%=datafile%>"><%=datapath%><%=datafile%></a></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">FSO 文本读写：<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">SA-FileUp 文件上传：<%If Not IsObjInstalled("SoftArtisans.FileUp") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
	<tr class="a4">
		<td width="50%" height="23">JMail 组件支持：<%If Not IsObjInstalled("JMail.Message") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%></td>
		<td width="50%">JPEG 组件支持：<%If Not IsObjInstalled("Persits.Jpeg") Then%><font color="Persits.Jpegred"><b>×</b></font><%else%><b>√</b><%end if%></td>
	</tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>管理快捷方式</b></td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">快速查找用户</td>
		<td width="80%" height="23">
		<form method="POST" action="Admin_User.asp?menu=User2">
			<input size="13" name="UserName"> <input type="submit" value="立刻查找">
		</td>
		</form>	</tr>
	<tr class="a4">
		<td width="20%" height="23">快捷功能链接</td>
		<td width="80%" height="23"><a href="Admin_bbs.asp?menu=classs">
		<font color="000000">建立论坛数据</font></a> |
		<a href="Admin_bbs.asp?menu=bbsManage"><font color="000000">管理论坛资料</font></a> 
		| <a href="Admin_bbs.asp?menu=upSiteSettingsok"><font color="000000">更新论坛资料</font></a></td>
	</tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr>
		<td class="a1" colspan="2" height="25" align="center"><b>授权信息</b></td>
	</tr>
	<tr class="a4">
		<td width="20%">注册网址</td>
		<td width="80%"><%=Request.ServerVariables("server_name")%></td>
</tr>
	<tr class="a4">
		<td width="20%">授权时间</td>
		<td width="80%"><span id="Licence">?</span></td>
</tr>
	<tr class="a4">
		<td width="20%">使用种类</td>
		<td width="80%"><span id="Class">?</span></td>
</tr>
	</table>
<br>

<table cellpadding="2" cellspacing="1" border="0" width="95%" align="center" class="a2">
	<tr class="a4">
		<td class="a1" colspan="2" height="25" align="center"><b>关于BBSXP论坛</b></td>
	</tr>
	<tr class="a4">
		<td width="20%">产品开发</td>
		<td width="80%">BBSXP开发组</td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23">产品负责</td>
		<td width="80%">YUZI网络有限公司</td>
	</tr>
	<tr class="a4">
		<td width="20%" height="23" valign="top">联系方法</td>
		<td width="80%">电话：0595-22205408<br>
		传真：0595-22205409<br>
		网址：<a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a><br>
		地址：中国福建省泉州市泉秀路农行大厦25楼I座</td>
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
				登录论坛管理</div>
			</td>
		</tr>
		<tr>
			<td height="19" width="328" valign="top" class="a3">
			<div align="center">
				请输入管理密码: <input size="15" name="pass" type="password"><br>
				<input type="submit" value=" 登录 ">　<input type="reset" value=" 取消 ">
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
			<td align="middle">系统设置</td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=variable">
						<font color="000000">常规设置</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=SiteStat">
						<font color="000000">统计设置</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=Showbanner">
						<font color="000000">广告设置</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_Setup.asp?menu=agreement">
						<font color="000000">注册用户协议</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			公告<span>管理</span></td>
		</tr>
		<tr>
			<td class="a3" id="TableTitleLink" height="25" align="center">
						<a target="main" href="Admin_Affiche.asp?menu=Affichelist">
						<font color="000000">社区公告管理</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>用户管理</span></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=Showall">
						<font color="000000">注册用户管理</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=UserRank">
						<font color="000000">用户等级名称</font></a></td>
		</tr>		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_User.asp?menu=Activation">
						<font color="000000">激活用户资料</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>论坛管理</span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=classs">
						<font color="000000">建立论坛数据</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=ApplyManage">
						<font color="000000">管理所有论坛</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=bbsManage">
						<font color="000000">管理论坛资料</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=activation&checkbox=1&type=Recycle"><font color="000000">回收站</font></a> <a target="main" href="Admin_bbs.asp?menu=activation&checkbox=1&type=Censorship"><font color="000000">审查区</font></a></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_bbs.asp?menu=upSiteSettings">
						<font color="000000">更新论坛资料</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>社区管理</span> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=sendMail">
						<font color="000000">群发邮件</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Message">
						<font color="000000">短讯息管理</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Link">
						<font color="000000">友情链接管理</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_club.asp?menu=Consortia">
						<font color="000000">社区公会管理</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>菜单管理</span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_menu.asp">
						<font color="000000">论坛菜单管理</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>数据库管理 </span></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=bak">
						<font color="000000">ACCESS数据库</font></a> </td>
		</tr>		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=statroom">
						<font color="000000">统计占用空间</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=files">
						<font color="000000">管理上传文件</font></a></td>
		</tr>

		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=PostAttachment">
						<font color="000000">管理数据库附件</font></a> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_fso.asp?menu=Posts">
						<font color="000000">帖子数据表</font></a> </td>
		</tr>

		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
			<span>其他功能</span> </td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=circumstance">
						<font color="000000">主机环境变量</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=discreteness">
						<font color="000000">组件支持情况</font></a></td>
		</tr>
		<tr class="a3" id="TableTitleLink" height="25">
			<td align="middle">
						<a target="main" href="Admin_other.asp?menu=Log">
						<font color="000000">系统操作记录</font></a></td>
		</tr>
		<tr class="a1" id="TableTitleLink" height="25">
			<td align="middle">
						 <a target="_top" href="?menu=out">退出管理</a></td>
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