<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
ForumID=int(Request("ForumID"))

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumLogo=Rs("ForumLogo")
followid=Rs("followid")
ForumName=Rs("ForumName")
moderated=Rs("moderated")
Rs.close

if ""&moderated&""="" then moderated="|"
moderated=split(moderated,"|")

if membercode<4 and moderated(0)<>CookieUserName then
error("<li>您的权限不够<li>只有正版主 <font color=red>"&moderated(0)&"</font> 、超级版主、管理员才拥有此权限")
end if



select case Request("menu")
case ""
error2("请选择您要操作的项目")

case "BatchRecycle"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsDel=0,lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&ho&" and IsDel=1 and ForumID="&ForumID&"")
next
Log("还原回收站内的帖子，主题ID："&Request.form("ThreadID")&"")
error2("成功还原回收站内的帖子")

case "BatchCensorship"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsDel=0 where id="&ho&" and IsDel=1 and ForumID="&ForumID&"")
next
Log("帖子通过审查，主题ID："&Request.form("ThreadID")&"")
error2("帖子已经成功通过审查")

case "BatchDel"
IsDel=int(Request.form("IsDel"))
for each ho in Request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsDel="&IsDel&",lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("批量删除，主题ID："&Request.form("ThreadID")&"")
error2("操作成功")

case "BatchGOOD"
IsGOOD=int(Request.form("IsGOOD"))
for each ho in Request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsGOOD="&IsGOOD&",lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("批量精华，主题ID："&Request.form("ThreadID")&"")
error2("操作成功")

case "BatchLocked"
IsLocked=int(Request.form("IsLocked"))
for each ho in Request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsLocked="&IsLocked&",lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("批量锁定，主题ID："&Request.form("ThreadID")&"")
error2("操作1"&IsLocked&"成功")

case "BatchSpecialTopic"
for each ho in Request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set SpecialTopic='"&Request.form("SpecialTopic")&"',lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("批量专题，主题ID："&Request.form("ThreadID")&"")
error2("操作成功")

case "BatchMoveTopic"
AimForumID=int(Request.form("AimForumID"))
if AimForumID="" then error("<li>您没有选择要将主题移动哪个论坛")
for each ho in Request.form("ThreadID")
ho=int(ho)
if Conn.Execute("Select ForumPass From [BBSXP_Forums] where id="&AimForumID&"")(0)=4 then error("<li>目标论坛为授权发帖状态")
Conn.execute("update [BBSXP_Threads] set ForumID="&AimForumID&",IsTop=0,IsGood=0,IsLocked=0 where id="&ho&" and ForumID="&ForumID&"")
next
Log("批量移动，主题ID："&Request.form("ThreadID")&"")
error2("操作成功")


case "Fix"
allarticle=Conn.execute("Select count(ID) from [BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&"")(0)
if allarticle>0 then
allrearticle=Conn.execute("Select sum(replies) from [BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&"")(0)
else
allrearticle=0
end if
Conn.execute("update [BBSXP_Forums] set ForumThreads="&allarticle&",ForumPosts="&allarticle+allrearticle&" where ID="&ForumID&"")
error2("修复论坛统计数据成功")


case "ForumDataUp"
ForumName=HTMLEncode(Request.Form("ForumName"))
TolSpecialTopic=HTMLEncode(Request.Form("TolSpecialTopic"))
ForumIcon=HTMLEncode(Request.Form("ForumIcon"))
ForumLogo=HTMLEncode(Request.Form("ForumLogo"))
moderated=HTMLEncode(Request.Form("moderated"))
ForumIntro=HTMLEncode(Request.Form("ForumIntro"))
ForumRules=HTMLEncode(Request.Form("ForumRules"))
if ForumName="" then error("<li>请输入论坛名称")
if Len(ForumName)>30 then  error("<li>论坛名称不能大于 30 个字符")
if Len(ForumIntro)>255 then  error("<li>论坛简介不能大于 255 个字符")
if instr(TolSpecialTopic,";") > 0 then error("<li>专题中不能含有特殊符号")

master=split(moderated,"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&master(i)&"'" ).eof and master(i)<>"" Then error("<li>"&master(i)&"的用户资料还未注册")
next

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Rs.Open sql,Conn,1,3
Rs("ForumName")=ForumName
Rs("moderated")=moderated
Rs("TolSpecialTopic")=TolSpecialTopic
Rs("ForumIcon")=ForumIcon
Rs("ForumLogo")=ForumLogo
Rs("ForumIntro")=ForumIntro
Rs("ForumRules")=ForumRules
Rs.update
Rs.close
Log("更新论坛（ID:"&ForumID&"）的信息！")
Message="<li>更新成功！<li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")

case "ForumData"
sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumIntro=replace(""&Rs("ForumIntro")&"","<br>",vbCrlf)
ForumRules=replace(""&Rs("ForumRules")&"","<br>",vbCrlf)
%>

<script>
if ("<%=Rs("ForumLogo")%>"!=''){Logo.innerHTML="<img border=0 src=<%=Rs("ForumLogo")%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(Rs("followid"))%><%=ForumTreeList%> <a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=Rs("ForumName")%></a> → 管理论坛</td>
		</tr>
	</table><br>
<table border="0" width="100%">
	<tr>
		<td valign="top" width="20%">
		


<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
	<tr class=a1>
		<td align="center">论坛信息 </td>
	</tr>
	<tr class=a3>
		<td>今日帖：<%=Rs("ForumToday")%></td>
	</tr>
	<tr class=a3>
		<td>主题数：<%=Rs("ForumThreads")%> </td>
	</tr>
	<tr class=a3>
		<td>帖子数：<%=Rs("ForumPosts")%> </td>
	</tr>
	<tr class=a1>
		<td align="center">管理选项</td>
	</tr>
	<tr class=a3>
		<td><a href="ShowForum.asp?ForumID=<%=ForumID%>&checkbox=1">批量管理帖子</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Censorship&ForumID=<%=ForumID%>&checkbox=1">主题审查</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Recycle&ForumID=<%=ForumID%>&checkbox=1">回 收 站</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Fix&ForumID=<%=ForumID%>">修复论坛统计数据</a></td>
	</tr>		
</table>


		</td>
		<td align="center">
		
		
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
<tr class=a1>
<td height="20" align="center" colspan="2"><b>论坛资料</b></td>
</tr>
<form name="form2" method="POST" action="?">
<input type=hidden name=menu value="ForumDataUp">
<input type=hidden name=ForumID value="<%=ForumID%>">
<tr class=a4>
<td align="right" valign="middle" width="20%">论坛名称：</td>
<td align="Left" valign="middle" width="78%">
<input type="text" name="ForumName" size="30" maxlength="12" value="<%=Rs("ForumName")%>">
</td>
</tr>
<tr class=a3>
<td align="right" valign="middle" width="20%">论坛版主：</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="moderated" value="<%=Rs("moderated")%>">
多版主添加请用“|”隔开，如：yuzi|裕裕
</td>
</tr>
<tr class=a4>
<td align="right" valign="middle" width="20%">帖子专题：</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="TolSpecialTopic" value="<%=Rs("TolSpecialTopic")%>"> 
添加请用“|”隔开，如：原创|转贴|贴图</td>
</tr>
<tr class=a3>
<td align="right" width="20%">论坛介绍：</td>
<td align="Left" valign="middle" width="78%">
<textarea name="ForumIntro" rows="4" cols="50"><%=ForumIntro%></textarea>&nbsp;
</td>
<tr class=a4>
<td align="right" width="20%">论坛规则：</td>
<td align="Left" valign="middle" width="78%">
<textarea name="ForumRules" rows="4" cols="50"><%=ForumRules%></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td align="right" valign="middle" width="20%">小图标URL：</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="ForumIcon" value="<%=Rs("ForumIcon")%>">　显示在社区首页论坛介绍右边</td>
</tr>
<tr class=a4>
<td align="right" valign="bottom" width="20%">大图标URL：</td>
<td align="Left" valign="bottom" width="78%">
<input size="30" name="ForumLogo" value="<%=Rs("ForumLogo")%>">　显示在论坛左上角</td>
</tr>
<tr class=a3>
<td align="right" valign="bottom" width="98%" colspan="2"><input type="submit" value=" 更 新 &gt;&gt;下 一 步 "></td>
</tr>
</table>
</form>
		
		
		
		
		</td>
	</tr>
</table>

<%
Rs.close
htmlend


end select

%>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a> → 
			<a href="?menu=<%=Request("menu")%>&ForumID=<%=ForumID%>&checkbox=1"><span id=menu>回收站</span></a></td>
		</tr>
	</table><br><form method="POST" action="?"><input type=hidden name=ForumID value=<%=ForumID%>>
<%

if Request("menu")="Recycle" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime<>lasttime and ForumID="&ForumID&" order by lasttime Desc"
response.write "<script>menu.innerText='回收站'</script>"
elseif Request("menu")="Censorship" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime=lasttime and ForumID="&ForumID&" order by lasttime Desc"
response.write "<script>menu.innerText='审查区'</script>"
end if
Rs.Open sql,Conn,1

PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数


%>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">主题</td>
<td align="center" width="10%">作者</td>
<td align="center" width="6%">回复</td>
<td align="center" width="6%">点击</td>
<td align="center" width="25%">最后更新</td>
</tr>
		<%

i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

ShowThread()
Rs.MoveNext
loop
Rs.Close
%>
	
</table>

<table cellspacing="0" cellpadding="1" width="100%" align="center" border="0">
	<tr height="25">
		<td width="100%" height="2">
		<table cellspacing="0" cellpadding="3" width="100%" border="0">
			<tr>
				<td height="2" valign="top"><input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">全选
<%
if Request("menu")="Recycle" then
%><input type="radio" value="BatchRecycle" name=menu checked>还原主题<%
elseif Request("menu")="Censorship" then
%><input type="radio" value="BatchCensorship" name=menu checked>通过审查 <input type="radio" value="BatchDel" name=menu><input type=hidden name=IsDel value=1>删除主题<%
end if
%>

<input onclick="checkclick('您确定执行本次操作?');" type="submit" value=" 执 行 ">




</form></td>

				<td align="right" height="2" valign="top">

<%ShowPage()%></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<%


htmlend
%>