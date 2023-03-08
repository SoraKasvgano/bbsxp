<!-- #include file="Setup.asp" --><%
id=cint(Request("id"))
if id<>empty then Response.Cookies("ForumID")=id
if ""&SiteSettings("AdminPassword")&""=empty then response.redirect "Install.asp"
top
Set Statistics=Conn.Execute("[BBSXP_Statistics_Site]")
%>
<table border="0" width="100%" align="center">
	<tr>
		<td valign="top">
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a2">
			<%
if CookieUserName<>"" then

sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs1=Conn.Execute(sql)
ShowRank()

%>
			<tr class=a3>
				<td width="18%" align="center" valign="bottom" height="100">
				<img src="<%=Rs1("Userface")%>" onload='javascript:if(this.width>60)this.width=60;if(this.height>60)this.height=60;'><br>
				<br>
				<a href="UserCp.asp">控制面板</a></td>
				<td align="center" valign="bottom" height="100%">
				<table width="100%" border="0" cellpadding="3" cellspacing="1">
					<tr>
						<td width="25%" align="Left" valign="middle">
						<a href="Profile.asp?UserName=<%=CookieUserName%>">我的资料</a></td>
						<td width="26%" align="Left" valign="middle">
						<a href="UpFace.asp">上传头像</a></td>
						<td width="28%" align="Left" valign="middle">
						<a href="UpPhoto.asp">上传照片</a></td>
					</tr>
					<tr>
						<td align="Left" valign="middle">
						<span style="text-decoration: none">
						<a href="Calendar.asp?UserName=<%=CookieUserName%>">我的日志</a></span></td>
						<td align="Left" valign="middle">
						<a href="ShowBBS.asp?menu=5&UserName=<%=CookieUserName%>">我的帖子</a></td>
						<td align="Left" valign="middle">
						<a href="MyFavorites.asp">我的收藏</a></td>
					</tr>
					<tr>
						<td align="Left" valign="middle">等级：<%=RankName%></td>
						<td align="Left" valign="middle">社区金币：<font face="Georgia, Times New Roman, Times, serif"><%=Rs1("UserMoney")%></font></td>
						<td align="Left" valign="middle">发表文章：<font face="Georgia, Times New Roman, Times, serif"><%=Rs1("PostTopic")%></font></td>
					</tr>
					<tr>
						<td align="Left" valign="middle">经验：<font face="Georgia, Times New Roman, Times, serif"><%=Rs1("experience")%></font></td>
						<td align="Left" valign="middle">社区存款：<font face="Georgia, Times New Roman, Times, serif"><%=Rs1("savemoney")%></font></td>
						<td align="Left" valign="middle">回复文章：<font face="Georgia, Times New Roman, Times, serif"><%=Rs1("Postrevert")%></font></td>
					</tr>
				</table>
				</td>
			</tr>
			<%
Set Rs1 = Nothing
else
%>
			<form action="Login.asp" method="POST">
				<input type="hidden" value="Default.asp" name="url">
				<tr class=a4>
					<td width="18%" align="center" height="100">
					<embed src="images/Clock.swf" width="100" height="80" wmode="transparent"></td>
					<td align="center" height="50">
					<table border="0" width="100%">
						<tr>
							<td rowspan="3">用户名称:
							<input size="20" name="UserName" value="<%=CookieUserName%>"><br>用户密码:
							<input type="password" size="20" value name="Userpass">
<%if sitesettings("EnableAntiSpamTextGenerateForLogin")=1 then%>
							<br>验 证 码: <input size="10" name="VerifyCode">&nbsp;&nbsp;<img src="VerifyCode.asp" alt="验证码,看不清楚?请点击刷新验证码" style=cursor:pointer onclick="this.src='VerifyCode.asp'">
<%end if%>							
							　</td>
							<td>
							<input type="checkbox" value="1" name="eremite" id="eremite"><label for="eremite">隐身登录</label>
							</td>
							<td><a href="CreateUser.asp">没有注册?</a></td>
						</tr>
						<tr>
							<td>
							<input type="checkbox" value="1" name="xuansave" id="xuansave"><label for="xuansave">记住密码</label>
							</td>
							<td><a href="RecoverPassword.asp">忘记密码?</a></td>
						</tr>
						<tr>
							<td colspan="2">
							<input type="submit" value=" 登录 ">&nbsp;
							<input type="reset" value=" 取消 "></td>
						</tr>
					</table>
					</td>
				</tr>
			</form>
			<%
end if


%>
			<tr class="a3">
				<td colspan="2" align="center" valign="bottom" height="25">
				注册会员：<%=Statistics("TotalUser")%>&nbsp;
				主题总数：<%=Statistics("TotalThread")%>&nbsp;
				帖子总数：<%=Statistics("TotalPost")%>&nbsp;
				今日帖数：<%=Statistics("TodayPost")%>&nbsp;
				新会员：<a href="Profile.asp?UserName=<%=Statistics("NewUser")%>"><%=Statistics("NewUser")%></a>
</td>
			</tr>
		</table>
		</td>
		<td width="223" valign="top">
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a2">
			<tr class="a1" id="TableTitleLink">
				<td height="9" align="Left" valign="middle">&nbsp;<a href="Affiche.asp">社区公告</a></td>
			</tr>
			<tr class="a3">
				<td height="100" align="Left" valign="top">
				<table border="0" width="100%">
					<%
Set Rs=Conn.ExeCute("Select top 5 * From [BBSXP_Affiche] Order By ID Desc")
Do While Not Rs.Eof
title=Rs("title")
if len(title)>17 then title=Left(""&title&"",14)&"..."
%>
					<tr>
						<td><a href="Affiche.asp?ID=<%=Rs("ID")%>"><%=title%></a></td>
					</tr>
					<%
	Rs.MoveNext
Loop
Rs.Close
%>
				</table>
				</td>
			</tr>
		</table>
		<br>
		</td>
	</tr>
</table>


<table width="100%" align="center" border="0" class="a2" cellspacing=1 cellpadding=0>
	<%
if Request.Cookies("forumid")<>empty then
sql="select * from [BBSXP_Forums] where id="&int(Request.Cookies("forumid"))&""
Set Rs1=Conn.Execute(sql)
if Rs1.eof then Response.Cookies("forumid")=""
ShowForum()
Set Rs1 = Nothing
else

sql="Select * From [BBSXP_Forums] where followid=0 and ForumHide=0 order by SortNum"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
ForumIntro=replace(Rs("ForumIntro"),"<br>",CHR(10))
%>
<tr class="a1" id="TableTitleLink">
<td colspan="7" height="25" title="<%=ForumIntro%>">　<b><a href="ShowForum.asp?ForumID=<%=Rs("id")%>"><%=Rs("ForumName")%></a></b></td></tr>
<%
i=0
sql="Select * From [BBSXP_Forums] where followid="&Rs("id")&" and ForumHide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)

do while not Rs1.eof
if SiteSettings("DisplayForumFloor")=1 then
ShowForumFloor
else
ShowForum
end if

Rs1.Movenext
loop
Set Rs1 = Nothing

if SiteSettings("DisplayForumFloor")=1 and i>0 then
for i = i to 2
response.write "<td class=a3></td>"
next
end if

Rs.Movenext
loop
Rs.close



end if

ForumName=SiteSettings("SiteName")
ForumID=0
%>
</table>
<br>
<!-- #include file="inc/line.asp" --><%


regOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where UserName<>''")(0)

if Statistics("BestOnline") < Onlinemany then
Conn.execute("update [BBSXP_Statistics_Site] set BestOnline="&Onlinemany&",BestOnlineTime="&SqlNowString&"")
end if
if SiteSettings("DisplayWhoIsOnline")=1 then
%>
<table cellspacing="1" cellpadding="0" width="100%" align="center" border="0" class="a2">
	<tr>
		<td height="25" class="a1" colspan="2">　<b>■ </b>在线统计</td>
	</tr>
	<tr>
		<td align="middle" width="5%" class=a3>
		<img src="images/totel.gif"></td>
		<td valign="top" class="a4">
		<table cellspacing="0" cellpadding="3" width="100%">
			<tr>
				<td height="15">&nbsp;<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=ForumID%>)"> 
				目前论坛总共 有 <b><%=Onlinemany%></b> 人在线。其中注册用户 <b><%=regOnline%></b> 
				人，访客 <b><%=Onlinemany-regOnline%></b> 人。最高在线 <font color="red">
				<b><%=Statistics("BestOnline")%></b></font> 人，发生在 <b><%=Statistics("BestOnlineTime")%></b>
				</td>
			</tr>
			<tr height="25" style="display:none" id="follow0">
				<td id="followTd0" align="Left" class="a4" width="94%" colspan="5">
				　Loading.....</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<br>
<%end if
if SiteSettings("DisplayLink")=1 then
%>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center">
	<tr>
		<td height="25" colspan="2" class="a1">　■<b> </b>友情链接</td>
	</tr>
	<tr>
		<td align="center" class=a3 width="5%" rowspan="2">
		<img src="images/shareforum.gif"></td>
		<td class="a4"><%
Rs.Open "[BBSXP_Link]",Conn
do while not Rs.eof
if Rs("Logo")="" or Rs("Logo")="http://" then
Link1=Link1+"<a title='"&Rs("Intro")&"' href="&Rs("url")&" target=_blank>"&Rs("name")&"</a> "
else
Link2=Link2+"<a title='"&Rs("name")&""&chr(10)&""&Rs("Intro")&"' href="&Rs("url")&" target=_blank><img src="&Rs("Logo")&" border=0 width=88 height=31></a> "
end if
Rs.Movenext
loop
Rs.close
%> <%=Link1%> </td>
	</tr>
	<tr>
		<td class="a4"><%=Link2%></td>
	</tr>
</table>
<br>
<%end if%>
<center>&nbsp;<img src="images/skins/<%=Request.Cookies("skins")%>/Board0.gif" alt="禁止浏览"> 关闭论坛&nbsp; 
<img src="images/skins/<%=Request.Cookies("skins")%>/Board1.gif" alt="任何人均可浏览"> 正规论坛&nbsp; 
<img src="images/skins/<%=Request.Cookies("skins")%>/Board2.gif" alt="游客禁止浏览"> 会员论坛&nbsp; 
<img src="images/skins/<%=Request.Cookies("skins")%>/Board3.gif" alt="需要授权才能浏览"> 特殊论坛</center>
<%
Set Statistics=Nothing


htmlend

sub ShowForumFloor()
ForumIntro=replace(Rs1("ForumIntro"),"<br>",CHR(10))
if rs1("Moderated")<>empty then
ModeratedList="版主："
filtrate=split(rs1("Moderated"),"|")
ModeratedList=ModeratedList&"<a href=Profile.asp?UserName="&filtrate(0)&">"&filtrate(0)&"</a> "
if ubound(filtrate)>0 then ModeratedList=ModeratedList&" <font color=gray>...</font>"
else
ModeratedList="　"
end if
i=i+1
if i=1 then response.write "<tr>"
%>
<td><table border=0 width=100% cellspacing=0 cellpadding=4><tr class=a3><td colspan=3 title='<%=ForumIntro%>'>『 <a href="ShowForum.asp?ForumID=<%=Rs1("id")%>"><%=Rs1("ForumName")%></a> 』</td></tr><tr class=a3><td><img src=images/Forum_nav.gif alt=今日> <font color=red><%=Rs1("ForumToday")%></font></td><td><img src=images/Forum_nav.gif alt=主题> <%=Rs1("ForumThreads")%></td><td><img src=images/Forum_nav.gif alt=帖子> <%=Rs1("ForumPosts")%></td></tr><tr class=a4><td colspan=3><%=ModeratedList%></td></tr></table></td>
<%
if i=3 then i=0:response.write "</tr>"

end sub
%>