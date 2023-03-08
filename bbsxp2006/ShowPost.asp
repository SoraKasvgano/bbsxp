<!-- #include file="Setup.asp" -->
<%
top

ThreadID=int(Request("ThreadID"))
ForumID=int(Request("ForumID"))

sql="select UserName From [BBSXP_UsersOnline] where Eremite=0"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
OnlineUserList=OnlineUserList&""&Rs("UserName")&"|"
Rs.MoveNext
loop
Rs.close

if Request("action")="Next" then
sql="select top 1 * from [BBSXP_Threads] where ID > "&ThreadID&" and ForumID="&ForumID&" and IsDel=0 order by id"
elseif Request("action")="Previous" then
sql="select top 1 * from [BBSXP_Threads] where ID < "&ThreadID&" and ForumID="&ForumID&" and IsDel=0 order by id Desc"
else
sql="select * from [BBSXP_Threads] where ID="&ThreadID&""
end if
Rs.Open SQL,Conn,1,3
if Rs.eof or Rs.bof then error"<li>系统不存在该帖子的资料"
if Rs("IsDel")=1 and membercode<4 then error"<li>该主题在回收站中！"
Rs("Views")=Rs("Views")+1
Rs.update
Topic=ReplaceText(Rs("Topic"),"<[^>]*>","")
Replies=Rs("Replies")
Views=Rs("Views")
IsVote=Rs("IsVote")
IsGood=Rs("IsGood")
IsTop=Rs("IsTop")
IsLocked=Rs("IsLocked")
PostsTableName=Rs("PostsTableName")
ThreadID=Rs("ID")
ForumID=Rs("ForumID")
Rs.close


sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
moderated=Rs("moderated")
ForumLogo=Rs("ForumLogo")
followid=Rs("followid")
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
Rs.Close
%>
<!-- #include file="inc/Validate.asp" -->
<script src="inc/birth.js"></script>
<title><%=Topic%> - Powered By BBSXP</title>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="?ThreadID=<%=ThreadID%>"><%=Topic%></a></td>
</tr>
</table><br>

<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0"><tr><td height="35" valign="bottom">
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href=NewTopic.asp?ForumID=<%=ForumID%>>发表新主题</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href=ReTopic.asp?ThreadID=<%=ThreadID%>>回复帖子</a>
</td><td align="right" height="35" valign="bottom"><font color="333333">您是本帖第 <b><%=Views%></b> 个阅读者</font>　　<a href="?action=Previous&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img height="12" alt="浏览上一篇主题" src="images/prethread.gif" width="52" border="0"></a>　<a style="text-decoration: none" href="javascript:location.reload()"><img height="12" alt="刷新本主题" src="images/refresh.gif" width="40" border="0"></a>　<a href="?action=Next&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img height="12" alt="浏览下一篇主题" src="images/nextthread.gif" width="52" border="0"></a></font></td></tr></table>

<table width="100%" border="0" cellspacing="1" class="a2" height="21" align="center"><tr class="a1">
	<td width="100%" height="9" colspan="4">
	<table border="0" width="100%" cellspacing="0">
		<tr class="a1">
			<td><b>&nbsp;主题</b>：<%=Topic%></td>
				<td align="right"> 
<a target=_blank href=Print.asp?ThreadID=<%=ThreadID%>><img alt="适合打印机打印的版本" src=images/Print.gif border=0></a>&nbsp; 
<script>document.write("<a target=_blank href='Mailto:?subject=<%=Topic%>&body="+location.href+"'>");</script><img alt="通过电子邮件发送此页面" src=images/sendMail.gif border=0></a>&nbsp; 
<a href="javascript:window.external.AddFavorite(location.href,document.title)"><img alt="添加加到IE收藏夹" src="images/favs.gif" border="0"></a>&nbsp; 
<script>document.write("<a style=cursor:hand onclick=\"javascript:open('Message.asp?menu=Post&report=1&moderated=<%=moderated%>&body=【问题帖子】："+location.href+"','','width=320,height=170')\">");</script><img alt="报告本帖" src="images/feedback.gif" border="0"></a>&nbsp; 
</td></tr></table>
</td></tr>
<%
'''''''投票''''''''
if IsVote=1 then
%>
<tr class="a4">
	<td width="40%" height="25" align="center">
	选项</td>
	<td width="10%" height="25" align="center">
	票数</td>
	<td width="50%" height="25" align="center" colspan="2">
	百分比</td></tr>

<form action=PostVote.asp?id=<%=ThreadID%> method=Post>
<%

sql="select * from [BBSXP_Vote] where ThreadID="&ThreadID&""
Set Rs=Conn.Execute(sql)

if Rs("Type")=1 then
multiplicity="checkbox"
else
multiplicity="radio"
end if

allticket=0
Result=split(Rs("Result"),"|")
for i = 0 to ubound(Result)
if not Result(i)="" then allticket=Result(i)+allticket
next

Vote=split(Rs("Items"),"|")
for i = 0 to ubound(Vote)
if not Vote(i)="" then

if allticket=0 then
Voteresult=0
Votepercent=0
else
Voteresult=result(i)/allticket*100
Votepercent=formatnumber(result(i)/allticket*100)
end if

%>
	<tr class="a3">
	<td width="45%" height="20"><input type=<%=multiplicity%> value=<%=i%> name=PostVote id=PostVote<%=i%>><label for=PostVote<%=i%>><%=Vote(i)%></label></td>
	<td width="5%" height="20" align="center"><%=Result(i)%></td>
	<td width="42%" height="20"><img src=images/bar/0.gif width=<%=Voteresult%>% height=10></td>
	<td width="8%" height="20" align="center"><%=Votepercent%>%</td>
	</tr>
<%
end if
next
%>
	<tr class="a3">
	<td height="25" align="center">
<%
if Rs("Expiry")< now() then
response.write "投票已过期"
elseif CookieUserName=empty then
response.write "登录后才能投票"
elseif instr(Rs("BallotUserList"),""&CookieUserName&"|")>0 then
response.write "您已经投过票了"
else
response.write "<INPUT type=submit value='投　票'>"
end if
%>
	</td>
	<td height="25" align="center">
总票数：<%=allticket%></td></form>
	<td colspan="2" align="center">截止投票时间：<%=Rs("Expiry")%></td>
	</tr>
	</table>
<%
Rs.Close
end if
'''''''投票 END''''''''


TotalCount=replies+1
PageSetup=SiteSettings("PostsPerPage") '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = cint(Request.QueryString("PageIndex")) '获取当前页
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


if PageCount<2 then
sql="select top "&pagesetup&" * from [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&" order by PostTime"
Set Rs=Conn.Execute(sql)
else
sql="select * from [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&" order by PostTime"
rs.Open sql,Conn,1
end if


if TotalPage>1 then RS.Move (PageCount-1) * pagesetup

i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

Set Rs1=Conn.Execute("[BBSXP_Users] where UserName='"&Rs("UserName")&"'")
if Rs1.EOF then
Conn.execute("Delete from [BBSXP_Posts"&PostsTableName&"] where UserName='"&Rs("UserName")&"'")
TotalCount=conn.Execute("Select count(ID) From [BBSXP_Posts"&PostsTableName&"] where ThreadID="&ThreadID&"")(0)
Conn.execute("update [BBSXP_Threads] set replies="&TotalCount&"-1 where id="&ThreadID&"")
end if
ShowRank()
%>

<table class=a2 cellPadding=5 width=100% align=center border=0 cellSpacing=1 style=TABLE-LAYOUT:fixed>
<tr class=a3><td width=156 align=center valign=top height=100%>



<table border=0 width=90%><tr><td><font style=font-size:10pt><b><%=Rs1("UserName")%></b></font><br><%=Rs1("UserHonor")%></td><td align=right valign=top>
<%if Rs1("UserSex")<>"" then%>
<img src=images/<%=Rs1("UserSex")%>.gif>&nbsp;
<%end if%>
<script>document.write(astro("<%=Rs1("birthday")%>"));</script>
</td></tr></table>
<%if Request.Cookies("DisabledShowFace")="" then%>
<img src='<%=Rs1("Userface")%>' onload='javascript:if(this.width>120)this.width=120;if(this.height>120)this.height=120;'>
<%end if%>
<br><br><img src=<%=RankIconUrl%>><table border=0 width=90%><tr><td><br>等　　级:<%=RankName%><br>
<%if Rs1("Consortia")<>"" then%>公　　会:<a target="_blank" href="Consortia.asp?menu=look&ConsortiaName=<%=Rs1("Consortia")%>"><%=Rs1("Consortia")%></a><br><%end if%>
<%if Rs1("Consort")<>"" then%>配　　偶:<%=Rs1("Consort")%><br><%end if%>
经 验 值:<%=Rs1("experience")%><br>
社区金币:<%=Rs1("UserMoney")%><br>
总发贴数:<%=Rs1("PostTopic")+Rs1("Postrevert")%><br>
注册时间:<%=split(Rs1("UserRegTime")," ")(0)%><br>
状　　态:<%if instr("|"&OnlineUserList&"","|"&Rs1("UserName")&"|")>0 then%>在线<%else%>离线<%end if%>
</td></tr></table>



</td><td><table cellSpacing=0 cellPadding=0 border=0 width=100% height=100%><tr>
		<td colspan=2> 
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/contact.gif) href=Profile.asp?UserName=<%=Rs("UserName")%> title='查看<%=Rs("UserName")%>的个人资料'>信息</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/message.gif) style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','','width=320,height=170')" title='发送短讯息给<%=Rs("UserName")%>'>短讯</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/mail.gif) href=Mailto:<%=Rs1("UserMail")%> title='发送电邮给 <%=Rs("UserName")%>'>邮箱</a>
<%if Rs1("userhome")<>"" and Rs1("userhome")<>"http://" then%>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/home.gif) target=_blank href='<%=Rs1("userhome")%>' title='访问 <%=Rs1("userhome")%> 的主页'>主页</a>
<%end if%>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/Friend.gif) href='Friend.asp?menu=add&UserName=<%=Rs("UserName")%>' title='把 <%=Rs("UserName")%> 加入好友'>好友</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/finds.gif) href='ShowBBS.asp?menu=5&UserName=<%=Rs("UserName")%>' title='搜索<%=Rs("UserName")%>发表过的所有主题'>搜索</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/Quote.gif) href='ReTopic.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>&quote=1' title='引用回复这个帖子'>引用</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href='ReTopic.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='回复帖子'>回复</a>
</td><td align=right>No.<font color=red><b><%=i+(PageCount-1)*PageSetup%></b></font></td></tr><tr vAlign=top><td colSpan=3><hr width=100% color=#777777 SIZE=1></td></tr>
<tr vAlign=top><td colSpan=3 height=100% style=word-break:break-all>

<%if instr("|"&SiteSettings("BannedUserPost")&"|"&Request.Cookies("BadUserList")&"|","|"&Rs("UserName")&"|")>0 then%>
==============================<br>　　　<font color=RED>该用户帖子已被过滤　　　</font><br>==============================
<%else%>
<b><%=Rs("Subject")%></b><br><br>
<%=Rs("content")%>
<%end if%></td></tr><tr vAlign=top><td colSpan=3 align=right>
<%if Rs1("UserSign")<>"" and Request.Cookies("DisabledShowSign")="" then%>
——————————<br><%=YbbEncode(Rs1("UserSign"))%>
<%end if%>

</td></tr><tr vAlign=top><td colSpan=3><hr width=100% color=#777777 SIZE=1></td></tr><tr vAlign=top><td><a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/edit.gif) href='EditPost.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='编辑帖子'>编辑</a> <a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/del.gif) href='Manage.asp?menu=IsDel&ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='删除帖子'>删除</a></td><td valign=bottom><img src=images/Posttime.gif> 发表时间：<%=Rs("Posttime")%>　</td>
			<td valign=bottom align="right"><img src=images/ip.gif> IP：
<%if SiteSettings("DisplayPostIP")=1 then%>
<%=Rs("PostIP")%>
<%else%>
<a href=Manage.asp?menu=lookip&ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>>已记录</a>
<%end if%>
</td></tr></table></td></tr>
</table>

<%
Set Rs1 = Nothing
Rs.MoveNext
loop
Rs.Close
act=topic
%>


<table cellspacing="1" cellpadding="0" width="100%" border="0" align=center><tr><td width="61%">
<%ShowPage()%>
</td><td width="39%" align="right">
<a href="MyFavorites.asp?menu=add&url=Topic&name=<%=ThreadID%>">收藏帖子</a> | <a href="MyFavorites.asp?menu=Del&url=Topic&name=<%=ThreadID%>">取消收藏</a> | <a href="#">返回页首</a>&nbsp;</td></tr></table>



<%if CookieUserName<>empty then%>
<form name="yuziform" method="POST" action="ReTopic.asp" onSubmit="return CheckForm(this);">
<input type="hidden" value="<%=ThreadID%>" name="ThreadID">
<input type="hidden" value="Re:<%=Topic%>" name="Subject">
<input name="content" type="hidden">

<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align=center>
<tr class="a1"><td width="20%" height="10"><b>快速回复主题</b></td><td width="80%"><b><%=act%></b></td></tr>

<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
<tr class="a3"><td width="20%" height="3"><b>验证码</b></td><td width="80%">
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="验证码,看不清楚?请点击刷新验证码" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></td></tr>
<%end if%>

<tr class="a3">
	<td width="20%" rowspan="2" valign="top" height="100%">
<TABLE cellSpacing=0 cellPadding=0 width=100% align=Left border=0 height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><br><B>文章内容</B><BR>
（<a href="javascript:CheckLength();">查看内容长度</a>）
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> 禁用YBB代码</label>
</TD></TR>
</TABLE>

	<td width="80%" height=200>

<SCRIPT src="inc/Post.js"></SCRIPT></td></tr><tr class=a3>
			<td><input type="submit" value="Ctrl+Enter 回复主题" name=EditSubmit>　　　<input type="reset" name="reset" value=" 重 置 "></td></tr></table>

</form>

<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0"><tr><td align="middle">管理选项: <%
response.write "<a href=Manage.asp?menu=MoveNew&ThreadID="&ThreadID&">拉前主题</a> | "
if IsLocked=1 then
response.write "<a href=Manage.asp?menu=DelIsLocked&ThreadID="&ThreadID&">解锁主题</a>"
else
response.write "<a href=Manage.asp?menu=IsLocked&ThreadID="&ThreadID&">锁定主题</a>"
end if
response.write " | "
if IsTop=2 then
response.write "<a href=Manage.asp?menu=untop&ThreadID="&ThreadID&">取消总固顶</a>"
else
response.write "<a href=Manage.asp?menu=top&ThreadID="&ThreadID&">主题总固顶</a>"
end if
response.write " | "
if IsTop=1 then
response.write "<a href=Manage.asp?menu=DelIsTop&ThreadID="&ThreadID&">取消固顶</a>"
else
response.write "<a href=Manage.asp?menu=IsTop&ThreadID="&ThreadID&">主题固顶</a>"
end if
response.write " | "
if IsGood=1 then
response.write "<a href=Manage.asp?menu=DelIsGood&ThreadID="&ThreadID&">取消精华帖</a>"
else
response.write "<a href=Manage.asp?menu=IsGood&ThreadID="&ThreadID&">加为精华帖</a>"
end if
%>
| <a href="Move.asp?ThreadID=<%=ThreadID%>">移动主题</a>
| <a title="修复帖子的回复数" href="Manage.asp?menu=Fix&ThreadID=<%=ThreadID%>">修复主题</a>

</td></tr></table><%end if%> 

<!-- #include file="inc/line.asp" -->
<%htmlend%>