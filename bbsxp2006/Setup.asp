<!-- #include file="Conn.asp" -->
<%
SiteSettings=Conn.Execute("[BBSXP_SiteSettings]")
CookieUserName=HTMLEncode(unescape(Request.Cookies("UserName")))

if CookieUserName<>empty then
''''''''''''''''''第一次来'''''''''''''''''''''''''''''
if Request.Cookies("Onlinetime")=empty then
Conn.execute("update [BBSXP_Users] set UserDegree=UserDegree+1,UserLandTime="&SqlNowString&",UserLastIP='"&Request.ServerVariables("REMOTE_ADDR")&"' where UserName='"&CookieUserName&"'")
Response.Cookies("Onlinetime")=now()
end if
''''''''''''''''''''''''''''
sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(SQL)
if Rs.eof then Response.Cookies("UserName")=""
if Request.Cookies("Userpass") <> Rs("Userpass") then Response.Cookies("UserName")=""
membercode=Rs("membercode")
Userface=""&Rs("Userface")&""
NewMessage=Rs("NewMessage")
set rs=nothing
end if

if SiteSettings("BannedIP")<>"" then
filtrate=split(SiteSettings("BannedIP"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then response.redirect "inc/BannedIP.htm"
next
end if

if Request.Cookies("skins")=empty then Response.Cookies("skins")=SiteSettings("DefaultSiteStyle")
if ""&SiteSettings("nowdate")&""<>""&date()&"" then
Conn.execute("update [BBSXP_SiteSettings] set Nowdate='"&date()&"'")
Conn.execute("update [BBSXP_Statistics_Site] set TodayPost=0")
Conn.execute("update [BBSXP_Forums] set ForumToday=0")
end if

dim toptrue,ForumsList,ForumTreeList,TotalPage,PageCount,RankName,RankIconUrl
ii=0
startime=timer()
Set rs = Server.CreateObject("ADODB.Recordset")
Server.ScriptTimeout=SiteSettings("Timeout")'设置脚本超时时间 单位:秒

response.write "<html><head><meta http-equiv=content-Type content=text/html;charset=gb2312><meta name=keywords content='"&SiteSettings("MetaKeywords")&"'><meta name=description content='"&SiteSettings("MetaDescription")&"'></head><script src=inc/BBSXP.js></script><script src=images/skins/"&Request.Cookies("skins")&"/bbs.js></script><Link href=images/skins/"&Request.Cookies("skins")&"/bbs.css rel=stylesheet>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function HTMLEncode(fString)
fString=replace(fString,";","&#59;")
fString=replace(fString,"<","&lt;")
fString=replace(fString,">","&gt;")
fString=replace(fString,"\","&#92;")
fString=replace(fString,"--","&#45;&#45;")
fString=replace(fString,CHR(9),"&#9;")
fString=replace(fString,CHR(10),"<br>")
fString=replace(fString,CHR(13),"")
fString=replace(fString,CHR(22),"&#22;")
fString=replace(fString,CHR(32),"&#32;")
fString=replace(fString,CHR(34),"&#34;")'双引号
fString=replace(fString,CHR(39),"&#39;")'单引号
HTMLEncode=fString
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function ContentEncode(fString)
fString=replace(fString,vbCrlf, "")
fString=replace(fString,"\","&#92;")
fString=replace(fString,"'","&#39;")
fString=replace(fString,"<A href=","<A target=_blank href=") '点链接打开新窗口
if SiteSettings("BannedHtmlLabel")<>"" then fString=ReplaceText(fString,"<(\/|)("&SiteSettings("BannedHtmlLabel")&")", "&lt;$1$2")
if SiteSettings("BannedHtmlEvent")<>"" then fString=ReplaceText(fString,"<(.[^>]*)("&SiteSettings("BannedHtmlEvent")&")", "&lt;$1$2")
if SiteSettings("BannedText")<>"" then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
fString=ReplaceText(fString,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if
contentEncode=fString
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function YbbEncode(str)
str=ReplaceText(str,"\[(\/|)(b|i|u|strike|center|marquee)\]","<$1$2>")
str=ReplaceText(str,"\[COLOR=([^[]*)\]","<FONT COLOR=$1>")
str=ReplaceText(str,"\[FONT=([^[]*)\]","<FONT face=$1>")
str=ReplaceText(str,"\[SIZE=([0-9]*)\]","<FONT size=$1>")
str=ReplaceText(str,"\[\/(SIZE|FONT|COLOR)\]","</FONT>")
str=ReplaceText(str,"\[QUOTE\]","<BLOCKQUOTE><strong>引用</strong>：<HR Size=1>")
str=ReplaceText(str,"\[\/QUOTE\]","<HR SIZE=1></BLOCKQUOTE>")
str=ReplaceText(str,"\[URL\]([^[]*)","<A TARGET=_blank HREF=$1>$1")
str=ReplaceText(str,"\[URL=([^[]*)\]","<A TARGET=_blank HREF=$1>")
str=ReplaceText(str,"\[\/URL\]","</A>")
str=ReplaceText(str,"\[EMAIL\](\S+\@[^[]*)(\[\/EMAIL\])","<a href=mailto:$1>$1</a>")
str=ReplaceText(str,"\[IMG\]([^[]*)(\[\/IMG\])","<img border=0 src=$1>")
YbbEncode=str
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
On Error GoTo 0
End Function

Function CheckSize(ByteSize)
if ByteSize=>1024000 then
ByteSize=formatnumber(ByteSize/1024000)&" MB"
elseif ByteSize=>1024 then
ByteSize=formatnumber(ByteSize/1024)&" KB"
else
ByteSize=ByteSize&" 字节"
end if
CheckSize=ByteSize
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sub top
%>
<script>BBSxpTop()</script>
<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2><tr class=a1><td><table cellspacing=0 cellpadding=4 width=100% border=0><tr class=a1><td id=TableTitleLink>&gt;&gt;欢迎您， <%
if CookieUserName=empty then
%>  <a href="Login.asp">请先登录</a> |  <a href="CreateUser.asp">
	注册</a> |  <a href="RecoverPassword.asp">
	忘记密码</a> |  <a href="Online.asp">
	在线情况</a> | <a href="Search.asp?ForumID=<%=Request("ForumID")%>">
	搜索</a> | <a href="Help.asp">
	帮助</a> </td></tr></table></td></tr></table><br>
<%else%>
<%=CookieUserName%>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=Cookies.asp?menu=Online>上线</a></div><div class=menuitems><a href=Cookies.asp?menu=eremite>隐身</a></div><div class=menuitems><a href=Login.asp?menu=out>退出</a></div>')" style=cursor:default>
<%
if Request.Cookies("eremite")="1" then
response.write "隐身"
else
response.write "上线"
end if
%></A>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=Profile.asp?UserName=<%=CookieUserName%>>我的资料</a></div><div class=menuitems><a href=UpFace.asp>上传头像</a></div><div class=menuitems><a href=UpPhoto.asp>上传照片</a></div><div class=menuitems><a href=Calendar.asp>社区日志</a></div><div class=menuitems><a href=Blog.asp?UserName=<%=CookieUserName%>>我的日志</a></div><div class=menuitems><a href=Login.asp>更换用户</a></div>')" style=cursor:default>个人服务</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=UserCp.asp>控制面板</a></div><div class=menuitems><a href=EditProfile.asp>资料修改</a></div><div class=menuitems><a href=EditProfile.asp?menu=pass>密码修改</a></div><div class=menuitems><a href=MySettings.asp>个性设置</a></div><div class=menuitems><a href=MyAttachment.asp>附件管理</a></div><div class=menuitems><a href=Message.asp>短信服务</a></div><div class=menuitems><a href=Friend.asp>好友列表</a></div>')" style=cursor:default>控制面板</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=ShowBBS.asp>最新帖子</a></div><div class=menuitems><a href=ShowBBS.asp?menu=1>人气帖子</a></div><div class=menuitems><a href=ShowBBS.asp?menu=2>热门帖子</a></div><div class=menuitems><a href=ShowBBS.asp?menu=3>精华帖子</a></div><div class=menuitems><a href=ShowBBS.asp?menu=4>投票帖子</a></div><div class=menuitems><a href=ShowBBS.asp?menu=5&UserName=<%=CookieUserName%>>我的帖子</a></div>')" style=cursor:default>帖子状态</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=Online.asp>在线情况</a></div><div class=menuitems><a href=Online.asp?menu=cutline>在线图例</a></div><div class=menuitems><a href=Online.asp?menu=UserSex>性别图例</a></div><div class=menuitems><a href=Online.asp?menu=TodayPage>今日图例</a></div><div class=menuitems><a href=Online.asp?menu=board>主题图例</a></div><div class=menuitems><a href=Online.asp?menu=ForumPosts>帖子图例</a></div><div class=menuitems><a href=UserTop.asp>会员列表</a></div><div class=menuitems><a href=Adminlist.asp>管理团队</a></div><div class=menuitems><a href=ApplyForum.asp>申请论坛</a></div>')" style=cursor:default>论坛状态</a>
 | <a href=Search.asp>搜索</a>
 | <a onmouseover="showmenu(event,'<div class=menuitems><a href=MyFavorites.asp?menu=Topic>帖子收藏夹</a></div><div class=menuitems><a href=MyFavorites.asp?menu=Forum>论坛收藏夹</a></div><div class=menuitems><a href=MyFavorites.asp>网站收藏夹</a></div>')" style=cursor:default>收藏</a>

<%
menu(0)

response.write " | <a href=Help.asp>帮助</a>"

if membercode=5 then response.write " | <a href=Admin.asp target=_top>管理</a>"

%></td><td align="right">已经停留 <b><%=DateDiff("n",Request.Cookies("Onlinetime"),Now())%></b> 分钟&nbsp; </td></tr></table></td></tr></table>
<%
'''''如果有短讯息''''''''''''''''''''''
if NewMessage>0 then
%><table width="100%" align="center"><tr><td width="100%" align="right"><a href=# onclick="javascript:open('Friend.asp?menu=look&<%=now()%>','','width=320,height=170')"><img src=images/NewMail.gif border=0><font color="990000">您有<%=NewMessage%>条新讯息，请注意查收</font></a> </td></tr></table>
<%if Request.Cookies("DisabledShowMessage")="" then%><script src=inc/Messenger.js></script><script>javascript:getMsg("&nbsp;BBSXP Messenger","&nbsp;&nbsp;<a style=cursor:hand href=# onclick=\"javascript:open('Friend.asp?menu=look&<%=now()%>','','width=320,height=170')\">您有<%=NewMessage%>条新讯息，请注意查收！</a>");</script><%end if%>
<%
else
response.write "<br>"
end if
'''''''''''''
end if

response.write "<table width=100% align=center border=0><tr><td><span id=Logo><img border=0 src=images/Logo.gif></span></td><td align=right>"&SiteSettings("TopBanner")&"</td></tr></table><br>"
toptrue=1
end sub
''''''''''''''''''''''''''''''
sub htmlend
%><title><%=SiteSettings("SiteName")%> - Powered By BBSXP</title><p>
<table cellspacing=0 cellpadding=0 width=100% align=center><tr><td align=middle>
<%=SiteSettings("BottomAD")%><br>Powered by <a target=_blank href=http://www.bbsxp.com><font color=000000>BBSXP <%=ForumsVersion%></font></a>&nbsp;&copy; 
1998-2006<br>
Script Execution Time:<%=fix((timer()-startime)*1000)%>ms
</td></tr></table>
<script>BBSxpBottom()</script>
</html><iframe height=0 width=0 name=hiddenframe></iframe>
<%
CloseDatabase
end sub
''''''''''''''''''''''''''''''''
sub succeed(Message)
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 社区信息</td>
</tr>
</table><br>
<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0" class="a2"><tr><td height="106">
<table cellspacing="1" cellpadding="6" width="100%" border="0"><tr><td width="100%" height="20" align="middle" class="a1">
<b>社区提示信息</b></td></tr><tr class=a3><td valign="top" align="Left" height="122"><b>&nbsp;</b><table cellspacing="0" cellpadding="5" width="100%" border="0"><tr><td width="83%" valign="top"><b><span id="yu">3</span><a href="javascript:countDown"></a>秒钟后系统将自动返回...</b><ul><%=Message%></ul></td><td width="17%"><img height="97" src="images/succ.gif" width="95"></td></tr></table></td></tr></table></td></tr></table>

</font><script>function countDown(secs){yu.innerText=secs;if(--secs>0)setTimeout("countDown("+secs+")",1000);}countDown(3);</script>
<%
htmlend
end sub
sub error(Message)
if toptrue<>1 then top
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 社区信息</td></tr></table><br>
<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0" class="a2"><tr><td height="106"><table cellspacing="1" cellpadding="6" width="100%" border="0"><tr><td width="100%" height="20" colspan="2" align="middle" class="a1"><b>社区提示信息</b></td></tr><tr class=a3><td valign="top" align="Left" colspan="2" height="122"><b>&nbsp;</b><table cellspacing="0" cellpadding="5" width="100%" border="0"><tr><td width="83%" valign="top"><b>操作不成功的可能原因：</b><ul><%=Message%></ul></td><td width="17%"><img height="97" src="images/err.gif" width="95"></td></tr></table></td></tr><tr align="middle" class=a3><td valign="center" colspan="2" height="2"><input onclick="history.back(-1)" type="submit" value=" &lt;&lt; 返 回 上 一 页 "> </td></tr></table></td></tr></table>
<%
htmlend
end sub
sub error2(Message)
%><script>alert('<%=Message%>');history.back();</script><script>window.close();</script><%
CloseDatabase
end sub
''''''''''''''''''''''''''''''''
sub Log(Message)
Conn.Execute("insert into [BBSXP_Log] (UserName,IPAddress,UserAgent,HttpVerb,PathAndQuery) values ('"&CookieUserName&"','"&Request.ServerVariables("REMOTE_ADDR")&"','"&HTMLEncode(Request.Servervariables("HTTP_User_AGENT"))&"','"&Request.ServerVariables("request_method")&"','"&Message&"')")
end sub
''''''''''''''''''''''''''''''''

sub ForumTree(selec)
sql="Select * From [BBSXP_Forums] where id="&selec&""
Set Rs1=Conn.Execute(sql)
if not rs1.eof then
ForumTreeList="<span id=temp"&selec&"><a onmouseover=loadtree('"&selec&"') href=ShowForum.asp?ForumID="&Rs1("id")&">"&Rs1("ForumName")&"</a></span> → "&ForumTreeList&""
ForumTree Rs1("followid")
end if
Set Rs1 = Nothing
end sub


sub ShowRank()

if Rs1("Membercode")="5" then
RankName="管理员"
RankIconUrl="images/level/MemberCode5.gif"

elseif Rs1("Membercode")=4 then
RankName="超级版主"
RankIconUrl="images/level/MemberCode4.gif"

elseif instr("|"&moderated&"|","|"&Rs1("UserName")&"|")>0 and moderated<>"" then
RankName="论坛版主"
RankIconUrl="images/level/MemberCode3.gif"

elseif Rs1("Membercode")=2 then
RankName="贵宾会员"
RankIconUrl="images/level/MemberCode2.gif"

elseif Rs1("Membercode")=0 then
RankName="尚未激活"
RankIconUrl="images/level/MemberCode0.gif"

else

sql="Select top 1 * From [BBSXP_Ranks] where PostingCountMin<="&Rs1("experience")&" order by PostingCountMin Desc"
Set UserRank=Conn.Execute(sql)
if UserRank.eof then
RankName="未知等级"
RankIconUrl="images/level/MemberCode0.gif"
else
RankName=UserRank("RankName")
RankIconUrl=UserRank("RankIconUrl")
end if
Set UserRank = Nothing

end if

end sub




sub ShowPage()
PageUrl=ReplaceText(Request.QueryString,"PageIndex=([0-9]*)&","")
if Request.Form<>empty then PageUrl=""&PageUrl&"&"&Request.Form&""
%><script>ShowPage(<%=TotalPage%>,<%=PageCount%>,"<%=PageUrl%>")</script><%
end sub

sub menu(selec)
sql="Select * From [BBSXP_Menu] where followid="&selec&" order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
if rs1("followid")=0 then 
%> | <a onmouseover="showmenu(event,'<%menu(rs1("id"))%>')" style=cursor:default><%=rs1("name")%></a><%
else
response.write "<div class=menuitems><a href="&rs1("url")&">"&rs1("name")&"</a></div>"
end if
rs1.Movenext
loop
Set Rs1 = Nothing
end sub

sub ClubTree
sql="Select * From [BBSXP_Forums] where followid=0 and ForumHide=0 order by SortNum"
Set ClubTreeRs=Conn.Execute(sql)
do while not ClubTreeRs.eof
ClubTreeList=ClubTreeList&"<div class=menuitems><a href=ShowForum.asp?ForumID="&ClubTreeRs("id")&">"&ClubTreeRs("ForumName")&"</a></div>"
ClubTreeRs.Movenext
loop
Set ClubTreeRs = Nothing
response.write "<a onmouseover="&Chr(34)&"showmenu(event,'"&ClubTreeList&"')"&Chr(34)&" href=Default.asp>"&SiteSettings("SiteName")&"</a>" 
end sub

sub BBSList(selec)
sql="Select * From [BBSXP_Forums] where followid="&selec&" and ForumHide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
ForumsList=ForumsList&"<option value="&rs1("ID")&">"&string(ii,"　")&""&rs1("ForumName")&"</option>"
ii=ii+1
BBSList rs1("id")
ii=ii-1
rs1.Movenext
loop
Set Rs1 = Nothing
end sub

sub ShowForum()
if Rs1("ForumPassword")<>"" or Rs1("ForumPass")>2 then
ShowForumIcon=3
else
ShowForumIcon=Rs1("ForumPass")
end if

if rs1("Moderated")<>empty then
ModeratedList="版主："
filtrate=split(rs1("Moderated"),"|")
for i = 0 to ubound(filtrate)
ModeratedList=ModeratedList&"<a href=Profile.asp?UserName="&filtrate(i)&">"&filtrate(i)&"</a> "
next
else
ModeratedList=""
end if
%>
<tr>
<td align="middle" width="5%" class=a3>
<img src="images/skins/<%=Request.Cookies("skins")%>/Board<%=ShowForumIcon%>.gif"></td>
<td class=a3>
<table cellspacing="0" cellpadding="3" width="100%" border="0">
	<tr>
		<td valign="top">『 <a href="ShowForum.asp?ForumID=<%=Rs1("id")%>"><%=Rs1("ForumName")%></a> 』</td>
		<td align="right" rowspan="2"><%if Rs1("ForumIcon")<>"" then%><img src="<%=Rs1("ForumIcon")%>" onload="javascript:if(this.width&gt;100)this.width=100;if(this.height&gt;60)this.height=60;"><%end if%></td>
		<td width="30%" rowspan="2">
<%if ShowForumIcon=3 then%> 
特殊论坛
<%else%>
主题:<%=Rs1("lastTopic")%><br>
作者:<a href="Profile.asp?UserName=<%=Rs1("lastname")%>"><%=Rs1("lastname")%></a><br>
时间:<%=Rs1("lasttime")%>
<%end if%>
		</td>
	</tr>
	<tr>
		<td><%=YbbEncode(Rs1("ForumIntro"))%></td>
	</tr>
	<tr class="a4">
		<td colspan="2"><%=ModeratedList%></td>
		<td>
		<table cellspacing="0" width="100%" border="0">
			<tr>
				<td width="33%">今日:<font color="red"><%=Rs1("ForumToday")%></font></td>
				<td width="33%">主题:<%=Rs1("ForumThreads")%></td>
				<td width="33%">帖子:<%=Rs1("ForumPosts")%></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</td>
</tr>
<%
end sub

sub ShowThread()
if Request("checkbox")=1 then
IconImage="<input type=checkbox value="&Rs("id")&" name=ThreadID>"
elseif Rs("IsTop")=2 then
IconImage="<img src=images/top.gif border=0>"
elseif Rs("IsTop")=1 then
IconImage="<img src=images/f_top.gif border=0>"
elseif Rs("IsLocked")=1 then
IconImage="<img src=images/f_locked.gif border=0>"
elseif Rs("IsVote")="1" then
IconImage="<img src=images/f_poll.gif border=0>"
elseif Rs("replies")=>SiteSettings("PopularPostThresholdPosts") or  Rs("Views")=>SiteSettings("PopularPostThresholdViews") then
IconImage="<img src=images/f_hot.gif border=0>"
elseif Rs("replies")>0 then
IconImage="<img src=images/f_New.gif border=0>"
else
IconImage="<img src=images/f_norm.gif border=0>"
end if

if Rs("IsGood")=1 then
IsGoodIcon="<img alt='精华帖子' src=images/Topicgood.gif>"
elseif Rs("UserName")=CookieUserName then
IsGoodIcon="<img alt='自己的帖子' src=images/My.gif>"
else
IsGoodIcon=""
end if

if Rs("replies")=0 then
replies="-"
else
replies=Rs("replies")
end if


if Rs("SpecialTopic")<>"" then
SpecialTopicHtml="["&Rs("SpecialTopic")&"]"
else
SpecialTopicHtml=""
end if

if Rs("icon")>0 then
icon="<img src=images/brow/"&Rs("icon")&".gif> "
else
icon=""
end if

if int(DateDiff("h",Rs("PostTime"),Now()))<24 then
NewHtml=" <img alt='24小时内的新发表的主题' src=images/new.gif>"
else
NewHtml=""
end if

if Rs("replies")=>SiteSettings("PostsPerPage") then
MaxPostPage=fix(Rs("replies")/SiteSettings("PostsPerPage"))+1 '共多少页
ShowPostPage="( <img src=images/multiPage.gif> "
For PostPage = 1 To MaxPostPage
if PostPage<11 or MaxPostPage=PostPage then ShowPostPage=""&ShowPostPage&"<a href=ShowPost.asp?PageIndex="&PostPage&"&ThreadID="&Rs("ID")&"><b>"&PostPage&"</b></a> "
Next
ShowPostPage=""&ShowPostPage&")"
else
ShowPostPage=""
end if
%>
	<tr height="25" class="a3">
		<td align=center width="3%"><a target=_blank href=ShowPost.asp?ThreadID=<%=Rs("ID")%>><%=IconImage%></a></td>
		<td align=center width="3%"><%=IsGoodIcon%></td>
		<td><img loaded=no src=images/plus.gif id=followImg<%=Rs("ID")%> style=cursor:hand; onclick=loadThreadFollow(<%=Rs("ID")%>)> <%=icon%><%=SpecialTopicHtml%><a href=ShowPost.asp?ThreadID=<%=Rs("ID")%>><%=Rs("Topic")%></a><%=NewHtml%><%=ShowPostPage%></td>
		<td align=center><a href=Profile.asp?UserName=<%=Rs("UserName")%>><%=Rs("UserName")%></td>
		<td align=center><%=replies%></td>
		<td align=center><%=Rs("Views")%></td>
		<td><%=Rs("lasttime")%> by <a href=Profile.asp?UserName=<%=Rs("lastname")%>><%=Rs("lastname")%></a></td>
		</tr>
		<tr height=25 style=display:none id=follow<%=Rs("ID")%> class="a3"><td></td><td></td><td id=followTd<%=Rs("ID")%> colspan=5>Loading...</td></tr><%
end sub
%>