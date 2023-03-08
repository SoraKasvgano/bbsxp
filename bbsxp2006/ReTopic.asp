<!-- #include file="Setup.asp" -->
<%
top

if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&CookieUserName&"'" ).eof Then error("<li>您被关进<a href=Prison.asp>监狱</a>")
ThreadID=int(Request("ThreadID"))



sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
if Rs("IsLocked")=1 then error("<li>此主题已经关闭，不接受新的回复")
ForumID=Rs("ForumID")
PostsTableName=Rs("PostsTableName")
Topic=Rs("Topic")
Subject=ReplaceText(Rs("Topic"),"<[^>]*>","")
Rs.close


sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=Rs("ForumLogo")
moderated=Rs("moderated")
followid=Rs("followid")
Rs.close

if membercode>1 or instr("|"&moderated&"|","|"&CookieUserName&"|")>0 then UserPopedomPass=1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then

if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>验证码错误"
end if

Subject=HTMLEncode(Request.Form("Subject"))
color=HTMLEncode(Request.Form("color"))
Content=ContentEncode(Request.Form("Content"))
if Request.Form("DisableYBBCode")<>1 then Content=YbbEncode(Content)

if Len(content)<2 then Message=Message&"<li>文章内容不能小于 2 字符"

if Message<>"" then error(""&Message&"")


if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
Subject=ReplaceText(Subject,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3

StopPostTime=int(DateDiff("s",Rs("UserLandTime"),Now()))
if StopPostTime < int(SiteSettings("DuplicatePostIntervalInMinutes")) then Message=Message&"<li>论坛限制一个人两次发帖间隔必须大于 "&SiteSettings("DuplicatePostIntervalInMinutes")&" 秒！<li>您必须再等待 "&SiteSettings("DuplicatePostIntervalInMinutes")-StopPostTime&" 秒！"

StopPostTime=int(DateDiff("s",Rs("UserRegTime"),Now()))
if StopPostTime < int(SiteSettings("RegUserTimePost")) then Message=Message&"<li>新注册用户必须等待 "&SiteSettings("RegUserTimePost")&" 秒后才能发帖！<li>您必须再等待 "&SiteSettings("RegUserTimePost")-StopPostTime&" 秒！"

if Message<>"" then error(""&Message&"")
Rs("Postrevert")=Rs("Postrevert")+1
Rs("UserMoney")=Rs("UserMoney")+SiteSettings("IntegralAddPost")
Rs("experience")=Rs("experience")+SiteSettings("IntegralAddPost")
Rs("UserLandTime")=now()
Rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
Rs.update
Rs.close

if Request.Form("UpFileID")<>"" then
UpFileID=split(Request.form("UpFileID"),",")
for i = 0 to ubound(UpFileID)-1
Conn.execute("update [BBSXP_PostAttachments] set ThreadID="&ThreadID&",Description='"&Subject&"' where id="&UpFileID(i)&" and ThreadID=0")
next
end if


if UserPopedomPass=1 and color<>"" then Subject="<font color="&color&">"&Subject&"</font>"

Conn.Execute("insert into [BBSXP_Posts"&PostsTableName&"] (ThreadID,UserName,Subject,content,Postip) values ('"&ThreadID&"','"&CookieUserName&"','"&Subject&"','"&content&"','"&Request.ServerVariables("REMOTE_ADDR")&"')")
Conn.execute("update [BBSXP_Threads] set lastname='"&CookieUserName&"',replies=replies+1,lasttime="&SqlNowString&" where ID="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set lastTopic='<a href=ShowPost.asp?ThreadID="&ThreadID&">"&Left(HTMLEncode(Request.Form("Subject")),15)&"</a>',lastname='"&CookieUserName&"',lasttime="&SqlNowString&",ForumToday=ForumToday+1,ForumPosts=ForumPosts+1 where id="&ForumID&"")
Conn.execute("update [BBSXP_Statistics_Site] set TodayPost=TodayPost+1,TotalPost=TotalPost+1")

Session("VerifyCode")=""

Message=Message&"<li>回复主题成功<li><a href=ShowPost.asp?ThreadID="&ThreadID&">返回主题</a><li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")


end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


if isnumeric(""&Request("PostID")&"") then
sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&Request("PostID")&""
Set Rs=Conn.Execute(sql)
Subject=ReplaceText(""&Rs("Subject")&"","<[^>]*>","")
if Request("quote")=1 then
quote="<blockquote><strong>引用</strong>：<hr>原文由 <b>"&Rs("UserName")&"</b> 发表于 <i>"&Rs("Posttime")&"</i> :<br>"&Rs("content")&""&vbCrlf&"<hr></blockquote>"
end if
Rs.close
end if

%>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
function title_color(color){document.yuziform.Subject.style.color = color;}
</script>


	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> → 回复帖子</td>
		</tr>
	</table><br>




<TABLE cellSpacing=1 cellPadding=5 width=100% border=0 class=a2 align="center">
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=quote%>'>
<input type=hidden name=ThreadID value=<%=ThreadID%>>
<input name="UpFileID" type="hidden">
<TR class=a1>
<TD vAlign=Left colSpan=2 height=25><b>回复帖子</b></TD></TR>


<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
	<tr>
<TD class=a3 height=6><b>验证码</b></TD>
<TD class=a3 height=6>
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="验证码,看不清楚?请点击刷新验证码" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></TD>
	</tr>
<%end if%>


<TR class=a4>
<TD width=180><B>文章标题 </B> 
</TD>
<TD class=a3 height=25>
<INPUT maxLength=50 size=60 name=Subject value="Re:<%=Subject%>">
<%if UserPopedomPass=1 then %>
<SELECT name=color onchange="title_color(this.options[this.selectedIndex].value)">
<option value="">颜色</option>
<option style=background-color:Black;color:Black value=Black>黑色</option>
<option style=background-color:green;color:green value=green>绿色</option>
<option style=background-color:red;color:red value=red>红色</option>
<option style=background-color:blue;color:blue value=blue>蓝色</option>
<option style=background-color:Navy;color:Navy value=Navy>深蓝</option>
<option style=background-color:Teal;color:Teal value=Teal>青色</option>
<option style=background-color:Purple;color:Purple value=Purple>紫色</option>
<option style=background-color:Fuchsia;color:Fuchsia value=Fuchsia>紫红</option>
<option style=background-color:Gray;color:Gray value=Gray>灰色</option>
<option style=background-color:Olive;color:Olive value=Olive>橄榄</option>
</SELECT>
<%end if%>
</TD></TR>



<TR>
<TD vAlign=top class=a3>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=Left border=0 height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><br><B>文章内容</B><BR>
（<a href="javascript:CheckLength();">查看内容长度</a>）<BR><BR><span id=UpFile></span>
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> 禁用YBB代码</label>
</TD></TR>
</TABLE></TD>
<TD class=a3 height=250>

<SCRIPT src="inc/Post.js"></SCRIPT>
</TD></TR>


<%if SiteSettings("UpFileOption")<>empty then%>
<TR>
<TD align=Left class=a4>
<IMG src=images/affix.gif alt="支持类型<%=SiteSettings("UpFileTypes")%>"><b>增加附件</b>（限制:<%=CheckSize(SiteSettings("MaxFileSize"))%></b>）</TD>
<TD align=Left class=a4><IFRAME src="PostUpFile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></TD></TR>
<%end if%>

<TR>
<TD align=middle class=a3 colSpan=2 height=27>
<INPUT type=submit value=回复主题 name=EditSubmit>&nbsp;   <INPUT type=reset value=" 重 置 "></TD></TR></FORM>
</TABLE>



<%

htmlend
%>