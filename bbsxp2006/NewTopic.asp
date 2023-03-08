<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&CookieUserName&"'" ).eof Then error("<li>您被关进<a href=Prison.asp>监狱</a>")
ForumID=int(Request("ForumID"))

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=Rs("ForumLogo")
moderated=Rs("moderated")
followid=Rs("followid")
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
TolSpecialTopic=Rs("TolSpecialTopic")
ForumPass=Rs("ForumPass")
Rs.close

if membercode>1 or instr("|"&moderated&"|","|"&CookieUserName&"|")>0 then UserPopedomPass=1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then



if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>验证码错误"
end if

color=HTMLEncode(Request.Form("color"))
icon=Request.Form("icon")
Subject=HTMLEncode(Request.Form("Subject"))
Content=ContentEncode(Request.Form("Content"))
if Request.Form("DisableYBBCode")<>1 then Content=YbbEncode(Content)

if Len(Subject)<2 then Message=Message&"<li>文章主题不能小于 2 字符"
if Len(content)<2 then Message=Message&"<li>文章内容不能小于 2 字符"

if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
Subject=ReplaceText(Subject,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if



'''''''''''''''''''''''''''''''

if Request.Form("Vote")<>"" then
Vote=Request("Vote")
if instr(Vote,"|") > 0 then error("<li>投票选项中不能含有“|”字符")
pollTopic=split(Vote,chr(13)&chr(10))
j=0
for i = 0 to ubound(pollTopic)
if not (pollTopic(i)="" or pollTopic(i)=" ") then
allpollTopic=""&allpollTopic&""&pollTopic(i)&"|"
j=j+1
end if
next

if j<SiteSettings("MinVoteOptions") or j>SiteSettings("MaxVoteOptions") then error("<li>投票选项不能少于 "&SiteSettings("MinVoteOptions")&" 个<li>投票选项超过 "&SiteSettings("MaxVoteOptions")&" 个")


for y = 1 to j
Votenum=""&Votenum&"0|"
next

end if
'''''''''''''''''''''''''''''''

if Message<>"" then error(""&Message&"")



sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3

StopPostTime=int(DateDiff("s",Rs("UserLandTime"),Now()))
if StopPostTime < int(SiteSettings("DuplicatePostIntervalInMinutes")) then Message=Message&"<li>论坛限制一个人两次发帖间隔必须大于 "&SiteSettings("DuplicatePostIntervalInMinutes")&" 秒！<li>您必须再等待 "&SiteSettings("DuplicatePostIntervalInMinutes")-StopPostTime&" 秒！"

StopPostTime=int(DateDiff("s",Rs("UserRegTime"),Now()))
if StopPostTime < int(SiteSettings("RegUserTimePost")) then Message=Message&"<li>新注册用户必须等待 "&SiteSettings("RegUserTimePost")&" 秒后才能发帖！<li>您必须再等待 "&SiteSettings("RegUserTimePost")-StopPostTime&" 秒！"

if Message<>"" then error(""&Message&"")

Rs("PostTopic")=Rs("PostTopic")+1
Rs("UserMoney")=Rs("UserMoney")+SiteSettings("IntegralAddThread")
Rs("experience")=Rs("experience")+SiteSettings("IntegralAddThread")
Rs("UserLandTime")=now()
Rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
Rs.update
Rs.close


if UserPopedomPass=1 and color<>"" then Subject="<font color="&color&">"&Subject&"</font>"

Rs.Open "select * from [BBSXP_Threads]",Conn,1,3
Rs.addNew
Rs("UserName")=CookieUserName
Rs("PostTime")=now()
Rs("lastname")=CookieUserName
Rs("lasttime")=now()
Rs("Topic")=Subject
Rs("ForumID")=ForumID
Rs("PostsTableName")=SiteSettings("DefaultPostsName")
if Request("SpecialTopic")<>"" then Rs("SpecialTopic")=Request("SpecialTopic")
if Request("icon")<>"" then Rs("icon")=Request("icon")
if Request("Vote")<>"" then Rs("isVote")=1
if Request("IsLocked")=1 then Rs("IsLocked")=1
if ForumPass=5 then Rs("IsDel")=1
Rs.update
ID=Rs("ID")
Rs.close

if Request.Form("Vote")<>"" then
multiplicity=int(Request.Form("multiplicity"))
Expiry=now()+int(Request.Form("Expiry"))
Conn.Execute("insert into [BBSXP_Vote] (ThreadID,Type,Items,Result,Expiry) values ('"&ID&"','"&multiplicity&"','"&HTMLEncode(allpollTopic)&"','"&Votenum&"','"&Expiry&"')")
end if

if Request.Form("UpFileID")<>"" then
UpFileID=split(Request.form("UpFileID"),",")
for i = 0 to ubound(UpFileID)-1
Conn.execute("update [BBSXP_PostAttachments] set ThreadID="&ID&",Description='"&Subject&"' where id="&UpFileID(i)&" and ThreadID=0")
next
end if

Conn.Execute("insert into [BBSXP_Posts"&SiteSettings("DefaultPostsName")&"] (ThreadID,IsTopic,UserName,Subject,content,Postip) values ('"&ID&"','1','"&CookieUserName&"','"&Subject&"','"&content&"','"&Request.ServerVariables("REMOTE_ADDR")&"')")
Conn.execute("update [BBSXP_Forums] set lastTopic='<a href=ShowPost.asp?ThreadID="&id&">"&Left(HTMLEncode(Request.Form("Subject")),15)&"</a>',lastname='"&CookieUserName&"',lasttime="&SqlNowString&",ForumToday=ForumToday+1,ForumThreads=ForumThreads+1,ForumPosts=ForumPosts+1 where id="&ForumID&"")
Conn.execute("update [BBSXP_Statistics_Site] set TodayPost=TodayPost+1,TotalPost=TotalPost+1,TotalThread=TotalThread+1")

Session("VerifyCode")=""

if ForumPass=5 then
EnableCensorship="由于论坛设有审查制度，您发表的帖子需要等待激活才能显示。"
else
EnableCensorship="<a href=ShowPost.asp?ThreadID="&id&">返回主题</a>"
end if

Message="<li>新主题发表成功<li>"&EnableCensorship&"<li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")

end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



%>
<!-- #include file="inc/Validate.asp" -->
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
function ShowADv(){
if (document.yuziform.advShow.checked == true) {
adv.style.display = "";
}else{
adv.style.display = "none";
}
}
function title_color(color){document.yuziform.Subject.style.color = color;}
</script>
	<table border="0" width="100%" align=center cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → 发表帖子</td>
		</tr>
	</table><br>


<TABLE cellSpacing=1 cellPadding=6 width=100% border=0 class=a2 align=center>
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden"><input name="UpFileID" type="hidden">
<input type=hidden name=ForumID value=<%=ForumID%>>
<TR>
<TD vAlign=Left colSpan=2 height=25 class=a1><b>发表帖子</b></TD></TR>
<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
	<tr>
<TD class=a3 height=6><b>验证码</b></TD>
<TD class=a3 height=6>
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="验证码,看不清楚?请点击刷新验证码" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></TD>
	</tr>
<%end if%>
	
<TR>
<TD class=a3 width="180"><B>文章标题 </B> 
<%
if TolSpecialTopic<>empty then
response.write "<SELECT name=SpecialTopic size=1><OPTION value='' selected>&nbsp;专题</OPTION>"
filtrate=split(TolSpecialTopic,"|")
for i = 0 to ubound(filtrate)
response.write "<OPTION value='"&filtrate(i)&"'>["&filtrate(i)&"]</OPTION>"
next
response.write "</SELECT>"
end if
%>

</TD>
<TD class=a3 height=13>
<INPUT maxLength=50 size=60 name=Subject>
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
<TD vAlign=top align=Left class=a4 height=23><B>您的表情</B></TD>
<TD class=a4>
<script>
for(i=1;i<=12;i++) {
document.write("<INPUT type=radio value="+i+" name=icon><IMG src=images/brow/"+i+".gif>　")
}
</script>

 </TD></TR>

<TR>
<TD vAlign=top class=a3>
<TABLE cellSpacing=0 cellPadding=0 align=Left border=0 width=100% height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><BR><B>文章内容</B><BR>
（<a href="javascript:CheckLength();">查看内容长度</a>）<BR>
<BR><span id=UpFile></span>


</TD></TR>

<TR><TD valign="bottom">
<INPUT id=LockMyPost name=IsLocked type=checkbox value=1><label for=LockMyPost> 主题锁定，不允许回复</label><br>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> 禁用YBB代码</label><br>
<INPUT id=advcheck name=advShow type=checkbox value=1 onclick=ShowADv()><label for=advcheck> 显示投票选项</label>

</TD></TR>

</TABLE></TD>

<TD class=a3 height=250>

<SCRIPT src="inc/Post.js"></SCRIPT>

</TD></TR>


<TR id=adv style=DISPLAY:none>
<TD vAlign=top align=Left class=a4>


<FONT color=000000><B>投票项目</B><BR>
每行一个投票项目<BR>
过期天数 <INPUT maxLength=3 size=2 name=Expiry value="7" onkeyup=if(isNaN(this.value))this.value=''> 天<br>

<INPUT type=radio CHECKED value=0 name=multiplicity id=multiplicity>
<label for=multiplicity>单选投票</label>
<BR><INPUT type=radio value=1 name=multiplicity id=multiplicity_1> <label for=multiplicity_1>多选投票</label></FONT> 
</TD>
<TD class=a4>
<TEXTAREA name=Vote rows=5 style="width:100%"></TEXTAREA>
</TD></TR>

<%if SiteSettings("UpFileOption")<>empty then%>
<TR>
<TD align=Left class=a4>
<IMG src=images/affix.gif alt="支持类型<%=SiteSettings("UpFileTypes")%>"><b>增加附件</b>（限制:<%=CheckSize(SiteSettings("MaxFileSize"))%></b>）</TD>
</TD>
<TD align=Left class=a4><IFRAME src="PostUpFile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></TD></TR>
<%end if%>

<TR>
<TD align=middle class=a3 colSpan=2 height=27>
<INPUT type=submit value=发表新主题 name=EditSubmit>&nbsp; <INPUT type=reset value=" 重 置 "></TD></TR></FORM>
</TABLE>





<%
htmlend
%>