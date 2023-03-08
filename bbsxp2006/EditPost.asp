<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
ThreadID=int(Request("ThreadID"))

sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
ForumID=Rs("ForumID")
PostsTableName=Rs("PostsTableName")
Topic=Rs("Topic")
Rs.close

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=Rs("ForumLogo")
moderated=Rs("moderated")
followid=Rs("followid")
Rs.close

if membercode>3 or instr("|"&moderated&"|","|"&CookieUserName&"|")>0 then UserPopedomPass=1


sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&Request("PostID")&""
Set Rs=Conn.Execute(sql)
if Rs.eof then error("<li>数据库中不存在此帖子的数据")
if Rs("UserName")<>CookieUserName and UserPopedomPass<>1 then error("<li>对不起，您的权限不够！")
Subject=ReplaceText(""&Rs("Subject")&"","<[^>]*>","")
content=Rs("content")
Rs.close


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.ServerVariables("request_method") = "POST" then
color=HTMLEncode(Request.Form("color"))
Subject=HTMLEncode(Request.Form("Subject"))
Content=ContentEncode(Request.Form("Content"))
if Request.Form("DisableYBBCode")<>1 then Content=YbbEncode(Content)


if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
Subject=ReplaceText(Subject,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


if content=empty then Message=Message&"<li>内容没有填写"
if sitesettings("DisplayEditNotes")=1 then content=""&content&"<p>［此帖子已被 "&CookieUserName&" 在 "&now()&" 编辑过］"

if Message<>"" then error(""&Message&"")

sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&Request("PostID")&""
Rs.Open sql,Conn,1,3
if UserPopedomPass=1 and color<>"" then Subject="<font color="&color&">"&Subject&"</font>"
if Rs("IsTopic")=1 then Conn.execute("update [BBSXP_Threads] set Topic='"&Subject&"' where ID="&ThreadID&"")
Rs("Subject")=Subject
Rs("content")=content
Rs.update
Rs.close

if Request.Form("UpFileID")<>"" then
UpFileID=split(Request.form("UpFileID"),",")
for i = 0 to ubound(UpFileID)-1
Conn.execute("update [BBSXP_PostAttachments] set ThreadID="&ThreadID&",Description='"&Subject&"' where id="&UpFileID(i)&" and ThreadID=0")
next
end if
 
Message="<li>修改帖子成功<li><a href=ShowPost.asp?ThreadID="&ThreadID&">返回主题</a><li><a href=ShowForum.asp?ForumID="&ForumID&">返回论坛</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")

end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
function title_color(color){document.yuziform.Subject.style.color = color;}
</script>


	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> → <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> → 编辑帖子</td>
		</tr>
	</table><br>


<TABLE cellSpacing=1 cellPadding=5 width=100% border=0 class=a2 align="center">
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=server.htmlencode(content)%>'>
<input name="UpFileID" type="hidden">
<TR>
<TD id=titlelarge vAlign=Left colSpan=2 height=25 class=a1><b>编辑帖子</b></TD></TR>
<TR class=a4>
<TD width="180" height=25><B>文章标题</B>&nbsp; </TD>
<TD height=25>
<INPUT maxLength=50 size=60 name=Subject value="<%=Subject%>">

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
（<a href="javascript:CheckLength();">查看内容长度</a>）<BR><BR>
<span id=UpFile></span>
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3><INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> 禁用YBB代码</label>
</TD></TR>
</TABLE></TD>
<TD class=a3 height=250>
<SCRIPT src="inc/Post.js"></SCRIPT>
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
<INPUT type=submit value=确定编辑 name=EditSubmit>&nbsp; <INPUT type=reset value=" 重 置 "></TD></TR></FORM>
</TABLE>





<%
htmlend
%>