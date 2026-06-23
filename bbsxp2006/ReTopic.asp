<!-- #include file="Setup.asp" -->
<%
top

if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&SqlString(CookieUserName)&"'" ).eof Then error("<li>�����ؽ�<a href=Prison.asp>����</a>")
ThreadID=RequestInt("ThreadID")



sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
if Rs("IsLocked")=1 then error("<li>�������Ѿ��رգ��������µĻظ�")
ForumID=Rs("ForumID")
PostsTableName=SafeTableSuffix(Rs("PostsTableName"))
if PostsTableName="" then PostsTableName="0"
Topic=Rs("Topic")
Subject=ReplaceText(Rs("Topic"),"<[^>]*>","")
Rs.close


sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=SafeUrl(Rs("ForumLogo"))
moderated=Rs("moderated")
followid=Rs("followid")
Rs.close

if membercode>1 or instr("|"&moderated&"|","|"&CookieUserName&"|")>0 then UserPopedomPass=1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then

if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>��֤�����"
end if

Subject=HTMLEncode(Request.Form("Subject"))
color=HTMLEncode(Request.Form("color"))
Content=ContentEncode(Request.Form("Content"))
if Request.Form("DisableYBBCode")<>1 then Content=YbbEncode(Content)

if Len(content)<2 then Message=Message&"<li>�������ݲ���С�� 2 �ַ�"

if Message<>"" then error(""&Message&"")


if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
Subject=ReplaceText(Subject,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


sql="select * from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Rs.Open sql,Conn,1,3

StopPostTime=int(DateDiff("s",Rs("UserLandTime"),Now()))
if StopPostTime < int(SiteSettings("DuplicatePostIntervalInMinutes")) then Message=Message&"<li>��̳����һ�������η������������� "&SiteSettings("DuplicatePostIntervalInMinutes")&" �룡<li>�������ٵȴ� "&SiteSettings("DuplicatePostIntervalInMinutes")-StopPostTime&" �룡"

StopPostTime=int(DateDiff("s",Rs("UserRegTime"),Now()))
if StopPostTime < int(SiteSettings("RegUserTimePost")) then Message=Message&"<li>��ע���û�����ȴ� "&SiteSettings("RegUserTimePost")&" �����ܷ�����<li>�������ٵȴ� "&SiteSettings("RegUserTimePost")-StopPostTime&" �룡"

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
UpFileItem=SafeLongValue(UpFileID(i),0)
if UpFileItem>0 then Conn.execute("update [BBSXP_PostAttachments] set ThreadID="&ThreadID&",Description='"&SqlString(Subject)&"' where id="&UpFileItem&" and ThreadID=0")
next
end if


if UserPopedomPass=1 and color<>"" then Subject="<font color="&color&">"&Subject&"</font>"

Conn.Execute("insert into [BBSXP_Posts"&PostsTableName&"] (ThreadID,UserName,Subject,content,Postip) values ('"&ThreadID&"','"&SqlString(CookieUserName)&"','"&SqlString(Subject)&"','"&SqlString(content)&"','"&SqlString(Request.ServerVariables("REMOTE_ADDR"))&"')")
Conn.execute("update [BBSXP_Threads] set lastname='"&SqlString(CookieUserName)&"',replies=replies+1,lasttime="&SqlNowString&" where ID="&ThreadID&"")
Conn.execute("update [BBSXP_Forums] set lastTopic='"&SqlString("<a href=ShowPost.asp?ThreadID="&ThreadID&">"&Left(HTMLEncode(Request.Form("Subject")),15)&"</a>")&"',lastname='"&SqlString(CookieUserName)&"',lasttime="&SqlNowString&",ForumToday=ForumToday+1,ForumPosts=ForumPosts+1 where id="&ForumID&"")
Conn.execute("update [BBSXP_Statistics_Site] set TodayPost=TodayPost+1,TotalPost=TotalPost+1")

Session("VerifyCode")=""

Message=Message&"<li>�ظ�����ɹ�<li><a href=ShowPost.asp?ThreadID="&ThreadID&">��������</a><li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")


end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


PostID=RequestInt("PostID")
if PostID>0 then
sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&PostID&""
Set Rs=Conn.Execute(sql)
Subject=ReplaceText(""&Rs("Subject")&"","<[^>]*>","")
if Request("quote")=1 then
quote="<blockquote><strong>����</strong>��<hr>ԭ���� <b>"&Rs("UserName")&"</b> ������ <i>"&Rs("Posttime")&"</i> :<br>"&Rs("content")&""&vbCrlf&"<hr></blockquote>"
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
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> �� �ظ�����</td>
		</tr>
	</table><br>




<TABLE cellSpacing=1 cellPadding=5 width=100% border=0 class=a2 align="center">
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=HTMLEncode(quote)%>'>
<input type=hidden name=ThreadID value=<%=ThreadID%>>
<input name="UpFileID" type="hidden">
<TR class=a1>
<TD vAlign=Left colSpan=2 height=25><b>�ظ�����</b></TD></TR>


<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
	<tr>
<TD class=a3 height=6><b>��֤��</b></TD>
<TD class=a3 height=6>
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="��֤��,�������?����ˢ����֤��" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></TD>
	</tr>
<%end if%>


<TR class=a4>
<TD width=180><B>���±��� </B> 
</TD>
<TD class=a3 height=25>
<INPUT maxLength=50 size=60 name=Subject value="Re:<%=Subject%>">
<%if UserPopedomPass=1 then %>
<SELECT name=color onchange="title_color(this.options[this.selectedIndex].value)">
<option value="">��ɫ</option>
<option style=background-color:Black;color:Black value=Black>��ɫ</option>
<option style=background-color:green;color:green value=green>��ɫ</option>
<option style=background-color:red;color:red value=red>��ɫ</option>
<option style=background-color:blue;color:blue value=blue>��ɫ</option>
<option style=background-color:Navy;color:Navy value=Navy>����</option>
<option style=background-color:Teal;color:Teal value=Teal>��ɫ</option>
<option style=background-color:Purple;color:Purple value=Purple>��ɫ</option>
<option style=background-color:Fuchsia;color:Fuchsia value=Fuchsia>�Ϻ�</option>
<option style=background-color:Gray;color:Gray value=Gray>��ɫ</option>
<option style=background-color:Olive;color:Olive value=Olive>���</option>
</SELECT>
<%end if%>
</TD></TR>



<TR>
<TD vAlign=top class=a3>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=Left border=0 height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><br><B>��������</B><BR>
��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<BR><BR><span id=UpFile></span>
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> ����YBB����</label>
</TD></TR>
</TABLE></TD>
<TD class=a3 height=250>

<SCRIPT src="inc/Post.js"></SCRIPT>
</TD></TR>


<%if SiteSettings("UpFileOption")<>empty then%>
<TR>
<TD align=Left class=a4>
<IMG src=images/affix.gif alt="֧������<%=SiteSettings("UpFileTypes")%>"><b>���Ӹ���</b>������:<%=CheckSize(SiteSettings("MaxFileSize"))%></b>��</TD>
<TD align=Left class=a4><IFRAME src="PostUpFile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></TD></TR>
<%end if%>

<TR>
<TD align=middle class=a3 colSpan=2 height=27>
<INPUT type=submit value=�ظ����� name=EditSubmit>&nbsp;   <INPUT type=reset value=" �� �� "></TD></TR></FORM>
</TABLE>



<%

htmlend
%>