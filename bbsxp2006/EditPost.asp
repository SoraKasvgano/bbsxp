<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
ThreadID=RequestInt("ThreadID")

sql="Select * From [BBSXP_Threads] where ID="&ThreadID&""
Rs.Open sql,Conn,1
if Rs.eof then error("<li>该主题不存在")
ForumID=Rs("ForumID")
PostsTableName=SafeTableSuffix(Rs("PostsTableName"))
if PostsTableName="" then PostsTableName="0"
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


sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&RequestInt("PostID")&""
Set Rs=Conn.Execute(sql)
if Rs.eof then error("<li>���ݿ��в����ڴ����ӵ�����")
if Rs("UserName")<>CookieUserName and UserPopedomPass<>1 then error("<li>�Բ�������Ȩ�޲�����")
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


if content=empty then Message=Message&"<li>����û����д"
if sitesettings("DisplayEditNotes")=1 then content=""&content&"<p>�۴������ѱ� "&CookieUserName&" �� "&now()&" �༭����"

if Message<>"" then error(""&Message&"")

sql="select * from [BBSXP_Posts"&PostsTableName&"] where id="&RequestInt("PostID")&""
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
 
Message="<li>�޸����ӳɹ�<li><a href=ShowPost.asp?ThreadID="&ThreadID&">��������</a><li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a><li><a href=Default.asp>����������ҳ</a>"
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
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> �� �༭����</td>
		</tr>
	</table><br>


<TABLE cellSpacing=1 cellPadding=5 width=100% border=0 class=a2 align="center">
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=server.htmlencode(content)%>'>
<input name="UpFileID" type="hidden">
<TR>
<TD id=titlelarge vAlign=Left colSpan=2 height=25 class=a1><b>�༭����</b></TD></TR>
<TR class=a4>
<TD width="180" height=25><B>���±���</B>&nbsp; </TD>
<TD height=25>
<INPUT maxLength=50 size=60 name=Subject value="<%=Subject%>">

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
��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<BR><BR>
<span id=UpFile></span>
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3><INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> ����YBB����</label>
</TD></TR>
</TABLE></TD>
<TD class=a3 height=250>
<SCRIPT src="inc/Post.js"></SCRIPT>
</TD></TR>
<%if SiteSettings("UpFileOption")<>empty then%>
<TR>
<TD align=Left class=a4>
<IMG src=images/affix.gif alt="֧������<%=SiteSettings("UpFileTypes")%>"><b>���Ӹ���</b>������:<%=CheckSize(SiteSettings("MaxFileSize"))%></b>��</TD>
</TD>
<TD align=Left class=a4><IFRAME src="PostUpFile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></TD></TR>
<%end if%>
<TR>
<TD align=middle class=a3 colSpan=2 height=27>
<INPUT type=submit value=ȷ���༭ name=EditSubmit>&nbsp; <INPUT type=reset value=" �� �� "></TD></TR></FORM>
</TABLE>





<%
htmlend
%>