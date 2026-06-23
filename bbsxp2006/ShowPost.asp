<!-- #include file="Setup.asp" -->
<%
top

ThreadID=RequestInt("ThreadID")
ForumID=RequestInt("ForumID")

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
if Rs.eof or Rs.bof then error"<li>ϵͳ�����ڸ����ӵ�����"
if Rs("IsDel")=1 and membercode<4 then error"<li>�������ڻ���վ�У�"
Rs("Views")=Rs("Views")+1
Rs.update
Topic=ReplaceText(Rs("Topic"),"<[^>]*>","")
Replies=Rs("Replies")
Views=Rs("Views")
IsVote=Rs("IsVote")
IsGood=Rs("IsGood")
IsTop=Rs("IsTop")
IsLocked=Rs("IsLocked")
PostsTableName=SafeTableSuffix(Rs("PostsTableName"))
if PostsTableName="" then PostsTableName="0"
ThreadID=Rs("ID")
ForumID=Rs("ForumID")
Rs.close


sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
moderated=Rs("moderated")
ForumLogo=SafeUrl(Rs("ForumLogo"))
followid=Rs("followid")
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
Rs.Close
%>
<!-- #include file="inc/Validate.asp" -->
<!-- Markdown Support -->
<script src="../bbsxp2008/js/marked.min.js"></script>
<script src="../bbsxp2008/js/dompurify.min.js"></script>
<script src="../bbsxp2008/js/markdown-handler.js"></script>
<link rel="stylesheet" href="../bbsxp2008/css/markdown-content.css">

<script src="inc/birth.js"></script>
<title><%=Topic%> - Powered By BBSXP</title>
<script>
if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� <a href="?ThreadID=<%=ThreadID%>"><%=Topic%></a></td>
</tr>
</table><br>

<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0"><tr><td height="35" valign="bottom">
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href=NewTopic.asp?ForumID=<%=ForumID%>>����������</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href=ReTopic.asp?ThreadID=<%=ThreadID%>>�ظ�����</a>
</td><td align="right" height="35" valign="bottom"><font color="333333">���Ǳ����� <b><%=Views%></b> ���Ķ���</font>����<a href="?action=Previous&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img height="12" alt="�����һƪ����" src="images/prethread.gif" width="52" border="0"></a>��<a style="text-decoration: none" href="javascript:location.reload()"><img height="12" alt="ˢ�±�����" src="images/refresh.gif" width="40" border="0"></a>��<a href="?action=Next&ForumID=<%=ForumID%>&ThreadID=<%=ThreadID%>"><img height="12" alt="�����һƪ����" src="images/nextthread.gif" width="52" border="0"></a></font></td></tr></table>

<table width="100%" border="0" cellspacing="1" class="a2" height="21" align="center"><tr class="a1">
	<td width="100%" height="9" colspan="4">
	<table border="0" width="100%" cellspacing="0">
		<tr class="a1">
			<td><b>&nbsp;����</b>��<%=Topic%></td>
				<td align="right"> 
<a target=_blank href=Print.asp?ThreadID=<%=ThreadID%>><img alt="�ʺϴ�ӡ����ӡ�İ汾" src=images/Print.gif border=0></a>&nbsp; 
<script>document.write("<a target=_blank href='Mailto:?subject=<%=Topic%>&body="+location.href+"'>");</script><img alt="ͨ�������ʼ����ʹ�ҳ��" src=images/sendMail.gif border=0></a>&nbsp; 
<a href="javascript:window.external.AddFavorite(location.href,document.title)"><img alt="���Ӽӵ�IE�ղؼ�" src="images/favs.gif" border="0"></a>&nbsp; 
<script>document.write("<a style=cursor:hand onclick=\"javascript:open('Message.asp?menu=Post&report=1&moderated=<%=moderated%>&body=���������ӡ���"+location.href+"','','width=320,height=170')\">");</script><img alt="���汾��" src="images/feedback.gif" border="0"></a>&nbsp; 
</td></tr></table>
</td></tr>
<%
'''''''ͶƱ''''''''
if IsVote=1 then
%>
<tr class="a4">
	<td width="40%" height="25" align="center">
	ѡ��</td>
	<td width="10%" height="25" align="center">
	Ʊ��</td>
	<td width="50%" height="25" align="center" colspan="2">
	�ٷֱ�</td></tr>

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
response.write "ͶƱ�ѹ���"
elseif CookieUserName=empty then
response.write "��¼�����ͶƱ"
elseif instr(Rs("BallotUserList"),""&CookieUserName&"|")>0 then
response.write "���Ѿ�Ͷ��Ʊ��"
else
response.write "<INPUT type=submit value='Ͷ��Ʊ'>"
end if
%>
	</td>
	<td height="25" align="center">
��Ʊ����<%=allticket%></td></form>
	<td colspan="2" align="center">��ֹͶƱʱ�䣺<%=Rs("Expiry")%></td>
	</tr>
	</table>
<%
Rs.Close
end if
'''''''ͶƱ END''''''''


TotalCount=replies+1
PageSetup=SiteSettings("PostsPerPage") '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = cint(Request.QueryString("PageIndex")) '��ȡ��ǰҳ
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
<br><br><img src=<%=RankIconUrl%>><table border=0 width=90%><tr><td><br>�ȡ�����:<%=RankName%><br>
<%if Rs1("Consortia")<>"" then%>��������:<a target="_blank" href="Consortia.asp?menu=look&ConsortiaName=<%=Rs1("Consortia")%>"><%=Rs1("Consortia")%></a><br><%end if%>
<%if Rs1("Consort")<>"" then%>�䡡��ż:<%=Rs1("Consort")%><br><%end if%>
�� �� ֵ:<%=Rs1("experience")%><br>
�������:<%=Rs1("UserMoney")%><br>
�ܷ�����:<%=Rs1("PostTopic")+Rs1("Postrevert")%><br>
ע��ʱ��:<%=split(Rs1("UserRegTime")," ")(0)%><br>
״����̬:<%if instr("|"&OnlineUserList&"","|"&Rs1("UserName")&"|")>0 then%>����<%else%>����<%end if%>
</td></tr></table>



</td><td><table cellSpacing=0 cellPadding=0 border=0 width=100% height=100%><tr>
		<td colspan=2> 
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/contact.gif) href=Profile.asp?UserName=<%=Rs("UserName")%> title='�鿴<%=Rs("UserName")%>�ĸ�������'>��Ϣ</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/message.gif) style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','','width=320,height=170')" title='���Ͷ�ѶϢ��<%=Rs("UserName")%>'>��Ѷ</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/mail.gif) href=Mailto:<%=Rs1("UserMail")%> title='���͵��ʸ� <%=Rs("UserName")%>'>����</a>
<%if Rs1("userhome")<>"" and Rs1("userhome")<>"http://" then%>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/home.gif) target=_blank href='<%=Rs1("userhome")%>' title='���� <%=Rs1("userhome")%> ����ҳ'>��ҳ</a>
<%end if%>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/Friend.gif) href='Friend.asp?menu=add&UserName=<%=Rs("UserName")%>' title='�� <%=Rs("UserName")%> �������'>����</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/finds.gif) href='ShowBBS.asp?menu=5&UserName=<%=Rs("UserName")%>' title='����<%=Rs("UserName")%>����������������'>����</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/Quote.gif) href='ReTopic.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>&quote=1' title='���ûظ��������'>����</a>
<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) href='ReTopic.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='�ظ�����'>�ظ�</a>
</td><td align=right>No.<font color=red><b><%=i+(PageCount-1)*PageSetup%></b></font></td></tr><tr vAlign=top><td colSpan=3><hr width=100% color=#777777 SIZE=1></td></tr>
<tr vAlign=top><td colSpan=3 height=100% style=word-break:break-all>

<%if instr("|"&SiteSettings("BannedUserPost")&"|"&Request.Cookies("BadUserList")&"|","|"&Rs("UserName")&"|")>0 then%>
==============================<br>������<font color=RED>���û������ѱ����ˡ�����</font><br>==============================
<%else%>
<b><%=Rs("Subject")%></b><br><br>
<div class="markdown-content"><%=Rs("content")%></div>
<%end if%></td></tr><tr vAlign=top><td colSpan=3 align=right>
<%if Rs1("UserSign")<>"" and Request.Cookies("DisabledShowSign")="" then%>
��������������������<br><%=YbbEncode(Rs1("UserSign"))%>
<%end if%>

</td></tr><tr vAlign=top><td colSpan=3><hr width=100% color=#777777 SIZE=1></td></tr><tr vAlign=top><td><a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/edit.gif) href='EditPost.asp?ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='�༭����'>�༭</a> <a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/del.gif) href='Manage.asp?menu=IsDel&ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>' title='ɾ������'>ɾ��</a></td><td valign=bottom><img src=images/Posttime.gif> ����ʱ�䣺<%=Rs("Posttime")%>��</td>
			<td valign=bottom align="right"><img src=images/ip.gif> IP��
<%if SiteSettings("DisplayPostIP")=1 then%>
<%=Rs("PostIP")%>
<%else%>
<a href=Manage.asp?menu=lookip&ThreadID=<%=Rs("ThreadID")%>&PostID=<%=Rs("id")%>>�Ѽ�¼</a>
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
<a href="MyFavorites.asp?menu=add&url=Topic&name=<%=ThreadID%>">�ղ�����</a> | <a href="MyFavorites.asp?menu=Del&url=Topic&name=<%=ThreadID%>">ȡ���ղ�</a> | <a href="#">����ҳ��</a>&nbsp;</td></tr></table>



<%if CookieUserName<>empty then%>
<form name="yuziform" method="POST" action="ReTopic.asp" onSubmit="return CheckForm(this);">
<input type="hidden" value="<%=ThreadID%>" name="ThreadID">
<input type="hidden" value="Re:<%=Topic%>" name="Subject">
<input name="content" type="hidden">

<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align=center>
<tr class="a1"><td width="20%" height="10"><b>���ٻظ�����</b></td><td width="80%"><b><%=act%></b></td></tr>

<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
<tr class="a3"><td width="20%" height="3"><b>��֤��</b></td><td width="80%">
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="��֤��,�������?����ˢ����֤��" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></td></tr>
<%end if%>

<tr class="a3">
	<td width="20%" rowspan="2" valign="top" height="100%">
<TABLE cellSpacing=0 cellPadding=0 width=100% align=Left border=0 height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><br><B>��������</B><BR>
��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��
</TD></TR>

<TR>
<TD vAlign=bottom align=Left width=100% class=a3>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> ����YBB����</label>
</TD></TR>
</TABLE>

	<td width="80%" height=200>

<SCRIPT src="inc/Post.js"></SCRIPT></td></tr><tr class=a3>
			<td><input type="submit" value="Ctrl+Enter �ظ�����" name=EditSubmit>������<input type="reset" name="reset" value=" �� �� "></td></tr></table>

</form>

<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0"><tr><td align="middle">����ѡ��: <%
response.write "<a href=Manage.asp?menu=MoveNew&ThreadID="&ThreadID&">��ǰ����</a> | "
if IsLocked=1 then
response.write "<a href=Manage.asp?menu=DelIsLocked&ThreadID="&ThreadID&">��������</a>"
else
response.write "<a href=Manage.asp?menu=IsLocked&ThreadID="&ThreadID&">��������</a>"
end if
response.write " | "
if IsTop=2 then
response.write "<a href=Manage.asp?menu=untop&ThreadID="&ThreadID&">ȡ���̶ܹ�</a>"
else
response.write "<a href=Manage.asp?menu=top&ThreadID="&ThreadID&">�����̶ܹ�</a>"
end if
response.write " | "
if IsTop=1 then
response.write "<a href=Manage.asp?menu=DelIsTop&ThreadID="&ThreadID&">ȡ���̶�</a>"
else
response.write "<a href=Manage.asp?menu=IsTop&ThreadID="&ThreadID&">����̶�</a>"
end if
response.write " | "
if IsGood=1 then
response.write "<a href=Manage.asp?menu=DelIsGood&ThreadID="&ThreadID&">ȡ��������</a>"
else
response.write "<a href=Manage.asp?menu=IsGood&ThreadID="&ThreadID&">��Ϊ������</a>"
end if
%>
| <a href="Move.asp?ThreadID=<%=ThreadID%>">�ƶ�����</a>
| <a title="�޸����ӵĻظ���" href="Manage.asp?menu=Fix&ThreadID=<%=ThreadID%>">�޸�����</a>

</td></tr></table><%end if%> 

<!-- #include file="inc/line.asp" -->
<%htmlend%>