<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&SqlString(CookieUserName)&"'" ).eof Then error("<li>�����ؽ�<a href=Prison.asp>����</a>")
ForumID=RequestInt(Request("ForumID"))

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumName=Rs("ForumName")
ForumLogo=SafeUrl(Rs("ForumLogo"))
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
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>��֤�����"
end if

color=HTMLEncode(Request.Form("color"))
icon=RequestInt("icon")
Subject=HTMLEncode(Request.Form("Subject"))
Content=ContentEncode(Request.Form("Content"))
if Request.Form("DisableYBBCode")<>1 then Content=YbbEncode(Content)

if Len(Subject)<2 then Message=Message&"<li>�������ⲻ��С�� 2 �ַ�"
if Len(content)<2 then Message=Message&"<li>�������ݲ���С�� 2 �ַ�"

if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
Subject=ReplaceText(Subject,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if



'''''''''''''''''''''''''''''''

if Request.Form("Vote")<>"" then
Vote=Request("Vote")
if instr(Vote,"|") > 0 then error("<li>ͶƱѡ���в��ܺ��С�|���ַ�")
pollTopic=split(Vote,chr(13)&chr(10))
j=0
for i = 0 to ubound(pollTopic)
if not (pollTopic(i)="" or pollTopic(i)=" ") then
allpollTopic=""&allpollTopic&""&pollTopic(i)&"|"
j=j+1
end if
next

if j<SiteSettings("MinVoteOptions") or j>SiteSettings("MaxVoteOptions") then error("<li>ͶƱѡ������� "&SiteSettings("MinVoteOptions")&" ��<li>ͶƱѡ��� "&SiteSettings("MaxVoteOptions")&" ��")


for y = 1 to j
Votenum=""&Votenum&"0|"
next

end if
'''''''''''''''''''''''''''''''

if Message<>"" then error(""&Message&"")



sql="select * from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Rs.Open sql,Conn,1,3

StopPostTime=int(DateDiff("s",Rs("UserLandTime"),Now()))
if StopPostTime < int(SiteSettings("DuplicatePostIntervalInMinutes")) then Message=Message&"<li>��̳����һ�������η������������� "&SiteSettings("DuplicatePostIntervalInMinutes")&" �룡<li>�������ٵȴ� "&SiteSettings("DuplicatePostIntervalInMinutes")-StopPostTime&" �룡"

StopPostTime=int(DateDiff("s",Rs("UserRegTime"),Now()))
if StopPostTime < int(SiteSettings("RegUserTimePost")) then Message=Message&"<li>��ע���û�����ȴ� "&SiteSettings("RegUserTimePost")&" �����ܷ�����<li>�������ٵȴ� "&SiteSettings("RegUserTimePost")-StopPostTime&" �룡"

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
if Request("SpecialTopic")<>"" then Rs("SpecialTopic")=HTMLEncode(Request("SpecialTopic"))
if Request("icon")<>"" then Rs("icon")=HTMLEncode(Request("icon"))
if Request("Vote")<>"" then Rs("isVote")=1
if Request("IsLocked")=1 then Rs("IsLocked")=1
if ForumPass=5 then Rs("IsDel")=1
Rs.update
ID=Rs("ID")
Rs.close

if Request.Form("Vote")<>"" then
multiplicity=SafeLongValue(Request.Form("multiplicity"),0)
Expiry=now()+SafeLongValue(Request.Form("Expiry"),0)
Conn.Execute("insert into [BBSXP_Vote] (ThreadID,Type,Items,Result,Expiry) values ('"&ID&"','"&multiplicity&"','"&SqlString(HTMLEncode(allpollTopic))&"','"&SqlString(Votenum)&"','"&SqlString(Expiry)&"')")
end if

if Request.Form("UpFileID")<>"" then
UpFileID=split(Request.form("UpFileID"),",")
for i = 0 to ubound(UpFileID)-1
UpFileItem=SafeLongValue(UpFileID(i),0)
if UpFileItem>0 then Conn.execute("update [BBSXP_PostAttachments] set ThreadID="&ID&",Description='"&SqlString(Subject)&"' where id="&UpFileItem&" and ThreadID=0")
next
end if

Conn.Execute("insert into [BBSXP_Posts"&SiteSettings("DefaultPostsName")&"] (ThreadID,IsTopic,UserName,Subject,content,Postip) values ('"&ID&"','1','"&SqlString(CookieUserName)&"','"&SqlString(Subject)&"','"&SqlString(content)&"','"&SqlString(Request.ServerVariables("REMOTE_ADDR"))&"')")
Conn.execute("update [BBSXP_Forums] set lastTopic='"&SqlString("<a href=ShowPost.asp?ThreadID="&id&">"&Left(HTMLEncode(Request.Form("Subject")),15)&"</a>")&"',lastname='"&SqlString(CookieUserName)&"',lasttime="&SqlNowString&",ForumToday=ForumToday+1,ForumThreads=ForumThreads+1,ForumPosts=ForumPosts+1 where id="&ForumID&"")
Conn.execute("update [BBSXP_Statistics_Site] set TodayPost=TodayPost+1,TotalPost=TotalPost+1,TotalThread=TotalThread+1")

Session("VerifyCode")=""

if ForumPass=5 then
EnableCensorship="������̳��������ƶȣ���������������Ҫ�ȴ����������ʾ��"
else
EnableCensorship="<a href=ShowPost.asp?ThreadID="&id&">��������</a>"
end if

Message="<li>�����ⷢ���ɹ�<li>"&EnableCensorship&"<li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a><li><a href=Default.asp>������̳��ҳ</a>"
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
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� ��������</td>
		</tr>
	</table><br>


<TABLE cellSpacing=1 cellPadding=6 width=100% border=0 class=a2 align=center>
<form name="yuziform" method="post" onSubmit="return CheckForm(this);">
<input name="content" type="hidden"><input name="UpFileID" type="hidden">
<input type=hidden name=ForumID value=<%=ForumID%>>
<TR>
<TD vAlign=Left colSpan=2 height=25 class=a1><b>��������</b></TD></TR>
<%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>
	<tr>
<TD class=a3 height=6><b>��֤��</b></TD>
<TD class=a3 height=6>
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="��֤��,�������?����ˢ����֤��" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></TD>
	</tr>
<%end if%>
	
<TR>
<TD class=a3 width="180"><B>���±��� </B> 
<%
if TolSpecialTopic<>empty then
response.write "<SELECT name=SpecialTopic size=1><OPTION value='' selected>&nbsp;ר��</OPTION>"
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
<TD vAlign=top align=Left class=a4 height=23><B>���ı���</B></TD>
<TD class=a4>
<script>
for(i=1;i<=12;i++) {
document.write("<INPUT type=radio value="+i+" name=icon><IMG src=images/brow/"+i+".gif>��")
}
</script>

 </TD></TR>

<TR>
<TD vAlign=top class=a3>
<TABLE cellSpacing=0 cellPadding=0 align=Left border=0 width=100% height="100%">

<TR>
<TD vAlign=top align=Left width=100% class=a3><BR><B>��������</B><BR>
��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<BR>
<BR><span id=UpFile></span>


</TD></TR>

<TR><TD valign="bottom">
<INPUT id=LockMyPost name=IsLocked type=checkbox value=1><label for=LockMyPost> �����������������ظ�</label><br>
<INPUT id=DisableYBBCode name=DisableYBBCode type=checkbox value=1><label for=DisableYBBCode> ����YBB����</label><br>
<INPUT id=advcheck name=advShow type=checkbox value=1 onclick=ShowADv()><label for=advcheck> ��ʾͶƱѡ��</label>

</TD></TR>

</TABLE></TD>

<TD class=a3 height=250>

<SCRIPT src="inc/Post.js"></SCRIPT>

</TD></TR>


<TR id=adv style=DISPLAY:none>
<TD vAlign=top align=Left class=a4>


<FONT color=000000><B>ͶƱ��Ŀ</B><BR>
ÿ��һ��ͶƱ��Ŀ<BR>
�������� <INPUT maxLength=3 size=2 name=Expiry value="7" onkeyup=if(isNaN(this.value))this.value=''> ��<br>

<INPUT type=radio CHECKED value=0 name=multiplicity id=multiplicity>
<label for=multiplicity>��ѡͶƱ</label>
<BR><INPUT type=radio value=1 name=multiplicity id=multiplicity_1> <label for=multiplicity_1>��ѡͶƱ</label></FONT> 
</TD>
<TD class=a4>
<TEXTAREA name=Vote rows=5 style="width:100%"></TEXTAREA>
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
<INPUT type=submit value=���������� name=EditSubmit>&nbsp; <INPUT type=reset value=" �� �� "></TD></TR></FORM>
</TABLE>





<%
htmlend
%>