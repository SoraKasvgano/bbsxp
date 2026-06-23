<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
ForumID=RequestInt("ForumID")

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumLogo=SafeUrl(Rs("ForumLogo"))
followid=Rs("followid")
ForumName=Rs("ForumName")
moderated=Rs("moderated")
Rs.close

if ""&moderated&""="" then moderated="|"
moderated=split(moderated,"|")

if membercode<4 and moderated(0)<>CookieUserName then
error("<li>����Ȩ�޲���<li>ֻ�������� <font color=red>"&moderated(0)&"</font> ����������������Ա��ӵ�д�Ȩ��")
end if



select case Request("menu")
case ""
error2("��ѡ����Ҫ��������Ŀ")

case "BatchRecycle"
for each ho in request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set IsDel=0,lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&ho&" and IsDel=1 and ForumID="&ForumID&"")
next
Log("��ԭ����վ�ڵ����ӣ�����ID��"&Request.form("ThreadID")&"")
error2("�ɹ���ԭ����վ�ڵ�����")

case "BatchCensorship"
for each ho in request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set IsDel=0 where id="&ho&" and IsDel=1 and ForumID="&ForumID&"")
next
Log("����ͨ����飬����ID��"&Request.form("ThreadID")&"")
error2("�����Ѿ��ɹ�ͨ�����")

case "BatchDel"
IsDel=int(Request.form("IsDel"))
for each ho in Request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set IsDel="&IsDel&",lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("����ɾ��������ID��"&Request.form("ThreadID")&"")
error2("�����ɹ�")

case "BatchGOOD"
IsGOOD=int(Request.form("IsGOOD"))
for each ho in Request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set IsGOOD="&IsGOOD&",lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("��������������ID��"&Request.form("ThreadID")&"")
error2("�����ɹ�")

case "BatchLocked"
IsLocked=SafeLongValue(Request.form("IsLocked"),0)
for each ho in Request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set IsLocked="&IsLocked&",lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("��������������ID��"&Request.form("ThreadID")&"")
error2("����1"&IsLocked&"�ɹ�")

case "BatchSpecialTopic"
for each ho in Request.form("ThreadID")
ho=SafeLongValue(ho,0)
Conn.execute("update [BBSXP_Threads] set SpecialTopic='"&SqlString(HTMLEncode(Request.form("SpecialTopic")))&"',lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&ho&" and ForumID="&ForumID&"")
next
Log("����ר�⣬����ID��"&Request.form("ThreadID")&"")
error2("�����ɹ�")

case "BatchMoveTopic"
AimForumID=SafeLongValue(Request.form("AimForumID"),0)
if AimForumID=0 then error("<li>��û��ѡ��Ҫ�������ƶ��ĸ���̳")
for each ho in Request.form("ThreadID")
ho=SafeLongValue(ho,0)
if Conn.Execute("Select ForumPass From [BBSXP_Forums] where id="&AimForumID&"")(0)=4 then error("<li>Ŀ����̳Ϊ��Ȩ����״̬")
Conn.execute("update [BBSXP_Threads] set ForumID="&AimForumID&",IsTop=0,IsGood=0,IsLocked=0 where id="&ho&" and ForumID="&ForumID&"")
next
Log("�����ƶ�������ID��"&Request.form("ThreadID")&"")
error2("�����ɹ�")


case "Fix"
allarticle=Conn.execute("Select count(ID) from [BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&"")(0)
if allarticle>0 then
allrearticle=Conn.execute("Select sum(replies) from [BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&"")(0)
else
allrearticle=0
end if
Conn.execute("update [BBSXP_Forums] set ForumThreads="&allarticle&",ForumPosts="&allarticle+allrearticle&" where ID="&ForumID&"")
error2("�޸���̳ͳ�����ݳɹ�")


case "ForumDataUp"
ForumName=HTMLEncode(Request.Form("ForumName"))
TolSpecialTopic=HTMLEncode(Request.Form("TolSpecialTopic"))
ForumIcon=HTMLEncode(Request.Form("ForumIcon"))
ForumLogo=HTMLEncode(SafeUrl(Request.Form("ForumLogo")))
moderated=HTMLEncode(Request.Form("moderated"))
ForumIntro=HTMLEncode(Request.Form("ForumIntro"))
ForumRules=HTMLEncode(Request.Form("ForumRules"))
if ForumName="" then error("<li>��������̳����")
if Len(ForumName)>30 then  error("<li>��̳���Ʋ��ܴ��� 30 ���ַ�")
if Len(ForumIntro)>255 then  error("<li>��̳��鲻�ܴ��� 255 ���ַ�")
if instr(TolSpecialTopic,";") > 0 then error("<li>ר���в��ܺ����������")

master=split(moderated,"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&master(i)&"'" ).eof and master(i)<>"" Then error("<li>"&master(i)&"���û����ϻ�δע��")
next

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Rs.Open sql,Conn,1,3
Rs("ForumName")=ForumName
Rs("moderated")=moderated
Rs("TolSpecialTopic")=TolSpecialTopic
Rs("ForumIcon")=ForumIcon
Rs("ForumLogo")=ForumLogo
Rs("ForumIntro")=ForumIntro
Rs("ForumRules")=ForumRules
Rs.update
Rs.close
Log("������̳��ID:"&ForumID&"������Ϣ��")
Message="<li>���³ɹ���<li><a href=ShowForum.asp?ForumID="&ForumID&">������̳</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?ForumID="&ForumID&">")

case "ForumData"
sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
ForumIntro=replace(""&Rs("ForumIntro")&"","<br>",vbCrlf)
ForumRules=replace(""&Rs("ForumRules")&"","<br>",vbCrlf)
%>

<script>
if ("<%=SafeJsString(SafeUrl(Rs("ForumLogo")))%>"!=''){Logo.innerHTML="<img border=0 src=<%=SafeUrl(Rs("ForumLogo"))%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(Rs("followid"))%><%=ForumTreeList%> <a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=Rs("ForumName")%></a> �� ������̳</td>
		</tr>
	</table><br>
<table border="0" width="100%">
	<tr>
		<td valign="top" width="20%">



<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
	<tr class=a1>
		<td align="center">��̳��Ϣ </td>
	</tr>
	<tr class=a3>
		<td>��������<%=Rs("ForumToday")%></td>
	</tr>
	<tr class=a3>
		<td>��������<%=Rs("ForumThreads")%> </td>
	</tr>
	<tr class=a3>
		<td>��������<%=Rs("ForumPosts")%> </td>
	</tr>
	<tr class=a1>
		<td align="center">����ѡ��</td>
	</tr>
	<tr class=a3>
		<td><a href="ShowForum.asp?ForumID=<%=ForumID%>&checkbox=1">������������</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Censorship&ForumID=<%=ForumID%>&checkbox=1">�������</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Recycle&ForumID=<%=ForumID%>&checkbox=1">�� �� վ</a></td>
	</tr>
	<tr class=a3>
		<td><a href="ForumManage.asp?menu=Fix&ForumID=<%=ForumID%>">�޸���̳ͳ������</a></td>
	</tr>
</table>


		</td>
		<td align="center">


<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
<tr class=a1>
<td height="20" align="center" colspan="2"><b>��̳����</b></td>
</tr>
<form name="form2" method="POST" action="?">
<input type=hidden name=menu value="ForumDataUp">
<input type=hidden name=ForumID value="<%=ForumID%>">
<tr class=a4>
<td align="right" valign="middle" width="20%">��̳���ƣ�</td>
<td align="Left" valign="middle" width="78%">
<input type="text" name="ForumName" size="30" maxlength="12" value="<%=Rs("ForumName")%>">
</td>
</tr>
<tr class=a3>
<td align="right" valign="middle" width="20%">��̳������</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="moderated" value="<%=Rs("moderated")%>">
������������á�|���������磺yuzi|ԣԣ
</td>
</tr>
<tr class=a4>
<td align="right" valign="middle" width="20%">����ר�⣺</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="TolSpecialTopic" value="<%=HTMLEncode(""&Rs("TolSpecialTopic")&"")%>">
�������á�|���������磺ԭ��|ת��|��ͼ</td>
</tr>
<tr class=a3>
<td align="right" width="20%">��̳���ܣ�</td>
<td align="Left" valign="middle" width="78%">
<textarea name="ForumIntro" rows="4" cols="50"><%=ForumIntro%></textarea>&nbsp;
</td>
<tr class=a4>
<td align="right" width="20%">��̳����</td>
<td align="Left" valign="middle" width="78%">
<textarea name="ForumRules" rows="4" cols="50"><%=ForumRules%></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td align="right" valign="middle" width="20%">Сͼ��URL��</td>
<td align="Left" valign="middle" width="78%">
<input size="30" name="ForumIcon" value="<%=Rs("ForumIcon")%>">����ʾ��������ҳ��̳�����ұ�</td>
</tr>
<tr class=a4>
<td align="right" valign="bottom" width="20%">��ͼ��URL��</td>
<td align="Left" valign="bottom" width="78%">
<input size="30" name="ForumLogo" value="<%=HTMLEncode(SafeUrl(Rs("ForumLogo")))%>">����ʾ����̳���Ͻ�</td>
</tr>
<tr class=a3>
<td align="right" valign="bottom" width="98%" colspan="2"><input type="submit" value=" �� �� &gt;&gt;�� һ �� "></td>
</tr>
</table>
</form>




		</td>
	</tr>
</table>

<%
Rs.close
htmlend


end select

%>
<script>
if ("<%=SafeJsString(SafeUrl(ForumLogo))%>"!=''){Logo.innerHTML="<img border=0 src=<%=SafeUrl(ForumLogo)%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <%ForumTree(followid)%><%=ForumTreeList%> <a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a> ��
			<a href="?menu=<%=HTMLEncode(Request("menu"))%>&ForumID=<%=ForumID%>&checkbox=1"><span id=menu>����վ</span></a></td>
		</tr>
	</table><br><form method="POST" action="?"><input type=hidden name=ForumID value=<%=ForumID%>>
<%

if Request("menu")="Recycle" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime<>lasttime and ForumID="&ForumID&" order by lasttime Desc"
response.write "<script>menu.innerText='����վ'</script>"
elseif Request("menu")="Censorship" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime=lasttime and ForumID="&ForumID&" order by lasttime Desc"
response.write "<script>menu.innerText='�����'</script>"
end if
Rs.Open sql,Conn,1

PageSetup=20 '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��


%>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
		<%

i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

ShowThread()
Rs.MoveNext
loop
Rs.Close
%>

</table>

<table cellspacing="0" cellpadding="1" width="100%" align="center" border="0">
	<tr height="25">
		<td width="100%" height="2">
		<table cellspacing="0" cellpadding="3" width="100%" border="0">
			<tr>
				<td height="2" valign="top"><input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">ȫѡ
<%
if Request("menu")="Recycle" then
%><input type="radio" value="BatchRecycle" name=menu checked>��ԭ����<%
elseif Request("menu")="Censorship" then
%><input type="radio" value="BatchCensorship" name=menu checked>ͨ����� <input type="radio" value="BatchDel" name=menu><input type=hidden name=IsDel value=1>ɾ������<%
end if
%>

<input onclick="checkclick('��ȷ��ִ�б��β���?');" type="submit" value=" ִ �� ">




</form></td>

				<td align="right" height="2" valign="top">

<%ShowPage()%></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
<%


htmlend
%>