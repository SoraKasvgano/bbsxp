<!-- #include file="Setup.asp" --><%
top

ForumID=RequestInt("ForumID")
order=HTMLEncode(Request("order"))
SpecialTopic=HTMLEncode(Request("SpecialTopic"))
SortOrder=RequestInt("SortOrder")
TimeLimit=RequestInt("TimeLimit")
if Len(order)>10 then error("<li>�Ƿ�����")
if Len(SpecialTopic)>20 then error("<li>ר������̫��")

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
if Rs.eof then error"<li>����̳�Ѿ���ɾ��"
ForumName=Rs("ForumName")
ForumThreads=Rs("ForumThreads")
moderated=Rs("moderated")
ForumLogo=SafeUrl(Rs("ForumLogo"))
followid=Rs("followid")
ForumHide=Rs("ForumHide")
ForumPass=Rs("ForumPass")
ForumRules=YbbEncode(Rs("ForumRules"))
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
ForumThreads=Rs("ForumThreads")
TolSpecialTopic=Rs("TolSpecialTopic")
Rs.close

%>
<!-- #include file="inc/Validate.asp" -->
<meta http-equiv="refresh" content="300">
<script>if ("<%=ForumLogo%>"!=''){Logo.innerHTML="<img border=0 src=<%=ForumLogo%> onload='javascript:if(this.height>60)this.height=60;'>"}</script>

<%if Request("checkbox")=1 then
BBSList(0)
%>
<form method="POST" action="ForumManage.asp">
<input type=hidden name=ForumID value=<%=ForumID%>>
<%end if%>


<title><%=ForumName%> - Powered By BBSXP</title>
<div align="center">
<table border="0" width="100%" cellspacing="1" class="a2">
	<tr class="a3">
		<td colspan="5">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" height="25">
			<tr>
				<td height="18">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
				�� <%ForumTree(followid)%><%=ForumTreeList%>
				<a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></td>
				<td height="18" align="right"><img src="images/jt.gif">
				<a href="ForumManage.asp?menu=ForumData&ForumID=<%=ForumID%>">�߼�����</a> </td>
			</tr>
		</table>
		</td>
	</tr>
	<%
IsShowForum=1
sql="Select * From [BBSXP_Forums] where followid="&ForumID&" and ForumHide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
if not Rs1.eof then
do while not Rs1.eof
ShowForum()
Rs1.Movenext
loop
IsShowForum=0
end if
Set Rs1 = Nothing
%>



</table>
</div>
<br>
<%if ForumRules<>"" then%>
<div class=a3 style=" PADDING-TOP: 10px; PADDING-BOTTOM: 10px;PADDING-LEFT: 10px;  PADDING-RIGHT: 10px;  BORDER-RIGHT:#ccc 1px dotted; BORDER-TOP:#ccc 1px dotted; BORDER-LEFT:#ccc 1px dotted; BORDER-BOTTOM:#ccc 1px dotted;">
<strong><font color="#ff0000">�� �� �� �� ��</font></strong><br><%=ForumRules%></div><br>
<%end if%>

<!-- #include file="inc/line.asp" --><%


if IsShowForum=1 or SiteSettings("SortShowForum")=1 then

ForumIDOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where ForumID="&ForumID&"")(0)
regForumIDOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where ForumID="&ForumID&" and UserName<>''")(0)
%><table cellspacing="1" cellpadding="0" width="100%" align="center" border="0" class="a2">
	<tr>
		<td width="93%" height="25" class="a1">��<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=ForumID%>)"> 
		Ŀǰ��̳������ <b><%=Onlinemany%></b> �ˣ�������̳���� <b><%=ForumIDOnline%></b> �����ߡ�����ע���û� 
		<b><%=regForumIDOnline%></b> �ˣ��ÿ� <b><%=ForumIDOnline-regForumIDOnline%></b> 
		�ˡ�</td>
		<td align="middle" width="7%" height="25" class="a1">
		<a href="javascript:this.location.reload()">
		<img src="images/refresh.gif" border="0"></a></td>
	</tr>
	<tr height="25" style="display:none" id="follow0">
		<td id="followTd0" align="Left" class="a4" width="94%" colspan="5">��Loading...</td>
	</tr>
	</tr>
</table>
<br>
<table height="30" cellspacing="3" cellpadding="0" width="100%" align="center" border="0">
	<tr>
		<td align="Left" width="20%">
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="NewTopic.asp?ForumID=<%=ForumID%>">
		����������</a> </td>
		<td align="right" width="80%">
		<img src="images/Showdigest.gif">
		<a onmouseover="showmenu(event,'&lt;div class=menuitems&gt;&lt;a href=MyFavorites.asp?menu=add&amp;url=forum&amp;name=<%=ForumID%>&gt;�ղ���̳&lt;/a&gt;&lt;/div&gt;&lt;div class=menuitems&gt;&lt;a href=MyFavorites.asp?menu=Del&amp;url=forum&amp;name=<%=ForumID%>&gt;ȡ���ղ�&lt;/a&gt;&lt;/div&gt;')" style="cursor:default">
		��̳�ղ�</a>
<%
if moderated<>empty then
filtrate=split(moderated,"|")
for i = 0 to ubound(filtrate)
ModeratedList=ModeratedList&"<div class=menuitems><a href=Profile.asp?UserName="&filtrate(i)&">"&filtrate(i)&"</a></div>"
next
%><img src=images/team.gif> <a onmouseover="showmenu(event,'<%=ModeratedList%>')" style=cursor:default>
		��̳����</a>
<%end if%> <a href="Rss.asp?ForumID=<%=ForumID%>"><img src="images/rss_button.gif" border="0" alt="RSS ���ĵ�ǰ��̳"></a>
		</td>
	</tr>
</table>



<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
	<%
if TolSpecialTopic<>empty then
response.write "<tr height=25 class=a3><td width=100% colSpan=7>ר�⣺"
filtrate=split(TolSpecialTopic,"|")
for i = 0 to ubound(filtrate)
response.write "<font face='Old English Text MT'><b>"&i+1&"</b></font>[<a href='ShowForum.asp?ForumID="&ForumID&"&SpecialTopic="&filtrate(i)&"'>"&filtrate(i)&"</a>] "
TolSpecialTopicOptionList=TolSpecialTopicOptionList&"<option value="&filtrate(i)&">"&filtrate(i)&"</option>"
next
response.write "</td></tr>"
end if
%>
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
	<%
if TimeLimit<>"" then SQLTimeLimit="and lasttime>"&SqlNowString&"-"&int(TimeLimit)&""
if SpecialTopic<>"" then SQLSpecialTopic="and SpecialTopic='"&SpecialTopic&"'"


if order="" then order="lasttime"
if SortOrder="1" then
SqlSortOrder=""
else
SqlSortOrder="Desc"
end if

topsql="[BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&" "&SQLSpecialTopic&" "&SQLTimeLimit&" or IsTop=2"

if Request("TimeLimit")<>"" or Request("SpecialTopic")<>"" then
TotalCount=conn.Execute("Select count(ID) From "&topsql&" ")(0) '��ȡ��������
else
TotalCount=ForumThreads  '��ȡ��������
end if

PageSetup=SiteSettings("ThreadsPerPage") '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = cint(Request.QueryString("PageIndex")) '��ȡ��ǰҳ
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


if PageCount<2 then
sql="select top "&pagesetup&" * from "&topsql&" order by IsTop Desc,"&order&" "&SqlSortOrder&""
Set Rs=Conn.Execute(sql)
else
sql=""&topsql&" order by IsTop Desc,"&order&" "&SqlSortOrder&""
rs.Open sql,Conn,1
end if


if TotalPage>1 then RS.Move (PageCount-1) * pagesetup

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1

ShowThread()

Rs.MoveNext
loop
Rs.Close
if Request("checkbox")=1 then


%><tr height="25" id="TableTitleLink" class="a3">
<td colspan="7"><input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">ȫѡ
<input type="radio" value="BatchDel" name=menu><select name=IsDel>
<option value="1">ɾ������</option>
<option value="0">ȡ��ɾ��</option>
</select>&nbsp;

<input type="radio" value="BatchGOOD" name=menu><select name=IsGOOD>
<option value="1">���뾫��</option>
<option value="0">ȡ������</option>
</select>&nbsp;

<input type="radio" value="BatchLocked" name=menu><select name=IsLocked>
<option value="1">����</option>
<option value="0">����</option>
</select>

<input type="radio" value="BatchSpecialTopic" name=menu><select name=SpecialTopic>
<option value="">ȡ��ר��</option>
<option selected value="">����ר��</option>
<%=TolSpecialTopicOptionList%>
</select>

<input type="radio" value="BatchMoveTopic" name=menu><select name=AimForumID>
<option selected value="">�ƶ���������̳</option>
<%=ForumsList%>
</select>&nbsp;


<input onclick="checkclick('��ȷ��ִ�б��β���?');" type="submit" value=" ִ �� ">
</td></form>
</tr>

<%end if%>
</table>

<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
	<tr>
		<td>
		<a onmousedown="ToggleMenuOnOff('ForumOption')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/ForumSettings.gif)" href="#ForumOption">
		ѡ��</a>
		<a onmousedown="ToggleMenuOnOff('ForumSearch')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/finds.gif)" href="#ForumSearch">
		����</a>
		<div id="ForumSearch" style="position:absolute;display:none;">
			<form name="form" action="Search.asp?menu=ok&ForumID=<%=ForumID%>&Search=key&sessionid=<%=session.sessionid%>" method="POST">
				<input name="content" size="20" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')">
				<input type="submit" value="����" id="btnadd" disabled>
			</form>
		</div>
		</td>
		<td align="right">
<%ShowPage()%>
</td>
	</tr>
	<tr id="ForumOption" style="display:none;">
		<td valign="top" colspan="2">
		<form name="form" action="ShowForum.asp?ForumID=<%=ForumID%>" method="POST">
			�������<select name="order">
			<option value="">������ʱ��</option>
			<option value="id">���ⷢ��ʱ��</option>
			<option value="IsGood">��������</option>
			<option value="IsVote">ͶƱ����</option>
			<option value="Topic">����</option>
			<option value="UserName">����</option>
			<option value="Views">�����</option>
			<option value="Replies">�ظ���</option>
			</select> ���� <select name="SortOrder">
			<option value="0" selected>����</option>
			<option value="1">����</option>
			</select> ����<br>
			���ڹ��ˣ�<select name="TimeLimit">
			<option value="">��ʾ����</option>
			<option value="1">һ������</option>
			<option value="2">��������</option>
			<option value="3">��������</option>
			<option value="7">һ��������</option>
			<option value="14">����������</option>
			<option value="30">һ��������</option>
			<option value="60">����������</option>
			<option value="90">����������</option>
			<option value="180">��������</option>
			</select><br>
			<input type="submit" value=" Ӧ�� "></form></td>
	</tr>
</table>
<%
end if


htmlend
%>