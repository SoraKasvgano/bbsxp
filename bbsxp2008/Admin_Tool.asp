<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")
LinkID=RequestInt("LinkID")
AdvertisementID=RequestInt("AdvertisementID")
TimeLimit=RequestInt("TimeLimit")
RoleID=RequestInt("RoleID")

UserName=HTMLEncode(Request("UserName"))
Body=HTMLEncode(Request("Body"))


select case Request("menu")
	case "Message"
		Message
	case "BatchSendMail"
		BatchSendMail
	case "BatchSendMailok"
		BatchSendMailok
	case "Messageok"
		if TimeLimit="" then error2("��û��ѡ�����ڣ�")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where DateDiff("&SqlChar&"d"&SqlChar&",CreateTime,"&SqlNowString&") > "&TimeLimit&"")
		Alert("�Ѿ���"&TimeLimit&"����ǰ�Ķ�ѶϢɾ���ˣ�")
	case "DelMessageUser"
		if UserName="" then Alert("��û�������û�����")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&UserName&"'")
		Execute("update ["&TablePrefix&"PrivateMessages] Set IsSenderDelete=1 where SenderUserName='"&UserName&"'")
		Alert("�Ѿ���"&UserName&"�Ķ�ѶϢȫ��ɾ���ˣ�")
	case "DelMessagekey"
		key=HTMLEncode(Request("key"))
		if key="" then Alert("��û������ؼ��ʣ�")
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where Body like '%"&key&"%'")
		Alert("�Ѿ��������а��� "&key&" �Ķ�ѶϢɾ���ˣ�")
		
	case "UpStatic"
		UpdateStatic
	case "UpUserRanks"
		UpUserRanks
	case "upSiteSettingsok"
		upSiteSettingsok
	case "DelApplication"
		Application.contents.ReMoveAll()
		Alert("�Ѿ���������������е�application���棡")
		
	case "Link"
		Link
	case "Linkok"
		Linkok
	case "DelLink"
		Execute("Delete from ["&TablePrefix&"Links] where LinkID="&LinkID&"")
		UpdateApplication"Links","Select name,Logo,url,Intro From ["&TablePrefix&"Links] where SortOrder>0 order by SortOrder"	'���»���
		Response.redirect"?menu=Link"
	case "AD"
		Advertisement
	case "Adok"
		Adok
	case "DelAd"
		Execute("Delete from ["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"")
		UpdateApplication"Advertisements","select Body from ["&TablePrefix&"Advertisements] where DateDiff("&SqlChar&"d"&SqlChar&","&SqlNowString&",ExpireDate) > 0"
		Response.redirect"?menu=AD"
end select

Sub BatchSendMail

if SiteConfig("SelectMailMode")="" then Alert("��û�п��������ʼ����ܣ�")


%>

<form name="form" method="post" action="?menu=BatchSendMailok" onsubmit="return CheckForm(this);">
<input name="Body" type="hidden" />
<table cellspacing="1" cellpadding="5" width="90%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan=2>Ⱥ���ʼ�</td>
	</tr>
	<tr class="CommonListCell">
		<td align=Left width="12%">�ռ��ˣ���</td>
		<td width="85%"><select name="RoleID">
<%
		sql="Select * from ["&TablePrefix&"Roles] where RoleID > 0 order by RoleID"
		Set Rs=Execute(sql)
			Do While Not Rs.EOF
			Response.Write("<option value='"&Rs("RoleID")&"'>"&Rs("Name")&"</option>")
				Rs.MoveNext
			loop
		Rs.Close
%>
</select>    
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align=Left>���⣺</td>
		<td width="85%"><input size="60" name="Subject" /></td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top"> ���ݣ�</td>
		<td height="250"> <script type="text/javascript" src="Editor/Post.js"></script>
		</td>
	</tr>
   <tr class="CommonListCell">
		<td align=center colspan=2> <input type="submit" value=" �� �� " name=EditSubmit /> <input type="reset" value=" �� �� " /><br /></td>
	</tr>
</table>
</form>

<%
End Sub

Sub BatchSendMailok
if Request("Subject")="" then Alert("����д�ʼ�����")
if Request("Body")="" then Alert("����д�ʼ�����")
	if RoleID>0 then
		sql="Select UserEmail from ["&TablePrefix&"Users] where UserRoleID="&RoleID&""
	else
		sql="Select UserEmail from ["&TablePrefix&"Users]"
	end if
	Set Rs=Execute(sql)
		do while not Rs.eof
			SendMail Rs("UserEmail"),Request("Subject"),Request("Body")
			Rs.Movenext
		loop
	Set Rs = Nothing
	Response.Write("�ʼ�Ⱥ�����ͳɹ���")
End Sub

Sub Message
%>
���ݿ⹲ <%=Execute("Select count(MessageID) from ["&TablePrefix&"PrivateMessages]")(0)%> ����ѶϢ
<br /><br />
<table cellspacing="1" cellpadding="5" width="70%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center">����ɾ������Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<form method="POST" action="?menu=DelMessageUser"><td align="center">����ɾ�� <input size="13" name="UserName" /> �Ķ�ѶϢ <input type="submit" value="ȷ��" /></td></form>
	</tr>
	<tr class="CommonListCell">
		<form method="POST" action="?menu=DelMessagekey">
		<td align="center">����ɾ�����ݺ��� <input size="20" name="key" /> �Ķ�ѶϢ <input type="submit" value="ȷ��" />
		</td>
		</form>
	</tr>


	<tr class="CommonListCell">
		<form method="POST" action="?menu=Messageok"><td align="center">ɾ�� <select name="TimeLimit">
			<option value="30">30</option>
			<option value="60">60</option>
			<option value="90">90</option>
			</select> ����ǰ�Ķ�ѶϢ
<input type="submit" value="ȷ��" />

		</td></form>
	</tr>

</table>
</form>
<%
End Sub

Sub Link
	if LinkID>0 then
		sql="Select * From ["&TablePrefix&"Links] where LinkID="&LinkID&""
		Set Rs=Execute(sql)
		if not Rs.eof then
			name=Rs("name")
			url=Rs("url")
			Logo=Rs("Logo")
			Intro=Rs("Intro")
			SortOrder=Rs("SortOrder")
		end if
		Rs.Close
	end if
%>
<form action=?menu=Linkok method=Post>
<input type=hidden name=LinkID value=<%=LinkID%> />
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align=center>
		<tr class=CommonListTitle><td colspan="2">���� �������ӹ���</td></tr>
		<tr class="CommonListCell"><td align=right>��վ���ƣ�</td><td><input size=40 name=name value="<%=name%>" /></td></tr>
		<tr class="CommonListCell"><td align=right>��ַ URL��</td><td><input size=40 name=url <%if ""&URL&""<>"" then Response.write("value='"&URL&"'") else Response.write(" value='http://'") end if%> /></td></tr>
		<tr class="CommonListCell"><td align=right>ͼ�� URL��</td><td><input size=40 name=Logo <%if ""&Logo&""<>"" then Response.write("value='"&Logo&"'") else Response.write(" value='http://'") end if%> /></td></tr>
		<tr class="CommonListCell"><td align=right>��վ��飺</td><td><input size=40 name=Intro value="<%=Intro%>" /></td></tr>
		<tr class="CommonListCell"><td align=right>�š�����</td><td><input size=5 name=SortOrder value="<%if ""&SortOrder&""<>"" then Response.write(SortOrder) else response.write("1") end if%>" /> ��С������������,Ϊ��0�������ش���������</td></tr>
		<tr class="CommonListCell"><td colspan="2" align="center"><input type=submit value=" ȷ �� " /> <input type="reset" value=" �� �� " /></td></tr>
</table>
</form>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="2">���� ��������</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="5%"><img src="images/shareforum.gif" /></td>
		<td>
<%
	Rs.Open "["&TablePrefix&"Links] order by SortOrder",Conn
	do while not Rs.eof
		if ""&Rs("Logo")&""="" or Rs("Logo")="http://" then
			Link1=Link1+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=Link&LinkID="&Rs("LinkID")&">�༭</a></div><div class=menuitems><a href=?menu=DelLink&LinkID="&Rs("LinkID")&">ɾ��</a></div>')"&Chr(34)&" title='"&Rs("Intro")&"' href="&Rs("url")&" target=_blank>"&Rs("name")&"</a>����"
		else
			Link2=Link2+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=Link&LinkID="&Rs("LinkID")&">�༭</a></div><div class=menuitems><a href=?menu=DelLink&LinkID="&Rs("LinkID")&">ɾ��</a></div>')"&Chr(34)&" title='"&Rs("name")&""&chr(10)&""&Rs("Intro")&"' href="&Rs("url")&" target=_blank><img src="&Rs("Logo")&" border=0 width=88 height=31></a>����"
		end if
		Rs.Movenext
	loop
	Rs.close
%>
<%=Link1%>
<br /><br />
<%=Link2%>
		</td>
	</tr>
</table>
<%
End Sub


Sub Linkok
	if Request("url")="http://" or Request("url")="" then Alert("��̳URLû����д")
	Rs.Open "["&TablePrefix&"Links]where LinkID="&LinkID&"",Conn,1,3
	If Rs.eof Then Rs.addNew
		Rs("name")=Request("name")
		Rs("url")=Request("url")
		Rs("Logo")=Request("Logo")
		Rs("Intro")=Request("Intro")
		Rs("SortOrder")=Requestint("SortOrder")
	Rs.update
	Rs.close
	UpdateApplication"Links","Select name,Logo,url,Intro From ["&TablePrefix&"Links] where SortOrder>0 order by SortOrder"	'���»���
	Response.redirect"?menu=Link"
End Sub


Sub Advertisement
	If AdvertisementID>0 Then
		Set Rs=Execute("Select * From ["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"")
		If Not Rs.eof Then
			Body=Rs("Body")
			ExpireDate=Rs("ExpireDate")
		End If
		Rs.close
	End If
	Body=replace(""&Body&"","<br>",vbCrlf)
	
if Not IsDate(ExpireDate) then ExpireDate=Date()+30

%>
<script type="text/javascript" src="Utility/calendar.js"></script>
<form name="form" action=?menu=Adok method="Post" onsubmit="return CheckForm(this);">
<input type=hidden name=AdvertisementID value=<%=AdvertisementID%> />

<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align=center>
		<tr class=CommonListTitle><td colspan="2">���� ���������</td></tr>

		<tr class="CommonListCell">
			<td align=right width="100">�����룺<br />��֧��HTML��</td><td><textarea name="Body" cols="80" rows="8"><%=Body%></textarea></td>
		</tr>
		<tr class="CommonListCell">
			<td align=right>�������ڣ�</td><td><input onclick="showcalendar(event, this)" onfocus="showcalendar(event, this)" size=24 name=ExpireDate value="<%=ExpireDate%>" /></td>
		</tr>
		<tr class="CommonListCell"><td align="center" colspan="2"><input type=submit value=" ȷ �� " name="EditSubmit" /> <input type="reset" value=" �� �� " /></td></tr>
</table>
</form>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="4">���� ������</td>
	</tr>
	<tr class=CommonListHeader align=center>
		<td>���</td>
		<td width=50%>����</td>
		<td>��������</td>
		<td>����</td>
	</tr>
	<%
	Rs.Open "["&TablePrefix&"Advertisements] order by ExpireDate Desc",Conn
	do while not Rs.eof
	%>
	<tr class="CommonListCell">
		<td align=center><%=Rs("AdvertisementID")%></td>
		<td><%=Rs("Body")%></td>
		<td align=center><%=Rs("ExpireDate")%></td>
		<td align=center><a href="?menu=AD&AdvertisementID=<%=Rs("AdvertisementID")%>">�༭</a> | <a href="?menu=DelAd&AdvertisementID=<%=Rs("AdvertisementID")%>" onclick="return window.confirm('ȷ��ִ�д˲�����')">ɾ��</a></td>
	</tr>
<%
		Rs.Movenext
	loop
	Rs.close
%>
</table>
<%
End Sub

Sub Adok
	Body=Request("Body")
	ExpireDate=Request("ExpireDate")
	if Body="" then  Alert("�������û����д")

	if Not IsDate(ExpireDate) then Alert("�������������")
	Rs.Open "["&TablePrefix&"Advertisements] where AdvertisementID="&AdvertisementID&"",Conn,1,3
	If Rs.eof Then Rs.addNew
		Rs("Body")=Body
		Rs("ExpireDate")=ExpireDate
	Rs.update
	Rs.close
	UpdateApplication"Advertisements","select Body from ["&TablePrefix&"Advertisements] where DateDiff("&SqlChar&"d"&SqlChar&","&SqlNowString&",ExpireDate) > 0"
	Response.redirect"?menu=AD"
End Sub



Sub UpdateStatic
%>
<form action="?menu=UpUserRanks" method=post>
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle><td align="center">�����û��ȼ�</td></tr>
	<tr class="CommonListCell"><td align="center">ÿ��ѭ���������û�����<input type=text name=PerCircleNum value="1000" /></td></tr>
	<tr class="CommonListCell"><td align="center"><input type=submit value=" �� �� " /></td></tr>
</table>
</form><br />
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan="2">������̳����</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" colspan="2">
			�˲�����������̳���ϣ��޸���̳ͳ�Ƶ���Ϣ<br /><a href="?menu=upSiteSettingsok">������������̳ͳ������</a><br /><a href="?menu=DelApplication">����������ϵ�application����</a><br />
		</td>
	</tr>
</table>
<br />
<%
End Sub

Sub UpUserRanks
	PerCircleNum=RequestInt("PerCircleNum")
	i=RequestInt("i")
	TotalCount=RequestInt("TotalCount")
	j=0:IsEnd=false
	if PerCircleNum>0 then
		if i=0 or TotalCount=0 then TotalCount=Execute("select count(UserID) from ["&TablePrefix&"Users]")(0)/PerCircleNum
		Rs.Open "["&TablePrefix&"Users]",Conn,1,3
			Rs.move(PerCircleNum*i)
			do while j<PerCircleNum
				if Rs.eof then
					IsEnd=true
					exit do
				end if
				Rs("UserRank")=UpUserRank()
				j=j+1
				Rs.Update
				Rs.movenext
			loop
		Rs.Close
		
		if IsEnd then	'���������
			Pause 0,0,0
		else			'����ִ��ѭ�����²���
			Pause PerCircleNum,i+1,TotalCount
		end if
	end if
End Sub


Sub Pause(PerCircleNum,i,TotalCount)
	if PerCircleNum>0 then
		RefreshURL="?menu=UpUserRanks&PerCircleNum="&PerCircleNum&"&i="&i&"&TotalCount="&TotalCount&""
		response.Write("<br /><p style='text-align:center'><img src=images/loading.gif align=absmiddle />���ڸ����û���Ϣ��<span style='color:#FF0000;font-weight:bold'>"&FormatPercent(i/TotalCount,,-1)&"</span>��...</p>")
%>
<script type="text/javascript">
	setTimeout('window.location.href="<%=RefreshURL%>"', 1000);
</script>
<%
	else
%>
<br />
<table cellspacing="1" cellpadding="5" width=80% class=CommonListArea>
	<tr class=CommonListTitle>
		<td width="100%" align="center">��̳��ʾ��Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<td valign="top" align="center">
			���²�������ɣ�<br /><br />��<a href='?menu=UpStatic'>�������</a>���ء�
		</td>
	</tr>
</table>
<%
	end if
End Sub


Sub upSiteSettingsok
	Rs.Open "["&TablePrefix&"Forums]",Conn
		do while not Rs.eof
			allarticle=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			if allarticle>0 then
				allrearticle=Execute("Select sum(TotalReplies) from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&Rs("ForumID")&"")(0)
			else
				allrearticle=0
			end if
			Execute("update ["&TablePrefix&"Forums] Set TotalThreads="&allarticle&",TotalPosts="&allarticle+allrearticle&" where ForumID="&Rs("ForumID")&"")
			Rs.Movenext
		loop
	Rs.close
	Response.Write("�����ɹ���<br /><br />�Ѿ�������̳������̳��ͳ������<br />")
End Sub

AdminFooter
%>