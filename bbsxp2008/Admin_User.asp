<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->

<%
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"

UserName=HTMLEncode(Request("UserName"))
RoleID=RequestInt("RoleID")
RoleName=HTMLEncode(Request("RoleName"))
Description=HTMLEncode(Request("Description"))

if Request("menu")="ChangePassword" then ChangePassword


AdminTop
Log("")

select case Request("menu")
	case "UserEdit"
		UserEdit
	case "SearchUser"
		SearchUser
	case "UserDelTopic"
		UserDelTopic
	case "UserDel"
		UserDel
	case "Userok"
		Userok
		
	case "ChangePassword"
		ChangePassword
		
	case "UserRankAdd"
		UserRankAdd	
	case "UserRank"
		UserRank
	case "UserRankUp"
		RankID=RequestInt("RankID")
		RoleID=RequestInt("RoleID")
		RankName=Request("RankName")
		PostingCountMin=RequestInt("PostingCountMin")
		if ""&RankName&""="" then Alert("��û������ȼ�����")
		
		
		Rs.Open "select * from ["&TablePrefix&"Ranks] where RankID="&RankID&"",Conn,1,3
			If Rs.eof Then Rs.Addnew()
			Rs("RoleID")=RoleID
			Rs("RankName")=RankName
			Rs("PostingCountMin")=PostingCountMin
		Rs.update()
		Rs.close

		response.write "���³ɹ�"
	case "UserRankDel"
		Execute("Delete from ["&TablePrefix&"Ranks] where RankID="&Request("RankID")&"")
		response.write "ɾ���ɹ�"
	case "AllRoles"
		AllRoles
	case "CreateRole"
		if RoleName=empty then Alert("��û�������ɫ����")
		Execute("insert into ["&TablePrefix&"Roles] (Name) values ('"&RoleName&"')")
		AllRoles
	case "ViewRole"
		ViewRole
	case "UpRole"
		if RoleName=empty then Alert("��û�������ɫ����")
		RoleMaxFileSize=RequestInt("RoleMaxFileSize")
		RoleMaxPostAttachmentsSize=RequestInt("RoleMaxPostAttachmentsSize")
		Execute("update ["&TablePrefix&"Roles] Set Name='"&RoleName&"',Description='"&Description&"',RoleMaxFileSize="&RoleMaxFileSize&",RoleMaxPostAttachmentsSize="&RoleMaxPostAttachmentsSize&" where RoleID="&RoleID&"")
		Response.Write("�༭�ɹ�")
	case "RolePermissions"
		RolePermissions
	case "RolePermissionsUP"
		for each ho in Request.Form("ForumID")
			Rs.Open "Select * from ["&TablePrefix&"ForumPermissions] where ForumID="&ho&" and RoleID="&RoleID&"",Conn,1,3
			if Rs.eof then
				Rs.Addnew()
				Rs("RoleID")=RoleID
				Rs("ForumID")=ho
			end if
			Rs("PermissionView")=RequestInt("PermissionView"&ho)
			Rs("PermissionRead")=RequestInt("PermissionRead"&ho)
			Rs("PermissionPost")=RequestInt("PermissionPost"&ho)
			Rs("PermissionReply")=RequestInt("PermissionReply"&ho)
			Rs("PermissionEdit")=RequestInt("PermissionEdit"&ho)
			Rs("PermissionDelete")=RequestInt("PermissionDelete"&ho)
			Rs("PermissionCreatePoll")=RequestInt("PermissionCreatePoll"&ho)
			Rs("PermissionVote")=RequestInt("PermissionVote"&ho)
			Rs("PermissionAttachment")=RequestInt("PermissionAttachment"&ho)
			Rs("PermissionManage")=RequestInt("PermissionManage"&ho)
			Rs.update
			Rs.close
		next
		Response.Write("����Ȩ�����óɹ�")
	case "DelRole"
		if Roleid<4 then Alert("�ý�ɫΪϵͳ���ý�ɫ���޷�ɾ��")
		if Not Execute("Select UserID From ["&TablePrefix&"Users] where UserRoleID="&RoleID&"" ).eof then Alert("ֻ��ɾ��û�г�Ա�Ľ�ɫ�飡")
		Execute("Delete from ["&TablePrefix&"Roles] where RoleID="&RoleID&"")
		Response.Write("ɾ���ɹ�")
	case else
		SearchUserok
end select


Sub ChangePassword
Response.clear
if Request_Method = "POST" then
	NewPassword1=Trim(Request("NewPassword1"))
	NewPassword2=Trim(Request("NewPassword2"))
	if NewPassword1<>NewPassword2 then  Alert("��2����������벻ͬ")
	if Len(NewPassword1)<6 then  Alert("���벻��С��6λ��")

	'--------------------  API Start  --------------------
	Message=""
	If SiteConfig("APIEnable")=1 Then APIUpdateUser UserName,NewPassword1,"","",""
	if Message<>"" then AlertForModal(""&Message&"")
	'--------------------  API End  ----------------------
	
	ModifyUserPassword UserName,NewPassword1,"","",""
%>
<script language="JavaScript" type="text/javascript">
	alert('�����޸ĳɹ�');
	parent.BBSXP_Modal.Close();
</script>
<%
end if
%>
<title>�޸�����</title>
<style>body,table{FONT-SIZE:9pt;}</style>
<form name=form action="Admin_User.asp?menu=ChangePassword" method="POST">
<input type=hidden name="UserName" value="<%=UserName%>">
�޸����� - (<%=UserName%>) <br /><br />
<table border="0" width="100%">
	<tr>
		<td>�����룺��</td>
		<td><input name="NewPassword1" type="password" maxlength="15" size="40" /></td>
	</tr>
	<tr>
		<td>�ٴ����������룺��</td>
		<td><input name="NewPassword2" type="password" maxlength="15" size="40" /></td>
	</tr>
</table>
<br />
<input type="submit" value=" �޸����� ">
</form>
<%
Response.End
End Sub


Sub SearchUser
%>
<SCRIPT type="text/javascript" src="Utility/calendar.js"></SCRIPT>
�û����ϣ�<b><font color=red><%=Execute("Select count(UserID) from ["&TablePrefix&"Users]")(0)%></font></b> ��
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<form method="POST" action="?menu=SearchUserok">
	<tr class=CommonListTitle>
		<td align=center>�û�����</td>
	</tr>
	<tr class="CommonListCell">
		<td><br />
			<div style="text-align:center"><input size="45" name="SearchText"> <input type="submit" value=" ��  �� "></div>
		<br /><br />
		<fieldset>
			<legend>��������</legend>
			<select name=MemberSortDropDown>
				<option value=UserName>�û���</option>
				<option value=UserEmail>�����ʼ�</option>
				<option value=TotalPosts>������</option>
				<option value=UserRegisterTime>ע������</option>
				<option value=UserActivityTime>�������</option>
			</select> 
			<select name=SortOrderDropDown><option value=desc>����</option><option value=asc>˳��</option></select>
		</fieldset><br />
		<fieldset>
			<legend>���ڹ���</legend>
			����ע��ʱ�䣺<select name="JoinedDateComparer" onchange="javascript:if(this.options[this.selectedIndex].value != ''){$('UserRegisterTime').style.display='';}else{$('UserRegisterTime').style.display='none';}">
					<option value="">----</option>
					<option value="<">�ڴ�֮ǰ</option>
					<option value="=">�ڴ�֮ʱ</option>
					<option value=">">�ڴ�֮��</option>
				</select> <span id=UserRegisterTime style="display:none"><input size="24" name="JoinedDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>"></span><br />
			���ʱ�䣺<select name="LastPostDateComparer" onchange="javascript:if(this.options[this.selectedIndex].value != ''){$('UserActivityTime').style.display='';}else{$('UserActivityTime').style.display='none';}">
					<option value="">----</option>
					<option value="<">�ڴ�֮ǰ</option>
					<option value="=">�ڴ�֮ʱ</option>
					<option value=">">�ڴ�֮��</option>
				</select> <span id="UserActivityTime" style="display:none"><input size="24" name="LastPostDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>"></span>
		</fieldset><br />
		<fieldset>
			<legend>�� ��</legend>
			������ɫ��<select name="SearchRole">
				<option value="">�����û�</option>
<%
	sql="Select * from ["&TablePrefix&"Roles] where RoleID > 0 order by RoleID"
	Set Rs=Execute(sql)
		Do While Not Rs.EOF
				Response.Write("<option value='"&Rs("RoleID")&"'>"&Rs("Name")&"</option>")
			Rs.MoveNext
		loop
	Rs.Close
%>
			</select><br />
			������Χ��<select name="SearchType">
				<option value="UserName">�û���������</option>
				<option value="UserEmail">���������</option>
				<option value="all">�û��������������</option>
			</select><br />
			����״̬��<select name="CurrentAccountStatus" size="1">
				<option value="">����״̬</option>
				<option value="0">���ȴ����</option>
				<option value="1">��ͨ�����</option>
				<option value="2">�ѽ���</option>
				<option value="3">δͨ�����</option>
			</select>
		</fieldset>
<hr noshade="noshade" size="1" color=#999999 />
		<fieldset>
			<legend>��ݷ�ʽ</legend>
			<li><a href="?MemberSortDropDown=TotalPosts&SortOrderDropDown=desc">��������</li></a>
			<li><a href="?LastPostDateComparer==&LastPostDate_picker=<%=date()%>">��ȥ 24 Сʱ�ڻ���û�</li></a>
			<li><a href="?JoinedDateComparer==&JoinedDate_picker=<%=date()%>">��ȥ 24 Сʱ��ע����û�</li></a>
			<li><a href="?CurrentAccountStatus=0">�ȴ���˵��û�</a></li>
		</fieldset>
	</form>
		</td>
	</tr>
</table><br />



<%
End Sub

Sub SearchUserok
%>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<TR align=center class=CommonListTitle>
		<TD>�û���</TD>
		<TD>Email</TD>
		<TD>������</TD>
		<TD>ע��ʱ��</TD>
		<TD>���ʱ��</TD>
		<TD>����</TD>
	</TR>
<%
	SearchType=HTMLEncode(Request("SearchType"))
	SearchText=HTMLEncode(Request("SearchText"))
	SearchRole=HTMLEncode(Request("SearchRole"))
	CurrentAccountStatus=HTMLEncode(Request("CurrentAccountStatus"))
	JoinedDateComparer=Left(Request("JoinedDateComparer"),1)
	LastPostDateComparer=Left(Request("LastPostDateComparer"),1)
	JoinedDate_picker=HTMLEncode(Request("JoinedDate_picker"))
	LastPostDate_picker=HTMLEncode(Request("LastPostDate_picker"))
	
	if SearchType="all" then SearchType="UserEmail like '%"&SearchText&"%' or UserName"
	if SearchText<>"" then item=item&" and ("&SearchType&" like '%"&SearchText&"%')"
	if JoinedDate_picker<>"" and JoinedDateComparer<>"" then item=item&" and DateDiff("&SqlChar&"d"&SqlChar&",'"&JoinedDate_picker&"',UserRegisterTime) "&JoinedDateComparer&" 0"
	if LastPostDate_picker<>"" and LastPostDateComparer<>"" then item=item&" and DateDiff("&SqlChar&"d"&SqlChar&",'"&LastPostDate_picker&"',UserActivityTime) "&LastPostDateComparer&" 0"
	if SearchRole <> "" then item=item&" and UserRoleID="&SearchRole&""
	if CurrentAccountStatus <> "" then item=item&" and UserAccountStatus="&CurrentAccountStatus&""

	if item<>"" then item=" where "&mid(item,5)

	sql="["&TablePrefix&"Users]"&item&""

	TotalCount=Execute("Select count(UserID) From "&sql&"")(0) '��ȡ��������
	PageSetup=20 '�趨ÿҳ����ʾ����
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
	PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	
	if Request("MemberSortDropDown")<>"" then item=item&" order by "&Request("MemberSortDropDown")&" "&Request("SortOrderDropDown")&""
	sql="["&TablePrefix&"Users]"&item&""
	if PageCount<11 then
		Set Rs=Execute(sql)
	else
		rs.Open sql,Conn,1
	end if
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	i=0
	Do While Not Rs.EOF and i<PageSetup
		i=i+1
%>
	<TR align=center class="CommonListCell">
		<TD><a href="Admin_User.asp?menu=UserEdit&UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a></TD>
		<TD><a href="mailto:<%=Rs("UserEmail")%>"><%=Rs("UserEmail")%></a></TD>
		<TD><%=Rs("TotalPosts")%></TD>
		<TD><%=Rs("UserRegisterTime")%></TD>
		<TD><%=Rs("UserActivityTime")%></TD>
		<TD><a href="Admin_User.asp?menu=UserEdit&UserName=<%=Rs("UserName")%>">�༭</a> | <a onclick="return window.confirm('��ȷ��Ҫɾ������ѡ�û���ȫ������?');" href="Admin_User.asp?menu=UserDel&UserID=<%=Rs("UserID")%>">ɾ��</a></TD>
	</TR>
<%
		Rs.MoveNext
	loop
	Rs.Close
%>
</TABLE>
<table border=0 width=100% align=center><tr><td><%ShowPage()%></tr></td></table>
<%
End Sub


Sub UserEdit
	sql="Select * from ["&TablePrefix&"Users] where UserName='"&HTMLEncode(UserName)&"'"
	Set Rs=Execute(sql)
	if Rs.eof then Alert(""&UserName&" ���û����ϲ�����")
		
		UserSign=replace(""&Rs("UserSign")&"","<br>", vbCrlf)
		UserBio=replace(""&Rs("UserBio")&"","<br>",vbCrlf)
		UserNote=replace(""&Rs("UserNote")&"","<br>",vbCrlf)
%>
<form method="POST" name=form action="?menu=Userok">
<input type=hidden name=UserName value="<%=Rs("UserName")%>">
<table cellSpacing="1" cellpadding="5" border="0" width="70%" class=CommonListArea align=center>
	<tr class=CommonListTitle>
		<td width="600" colspan="4" align="center"><font color="000000"><a target="_blank" href="Profile.asp?UID=<%=Rs("UserID")%>">�鿴��<%=Rs("UserName")%>������ϸ����</a></font></td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2">�û����ƣ�<%=Rs("UserName")%></td>
		<td width="600" colspan="2">�û����룺<a href="javascript:BBSXP_Modal.Open('?menu=ChangePassword&UserName=<%=Rs("UserName")%>',500,160);">�޸�����</a></td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2">�û���ɫ��<select name="UserRoleID">
<%
	RoleGetRow = FetchEmploymentStatusList("Select RoleID,Name from ["&TablePrefix&"Roles] where RoleID > 0 order by RoleID")
	For i=0 To Ubound(RoleGetRow,2)
	%><option value="<%=RoleGetRow(0,i)%>" <%if Rs("UserRoleID")=RoleGetRow(0,i) then%>selected<%end if%>><%=RoleGetRow(1,i)%></option><%
	Next
%>
			</select>		</td>
		<td width="600" colspan="2">�ʺ�״̬��<select name="UserAccountStatus" size="1">
				<option value="0" <%if Rs("UserAccountStatus")=0 then%>selected<%end if%>>���ȴ����</option>
				<option value="1" <%if Rs("UserAccountStatus")=1 then%>selected<%end if%>>��ͨ�����</option>
				<option value="2" <%if Rs("UserAccountStatus")=2 then%>selected<%end if%>>�ѽ���</option>
				<option value="3" <%if Rs("UserAccountStatus")=3 then%>selected<%end if%>>δͨ�����</option>
			</select>		</td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2">�û�ͷ�Σ�<input size="10" name="UserTitle" value="<%=Rs("UserTitle")%>"></td>
		<td width="600" colspan="2">���εȼ���<select name="ModerationLevel" size="1">
			<option value="1" <%if Rs("ModerationLevel")=True then%>selected<%end if%>>�����û�</option>
			<option value="0" <%if Rs("ModerationLevel")=False then%>selected<%end if%>>�������û�</option>
		</select></td>
	</tr>
	<tr class="CommonListCell">
		<td width="600" colspan="2">�� �� ����<input size="10" name="TotalPosts" value="<%=Rs("TotalPosts")%>"></td>
		<td colspan="2">�û�������<input size="10" name="Reputation" value="<%=Rs("Reputation")%>"></td>
	</tr>
	
	<tr class="CommonListCell">
		<td width="600" colspan="2">��Ծ������<input size="10" name="UserActivityDay" value="<%=Rs("UserActivityDay")%>"></td>
		<td colspan="2">�û��ȼ���<%=Rs("UserRank")%></td>
	</tr>
	
	<tr class="CommonListCell">
		<td colspan="2">�𡡡�Ǯ��<input size="10" name="UserMoney" value="<%=Rs("UserMoney")%>" /></td>
		<td width="600" colspan="2">�� �� �ˣ�<input size="20" name="ReferrerName" value="<%=Rs("ReferrerName")%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2">�� �� ֵ��<input size="10" name="experience" value="<%=Rs("experience")%>" /></td>
		<td colspan="2">�û�ͷ��<input size="20" name="UserFaceUrl" value="<%=Rs("UserFaceUrl")%>"></td>
	</tr>
	<tr class="CommonListCell">
		<td width="1200" colspan="4">ע�����ڣ�<%=Rs("UserRegisterTime")%> ���ɣУ�<%=Rs("UserRegisterIP")%>��</td>
	</tr>
	<tr class="CommonListCell">
		<td width="1200" colspan="4">����ڣ�<%=Rs("UserActivityTime")%> ���ɣУ�<%=Rs("UserActivityIP")%>��</td>
	</tr>
	<tr class=CommonListTitle>
		<td colspan="4" align="center">��������</td>
	</tr>
	<tr class="CommonListCell">
	  <td width="600" colspan="2">���֣� <input type="text" name="RealName" size="20" value="<%=Rs("RealName")%>" /></td>
		<td width="600" height="3" colspan="2">�Ա� <select size=1 name=UserSex>
				<option value=0 selected>[��ѡ��]</option>
				<option value=1 <%if Rs("UserSex")=1 then%>selected<%end if%>>��</option>
				<option value=2 <%if Rs("UserSex")=2 then%>selected<%end if%>>Ů</option>
	  </select></td>
	</tr>
	<tr class="CommonListCell">
		<td width="600" colspan="2">���գ� <input type="text" name="birthday" size="20" value="<%=Rs("birthday")%>"></td>
		<td width="600" colspan="2">��ַ�� <input type="text" name="Address" size="20" value="<%=Rs("Address")%>"></td>
	</tr>
	<tr class="CommonListCell">
	  <td colspan="2">ְҵ�� <input type="text" name="Occupation" size="20" value="<%=Rs("Occupation")%>"></td>
	  <td colspan="2">��Ȥ�� <input type="text" name="Interests" size="20" value="<%=Rs("Interests")%>"></td>
    </tr>
	<tr class="CommonListCell">
	  <td colspan="2">���䣺 <input type="text" name="UserEmail" size="20" value="<%=Rs("UserEmail")%>"></td>
	  <td colspan="2">��ҳ�� <input type="text" name="WebAddress" size="20" value="<%=Rs("WebAddress")%>"></td>
    </tr>
	<tr class="CommonListCell">
	  <td colspan="2">���ͣ� <input type="text" name="WebLog" size="20" value="<%=Rs("WebLog")%>"></td>
	  <td colspan="2">��᣺ <input type="text" name="WebGallery" size="20" value="<%=Rs("WebGallery")%>"></td>
    </tr>
    
	<tr class=CommonListTitle>
		<td colspan="4" align="center">��ʱͨ��</td>
	</tr>
    
	<tr class="CommonListCell">
	  <td colspan="2">QQ ����<input type="text" name="QQ" size="20" value="<%=Rs("QQ")%>"></td>
	  <td colspan="2">ICQ����<input type="text" name="ICQ" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=Rs("ICQ")%>"></td>
    </tr>
	<tr class="CommonListCell">
	  <td colspan="2">AIM����<input type="text" name="AIM" size="20" value="<%=Rs("AIM")%>"></td>
	  <td colspan="2">MSN����<input type="text" name="MSN" size="20" value="<%=Rs("MSN")%>"></b></td>
    </tr>
	<tr class="CommonListCell">
	  <td colspan="2">Yahoo��<input type="text" name="Yahoo" size="20" value="<%=Rs("Yahoo")%>"></td>
	  <td colspan="2">Skype��<input type="text" name="Skype" size="20" value="<%=Rs("Skype")%>"></b></td>
    </tr>
    
	<tr class=CommonListTitle>
		<td colspan="4" align="center">ǩ������飦��ע</td>
	</tr>
    
    
	<tr class="CommonListCell">
		<td width="600" colspan="4">ǩ����<textarea name="UserSign" rows="4" cols="60"><%=UserSign%></textarea></td>
	</tr>
	<tr class="CommonListCell">
		<td width="600" colspan="4">��飺<textarea name="UserBio" rows="4" cols="60"><%=UserBio%></textarea></td>
	</tr>
	
	<tr class="CommonListCell">
		<td width="600" colspan="4">��ע��<textarea name="UserNote" rows="4" cols="60"><%=UserNote%></textarea></td>
	</tr>
	
	
	<tr class=CommonListTitle>
		<td width="600" align="center" ><a onclick="return window.confirm('��ȷ��Ҫɾ�����û����з����������?');" href="?menu=UserDelTopic&UserName=<%=Rs("UserName")%>">ɾ�����û�����������</a></td>
		<td width="600" colspan="2" align="center" ><input type="submit" value=" �� �� "></td>
		<td width="600" align="center" ><a onclick="return window.confirm('��ȷ��Ҫɾ�����û�����������?');" href="?menu=UserDel&UserID=<%=Rs("UserID")%>">ɾ�����û�����������</a></td>
	</tr>
</table>
</form>
<%
End Sub

Sub UserDelTopic
	ThreadGetRow=FetchEmploymentStatusList("select ThreadID from ["&TablePrefix&"Threads] where PostAuthor='"&UserName&"'")
	if IsArray(ThreadGetRow) then
		for i=0 to ubound(ThreadGetRow,2)
			sql="select PostID from ["&TablePrefix&"Posts] where ThreadID="&ThreadGetRow(0,i)&""
			Rs.open sql,Conn,1,1
				do while not Rs.eof
					DelAttachments"Select * from ["&TablePrefix&"PostAttachments] where PostID="&Rs("PostID")&""
					Rs.movenext
				loop
			Rs.Close
		next
	end if
	ThreadGetRow=null
			
	Execute("Delete from ["&TablePrefix&"Threads] where PostAuthor='"&UserName&"'")
	Response.Write("�Ѿ��� "&UserName&" ���з����������ȫ��ɾ��")
End Sub

Sub UserDel
	for each ho in Request("UserID")
		ho=int(ho)
		if ho=CookieUserID then Alert("�����Լ�ɾ���Լ�")
		Rs.open "select UserName from ["&TablePrefix&"Users] where UserID="&ho&"",Conn,1,3
		if not Rs.Eof Then
			UserName=Rs("UserName")
			'--------------------  API Start  --------------------
			Message=""
			If SiteConfig("APIEnable")=1 and UserName<>"" Then APIDeleteUser()
			if Message<>"" then Alert(""&Message&"")
			'--------------------  API End  ----------------------
		
			DelAttachments"Select * from ["&TablePrefix&"PostAttachments] where UserName='"&UserName&"'"
			
			Rs.Delete
			Rs.Update
		end if
		Rs.Close
	next
	Response.Write("�Ѿ��ɹ�ɾ���û�IDΪ��"&Request("UserID")&"������������")
End Sub

Sub Userok
	birthday=HTMLEncode(Request.Form("birthday"))
	if Not IsDate(birthday) then birthday=null
	
	
	sql="Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'"
	Rs.Open sql,Conn,1,3
		Rs("UserFaceUrl")=Request("UserFaceUrl")
		Rs("UserRoleID")=Request("UserRoleID")
		Rs("UserEmail")=Request("UserEmail")
		Rs("TotalPosts")=Request("TotalPosts")
		Rs("UserActivityDay")=Request("UserActivityDay")
		Rs("UserRank")=Request("UserRank")
		Rs("experience")=Request("experience")
		Rs("UserMoney")=Request("UserMoney")
		Rs("ReferrerName")=Request("ReferrerName")
		Rs("TotalPosts")=Request("TotalPosts")
		Rs("UserTitle")=Request("UserTitle")
		Rs("Reputation")=RequestInt("Reputation")
		Rs("UserSign")=HTMLEncode(Request.Form("UserSign"))
		Rs("Interests")=HTMLEncode(Request.Form("Interests"))
		Rs("UserBio")=HTMLEncode(Request.Form("UserBio"))
		Rs("UserNote")=HTMLEncode(Request.Form("UserNote"))
		
		Rs("UserSex")=RequestInt("UserSex")
		Rs("UserAccountStatus")=Request("UserAccountStatus")
		Rs("ModerationLevel")=Request("ModerationLevel")
		Rs("birthday")=birthday
		
		Rs("QQ")=HTMLEncode(Request("QQ"))
		Rs("ICQ")=HTMLEncode(Request("ICQ"))
		Rs("AIM")=Request("AIM")
		Rs("MSN")=Request("MSN")
		Rs("Yahoo")=Request("Yahoo")
		Rs("Skype")=Request("Skype")
		
		Rs("RealName")=Request("RealName")
		Rs("Occupation")=Request("Occupation")
		Rs("Address")=Request("Address")
		Rs("WebAddress")=Request("WebAddress")
		Rs("WebLog")=Request("WebLog")
		Rs("WebGallery")=Request("WebGallery")
	Rs.update
	Rs.close
	
	UpdateApplication"UserBirthday","Select UserID,UserName From ["&TablePrefix&"Users] where Month(Birthday)="&Month(now())&" and day(Birthday)="&day(now())&""
	
	Response.Write("���³ɹ�")
End Sub



Sub UserRank


%>
<table border="0" width="90%" align="center">
	<tr>
		<td align=right><a href="?menu=UserRankAdd" class="CommonTextButton">����û��ȼ�</a>
</td>
	</tr>
</table>

<%
	RoleGetRow = FetchEmploymentStatusList("Select RoleID,Name from ["&TablePrefix&"Roles] where RoleID >= 0 order by RoleID")
	If IsArray(RoleGetRow) Then
		For i=0 to ubound(RoleGetRow,2)
			Set Rs=Execute("select * from ["&TablePrefix&"Ranks] where RoleID="&RoleGetRow(0,i)&"")
			If not Rs.eof Then
%>
<table border="0" cellpadding="5" cellspacing="1" class=CommonListArea width=90% align="center">
	<tr class=CommonListTitle>
		<td colspan="4" align="center"><%=RoleGetRow(1,i)%></td>
	</tr>
	<tr class=CommonListHeader>
		<td width="100" align="center">ID</td>
		<td>�û��ȼ�</td>
		<td width="150">���ٷ���</td>
		<td width="150" align="center">����</td>
	</tr>
<%
			Do While Not Rs.eof
%>
	<tr class="CommonListCell">
		<td align="center"><%=Rs("RankID")%></td>
		<td><%=Rs("RankName")%></td>
		<td><%=Rs("PostingCountMin")%></td>
		<td align="center">
			<a href="?menu=UserRankAdd&RankID=<%=Rs("RankID")%>">�༭</a> | 
			<a onclick="return window.confirm('��ȷ��Ҫִ�иò���?');" href="?menu=UserRankDel&RankID=<%=Rs("RankID")%>">ɾ��</a>
		</td>
	</tr>
<%
				Rs.movenext
			Loop
%>
</table>
<%
			End If
		Next
	End If
End Sub

Sub UserRankAdd
	RankID=RequestInt("RankID")
	If RankID>0 Then
		Set Rs=Execute("select * from ["&TablePrefix&"Ranks] where RankID="&RankID&"")
		If Not Rs.eof Then
			RoleID=Rs("RoleID")
			RankName=Rs("RankName")
			PostingCountMin=Rs("PostingCountMin")
		End If
		Rs.close
	End If
%>
<form method="POST" action=?menu=UserRankUp>
<input type="hidden" name="RankID" value="<%=RankID%>">
<table border="0" cellpadding="5" cellspacing="1" class=CommonListArea width=100% align="center">
	<tr class=CommonListTitle>
		<td colspan="2" align="center">���/�༭�û��ȼ�</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width=30%>�û���ɫ��</td>
		<td>
		<select name="RoleID">
		<option value="0">������</option>
<%
	RoleGetRow = FetchEmploymentStatusList("Select RoleID,Name from ["&TablePrefix&"Roles] where RoleID > 0 order by RoleID")
	For i=0 To Ubound(RoleGetRow,2)
	%><option value="<%=RoleGetRow(0,i)%>" <%if RoleID=RoleGetRow(0,i) then%>selected<%end if%>><%=RoleGetRow(1,i)%></option><%
	Next
%>
		</select>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width=30%>���ٷ�����</td>
		<td><input size="20" name="PostingCountMin" value="<%=PostingCountMin%>"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width=30%>�û��ȼ���Ӧ�����ֻ�ͼƬHTML���룺<br /><font color=#FF0000>��֧��HTML��</font></td>
		<td><input size=50 name="RankName" value='<%=server.HTMLEncode(RankName)%>' /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" colspan=2><input type="submit" value=" ȷ �� "></td>
	</tr>
</table>
</form>
<%
End Sub

Sub AllRoles
%><b>��ɫ����</b><br />����/�༭��ɫ����������һ���û����ƶ�����Ȩ�ޡ�<br />
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center" colspan="2">��ɫ����</td>
	</tr>
<%
	sql="Select * from ["&TablePrefix&"Roles] order by RoleID"
	Set Rs=Execute(sql)
		Do While Not Rs.EOF 
%>
	<tr class="CommonListCell">
		<td><a href="?menu=ViewRole&RoleID=<%=Rs("RoleID")%>"><b><%=Rs("Name")%></b></a><br /><%=Rs("Description")%></td>
		<td align="center" width="200"><a class="CommonTextButton" href="?menu=ViewRole&RoleID=<%=Rs("RoleID")%>">�༭��ɫ����</a> <a class="CommonTextButton" href="?menu=RolePermissions&RoleID=<%=Rs("RoleID")%>">�༭���Ȩ��</a></td>
	</tr>
<%
			Rs.MoveNext
		loop
	Rs.Close
%>
	<tr class="CommonListCell">
		<td colspan="2">
		<form name="form" method="POST" action="?menu=CreateRole">
			<input name="RoleName" onkeyup="ValidateTextboxAdd(this, 'RoleName1')" size="50"><input type="submit" value="����" id="RoleName1" disabled></form>
		</td>
	</tr>
</table>
<%
End Sub

Sub ViewRole
	if  RoleId<4 then PostDisabled="disabled"
	Rs.Open "Select * from ["&TablePrefix&"Roles] where RoleID="&RoleId&"",Conn,1
%>
<form name="form" method="POST" action="?menu=UpRole&RoleID=<%=RoleID%>" style="margin:0px;padding:0px;">
	<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align="center">
		<tr class="CommonListTitle">
			<td>RoleID</td>
			<td width="60%"><%=Rs("RoleID")%></td>
		</tr>
		<tr class="CommonListCell">
			<td>����</td>
			<td width="60%"><input name="RoleName" size="50" value="<%=Rs("Name")%>"></td>
		</tr>
		<tr class="CommonListCell">
			<td>����</td>
			<td width="60%"><input name="Description" size="50" value="<%=Rs("Description")%>"></td>
		</tr>
		<tr class="CommonListTitle">
			<td colspan="2">���øý�ɫ�ϴ����������ѡ�</td>
		</tr>
		<tr class="CommonListCell">
			<td><b>���������Ӹ����Ĵ�С��KB��</b><br />����ϵͳĬ�����ã�������0</td>
			<td><input size="20" name="RoleMaxFileSize" value="<%=Rs("RoleMaxFileSize")%>" /></td>
		</tr>
		<tr class="CommonListCell">
			<td><b>�����û��ϴ��ļ��е����������KB��</b><br />����ϵͳĬ�����ã�������0</td>
			<td><input size="20" name="RoleMaxPostAttachmentsSize" value="<%=Rs("RoleMaxPostAttachmentsSize")%>" /></td>
		</tr>
		<tr class="CommonListCell">
			<td colspan="2" align="center"><input <%=PostDisabled%> type="button" value="ɾ��" onclick="document.location.href='?menu=DelRole&amp;RoleID=<%=RoleID%>'"> <input type="submit" value="����"> </td>
		</tr>
        
	</table>
</form>
<%
	Rs.close
End Sub




Sub RolePermissions
%>
<form name="form" method="POST" action="?menu=RolePermissionsUP&RoleID=<%=RoleID%>" style="margin:0px;padding:0px;">
	<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea align="center">
	<tr class="CommonListTitle">
		<td colspan="11" align="center">�༭��ɫ��<%=Execute("select Name from ["&TablePrefix&"Roles] where RoleID="&RoleId&"")(0)%>���ڸ�����Ȩ��</td>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>���Ȩ��</td>
		<td width="8%">���</td>
		<td width="8%">�Ķ�</td>
		<td width="8%">����</td>
		<td width="8%">�ظ�</td>
		<td width="8%">�༭</td>
		<td width="8%">ɾ��</td>
		<td width="10%">����ͶƱ</td>
		<td width="8%">ͶƱ</td>
		<td width="8%">����</td>
		<td width="9%">����</td>
	</tr>
<%
	Sql="Select ForumID,ForumName from ["&TablePrefix&"Forums]"
	Set Rs=Execute(Sql)
	do while Not Rs.eof
		Set Rs1=Execute("Select * from ["&TablePrefix&"ForumPermissions] where RoleID="&RoleID&" and ForumID="&Rs("ForumID")&"")
			if not Rs1.eof then
				PermissionView=Rs1("PermissionView")
				PermissionRead=Rs1("PermissionRead")
				PermissionPost=Rs1("PermissionPost")
				PermissionReply=Rs1("PermissionReply")
				PermissionEdit=Rs1("PermissionEdit")
				PermissionDelete=Rs1("PermissionDelete")
				PermissionCreatePoll=Rs1("PermissionCreatePoll")
				PermissionVote=Rs1("PermissionVote")
				PermissionAttachment=Rs1("PermissionAttachment")
				PermissionManage=Rs1("PermissionManage")
			end if
		Set Rs1=nothing
%>
	<input type="hidden" name="ForumID" value="<%=Rs("ForumID")%>" />
	<tr align="center" class="CommonListCell">
		<td><%=Rs("ForumName")%></td>
		<td><input type="checkbox" value="1" name="PermissionView<%=Rs("ForumID")%>"<%if PermissionView=1 then%> checked<%end if%> /></td>
		<td><input type="checkbox" value="1" name="PermissionRead<%=Rs("ForumID")%>"<%if PermissionRead=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionPost<%=Rs("ForumID")%>"<%if PermissionPost=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionReply<%=Rs("ForumID")%>"<%if PermissionReply=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionEdit<%=Rs("ForumID")%>"<%if PermissionEdit=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionDelete<%=Rs("ForumID")%>"<%if PermissionDelete=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionCreatePoll<%=Rs("ForumID")%>"<%if PermissionCreatePoll=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionVote<%=Rs("ForumID")%>"<%if PermissionVote=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionAttachment<%=Rs("ForumID")%>"<%if PermissionAttachment=1 then%> checked /><%end if%></td>
		<td><input type="checkbox" value="1" name="PermissionManage<%=Rs("ForumID")%>"<%if PermissionManage=1 then%> checked /><%end if%></td>
	</tr>
<%
		Rs.movenext
	loop
%>
	<tr class="CommonListCell">
		<td colspan="11" align="right"><input type="submit" value=" ���� " /></td>
	</tr>
</table></form>


<%
	Rs.close
End Sub




AdminFooter
%>