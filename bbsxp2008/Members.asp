<!-- #include file="Setup.asp" -->
<%
HtmlTop

if SiteConfig("PublicMemberList")=0 then error("ϵͳ�ѽ�ֹ��ʾ��Ա�б�")

if SiteConfig("MemberListAdvancedSearch")=1 then
	SearchType=HTMLEncode(Request("SearchType"))
	SearchText=HTMLEncode(Request("SearchText"))
	SearchRole=RequestInt("SearchRole")
	CurrentAccountStatus=HTMLEncode(Request("CurrentAccountStatus"))
	JoinedDateComparer=Left(Request("JoinedDateComparer"),1)
	LastPostDateComparer=Left(Request("LastPostDateComparer"),1)
	JoinedDate_picker=HTMLEncode(Request("JoinedDate_picker"))
	LastPostDate_picker=HTMLEncode(Request("LastPostDate_picker"))
	
	if SearchType="all" then SearchType="UserEmail like '%"&SearchText&"%' or UserName"
	if SearchText<>"" and len(HTMLEncode(Request("SearchType")))<10 then item=item&" and ("&SearchType&" like '%"&SearchText&"%')"
	if JoinedDate_picker<>"" and JoinedDateComparer<>"" then item=item&" and DateDiff("&SqlChar&"d"&SqlChar&",'"&JoinedDate_picker&"',UserRegisterTime) "&JoinedDateComparer&" 0"
	if LastPostDate_picker<>"" and LastPostDateComparer<>"" then item=item&" and DateDiff("&SqlChar&"d"&SqlChar&",'"&LastPostDate_picker&"',UserActivityTime) "&LastPostDateComparer&" 0"
	if SearchRole >0 then item=item&" and UserRoleID="&SearchRole&""
	if CurrentAccountStatus <> "" and len(CurrentAccountStatus)<3 then item=item&" and UserAccountStatus="&CurrentAccountStatus&""

	if item<>"" then item=" where "&mid(item,5)
end if

%>
<script type="text/javascript" src="Utility/calendar.js"></script>


<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <a href="?">��Ա�б�</a></div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr align=center class=CommonListTitle>
		<td><a href="?MemberSortDropDown=UserName">�û���</a></td>
		<td><a href="?MemberSortDropDown=UserRoleID">��ɫ</a></td>
		<td>��ѶϢ</td>
		<td><a href="?MemberSortDropDown=UserRegisterTime">ע��ʱ��</a></td>
		<td><a href="?MemberSortDropDown=Reputation">����</a></td>
		<td><a href="?MemberSortDropDown=TotalPosts">������</a></td>
		<td><a href="?MemberSortDropDown=UserMoney">���</a></td>
		<td><a href="?MemberSortDropDown=experience">����ֵ</a></td>
		<td><a href="?MemberSortDropDown=UserActivityTime">���ʱ��</a></td>
	</tr>
<%
MemberSortDropDown=HTMLEncode(Request("MemberSortDropDown"))
SortOrderDropDown=HTMLEncode(Request("SortOrderDropDown"))
if Len(MemberSortDropDown)>20 then error("�Ƿ�����")
if len(SortOrderDropDown)>4 then error("�Ƿ�����")

if SortOrderDropDown="" then SortOrderDropDown="Desc"
if MemberSortDropDown<>"" then SqlOrder=" order by "&MemberSortDropDown&" "&SortOrderDropDown
if MemberSortDropDown="UserRoleID" then SqlOrder=" order by "&MemberSortDropDown&""

TotalCount=Execute("Select count(UserID) From ["&TablePrefix&"Users]"&item)(0) '��ȡ��������
PageSetup=SiteConfig("MemberListPageSize") '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage

sql="Select * from ["&TablePrefix&"Users] "&item&" "&SqlOrder&""

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
	<tr align=center class="CommonListCell">
		<td><a href=Profile.asp?UID=<%=Rs("UserID")%>><%=Rs("UserName")%></a></td>
		<td><%=ShowRole(RS("UserRoleID"))%></td>
		<td><a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&RecipientUserName=<%=Rs("UserName")%>',600, 350);"><img border="0" src="images/message.gif" /></a></td>
		<td><%=Rs("UserRegisterTime")%></td>
		<td><%=Rs("Reputation")%></td>
		<td><%=Rs("TotalPosts")%></td>
		<td><%=Rs("UserMoney")%></td>
		<td><%=Rs("experience")%></td>
		<td><%=Rs("UserActivityTime")%></td>
		</tr>
<%
Rs.MoveNext
loop
Set Rs=Nothing
%>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
<tr>
	<td><%ShowPage()%></td>
<%if SiteConfig("MemberListAdvancedSearch")=1 then%>	
	<td align="right">
		<a onmousedown="ToggleMenuOnOff('ForumOption')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/ForumSettings.gif)">
		�߼�����</a>
	</td>
<%end if%>
</tr>
</table>

<div id="ForumOption" style="display:none;">
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<form method="POST" action="Members.asp" method=post>
	<tr class=CommonListTitle>
		<td align=center colspan=2>�û�����</td>
	</tr>
	<tr class="CommonListCell">
		<td colspan=2><br />
			<div style="text-align:center"><input size="45" name="SearchText" /> <input type="submit" value=" ��  �� " /></div>
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
				</select> <span id=UserRegisterTime style="display:none">
			<input size="24" name="JoinedDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>" /></span><br />
			���ʱ�䣺<select name="LastPostDateComparer" onchange="javascript:if(this.options[this.selectedIndex].value != ''){$('UserActivityTime').style.display='';}else{$('UserActivityTime').style.display='none';}">
					<option value="">----</option>
					<option value="<">�ڴ�֮ǰ</option>
					<option value="=">�ڴ�֮ʱ</option>
					<option value=">">�ڴ�֮��</option>
				</select> <span id="UserActivityTime" style="display:none">
			<input size="24" name="LastPostDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>" /></span>
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
		</td>
	</tr>
	</form>
</table>
</div>

<%
HtmlBottom
%>