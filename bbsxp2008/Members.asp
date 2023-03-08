<!-- #include file="Setup.asp" -->
<%
HtmlTop

if SiteConfig("PublicMemberList")=0 then error("系统已禁止显示成员列表。")

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


<div class="CommonBreadCrumbArea"><%=ClubTree%> → <a href="?">成员列表</a></div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr align=center class=CommonListTitle>
		<td><a href="?MemberSortDropDown=UserName">用户名</a></td>
		<td><a href="?MemberSortDropDown=UserRoleID">角色</a></td>
		<td>短讯息</td>
		<td><a href="?MemberSortDropDown=UserRegisterTime">注册时间</a></td>
		<td><a href="?MemberSortDropDown=Reputation">声望</a></td>
		<td><a href="?MemberSortDropDown=TotalPosts">发帖数</a></td>
		<td><a href="?MemberSortDropDown=UserMoney">金币</a></td>
		<td><a href="?MemberSortDropDown=experience">经验值</a></td>
		<td><a href="?MemberSortDropDown=UserActivityTime">最后活动时间</a></td>
	</tr>
<%
MemberSortDropDown=HTMLEncode(Request("MemberSortDropDown"))
SortOrderDropDown=HTMLEncode(Request("SortOrderDropDown"))
if Len(MemberSortDropDown)>20 then error("非法操作")
if len(SortOrderDropDown)>4 then error("非法操作")

if SortOrderDropDown="" then SortOrderDropDown="Desc"
if MemberSortDropDown<>"" then SqlOrder=" order by "&MemberSortDropDown&" "&SortOrderDropDown
if MemberSortDropDown="UserRoleID" then SqlOrder=" order by "&MemberSortDropDown&""

TotalCount=Execute("Select count(UserID) From ["&TablePrefix&"Users]"&item)(0) '获取数据数量
PageSetup=SiteConfig("MemberListPageSize") '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = RequestInt("PageIndex") '获取当前页
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
		高级搜索</a>
	</td>
<%end if%>
</tr>
</table>

<div id="ForumOption" style="display:none;">
<table cellspacing="1" cellpadding="5" width="100%" border="0" class=CommonListArea>
	<form method="POST" action="Members.asp" method=post>
	<tr class=CommonListTitle>
		<td align=center colspan=2>用户搜索</td>
	</tr>
	<tr class="CommonListCell">
		<td colspan=2><br />
			<div style="text-align:center"><input size="45" name="SearchText" /> <input type="submit" value=" 搜  索 " /></div>
		<br /><br />
		<fieldset>
			<legend>根据排序</legend>
			<select name=MemberSortDropDown>
				<option value=UserName>用户名</option>
				<option value=UserEmail>电子邮件</option>
				<option value=TotalPosts>发帖数</option>
				<option value=UserRegisterTime>注册日期</option>
				<option value=UserActivityTime>最后活动日期</option>
			</select> 
			<select name=SortOrderDropDown><option value=desc>倒序</option><option value=asc>顺序</option></select>
		</fieldset><br />
		<fieldset>
			<legend>日期过滤</legend>
			　　注册时间：<select name="JoinedDateComparer" onchange="javascript:if(this.options[this.selectedIndex].value != ''){$('UserRegisterTime').style.display='';}else{$('UserRegisterTime').style.display='none';}">
					<option value="">----</option>
					<option value="<">在此之前</option>
					<option value="=">在此之时</option>
					<option value=">">在此之后</option>
				</select> <span id=UserRegisterTime style="display:none">
			<input size="24" name="JoinedDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>" /></span><br />
			最后活动时间：<select name="LastPostDateComparer" onchange="javascript:if(this.options[this.selectedIndex].value != ''){$('UserActivityTime').style.display='';}else{$('UserActivityTime').style.display='none';}">
					<option value="">----</option>
					<option value="<">在此之前</option>
					<option value="=">在此之时</option>
					<option value=">">在此之后</option>
				</select> <span id="UserActivityTime" style="display:none">
			<input size="24" name="LastPostDate_picker" onclick="showcalendar(event, this)" value="<%=date()%>" /></span>
		</fieldset><br />
		<fieldset>
			<legend>过 滤</legend>
			　　角色：<select name="SearchRole">
				<option value="">所有用户</option>
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
			搜索范围：<select name="SearchType">
				<option value="UserName">用户名包含有</option>
				<option value="UserEmail">邮箱包含有</option>
				<option value="all">用户名或邮箱包含有</option>
			</select><br />
			　　状态：<select name="CurrentAccountStatus" size="1">
				<option value="">所有状态</option>
				<option value="0">正等待审核</option>
				<option value="1">已通过审核</option>
				<option value="2">已禁用</option>
				<option value="3">未通过审核</option>
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