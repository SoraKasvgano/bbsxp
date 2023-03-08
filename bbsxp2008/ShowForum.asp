<!-- #include file="Setup.asp" --><%

ForumID=RequestInt("ForumID")
SortOrder=RequestInt("SortOrder")
TimeLimit=RequestInt("TimeLimit")
GoodTopic=RequestInt("GoodTopic")
VoteTopic=RequestInt("VoteTopic")
order=HTMLEncode(Request("order"))
Category=HTMLEncode(Request("Category"))
if ""&order&""<>"" then
	if Len(order)>12 then error("非法操作")
	if instr("|threadid|lasttime|posttime|lastvieweddate|IsGood|totalreplies|","|"&lcase(order)&"|")<1 then error("无效的参数值")
end if
if Len(Category)>25 then error("类别名字太长")


%><!-- #include file="Utility/ForumPermissions.asp" --><%
HtmlHeadTitle=ForumName

HtmlTop




if TotalCategorys<>"" then
	filtrate=split(TotalCategorys,"|")
	CategorysList="[<a href='ShowForum.asp?ForumID="&ForumID&"'>全部</a>] "
	for i = 0 to ubound(filtrate)
		CategorysList=CategorysList&"[<a href='ShowForum.asp?ForumID="&ForumID&"&Category="&filtrate(i)&"'>"&filtrate(i)&"</a>] "
	next
end if

%>

<meta http-equiv="refresh" content="300">


<div class="CommonBreadCrumbArea">
	<div style="float:left"><%=ClubTree%> → <%=ForumTree(ParentID)%><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></div>
<%
if Moderated<>empty then
	filtrate=split(Moderated,"|")
	for i = 0 to ubound(filtrate)
		ModeratedList=ModeratedList&"<div class=menuitems><a href=Profile.asp?UserName="&filtrate(i)&">"&filtrate(i)&"</a></div>"
	next
%>
	<div style="float:right"><img src="images/team.gif" /> <a onmouseover="showmenu(event,'<%=ModeratedList%>')" style="cursor:default">论坛版主</a></div>
<%end if%>

</div>
<%
sql="Select * From ["&TablePrefix&"Forums] where ParentID="&ForumID&" and SortOrder>0 and IsActive=1 order by SortOrder"
Set Rs=Execute(sql)
	if not Rs.eof then
	%><table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="7"><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></td>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td width="30"></td>
		<td>论坛</td>
		<td width="50">主题</td>
		<td width="50">帖数</td>
		<td width="150">最后发表</td>
		<td width="100">版主</td>
	</tr>
	<%
	do while not Rs.eof
		ShowForum()
		Rs.Movenext
	loop
	response.write "</table><br />"
	end if
Rs.close


if ForumRules<>"" then
%>

<table cellpadding="8" class="CommonListCell" width="100%" style=" BORDER-RIGHT:#ccc 1px dotted; BORDER-TOP:#ccc 1px dotted; BORDER-LEFT:#ccc 1px dotted; BORDER-BOTTOM:#ccc 1px dotted;">
	<tr>
		<td><strong><font color="#ff0000">版 规 和 导 读</font></strong><br />
		<%=ForumRules%></td>
	</tr>
</table><br />
<%end if%>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="Left"><%if PermissionPost=1 then%>
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="AddTopic.asp?ForumID=<%=ForumID%>">
		发表新帖</a> <%end if%> <%if PermissionCreatePoll=1 then%>
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/Poll.gif)" href="AddTopic.asp?ForumID=<%=ForumID%>&Poll=1">
		发起投票</a> <%end if%> </td>
		<td align="right"><%if ModerateNewThread=1 then%>
		<a href="?ForumID=<%=ForumID%>&Order=Visible&SortOrder=1">
		主题审核(<%=Execute("Select count(ThreadID) From ["&TablePrefix&"Threads] where Visible=0 and ForumID="&ForumID&"")(0)%>)</a> 
		| <%end if%> <a href="?ForumID=<%=ForumID%>&GoodTopic=1">精华主题</a> | 
        <a href="?ForumID=<%=ForumID%>&VoteTopic=1">投票主题</a> |
		<a href="ShowForumPermissions.asp?ForumID=<%=ForumID%>">论坛权限</a> |
		<a href="ForumManage.asp?menu=ForumData&ForumID=<%=ForumID%>">版块管理</a> |
<script language="JavaScript" type="text/javascript">
var AllForumNameList=getCookie("ForumNameList");
var NowForumName="<option value='ShowForum.asp?ForumID=<%=ForumID%>'><%=ForumName%></option>";
if(AllForumNameList.indexOf(NowForumName)==-1){document.cookie= "ForumNameList" + "=" + escape(NowForumName+AllForumNameList);}
document.write("<select onchange='location=this.options[this.selectedIndex].value;'><option>最近访问的版块...</option>" + AllForumNameList+ "</select>");
</script>
		</td>
	</tr>
</table>


<%if PermissionManage=1 then%>
<form style="margin:0px;" method="POST" action="Manage.asp">
<input type="hidden" name="ForumID" value="<%=ForumID%>" />
<%end if%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="5">
		<div style="float:left">主题</div>
		<div style="float:right"><%=CategorysList%></div>
		</td>
		<%
		if PermissionManage=1 then
			Visible=" and (Visible=1 or Visible=0)"
			response.write("<td width=50 align=center><input type=checkbox name=chkall onclick='CheckAll(this.form)' value=ON></td>")
		else
			Visible=" and Visible=1"
		end if
		%>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>标题</td>
		<td>作者</td>
		<td>回复</td>
		<td>查看</td>
		<td>最后更新</td>
		<%if PermissionManage=1 then response.write("<td width=50>&nbsp;</td>")%>
	</tr>
	<%
	if RequestInt("PageIndex") < 2 then
		sql="["&TablePrefix&"Threads] where (ThreadTop=2 or ThreadTop=1 and ForumID="&ForumID&")"&Visible&" order by ThreadTop DESC"
		Set Rs=Execute(sql)
	
		if Not RS.EOF then
			Do While Not RS.EOF
				ShowThread()
				Rs.MoveNext
			loop
			if PermissionManage=1 then
				response.write("<tr class=CommonListHeader><td colspan=5>版块主题</td><td width=50>&nbsp;</td></tr>")
			else
				response.write "<tr class=CommonListHeader><td colspan=5>版块主题</td></tr>"
			end if
		end if

		Rs.Close
	end if

	if Order=empty then Order="lasttime"
	if Category<>"" then SQLCategory="and Category='"&Category&"'"
	if GoodTopic > 0 then SQLGoodTopic="and IsGood=1"
	if VoteTopic > 0 then SQLVoteTopic="and IsVote=1"
			
	if TimeLimit > 0 then SQLTimeLimit="and DateDiff("&SqlChar&"d"&SqlChar&",lasttime,"&SqlNowString&") < "&TimeLimit&""
		
	if SortOrder="1" then
		SqlSortOrder=""
	else
		SqlSortOrder="Desc"
	end if

	topsql="["&TablePrefix&"Threads] where ForumID="&ForumID&Visible&" and ThreadTop=0 "&SQLCategory&" "&SQLTimeLimit&" "&SQLGoodTopic&" "&SQLVoteTopic&""
	TotalCount=Execute("Select count(ThreadID) From "&topsql&" ")(0) '获取数据数量
	PageSetup=SiteConfig("ThreadsPerPage") '设定每页的显示数量
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
	PageCount = RequestInt("PageIndex") '获取当前页
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	
	sql="Select * from "&topsql&" order by "&order&" "&SqlSortOrder&""

	if PageCount<11 then
		Set Rs=Execute(sql)
	else
		rs.Open sql,Conn,1
	end if
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	i=0
	Do While Not RS.EOF and i<PageSetup
		i=i+1
		ShowThread()
		Rs.MoveNext
	loop
	Rs.Close
%>
</table>
<%if PermissionManage=1 then%>
<table cellspacing="0" cellpadding="5" border="0" width="100%">
	<tr>
		<td align="right">
		管理选项：<select name=menu size=1>
        	<%if BestRole = 1 then%>
			<optgroup label="公告">
				<option value="Top">设为公告</option>
				<option value="UnTop">取消公告</option>
			</optgroup>
            <%end if%>
			<optgroup label="解锁">
				<option value="IsLocked">锁定主题</option>
				<option value="DelIsLocked">解锁主题</option>
			</optgroup>
			<optgroup label="精华">
				<option value="IsGood">精华主题</option>
				<option value="DelIsGood">取消精华</option>
			</optgroup>
			<optgroup label="审核">
				<option value="Visible">审核通过</option>
				<option value="InVisible">审核失败</option>
			</optgroup>
			<optgroup label="置顶">
				<option value="ThreadTop">置顶主题</option>
				<option value="DelTop">取消置顶</option>
			</optgroup>
			<optgroup label="删除">
				<option value="DelThread">删除主题</option>
				<option value="UnDelThread">取消删除</option>
			</optgroup>
			<optgroup label="其它">
				<option value="Fix">修复主题</option>
				<option value="MoveNew">拉前主题</option>
				<option value="MoveThread">移动主题</option>
			</optgroup>
		</select>　<input type="submit" value=" 执 行 " onclick="return VerifyRadio('Item');" />
</td></form>
	</tr>
</table>
<%end if%>
<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td>
		<a onmousedown="ToggleMenuOnOff('ForumOption')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/ForumSettings.gif)" href="#ForumOption">
		选项</a>
		</td>
		<td align="right"><%ShowPage()%></td>
	</tr>
</table>
<div id="ForumOption" style="display:none;">
	<fieldset>
	<legend>选项</legend>
	<table border="0" width="100%">
		<tr>
			<td valign="top">
			<form name="form" action="ShowForum.asp?ForumID=<%=ForumID%>" method="POST">
				排序规则：<select name="order">
				<option value="">最后更新时间</option>
				<option value="ThreadID">主题发表时间</option>
				<option value="IsGood">精华帖子</option>
				<option value="IsVote">投票帖子</option>
				<option value="Topic">主题</option>
				<option value="PostAuthor">作者</option>
				<option value="TotalViews">点击数</option>
				<option value="TotalReplies">回复数</option>
				</select> 根据 <select name="SortOrder">
				<option value="0" selected="selected">降序</option>
				<option value="1">升序</option>
				</select> 排列<br />
				日期过滤：<select name="TimeLimit">
				<option value="1" >昨天以来</option>
				<option value="2" >2 天以来</option>
				<option value="7" >1 周以来</option>
				<option value="10" >10 天以来</option>
				<option value="14" >2 周以来</option>
				<option value="30" >1 个月以来</option>
				<option value="45" >45 天以来</option>
				<option value="60" >2 个月以来</option>
				<option value="75" >75 天以来</option>
				<option value="100" >100 天以来</option>
				<option value="365">1 年以来</option>
				<option value="-1">任何时间</option>
				</select><br />
				<input type="submit" value=" 应用 " /></form>
			</td>
			<td valign="top" align="right">您 <b><%if PermissionPost=0 then%>不<%end if%>能</b> 
			在此版发表新主题<br />
			您 <b><%if PermissionReply=0 then%>不<%end if%>能</b> 在此版回复主题<br />
			您 <b><%if PermissionEdit=0 then%>不<%end if%>能</b> 在此版修改您发表的帖子<br />
			您 <b><%if PermissionDelete=0 then%>不<%end if%>能</b> 在此版删除您发表的帖子<br />
			您 <b><%if PermissionCreatePoll=0 then%>不<%end if%>能</b> 在此版发起投票<br />
			您 <b><%if PermissionVote=0 then%>不<%end if%>能</b> 在此版参与投票<br />
			您 <b><%if PermissionAttachment=0 then%>不<%end if%>能</b> 在此版上传附件<br />
			该版主题<b><%if ModerateNewThread=0 then%>不<%end if%>需要审核</b><br />
			该版帖子<b><%if ModerateNewPost=0 then%>不<%end if%>需要审核</b></td>
		</tr>
	</table>
	</fieldset> </div>
	
	
<!-- #include file="Utility/OnLine.asp" -->
<%
if SiteConfig("DisplayForumUsers")=1 then
	ForumIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ForumID="&ForumID&"")(0)
	regForumIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ForumID="&ForumID&" and UserName<>''")(0)
%>
	
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>用户在线信息</td>
	</tr>
	<tr class="CommonListCell">
		<td>
		<img src="images/plus.gif" id="followImg" style="cursor:pointer;" onclick="loadThreadFollow('ForumID=<%=ForumID%>')" /> 
		目前论坛总在线 <b><%=Onlinemany%></b> 人，本分论坛共有 <b><%=ForumIDOnline%></b> 人在线。其中注册用户 
		<b><%=regForumIDOnline%></b> 人，访客 <b><%=ForumIDOnline-regForumIDOnline%></b> 
		人。<div style="display:none" id="follow">
			<hr width="90%" size="1" align="left"><span id="followTd" class="UserList">
			<img src="images/loading.gif" />正在加载...</span></div>
		</td>
	</tr>
</table>
<br />
<%end if%>
<table border="0" width="90%" align="center">
	<tr>
		<td><img src="images/topic-announce.gif" border="0" align="absmiddle" /> 公告主题</td>
		<td><img src="images/topic-pinned.gif" border="0" align="absmiddle" /> 置顶主题</td>
		<td><img src="images/topic-popular.gif" align="absmiddle" /> 精华主题</td>
		<td><img src="images/topic-locked.gif" border="0" align="absmiddle" /> 锁定主题</td>
		<td><img src="images/topic-poll.gif" border="0" align="absmiddle" /> 投票主题</td>
		<td><img src="images/topic-hot.gif" border="0" title="回复数达到 <%=SiteConfig("PopularPostThresholdPosts")%> 或者点击数达到 <%=SiteConfig("PopularPostThresholdViews")%>" align="absmiddle" /> 热门主题</td>
		<td><img src="images/topic.gif" border="0" align="absmiddle" /> 普通主题</td>
	</tr>
</table>
<script language="JavaScript" type="text/javascript">
function VerifyRadio() {
	objYN=false;
	if (window.confirm('您确定执行本次操作?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
		if (objYN==false) {alert ('请选择您要操作的主题！');return false;}
	}
	return objYN;
}
</script>
<%
HtmlBottom
%>