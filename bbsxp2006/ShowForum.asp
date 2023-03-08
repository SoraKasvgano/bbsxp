<!-- #include file="Setup.asp" --><%
top

ForumID=int(Request("ForumID"))
order=HTMLEncode(Request("order"))
SpecialTopic=HTMLEncode(Request("SpecialTopic"))
SortOrder=Request("SortOrder")
TimeLimit=Request("TimeLimit")
if Len(order)>10 then error("<li>非法操作")
if Len(SpecialTopic)>20 then error("<li>专题名字太长")

sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
if Rs.eof then error"<li>该论坛已经被删除"
ForumName=Rs("ForumName")
ForumThreads=Rs("ForumThreads")
moderated=Rs("moderated")
ForumLogo=Rs("ForumLogo")
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
				→ <%ForumTree(followid)%><%=ForumTreeList%>
				<a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></td>
				<td height="18" align="right"><img src="images/jt.gif">
				<a href="ForumManage.asp?menu=ForumData&ForumID=<%=ForumID%>">高级管理</a> </td>
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
<strong><font color="#ff0000">版 规 和 导 读</font></strong><br><%=ForumRules%></div><br>
<%end if%>

<!-- #include file="inc/line.asp" --><%


if IsShowForum=1 or SiteSettings("SortShowForum")=1 then

ForumIDOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where ForumID="&ForumID&"")(0)
regForumIDOnline=Conn.execute("Select count(sessionid)from [BBSXP_UsersOnline] where ForumID="&ForumID&" and UserName<>''")(0)
%><table cellspacing="1" cellpadding="0" width="100%" align="center" border="0" class="a2">
	<tr>
		<td width="93%" height="25" class="a1">　<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=ForumID%>)"> 
		目前论坛总在线 <b><%=Onlinemany%></b> 人，本分论坛共有 <b><%=ForumIDOnline%></b> 人在线。其中注册用户 
		<b><%=regForumIDOnline%></b> 人，访客 <b><%=ForumIDOnline-regForumIDOnline%></b> 
		人。</td>
		<td align="middle" width="7%" height="25" class="a1">
		<a href="javascript:this.location.reload()">
		<img src="images/refresh.gif" border="0"></a></td>
	</tr>
	<tr height="25" style="display:none" id="follow0">
		<td id="followTd0" align="Left" class="a4" width="94%" colspan="5">　Loading...</td>
	</tr>
	</tr>
</table>
<br>
<table height="30" cellspacing="3" cellpadding="0" width="100%" align="center" border="0">
	<tr>
		<td align="Left" width="20%">
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="NewTopic.asp?ForumID=<%=ForumID%>">
		发表新主题</a> </td>
		<td align="right" width="80%">
		<img src="images/Showdigest.gif">
		<a onmouseover="showmenu(event,'&lt;div class=menuitems&gt;&lt;a href=MyFavorites.asp?menu=add&amp;url=forum&amp;name=<%=ForumID%>&gt;收藏论坛&lt;/a&gt;&lt;/div&gt;&lt;div class=menuitems&gt;&lt;a href=MyFavorites.asp?menu=Del&amp;url=forum&amp;name=<%=ForumID%>&gt;取消收藏&lt;/a&gt;&lt;/div&gt;')" style="cursor:default">
		论坛收藏</a>
<%
if moderated<>empty then
filtrate=split(moderated,"|")
for i = 0 to ubound(filtrate)
ModeratedList=ModeratedList&"<div class=menuitems><a href=Profile.asp?UserName="&filtrate(i)&">"&filtrate(i)&"</a></div>"
next
%><img src=images/team.gif> <a onmouseover="showmenu(event,'<%=ModeratedList%>')" style=cursor:default>
		论坛版主</a>
<%end if%> <a href="Rss.asp?ForumID=<%=ForumID%>"><img src="images/rss_button.gif" border="0" alt="RSS 订阅当前论坛"></a>
		</td>
	</tr>
</table>



<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
	<%
if TolSpecialTopic<>empty then
response.write "<tr height=25 class=a3><td width=100% colSpan=7>专题："
filtrate=split(TolSpecialTopic,"|")
for i = 0 to ubound(filtrate)
response.write "<font face='Old English Text MT'><b>"&i+1&"</b></font>[<a href='ShowForum.asp?ForumID="&ForumID&"&SpecialTopic="&filtrate(i)&"'>"&filtrate(i)&"</a>] "
TolSpecialTopicOptionList=TolSpecialTopicOptionList&"<option value="&filtrate(i)&">"&filtrate(i)&"</option>"
next
response.write "</td></tr>"
end if
%>
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">主题</td>
<td align="center" width="10%">作者</td>
<td align="center" width="6%">回复</td>
<td align="center" width="6%">点击</td>
<td align="center" width="25%">最后更新</td>
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
TotalCount=conn.Execute("Select count(ID) From "&topsql&" ")(0) '获取数据数量
else
TotalCount=ForumThreads  '获取数据数量
end if

PageSetup=SiteSettings("ThreadsPerPage") '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = cint(Request.QueryString("PageIndex")) '获取当前页
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
<td colspan="7"><input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">全选
<input type="radio" value="BatchDel" name=menu><select name=IsDel>
<option value="1">删除帖子</option>
<option value="0">取消删除</option>
</select>&nbsp;

<input type="radio" value="BatchGOOD" name=menu><select name=IsGOOD>
<option value="1">加入精华</option>
<option value="0">取消精华</option>
</select>&nbsp;

<input type="radio" value="BatchLocked" name=menu><select name=IsLocked>
<option value="1">锁定</option>
<option value="0">解锁</option>
</select>

<input type="radio" value="BatchSpecialTopic" name=menu><select name=SpecialTopic>
<option value="">取消专题</option>
<option selected value="">更改专题</option>
<%=TolSpecialTopicOptionList%>
</select>

<input type="radio" value="BatchMoveTopic" name=menu><select name=AimForumID>
<option selected value="">移动到以下论坛</option>
<%=ForumsList%>
</select>&nbsp;


<input onclick="checkclick('您确定执行本次操作?');" type="submit" value=" 执 行 ">
</td></form>
</tr>

<%end if%>
</table>

<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
	<tr>
		<td>
		<a onmousedown="ToggleMenuOnOff('ForumOption')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/ForumSettings.gif)" href="#ForumOption">
		选项</a>
		<a onmousedown="ToggleMenuOnOff('ForumSearch')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/finds.gif)" href="#ForumSearch">
		搜索</a>
		<div id="ForumSearch" style="position:absolute;display:none;">
			<form name="form" action="Search.asp?menu=ok&ForumID=<%=ForumID%>&Search=key&sessionid=<%=session.sessionid%>" method="POST">
				<input name="content" size="20" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')">
				<input type="submit" value="搜索" id="btnadd" disabled>
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
			排序规则：<select name="order">
			<option value="">最后更新时间</option>
			<option value="id">主题发表时间</option>
			<option value="IsGood">精华帖子</option>
			<option value="IsVote">投票帖子</option>
			<option value="Topic">主题</option>
			<option value="UserName">作者</option>
			<option value="Views">点击数</option>
			<option value="Replies">回复数</option>
			</select> 根据 <select name="SortOrder">
			<option value="0" selected>降序</option>
			<option value="1">升序</option>
			</select> 排列<br>
			日期过滤：<select name="TimeLimit">
			<option value="">显示所有</option>
			<option value="1">一天以来</option>
			<option value="2">两天以来</option>
			<option value="3">三天以来</option>
			<option value="7">一星期以来</option>
			<option value="14">两星期以来</option>
			<option value="30">一个月以来</option>
			<option value="60">两个月以来</option>
			<option value="90">三个月以来</option>
			<option value="180">半年以来</option>
			</select><br>
			<input type="submit" value=" 应用 "></form></td>
	</tr>
</table>
<%
end if


htmlend
%>