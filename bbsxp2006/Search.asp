<!-- #include file="Setup.asp" --><%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")


if Request("menu")="ok" then

if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")

Search=Request("Search")
ForumID=Request("ForumID")
TimeLimit=Request("TimeLimit")
content=HTMLEncode(Request("content"))
Searchxm=HTMLEncode(Request("Searchxm"))


if content="" then error("<li>您没有输入关键词")


if isnumeric(""&ForumID&"") then ForumIDor="ForumID="&ForumID&" and"

if Search="author" then
if Len(Searchxm)>9 then error("<li>非法操作")
item=""&Searchxm&"='"&content&"'"
elseif Search="key" then
item="Topic like '%"&content&"%'"
end if

if TimeLimit<>"" then TimeLimitList="and lasttime>"&SqlNowString&"-"&int(TimeLimit)&""


sql="select * from [BBSXP_Threads] where IsDel=0 and "&ForumIDor&" "&item&" "&TimeLimitList&" order by lasttime Desc"
Rs.Open sql,Conn,1

count=Rs.recordcount    '数据总条数
if Count=0 then error("<li>对不起，没有找到您要查询的内容")

PageSetup=SiteSettings("ThreadsPerPage") '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数


%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ 搜索结果</td>
	</tr>
</table>
<br>
<table width="100%" align="center">
	<tr>
		<td width="100%" align="right">搜索关键词：<font color="FF0000"><%=content%></font>　　共找到了 
		<b><font color="FF0000"><%=Count%></font></b> 篇相关帖子</td>
	</tr>
</table>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">主题</td>
<td align="center" width="10%">作者</td>
<td align="center" width="6%">回复</td>
<td align="center" width="6%">点击</td>
<td align="center" width="25%">最后更新</td>
</tr>
	<%
i=0
Do While Not RS.EOF and i<pagesetup
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
				<td height="2">
				<%
ShowPage()%>
</td>
				<form name="form" action="Search.asp?menu=ok&ForumID=<%=ForumID%>&Search=key" method="POST"><input type=hidden name=sessionid value=<%=session.sessionid%>>
					<td height="2" align="right">快速搜索：<input name="content" size="20" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="关键词" Helptext="关键词">
					<input type="submit" value="搜索" id="btnadd" disabled>
					</td>
				</form>
			</tr>
		</table>
		</td>
	</tr>
</table>

<%
htmlend


end if
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ 搜索帖子</td>
	</tr>
</table>
<br>

<table height="207" cellspacing="1" cellpadding="3" width="100%" class="a2" border="0" align="center">
	<form method="POST" action="Search.asp?menu=ok" name="form"><input type=hidden name=sessionid value=<%=session.sessionid%>>
		<tr>
			<td colspan="2" height="25" class="a1" align="center">请输入要搜索的关键词</td>
		</tr>
		<tr class=a3>
			<td valign="top" colspan="2" height="8">
			<p align="center">
			<input size="40" name="content" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')"></p>
			</td>
		</tr>
		<tr>
			<td class="a1" colspan="2" height="25" align="center">搜索选项</td>
		</tr>
		<tr class=a3>
			<td width="41%" height="24" align="right">关键词搜索<input type="radio" value="key" name="Search" checked></td>
			<td height="25" width="58%">&nbsp;<select size=1>
			<option>在主题中搜索关键词</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="21" align="right">作者搜索<input type="radio" value="author" name="Search" id="Search"></td>
			<td height="25" width="58%">&nbsp;<select size="1" name="Searchxm">
			<option selected value="UserName">搜索主题作者</option>
			<option value="lastname">搜索最后回复作者</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="23" align="right">日期范围</td>
			<td height="25" width="58%">&nbsp;<select size="1" name="TimeLimit">
			<option value="">所有日期</option>
			<option value="1">昨天以来</option>
			<option value="5" selected>5天以来</option>
			<option value="10">10天以来</option>
			<option value="30">30天以来</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="26" align="right">请选择要搜索的论坛</td>
			<td height="26" width="58%">&nbsp;<select name="ForumID" size="1">
			<option value="" selected>全部论坛</option>
			<%BBSList(0)%><%=ForumsList%></select>
			<input type="submit" value="开始搜索" id="btnadd" disabled></td>
		</tr>
	</form>
</table>

</center><%
htmlend
%>