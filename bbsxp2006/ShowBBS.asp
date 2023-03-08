<!-- #include file="Setup.asp" -->
<%
UserName=HTMLEncode(Request("UserName"))
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
%>


<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <span id=bbsxpname></span></td>
</tr>
</table><br>

<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><a href="?">最新帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=1">本周人气帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=2">本周热门帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=3">精华帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=4">投票帖子</a></td>
      <TD width="14%" align="center"><a href="?menu=5">我的帖子</a></td>
      </TR></TABLE>


<br>


<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">主题</td>
<td align="center" width="10%">作者</td>
<td align="center" width="6%">回复</td>
<td align="center" width="6%">点击</td>
<td align="center" width="25%">最后更新</td>
</tr>
<%
select case Request("menu")
case ""
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="最新帖子"</SCRIPT><%
case "1"
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 and PostTime>"&SqlNowString&"-7 order by Views Desc"
%><SCRIPT>bbsxpname.innerText="本周人气帖子"</SCRIPT><%
case "2"
sql="select top 200 * from [BBSXP_Threads] where IsDel=0 and PostTime>"&SqlNowString&"-7 order by replies Desc"
%><SCRIPT>bbsxpname.innerText="本周热门帖子"</SCRIPT><%
case "3"
sql="select top 200 * from [BBSXP_Threads] where IsGood=1 and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="精华帖子"</SCRIPT><%
case "4"
sql="select top 200 * from [BBSXP_Threads] where IsVote=1 and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="投票帖子"</SCRIPT><%
case "5"
sql="select top 200 * from [BBSXP_Threads] where UserName='"&UserName&"' and IsDel=0 order by id Desc"
%><SCRIPT>bbsxpname.innerText="<%=UserName%> 的帖子"</SCRIPT><%
end select

Rs.Open sql,Conn,1

PageSetup=SiteSettings("ThreadsPerPage") '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
ShowThread()
Rs.MoveNext
loop
Rs.Close
%></table>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td></tr></table>
<%
htmlend
%>