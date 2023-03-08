<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
UserName=Trim(HTMLEncode(Request("UserName")))

if Request.Form("menu")="ok" then

if membercode < 4 then error("<li>您的权限不够，无法抓人入狱！")

If Conn.Execute("Select UserName From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then error("<li>"&UserName&"的资料不存在")
If not Conn.Execute("Select UserName From [BBSXP_Prison] where UserName='"&UserName&"'" ).eof Then error("<li>"&UserName&"已经被关进监狱")


if UserName="" then error("<li>犯人的名称没有添写")


causation=HTMLEncode(Request("causation"))
PrisonDay=HTMLEncode(Request("PrisonDay"))

if causation="" then error("<li>您没有输入入狱原因")
if PrisonDay>1000 then error("<li>坐牢时间不能超过1000天")

sql="insert into [BBSXP_Prison] (UserName,causation,constable,PrisonDay) values ('"&UserName&"','"&causation&"','"&CookieUserName&"','"&PrisonDay&"')"
Conn.Execute(SQL)



end if

if Request("menu")="release" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")

if membercode < 4 then error("<li>您的权限不够，无法释放犯人！")

Conn.execute("Delete from [BBSXP_Prison] where UserName='"&UserName&"'")

Log("将 "&UserName&" 释放出监狱！")

error2("已经将 "&UserName&" 释放出监狱！")
end if


if Request("menu")="look" then
sql="select * from [BBSXP_Prison] where UserName='"&UserName&"'"
Set Rs=Conn.Execute(sql)


%>
<title>探 监 - Powered By BBSXP</title>
<b><%=Rs("UserName")%></b>
<SCRIPT>
var tips=["斜着眼睛瞟了一眼看守,嘟哝着:最近的点儿太背!","两眼汪汪的说:都是我不好!对不起大家了!","脸上露出诡异的笑容:嘿嘿!要不要进来看看!","感慨万分道:一失足,成千古恨!我一定重新做人!","望着布满电网和铁丝网的高墙,摇头叹息着!"]
index = Math.floor(Math.random() * tips.length);
document.write("" + tips[index] + "");
  </SCRIPT><br><br>
入狱原因：<%=Rs("causation")%><br><br>
入狱时间：<%=Rs("cometime")%><br><br>
出狱时间：<%=Rs("cometime")+Rs("PrisonDay")%><br><br>
执行警官：<%=Rs("constable")%>

<%
Rs.close
CloseDatabase

end if

top

Conn.execute("Delete from [BBSXP_Prison] where cometime<"&SqlNowString&"-PrisonDay")
%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="Prison.asp">社区监狱</a></td>
</tr>
</table><br>



<TABLE cellSpacing=1 cellPadding=3 border=0 width=100% align=center class=a2><TR class=a1 height="25">
	<TD align=center width="10%"><b>用户名</b></TD>
<TD align=center><b>入狱原因</b></TD>
<TD align=center width="20%"><b>入狱时间</b></TD>
<TD align=center width="10%"><b>执行警官</b></TD>
<TD align=center width="15%"><b>动作</b></TD>
</TR>

<%
Rs.Open "[BBSXP_Prison] order by cometime Desc",Conn,1
PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1


response.write "<tr class=a4><TD align=center><a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a></TD><TD>"&Rs("causation")&"</TD><TD align=center>"&Rs("cometime")&"</TD><TD align=center><a href=Profile.asp?UserName="&Rs("constable")&">"&Rs("constable")&"</a></TD><TD align=center><a href=?menu=release&sessionid="&session.sessionid&"&UserName="&Rs("UserName")&">释 放</a> | <a href=# onClick=javascript:open('Prison.asp?menu=look&UserName="&Rs("UserName")&"','','resizable,scrollbars,width=220,height=180')>探 监</a></TD></tr>"


Rs.MoveNext
loop
Rs.Close

%>

<form METHOD=Post><input type=hidden name=menu value=ok><tr>
  <TD align=center colspan="5" class=a3>将
<input name="UserName" size="13"> 抓入监狱　　坐牢天数：<input name="PrisonDay" size="2" value="15"><br>
入狱原因：<input name="causation" size="33"> <input type="submit" value=" 确 定 "></TD>
  			</tr></form>
</TABLE>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>



<%
htmlend
%>