<!-- #include file="Setup.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))


response.write "<center>"

select case Request("menu")

case "Affichelist"
Affichelist


case "addAffiche"
addAffiche

case "addAfficheok"
sql="select * from [BBSXP_Affiche] where id="&id&""
Rs.Open sql,Conn,1,3
if Rs.eof then Rs.addNew
Rs("title")=""&Request("Subject")&""
Rs("content")=replace(replace(Request("content"),vbCrlf,""),"'","&#39;")
Rs("UserName")=""&CookieUserName&""
Rs("Posttime")=date()
Rs.update
Rs.close
sql="select top 2 * from [BBSXP_Affiche] order by Posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
Affiche=Affiche&"<b>"&Rs("title")&"</b> ("&Rs("Posttime")&")　　　"
Rs.MoveNext
loop
Set Rs = Nothing

%> 发布成功<br><br><a href=javascript:history.back()>返 回</a><%

case "DelAffiche"
Conn.execute("Delete from [BBSXP_Affiche] where id="&id&"")
sql="select top 2 * from [BBSXP_Affiche] order by Posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
Affiche=Affiche&"<b>"&Rs("title")&"</b> ("&Rs("Posttime")&")　　　"
Rs.MoveNext
loop
Set Rs = Nothing

%> 删除成功<br><br><a href=javascript:history.back()>返 回</a><%


end select

sub Affichelist
%>

<a href="?menu=addAffiche">发布公告</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="javascript:this.location.reload()">刷新列表</a><br>
　<table border="0" cellpadding="5" cellspacing="1" class=a2 width=100%>
<tr>
<td width="5%" align="center" height="25" class=a1>ID</td>
<td width="35%" align="center" height="25" class=a1>标题</td>
<td width="10%" align="center" height="25" class=a1>发布人</td>
<td width="15%" align="center" height="25" class=a1>发布时间</td>
<td width="15%" align="center" height="25" class=a1>管理</td>
</tr>
<%
sql="select * from [BBSXP_Affiche] order by Posttime Desc"
Rs.Open sql,Conn,1
PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1%>
<tr class=a3>
<td height="25" align="center"> <%=Rs("id")%>
<td height="25" align="center"><%=Rs("title")%>
<td height="25" align="center"><a target="_blank" href="Profile.asp?UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a>
<td height="25" align="center"><%=Rs("Posttime")%>
<td height="25" align="center"><a href="?menu=addAffiche&id=<%=Rs("id")%>">修改公告</a> <a onclick=checkclick('您确定要删除该公告？') href="?menu=DelAffiche&id=<%=Rs("id")%>">删除公告</a></td>
</tr>
<%
Rs.MoveNext
loop
Rs.Close
%>
</table>
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>
<%
end sub
sub addAffiche


if Request("id")<>empty then
sql="select * from [BBSXP_Affiche] where id="&id&""
Set Rs=Conn.Execute(sql)
content=Rs("content")
title=Rs("title")
end if

%>
<form name="yuziform" method="POST" action="?menu=addAfficheok" onSubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=content%>'>
<input name="id" type="hidden" value='<%=id%>'>

<table cellspacing="1" cellpadding="2" width="90%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>发布公告</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
标题：</td>
    <td class=a3 width="82%">
<input type="text" name="Subject" size="60" value="<%=title%>"></td></tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
内容：</td>
    <td class=a3 width="82%" height="250">
    
    <SCRIPT src="inc/Post.js"></SCRIPT>

</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<input type="submit" value=" 发 送 " name=EditSubmit>&nbsp;
<input type="reset" value=" 重 置 ">
</td></tr></table></form>
<a href=javascript:history.back()>返 回</a>
<%
end sub


htmlend

%>