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
Affiche=Affiche&"<b>"&Rs("title")&"</b> ("&Rs("Posttime")&")������"
Rs.MoveNext
loop
Set Rs = Nothing

%> �����ɹ�<br><br><a href=javascript:history.back()>�� ��</a><%

case "DelAffiche"
Conn.execute("Delete from [BBSXP_Affiche] where id="&id&"")
sql="select top 2 * from [BBSXP_Affiche] order by Posttime Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
Affiche=Affiche&"<b>"&Rs("title")&"</b> ("&Rs("Posttime")&")������"
Rs.MoveNext
loop
Set Rs = Nothing

%> ɾ���ɹ�<br><br><a href=javascript:history.back()>�� ��</a><%


end select

sub Affichelist
%>

<a href="?menu=addAffiche">��������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="javascript:this.location.reload()">ˢ���б�</a><br>
��<table border="0" cellpadding="5" cellspacing="1" class=a2 width=100%>
<tr>
<td width="5%" align="center" height="25" class=a1>ID</td>
<td width="35%" align="center" height="25" class=a1>����</td>
<td width="10%" align="center" height="25" class=a1>������</td>
<td width="15%" align="center" height="25" class=a1>����ʱ��</td>
<td width="15%" align="center" height="25" class=a1>����</td>
</tr>
<%
sql="select * from [BBSXP_Affiche] order by Posttime Desc"
Rs.Open sql,Conn,1
PageSetup=20 '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1%>
<tr class=a3>
<td height="25" align="center"> <%=Rs("id")%>
<td height="25" align="center"><%=Rs("title")%>
<td height="25" align="center"><a target="_blank" href="Profile.asp?UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a>
<td height="25" align="center"><%=Rs("Posttime")%>
<td height="25" align="center"><a href="?menu=addAffiche&id=<%=Rs("id")%>">�޸Ĺ���</a> <a onclick=checkclick('��ȷ��Ҫɾ���ù��棿') href="?menu=DelAffiche&id=<%=Rs("id")%>">ɾ������</a></td>
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
    <td class=a1 align=middle colspan=2>��������</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
���⣺</td>
    <td class=a3 width="82%">
<input type="text" name="Subject" size="60" value="<%=title%>"></td></tr>
   <tr height=25>
    <td class=a3 align=middle width="16%">
���ݣ�</td>
    <td class=a3 width="82%" height="250">
    
    <SCRIPT src="inc/Post.js"></SCRIPT>

</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<input type="submit" value=" �� �� " name=EditSubmit>&nbsp;
<input type="reset" value=" �� �� ">
</td></tr></table></form>
<a href=javascript:history.back()>�� ��</a>
<%
end sub


htmlend

%>