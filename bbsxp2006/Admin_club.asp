<!-- #include file="Setup.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))

TimeLimit=HTMLEncode(Request("TimeLimit"))
UserName=HTMLEncode(Request("UserName"))
membercode=HTMLEncode(Request("membercode"))


response.write "<center>"

select case Request("menu")


case "Message"
Message

case "broadcast"
broadcast

case "sendMail"
sendMail

case "sendMailok"
sendMailok



case "Messageok"
if TimeLimit="" then error2("您没有选择日期！")
Conn.execute("Delete from [BBSXP_Messages] where DateCreated<"&SqlNowString&"-"&TimeLimit&"")
error2("已经将"&TimeLimit&"天以前的短讯息删除了！")


case "DelMessageUser"
if UserName="" then error2("您没有输入用户名！")
Conn.execute("Delete from [BBSXP_Messages] where UserName='"&UserName&"' or incept='"&UserName&"'")
error2("已经将"&UserName&"的短讯息全部删除了！")

case "DelMessagekey"
key=HTMLEncode(Request("key"))
if key="" then error2("您没有输入关键词！")
Conn.execute("Delete from [BBSXP_Messages] where content like '%"&key&"%'")
error2("已经将内容中包含 "&key&" 的短讯息删除了！")


case "Consortia"
Consortia

case "editConsortia"
editConsortia

case "editConsortiaok"
editConsortiaok

case "DelConsortia"
DelConsortia


case "Link"
Link

case "Linkok"
Linkok

case "editLink"
editLink


case "editLinkok"
Rs.Open "select * from [BBSXP_Link] where id="&id&"",Conn,1,3
Rs("name")=Request("name")
Rs("url")=Request("url")
Rs("Logo")=Request("Logo")
Rs("Intro")=Request("Intro")
Rs.update
Rs.close
%>编 辑 成 功 ！<p><a href=javascript:history.back()>< 返 回 ></a><%


case "DelLink"
Conn.execute("Delete from [BBSXP_Link] where id="&id&"")
Link



end select


sub sendMail
%>

<form method="POST" action="?menu=sendMailok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>群发邮件</td>
  </tr>
  <tr height=25>
    <td class=a3 align=Left>&nbsp;&nbsp; 主题：<input size="40" name="title"></td>
    <td class=a3 align=middle>接收对象：
<select name=membercode>
<option value="">所有会员</option>
<option value="1">普通会员</option>
<option value="2">贵宾会员</option>
<option value="4">超级版主</option>
<option value="5">管理员</option>
</select>&nbsp;&nbsp;&nbsp; </td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <input type="submit" value=" 发 送 ">
<input type="reset" value=" 重 填 "><br></td></tr></table></form>

<%
end sub


sub sendMailok

if Request("title")="" then error2("请填写邮件主题")
if Request("content")="" then error2("请填写邮件内容")


if membercode<>"" then
sql="select UserMail from [BBSXP_Users] where membercode="&membercode&""
else
sql="select UserMail from [BBSXP_Users]"
end if

Set Rs=Conn.Execute(sql)
do while not Rs.eof

Mailaddress=""&Rs("UserMail")&""
MailTopic=Request("title")
body=""&Request("content")&""&vbCrlf&""&vbCrlf&"该邮件通过 BBSXP 群发系统发送　程序制作：YUZI工作室(http://www.yuzi.net)"
%><!-- #include file="inc/Mail.asp" --><%

Rs.Movenext
loop
Rs.close

response.write "邮件发送成功！"

end sub


sub Message
%>




数据库共 <%=Conn.execute("Select count(id)from [BBSXP_Messages]")(0)%> 条短讯息
<br><br>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
		<td align="center">批量删除短消息</td>
	</tr>
	<tr class=a3>
		<td align="center"><form method="POST" action="?menu=DelMessageUser">批量删除 <input size="13" name="UserName" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="用户名" Helptext="用户名"> 的短讯息
<input type="submit" value="确定" id='btnadd' disabled>
		</td></form>
	</tr>
	
	<tr class=a3>
		<td align="center"><form method="POST" action="?menu=DelMessagekey">批量删除内容含有 <input size="20" name="key" onkeyup="ValidateTextboxAdd(this, 'nrkey')" onpropertychange="ValidateTextboxAdd(this, 'nrkey')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="关键词" Helptext="关键词"> 的短讯息
<input type="submit" value="确定" id='nrkey' disabled>
		</td></form>
	</tr>
	
		<tr class=a3>
		<td align="center"><form method="POST" action="?menu=Messageok">删除 <INPUT size=2 name=TimeLimit value="30">
			天以前的短讯息
<input type="submit" value="确定">

		</td></form>
	</tr>
	
</table>
</form>
<form method="POST" action="?menu=broadcast">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle width="50%">系统广播</td>
    <td class=a1 align=middle width="50%">接收对象：
<select name=membercode>
<option value="">在线会员</option>
<option value="1">普通会员</option>
<option value="2">贵宾会员</option>
<option value="4">超级版主</option>
<option value="5">管理员</option>
</select>
</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<input type="submit" value=" 发 送 ">
<input type="reset" value=" 重 填 "></td></tr></table></form>
<%
end sub

sub broadcast
content=HTMLEncode(Request.Form("content"))

if content="" then error2("请填写广播内容!")

if membercode<>"" then
sql="select UserName from [BBSXP_Users] where membercode="&membercode&""
else
sql="select UserName from [BBSXP_UsersOnline] where UserName<>''"
end if

Set Rs=Conn.Execute(sql)
do while not Rs.eof
Count=Count+1
Conn.Execute("insert into [BBSXP_Messages] (UserName,incept,content) values ('"&CookieUserName&"','"&Rs("UserName")&"','<font color=0000FF>【系统广播】："&content&"</font>')")
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&Rs("UserName")&"'")
Rs.Movenext
loop
Rs.close

%>
发布成功
<br><br>
共发送给 <%=Count%> 位在线用户<br><br>
<a href=javascript:history.back()>返 回</a>
<%
end sub




sub Link
%>
<FORM action=?menu=Linkok method=Post>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align=center><tr>
	<td height="7" class="a1" colspan="2">　■<b> </b>友情链接管理</td></tr><tr>
	<td height="6" class="a3">网站名称：<INPUT size=40 name=name></td>
	<td height="6" class="a3">地址URL：<INPUT size=40 name=url value="http://"></td></tr><tr>
	<td height="6" class="a3">网站简介：<INPUT size=40 name=Intro></td>
	<td height="6" class="a3">图标URL：<INPUT size=40 name=Logo value="http://"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" 添 加 ">
<input type="reset" value=" 重 填 ">

</td></tr></table>
</FORM>


<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center"><tr><td height="25" colspan="2" class="a1">　■<b> </b>友情链接</td></tr>
<tr class=a3>
<td align="center" width="5%"><img src="images/shareforum.gif"></td>
<td class="a4"><%
Rs.Open "[BBSXP_Link]",Conn
do while not Rs.eof
if Rs("Logo")="" or Rs("Logo")="http://" then
Link1=Link1+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editLink&id="&Rs("id")&">编辑</a></div><div class=menuitems><a href=?menu=DelLink&id="&Rs("id")&">删除</a></div>')"&Chr(34)&" title='"&Rs("Intro")&"' href="&Rs("url")&" target=_blank>"&Rs("name")&"</a>　　"
else
Link2=Link2+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editLink&id="&Rs("id")&">编辑</a></div><div class=menuitems><a href=?menu=DelLink&id="&Rs("id")&">删除</a></div>')"&Chr(34)&" title='"&Rs("name")&""&chr(10)&""&Rs("Intro")&"' href="&Rs("url")&" target=_blank><img src="&Rs("Logo")&" border=0 width=88 height=31></a>　　"
end if
Rs.Movenext
loop
Rs.close
%>
<%=Link1%>
<br><br>
<%=Link2%>
</td></tr></table>


<%


end sub



sub editLink

sql="Select * From [BBSXP_Link] where id="&id&""
Set Rs=Conn.Execute(sql)
%>
<FORM action=?menu=editLinkok method=Post>
<input type=hidden name=id value=<%=id%>>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align=center><tr>
	<td height="7" class="a1" colspan="2">　■<b> </b>友情链接管理</td></tr><tr>
	<td height="6" class="a3">网站名称：<INPUT size=40 name=name value="<%=Rs("name")%>"></td>
	<td height="6" class="a3">地址URL：<INPUT size=40 name=url value="<%=Rs("url")%>"></td></tr><tr>
	<td height="6" class="a3">网站简介：<INPUT size=40 name=Intro value="<%=Rs("Intro")%>"></td>
	<td height="6" class="a3">图标URL：<INPUT size=40 name=Logo value="<%=Rs("Logo")%>"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" 编 辑 ">
<input type="reset" value=" 重 填 ">
</td></tr></table>
</FORM><p><a href=javascript:history.back()>< 返 回 ></a>
<%
end sub



sub Linkok

if Request("url")="http://" or Request("url")="" then error2("论坛URL没有填写")

Rs.Open "[BBSXP_Link]",Conn,1,3
Rs.addNew
Rs("name")=Request("name")
Rs("url")=Request("url")
Rs("Logo")=Request("Logo")
Rs("Intro")=Request("Intro")
Rs.update
Rs.close

Link
end sub


sub Consortia
%>
<table border="0" cellpadding="5" cellspacing="1" class=a2 width=100%>
<tr>
<td width="15%" align="center" height="25" class=a1>公会</td>
<td width="40%" align="center" height="25" class=a1>公告</td>
<td width="15%" align="center" height="25" class=a1>创始人</td>
<td width="20%" align="center" height="25" class=a1>管理</td>
</tr>
<%
sql="select * from [BBSXP_Consortia] order by DateCreated Desc"
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
<td width="10%" align="center" height="25"> <a target="_blank" href=Consortia.asp?menu=look&ConsortiaName=<%=Rs("ConsortiaName")%>><%=Rs("ConsortiaName")%></a>
<td width="50%" align="center" height="25"><%=Rs("tenet")%>
<td width="10%" align="center" height="25"><a target="_blank" href="Profile.asp?UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a>
<td width="20%" align="center" height="25"><a href="?menu=editConsortia&id=<%=Rs("id")%>">修改资料</a> <a onclick=checkclick('您确定要删除该公会？') href="?menu=DelConsortia&id=<%=Rs("id")%>">删除公会</a></td>
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

sub editConsortia
sql="select * from [BBSXP_Consortia] where id="&id&""
Set Rs=Conn.Execute(SQL)
%>
<form method=Post action=?menu=editConsortiaok&id=<%=Rs("id")%>>
<table cellpadding="2" cellspacing="1" width="70%" border="0" class=a2>
<tr>
<td colspan="2" height="25" align="center" class=a1>公会设定</td>
</tr>
<tr class=a3>
<td>　　公会简称： </td>
<td>
<input size="20" maxlength=7 name="Consortianame" value="<%=Rs("Consortianame")%>"> 
最多7个字符</td>
</tr>
<tr class=a3>
<td>　　公会名称： </td>
<td><input size="30" name="FullName" value="<%=Rs("FullName")%>"> </td>
</tr>
<tr class=a3>
<td>　　会长名称： </td>
<td><input size="30" name="UserName" value="<%=Rs("UserName")%>"> </td>
</tr>
<tr class=a3>
<td>　　公会公告： </td>
<td><input size="60" name="tenet" value="<%=Rs("tenet")%>"> </td>
</tr>
<tr class=a3>
<td colSpan="2">
<div align="center">
<input type="submit" value=" 修 改 ">
<input type="reset" value=" 重 填 ">
</div>
</td>
</tr>
</table>
</form><p><a href=javascript:history.back()>< 返 回 ></a>
<%
end sub


sub editConsortiaok
Consortianame=HTMLEncode(Request("Consortianame"))
FullName=HTMLEncode(Request("FullName"))
tenet=HTMLEncode(Request("tenet"))
UserName=HTMLEncode(Request("UserName"))

if Consortianame="" then error2("公会简称没有填写")
if FullName="" then error2("公会名称没有填写")
if UserName="" then error2("会长不能没有填写")

sql="select * from [BBSXP_Consortia] where id="&id&""
Rs.Open sql,Conn,1,3
oldConsortianame=Rs("Consortianame")
Rs("Consortianame")=Consortianame
Rs("FullName")=FullName
Rs("UserName")=UserName
Rs("tenet")=tenet
Rs.update
Rs.close
Conn.execute("update [BBSXP_Users] set Consortia='"&Consortianame&"' where Consortia='"&oldConsortianame&"'")
error2("修改成功")
end sub

sub DelConsortia
sql="select * from [BBSXP_Consortia] where id="&id&""
Set Rs1=Conn.Execute(sql)
Conn.execute("update [BBSXP_Users] set Consortia='' where Consortia='"&rs1("Consortianame")&"'")
set rs1=nothing
Conn.execute("Delete from [BBSXP_Consortia] where id="&id&"")
error2("删除成功")
end sub


htmlend

%>