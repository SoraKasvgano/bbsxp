<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

id=int(Request("id"))
incept=HTMLEncode(Request("incept"))
UserName=HTMLEncode(Request("UserName"))
UserHonor=HTMLEncode(Request("UserHonor"))
ConsortiaName=HTMLEncode(Request("ConsortiaName"))

sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(sql)
Consortia=Rs("Consortia")
experience=Rs("experience")
UserMoney=Rs("UserMoney")
Rs.close
top

if Request.form("menu")="Consortiaadd" then
if Consortia<>"" then error2("您已经加入 "&Consortia&" 了，不能再加入其他公会！")
Consortianame=Conn.Execute("Select Consortianame From [BBSXP_Consortia] where id="&id&"")(0)
if Conn.execute("Select count(id) from [BBSXP_Users] where Consortia='"&Consortianame&"'")(0)>99 then error2("该公会已经超过100名会员，无法再加入新会员")
Conn.execute("Delete from [BBSXP_Messages] where id="&int(Request("Messageid"))&" and incept='"&CookieUserName&"'")
Conn.execute("update [BBSXP_Users] set Consortia='"&Consortianame&"' where UserName='"&CookieUserName&"'")
error2("加入公会成功")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="Consortiaout" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Consortia=empty then error("<li>您目前没有加入任何公会！")
If not Conn.Execute("Select id From [BBSXP_Consortia] where UserName='"&CookieUserName&"'").eof Then error("<li>要退出请先解散公会")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&CookieUserName&"'")
Message=Message&"<li>退出公会成功<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="invite" then
if UserName="" then error("<li>请填写受邀人的名称")
if CookieUserName=UserName then error("<li>不能自己邀请自己")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>只有会长才有权限执行该操作")
if Conn.execute("Select count(id) from [BBSXP_Users] where Consortia='"&Consortia&"'")(0)>99 then error2("公会已经超过100名会员，无法再加入新会员")
if Conn.Execute("Select Consortia From [BBSXP_Users] where UserName='"&UserName&"'")(0)<>"" then error("<li>对方已经加入了其他公会")
Messageid=Conn.execute("select Max(ID)+1 From [BBSXP_Messages]")(0)
Conn.Execute("insert into [BBSXP_Messages] (UserName,incept,content) values ('"&CookieUserName&"','"&UserName&"','<form name=ConsortiaAdd"&Messageid&" method=Post action=Consortia.asp?id="&id&"&Messageid="&Messageid&"><input type=hidden name=menu value=Consortiaadd></form><font color=0000FF>【系统消息】："&CookieUserName&" 邀请您加入 "&Consortia&" 公会<br><br><center><a href=javascript:ConsortiaAdd"&Messageid&".submit()>同意</a>　　　<a href=Message.asp?menu=Del&id="&Messageid&">拒绝</a></font></center>')")
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&UserName&"'")
Message=Message&"<li>邀请已经成功发出<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaDel" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>只有会长才有权限执行该操作")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where Consortia='"&Consortia&"'")
Conn.execute("Delete from [BBSXP_Consortia] where id="&id&"")
Message=Message&"<li>解散公会成功<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaUserOut" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if CookieUserName=UserName then error("<li>不能自己开除自己")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>只有会长才有权限执行该操作")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&UserName&"' and Consortia='"&Consortia&"'")
Message=Message&"<li>已经将 "&UserName&" 从公会中开除了<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaUserUserHonor" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Len(UserHonor)>7 then error("<li>头衔长度不能超多7个字符")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>只有会长才有权限执行该操作")
Conn.execute("update [BBSXP_Users] set UserHonor='"&UserHonor&"' where UserName='"&UserName&"' and Consortia='"&Consortia&"'")
Message=Message&"<li> "&UserName&" 已经获得 "&UserHonor&" 的头衔<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="addok" then
FullName=HTMLEncode(Request.Form("FullName"))
tenet=HTMLEncode(Request.Form("tenet"))
if instr(Consortianame,";")>0 then Message=Message&"<li>公会简称中不能含有特殊符号"
if Consortia<>empty then Message=Message&"<li>您已经加入了其他公会！"
if experience< 10000 then Message=Message&"<li>您的经验值小于 10000 ！"
if UserMoney< 10000 then Message=Message&"<li>您的金币少于 10000 ！"
if Consortianame="" then Message=Message&"<li>公会简称没有填写"
if Len(Consortianame)>7 then Message=Message&"<li>公会简称最多7个字符"
if FullName="" then Message=Message&"<li>公会全称没有填写"
If not Conn.Execute("Select id From [BBSXP_Consortia] where Consortianame='"&Consortianame&"' or UserName='"&CookieUserName&"'").eof Then  Message=Message&"<li>社区中已存在同名公会<li>您已经建立过公会"
if Message<>"" then error(""&Message&"")
Conn.Execute("insert into [BBSXP_Consortia] (Consortianame,FullName,tenet,UserName) values ('"&Consortianame&"','"&FullName&"','"&tenet&"','"&CookieUserName&"')")
Conn.execute("update [BBSXP_Users] set Consortia='"&Consortianame&"',[UserMoney]=[UserMoney]-10000 where UserName='"&CookieUserName&"'")
Message=Message&"<li>创建公会成功<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="xiuok" then
FullName=HTMLEncode(Request.Form("FullName"))
tenet=HTMLEncode(Request.Form("tenet"))
if FullName="" then Message=Message&"<li>公会全称没有填写"
if Message<>"" then error(""&Message&"")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>只有会长才有权限执行该操作")
Conn.execute("update [BBSXP_Consortia] set FullName='"&FullName&"',tenet='"&tenet&"' where id="&id&"")
Message=Message&"<li>修改公会成功<li><a href=Consortia.asp>返回社区公会</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="look" then
sql="select * from [BBSXP_Consortia] where ConsortiaName='"&ConsortiaName&"'"
Set Rs=Conn.Execute(sql)
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="Consortia.asp">社区公会</a></td>
</tr>
</table><br>
<table width="82%" border="0" align="center" cellspacing="1" cellpadding="2"  class=a2 height="150">
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>公会简称:</b></font>
</td>
<td width="82%"><%=Rs("Consortianame")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>公会全称:</b></font>
</td>
<td width="82%"><%=Rs("FullName")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>公会公告:</b></font>
</td>
<td width="82%"><%=Rs("tenet")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>创建时间:</b></font>
</td>
<td width="82%"><%=Rs("DateCreated")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>公会会长:</b></font>
</td>
<td width="82%"><%=Rs("UserName")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>现有会员:</b></font>
</td>
<td width="82%">
<%
sql="select UserName from [BBSXP_Users] where Consortia='"&Rs("Consortianame")&"'"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
i=i+1
list=list&"<a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a> "
Rs.MoveNext
loop
%><%=i%>人</td>
</tr>

<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>会员名单:</b></font>
</td>
<td width="82%">
<%=list%>
</td>
</tr>

</table>
<br><center><INPUT onclick=history.back(-1) type=button value=" << 返 回 ">
<%

htmlend
end if



%>
<center>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="Consortia.asp">社区公会</a></td>
</tr>
</table><br>
<%
if Request("menu")="add" then

%>
<form method=Post name=form action=Consortia.asp?menu=addok>

<table cellspacing=1 cellpadding=2 width=442 border=0 align="center" class=a2>
<tr class=a1>
<td width=526 colspan="2" align="center" height="25">
创建公会</td>
</tr>
<tr class=a3>
<td width=187 align="right">
<b><font color="0033CC">公会简称：</font></b></td>
<td width=339>
<input maxlength=7 name=Consortianame size="10"> 最多7个字符</td>
</tr>
<tr class=a3>
<td width=187>
<div align="right"><b><font color="0033CC">公会全称： </font></b></div>
</td>
<td width=339>
<input size=30 name=FullName>
</td>
</tr>
<tr class=a3>
<td width=187 height=15>
<div align="right"><b><font color="0033CC">公会公告： </font></b></div>
</td>
<td width=339 height=15>
<input size=40 name=tenet>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=8>
<div align=center>
<input type=submit value=" 创 建 ">
<input type=reset value=" 重 填 ">
</div>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=7>
<ol>
创建公会的注意事项：
<li>您的经验值必须 10000 以上
<li>需要扣除您身上 10000 金币作为公会基金 </li>
<li>公会最多只能容纳 100 名会员</td>
</tr>
</table>


</form>



<%
elseif Request("menu")="xiu" then
sql="select * from [BBSXP_Consortia] where id="&id&""
Set Rs=Conn.Execute(SQL)
%>

<form method=Post action=Consortia.asp?menu=xiuok&id=<%=Rs("id")%>>
<table cellpadding="2" cellspacing="1" width="70%" border="0" class=a2>

<tr>
<td colspan="2" height="25" align="center" class=a1>　　公会设定</td>
</tr>


<tr class=a3>
<td>　　公会简称： </td>
<td>
<%=Rs("Consortianame")%></td>
</tr>
<tr class=a3>
<td>　　公会全称： </td>
<td><input size="30" name="FullName" value="<%=Rs("FullName")%>"> </td>
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
</form>
<%
else
%>

		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="?menu=add">创建公会</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="?menu=Consortiaout&sessionid=<%=session.sessionid%>">退出公会</a>
<br><br>




<table width="100%" border="0" align="center" cellspacing="1" cellpadding="5"  class=a2>



<%


if Consortia<>"" then
sql="select * from [BBSXP_Consortia] where Consortianame='"&Consortia&"'"
Set Rs=Conn.Execute(sql)
if Rs.eof then Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&CookieUserName&"'")
sql="select UserName from [BBSXP_Users] where Consortia='"&Rs("Consortianame")&"'"
Set Rs1=Conn.Execute(sql)
if Rs1.eof then Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&CookieUserName&"'")
Do While Not RS1.EOF
i=i+1
list=list&"<input type=radio value='"&rs1("UserName")&"' name=UserName id="&i&"><label for="&i&">"&rs1("UserName")&"</label> "
RS1.MoveNext
loop
Set Rs1 = Nothing

%>
<SCRIPT>
function ConsortiaUser(UserName,act){
var UserNames
if (UserName.length > 1){
for(iIndex=0;iIndex<UserName.length;iIndex++){if(UserName[iIndex].checked==true){UserNames = UserName[iIndex].value;}}
}
else{if(UserName.checked==true){UserNames=UserName.value;}}
if(UserNames==undefined){alert('请选择对象');return;}
if(act=="add"){document.location='Friend.asp?menu=add&UserName='+UserNames+'';}
if(act=="Post"){open('Friend.asp?menu=Post&incept='+UserNames+'','','width=320,height=170')}
if(act=="look"){open('Profile.asp?UserName='+UserNames+'')}
if(act=="out"){document.location='Consortia.asp?menu=ConsortiaUserOut&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+UserNames+'';}
if(act=="UserHonor"){var id=prompt("请输入该会员的头衔！","");if(id){document.location='Consortia.asp?menu=ConsortiaUserUserHonor&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+UserNames+'&UserHonor='+id+'';}}
}

function invite(){
var id=prompt("请输入邀请加入本公会的会员名称！","");if(id){document.location='Consortia.asp?menu=invite&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+id+'';}
}

</SCRIPT>

<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>公会简称</b></font><font color="#000066"><b>:</b></font>
</td>
<td width="28%"><%=Rs("Consortianame")%></td>
<td width="10%" align="center"><font color="000066"><b>公会全称:</b></font></td>
<td width="27%"><%=Rs("FullName")%></td>
</tr>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>公会会长:</b></font>
</td>
<td width="28%"><%=Rs("UserName")%></td>
<td width="10%" align="center"><font color="000066"><b>创建时间:</b></font></td>
<td width="27%"><%=Rs("DateCreated")%></td>
</tr>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>公会公告:</b></font>
</td>
<td width="82%" colspan="3"><%=Rs("tenet")%></td>
</tr>
<form method=Post name=Consortia>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>会员名单:</b></font><br>
<font size="1">共 <font color=red><%=i%></font> 人</font>
</td>
<td width="65%" colspan="3">

<%=list%>
</td>
</tr>

<tr>
<td width="38%" colspan="2" align="center" class=a1>
会员操作</td>
<td width="37%" colspan="2" align="center" class=a1>
会长管理</td>
</tr>
<tr class=a3>
<td width="38%" colspan="2" align="center">
<a onclick="ConsortiaUser(document.Consortia.UserName,'add')" href=#>加为好友</a>　　<a onclick="ConsortiaUser(document.Consortia.UserName,'Post')" href=#>发送信息</a>　　<a onclick="ConsortiaUser(document.Consortia.UserName,'look')" href=#>查看资料</a></td>
<td width="37%" colspan="2" align="center">


<a onclick="invite()" href=#>添加会员</a>　　<a onclick="ConsortiaUser(document.Consortia.UserName,'out')" href=#>开除会员</a>　　<a onclick="ConsortiaUser(document.Consortia.UserName,'UserHonor')" href=#>赐予头衔</a>　　<a href="?menu=xiu&id=<%=Rs("id")%>">修改资料</a>　　<a onclick=checkclick('您确定要解散该公会？') href="?menu=ConsortiaDel&sessionid=<%=session.sessionid%>&id=<%=Rs("id")%>">解散公会</a></td>
</tr>
</form>
<%
Rs.Close


else
%>
<tr class=a3><td align="center">没有创建或者加入任何公会</td></tr>
<%
end if
%>
</table>






<%
end if


htmlend
%>