<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error2("您还未登录论坛")

id=int(Request("id"))
UserName=HTMLEncode(Trim(Request("UserName")))



if Request("menu")="Post" then
%>
<title>发送讯息 - Powered By BBSXP</title>
<SCRIPT>
function check(theForm) {

if(theForm.incept.value == "" ) {
alert("昵称不能没有填写！");
return false;
}

if(theForm.incept.value == "<%=CookieUserName%>" ) {
alert("请输入您要发送的对象，不能发讯息给自己！");
return false;
}

if(theForm.content.value == "" ) {
alert("不能发空讯息！");
return false;
}

if (theForm.content.value.length > 255){
alert("对不起，您的讯息不能超过 255 个字节！");
return false;
}
}

function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}


function DoTitle(addTitle) {
document.form.incept.value=addTitle
}

function Check(){var Name=document.form.incept.value;
if(Name == ""){alert("昵称不能为空！");return false;}
window.open("Friend.asp?menu=Post&Log=1&incept="+Name+"&content="+document.form.content.value+"","_blank","width=320,height=270");window.close()
}

</SCRIPT><body topmargin=0>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="Friend.asp" method="POST">
<input type="hidden" name="menu" value="addPost">
<TD height="35">
&nbsp;昵称：<input name="incept" type="text" size="10"></TD>

<TD align=right height="35">

<select onchange=DoTitle(this.options[this.selectedIndex].value)>


<%
if Request("report")="1" then
%>
<option>版主列表</option>
<SCRIPT>
var moderated="<%=Request("moderated")%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i] !=""){document.write("<option value="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
<%
else

%>
<option>好友列表</option>
<SCRIPT>
var moderated="<%=Conn.Execute("Select UserFriend From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i] !=""){document.write("<option value="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
<%
end if

%>



</select>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="3" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey()><%=Request("body")%></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD>
<input onclick=javascript:Check() type="button" value="聊天记录">
</td><TD><input type="reset" value="取消发送" OnClick="window.close();"> </td><TD><input type="submit" value="发送讯息" onclick="return check(this.form)"></td>
</TR></form>
</TABLE>
<%
CloseDatabase

end if


if Request("menu")="Del" then
Conn.execute("Delete from [BBSXP_Messages] where id="&id&" and incept='"&CookieUserName&"'")
error2("删除成功")
elseif Request("menu")="allDel" then
Conn.execute("Delete from [BBSXP_Messages] where incept='"&CookieUserName&"'")
error2("已经成功清空收件箱")
end if




top

%>



<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 控制面板</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">控制面板</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">密码修改</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">个性设置</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">附件管理</a></b></td>
      <TD align="center"><b><a href="Message.asp">短信服务</a></b></td>
      <TD align="center"><b><a href="Friend.asp">好友列表</a></b></td>
      </TR></TABLE>
<br>

<center>


<TABLE  cellSpacing=1 border=0 width="100%" class=a2 cellpadding="3">
<TR class=a3><TD align="center" colspan=5>
          <a style="text-decoration: none; color: #000000" href="Message.asp">
          <img alt src="images/m_inbox.gif" border="0" dypop="收件箱"></a>&nbsp;
          <a style="text-decoration: none; color: #000000" href="Message.asp?send=1">
          <img alt src="images/M_issend.gif" border="0" dypop="已发送邮件"></a>&nbsp;
          <a href="Message.asp?menu=allDel" onclick=checkclick('您确定要清空收件箱?')>
          <img alt src="images/Recycle.gif" border="0" dypop="清空收件箱"></a>&nbsp;
          <a href="Friend.asp">
          <img alt src="images/M_address.gif" border="0" dypop="发送讯息"></a>&nbsp;
          <a style=cursor:hand onclick="javascript:open('?menu=Post','','width=320,height=170')"><img alt src="images/m_write.gif" border="0" dypop="发送讯息"></a></td></TR>

  <TR>
          
    <TD vAlign=center align=middle 
          width=5% class=a1 height="20">　</TD>
          
    <TD vAlign=center align=middle 
          width=12% class=a1 height="20"><B><span id=send>发件人</span></B></TD>
          
    <TD vAlign=center align=middle 
          width=50% class=a1 height="20"><B>内容</B></TD>
          
    <TD align=middle 
          width=17% class=a1 height="20"><B>日期</B></TD>
          
    <TD vAlign=center align=middle 
          width=16% class=a1 height="20"><B>操作</B></TD>
  </TR>



<%



if Request("send")="1" then
sql="select * from [BBSXP_Messages] where UserName='"&CookieUserName&"' order by id Desc"
response.write "<script>send.innerText='收件人'</script>"
else
sql="select * from [BBSXP_Messages] where incept='"&CookieUserName&"' order by id Desc"

end if

Rs.Open sql,Conn,1

PageSetup=10 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数

i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1


if Request("send")="1" then
UserName=Rs("incept")
Del="<a style=cursor:hand onclick=javascript:open('Friend.asp?menu=Post&Log=1&incept="&UserName&"','','width=320,height=270')>聊天记录</a>"

else
UserName=Rs("UserName")
Del="<a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/NewPost.gif) style=cursor:hand onclick=javascript:open('Friend.asp?menu=Post&incept="&Rs("UserName")&"','','width=320,height=170') title='回复讯息'>回复</a><a class=CommonImageTextButton style=BACKGROUND-IMAGE:url(images/del.gif) onclick=checkclick('您确定要删除此条讯息?') href=?menu=Del&id="&Rs("id")&" title='删除讯息'>删除</a>"
end if
%>
<TR class=a3>
    <TD vAlign=center align=middle width="5%"><IMG src="images/f_norm.gif"></TD>
    <TD vAlign=center align=middle width="12%"><A href="Profile.asp?UserName=<%=UserName%>"><%=UserName%></A> </TD>
    <TD width="50%"><%=Rs("content")%></TD>
    <TD align="center" width="17%"><%=Rs("DateCreated")%></TD>
    <TD align="center" width="16%"><%=Del%></TD>
</TR>
<%
Rs.MoveNext
loop
Rs.Close
if NewMessage>0 then Conn.execute("update [BBSXP_Users] set NewMessage=0 where UserName='"&CookieUserName&"'")
%>
</TD></TR></TABLE>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td>
</tr></table>




<%htmlend%>