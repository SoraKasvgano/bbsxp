<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
if Request.ServerVariables("request_method") = "POST" then
Response.Cookies("DisabledShowFace")=Request("DisabledShowFace")
Response.Cookies("DisabledShowSign")=Request("DisabledShowSign")
Response.Cookies("DisabledShowMessage")=Request("DisabledShowMessage")
Response.Cookies("BadUserList")=Request("BadUserList")
Response.Cookies("DisabledShowFace").Expires=date+9999
Response.Cookies("DisabledShowSign").Expires=date+9999
Response.Cookies("DisabledShowMessage").Expires=date+9999
Response.Cookies("BadUserList").Expires=date+9999
message=message&"<li>个性设置成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end if

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
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="?">
<tr>
<td height="10" align="center" colspan="2" valign="bottom" class=a1>
<b>个性设置</b></td>
</tr>
<tr class=a3>
    <td height="2" align="right" width="45%"><b>显示用户头像：<br>
	</b>在帖子中显示用户头像</td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowFace" value="" <%if Request.Cookies("DisabledShowFace")="" then%>checked<%end if%>>是
<input type=radio name="DisabledShowFace" value="1" <%if Request.Cookies("DisabledShowFace")="1" then%>checked<%end if%>>否</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>显示用户签名：<br>
	</b>在帖子中显示用户签名</td>
    <td height="1" align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowSign" value="" <%if Request.Cookies("DisabledShowSign")="" then%>checked<%end if%>>是
<input type=radio name="DisabledShowSign" value="1" <%if Request.Cookies("DisabledShowSign")="1" then%>checked<%end if%>>否</td>
</tr>
<tr class=a3>
    <td align="right" width="45%"><b>短讯提示：<br>
	</b>有短讯息时跳出窗口提示</td>
    <td align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowMessage" value="" <%if Request.Cookies("DisabledShowMessage")="" then%>checked<%end if%>>是
<input type=radio name="DisabledShowMessage" value="1" <%if Request.Cookies("DisabledShowMessage")="1" then%>checked<%end if%>>否</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>过滤用户的名单：</b><br>用户发表的帖子将被过滤</td>
    <td height="1" align="left" width="55%">  
<input size=30 name=BadUserList value="<%=Request.Cookies("BadUserList")%>">
	<font color="FF0000">多人请用“|”分隔</font> </td>
</tr>
<tr class=a3>
    <td height="1" align="center" width="100%" colspan="2">
<input type="submit" name="Submit2" value=" 确 定 "></td>
</tr>


</table>

<%
htmlend
%>