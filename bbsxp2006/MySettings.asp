<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
if Request.ServerVariables("request_method") = "POST" then
Response.Cookies("DisabledShowFace")=Request("DisabledShowFace")
Response.Cookies("DisabledShowSign")=Request("DisabledShowSign")
Response.Cookies("DisabledShowMessage")=Request("DisabledShowMessage")
Response.Cookies("BadUserList")=Request("BadUserList")
Response.Cookies("DisabledShowFace").Expires=date+9999
Response.Cookies("DisabledShowSign").Expires=date+9999
Response.Cookies("DisabledShowMessage").Expires=date+9999
Response.Cookies("BadUserList").Expires=date+9999
message=message&"<li>�������óɹ�<li><a href=usercp.asp>���������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end if

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �������</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">�������</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">�����޸�</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">�����޸�</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">��������</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">��������</a></b></td>
      <TD align="center"><b><a href="Message.asp">���ŷ���</a></b></td>
      <TD align="center"><b><a href="Friend.asp">�����б�</a></b></td>
      </TR></TABLE>
<br>
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="?">
<tr>
<td height="10" align="center" colspan="2" valign="bottom" class=a1>
<b>��������</b></td>
</tr>
<tr class=a3>
    <td height="2" align="right" width="45%"><b>��ʾ�û�ͷ��<br>
	</b>����������ʾ�û�ͷ��</td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowFace" value="" <%if Request.Cookies("DisabledShowFace")="" then%>checked<%end if%>>��
<input type=radio name="DisabledShowFace" value="1" <%if Request.Cookies("DisabledShowFace")="1" then%>checked<%end if%>>��</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>��ʾ�û�ǩ����<br>
	</b>����������ʾ�û�ǩ��</td>
    <td height="1" align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowSign" value="" <%if Request.Cookies("DisabledShowSign")="" then%>checked<%end if%>>��
<input type=radio name="DisabledShowSign" value="1" <%if Request.Cookies("DisabledShowSign")="1" then%>checked<%end if%>>��</td>
</tr>
<tr class=a3>
    <td align="right" width="45%"><b>��Ѷ��ʾ��<br>
	</b>�ж�ѶϢʱ����������ʾ</td>
    <td align="left" width="55%"> &nbsp; 
<input type=radio name="DisabledShowMessage" value="" <%if Request.Cookies("DisabledShowMessage")="" then%>checked<%end if%>>��
<input type=radio name="DisabledShowMessage" value="1" <%if Request.Cookies("DisabledShowMessage")="1" then%>checked<%end if%>>��</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>�����û���������</b><br>�û���������ӽ�������</td>
    <td height="1" align="left" width="55%">  
<input size=30 name=BadUserList value="<%=Request.Cookies("BadUserList")%>">
	<font color="FF0000">�������á�|���ָ�</font> </td>
</tr>
<tr class=a3>
    <td height="1" align="center" width="100%" colspan="2">
<input type="submit" name="Submit2" value=" ȷ �� "></td>
</tr>


</table>

<%
htmlend
%>