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
if TimeLimit="" then error2("��û��ѡ�����ڣ�")
Conn.execute("Delete from [BBSXP_Messages] where DateCreated<"&SqlNowString&"-"&TimeLimit&"")
error2("�Ѿ���"&TimeLimit&"����ǰ�Ķ�ѶϢɾ���ˣ�")


case "DelMessageUser"
if UserName="" then error2("��û�������û�����")
Conn.execute("Delete from [BBSXP_Messages] where UserName='"&UserName&"' or incept='"&UserName&"'")
error2("�Ѿ���"&UserName&"�Ķ�ѶϢȫ��ɾ���ˣ�")

case "DelMessagekey"
key=HTMLEncode(Request("key"))
if key="" then error2("��û������ؼ��ʣ�")
Conn.execute("Delete from [BBSXP_Messages] where content like '%"&key&"%'")
error2("�Ѿ��������а��� "&key&" �Ķ�ѶϢɾ���ˣ�")


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
%>�� �� �� �� ��<p><a href=javascript:history.back()>< �� �� ></a><%


case "DelLink"
Conn.execute("Delete from [BBSXP_Link] where id="&id&"")
Link



end select


sub sendMail
%>

<form method="POST" action="?menu=sendMailok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>Ⱥ���ʼ�</td>
  </tr>
  <tr height=25>
    <td class=a3 align=Left>&nbsp;&nbsp; ���⣺<input size="40" name="title"></td>
    <td class=a3 align=middle>���ն���
<select name=membercode>
<option value="">���л�Ա</option>
<option value="1">��ͨ��Ա</option>
<option value="2">�����Ա</option>
<option value="4">��������</option>
<option value="5">����Ա</option>
</select>&nbsp;&nbsp;&nbsp; </td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <input type="submit" value=" �� �� ">
<input type="reset" value=" �� �� "><br></td></tr></table></form>

<%
end sub


sub sendMailok

if Request("title")="" then error2("����д�ʼ�����")
if Request("content")="" then error2("����д�ʼ�����")


if membercode<>"" then
sql="select UserMail from [BBSXP_Users] where membercode="&membercode&""
else
sql="select UserMail from [BBSXP_Users]"
end if

Set Rs=Conn.Execute(sql)
do while not Rs.eof

Mailaddress=""&Rs("UserMail")&""
MailTopic=Request("title")
body=""&Request("content")&""&vbCrlf&""&vbCrlf&"���ʼ�ͨ�� BBSXP Ⱥ��ϵͳ���͡�����������YUZI������(http://www.yuzi.net)"
%><!-- #include file="inc/Mail.asp" --><%

Rs.Movenext
loop
Rs.close

response.write "�ʼ����ͳɹ���"

end sub


sub Message
%>




���ݿ⹲ <%=Conn.execute("Select count(id)from [BBSXP_Messages]")(0)%> ����ѶϢ
<br><br>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
		<td align="center">����ɾ������Ϣ</td>
	</tr>
	<tr class=a3>
		<td align="center"><form method="POST" action="?menu=DelMessageUser">����ɾ�� <input size="13" name="UserName" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="�û���" Helptext="�û���"> �Ķ�ѶϢ
<input type="submit" value="ȷ��" id='btnadd' disabled>
		</td></form>
	</tr>
	
	<tr class=a3>
		<td align="center"><form method="POST" action="?menu=DelMessagekey">����ɾ�����ݺ��� <input size="20" name="key" onkeyup="ValidateTextboxAdd(this, 'nrkey')" onpropertychange="ValidateTextboxAdd(this, 'nrkey')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="�ؼ���" Helptext="�ؼ���"> �Ķ�ѶϢ
<input type="submit" value="ȷ��" id='nrkey' disabled>
		</td></form>
	</tr>
	
		<tr class=a3>
		<td align="center"><form method="POST" action="?menu=Messageok">ɾ�� <INPUT size=2 name=TimeLimit value="30">
			����ǰ�Ķ�ѶϢ
<input type="submit" value="ȷ��">

		</td></form>
	</tr>
	
</table>
</form>
<form method="POST" action="?menu=broadcast">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle width="50%">ϵͳ�㲥</td>
    <td class=a1 align=middle width="50%">���ն���
<select name=membercode>
<option value="">���߻�Ա</option>
<option value="1">��ͨ��Ա</option>
<option value="2">�����Ա</option>
<option value="4">��������</option>
<option value="5">����Ա</option>
</select>
</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<input type="submit" value=" �� �� ">
<input type="reset" value=" �� �� "></td></tr></table></form>
<%
end sub

sub broadcast
content=HTMLEncode(Request.Form("content"))

if content="" then error2("����д�㲥����!")

if membercode<>"" then
sql="select UserName from [BBSXP_Users] where membercode="&membercode&""
else
sql="select UserName from [BBSXP_UsersOnline] where UserName<>''"
end if

Set Rs=Conn.Execute(sql)
do while not Rs.eof
Count=Count+1
Conn.Execute("insert into [BBSXP_Messages] (UserName,incept,content) values ('"&CookieUserName&"','"&Rs("UserName")&"','<font color=0000FF>��ϵͳ�㲥����"&content&"</font>')")
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&Rs("UserName")&"'")
Rs.Movenext
loop
Rs.close

%>
�����ɹ�
<br><br>
�����͸� <%=Count%> λ�����û�<br><br>
<a href=javascript:history.back()>�� ��</a>
<%
end sub




sub Link
%>
<FORM action=?menu=Linkok method=Post>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align=center><tr>
	<td height="7" class="a1" colspan="2">����<b> </b>�������ӹ���</td></tr><tr>
	<td height="6" class="a3">��վ���ƣ�<INPUT size=40 name=name></td>
	<td height="6" class="a3">��ַURL��<INPUT size=40 name=url value="http://"></td></tr><tr>
	<td height="6" class="a3">��վ��飺<INPUT size=40 name=Intro></td>
	<td height="6" class="a3">ͼ��URL��<INPUT size=40 name=Logo value="http://"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" �� �� ">
<input type="reset" value=" �� �� ">

</td></tr></table>
</FORM>


<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center"><tr><td height="25" colspan="2" class="a1">����<b> </b>��������</td></tr>
<tr class=a3>
<td align="center" width="5%"><img src="images/shareforum.gif"></td>
<td class="a4"><%
Rs.Open "[BBSXP_Link]",Conn
do while not Rs.eof
if Rs("Logo")="" or Rs("Logo")="http://" then
Link1=Link1+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editLink&id="&Rs("id")&">�༭</a></div><div class=menuitems><a href=?menu=DelLink&id="&Rs("id")&">ɾ��</a></div>')"&Chr(34)&" title='"&Rs("Intro")&"' href="&Rs("url")&" target=_blank>"&Rs("name")&"</a>����"
else
Link2=Link2+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editLink&id="&Rs("id")&">�༭</a></div><div class=menuitems><a href=?menu=DelLink&id="&Rs("id")&">ɾ��</a></div>')"&Chr(34)&" title='"&Rs("name")&""&chr(10)&""&Rs("Intro")&"' href="&Rs("url")&" target=_blank><img src="&Rs("Logo")&" border=0 width=88 height=31></a>����"
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
	<td height="7" class="a1" colspan="2">����<b> </b>�������ӹ���</td></tr><tr>
	<td height="6" class="a3">��վ���ƣ�<INPUT size=40 name=name value="<%=Rs("name")%>"></td>
	<td height="6" class="a3">��ַURL��<INPUT size=40 name=url value="<%=Rs("url")%>"></td></tr><tr>
	<td height="6" class="a3">��վ��飺<INPUT size=40 name=Intro value="<%=Rs("Intro")%>"></td>
	<td height="6" class="a3">ͼ��URL��<INPUT size=40 name=Logo value="<%=Rs("Logo")%>"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" �� �� ">
<input type="reset" value=" �� �� ">
</td></tr></table>
</FORM><p><a href=javascript:history.back()>< �� �� ></a>
<%
end sub



sub Linkok

if Request("url")="http://" or Request("url")="" then error2("��̳URLû����д")

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
<td width="15%" align="center" height="25" class=a1>����</td>
<td width="40%" align="center" height="25" class=a1>����</td>
<td width="15%" align="center" height="25" class=a1>��ʼ��</td>
<td width="20%" align="center" height="25" class=a1>����</td>
</tr>
<%
sql="select * from [BBSXP_Consortia] order by DateCreated Desc"
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
<td width="10%" align="center" height="25"> <a target="_blank" href=Consortia.asp?menu=look&ConsortiaName=<%=Rs("ConsortiaName")%>><%=Rs("ConsortiaName")%></a>
<td width="50%" align="center" height="25"><%=Rs("tenet")%>
<td width="10%" align="center" height="25"><a target="_blank" href="Profile.asp?UserName=<%=Rs("UserName")%>"><%=Rs("UserName")%></a>
<td width="20%" align="center" height="25"><a href="?menu=editConsortia&id=<%=Rs("id")%>">�޸�����</a> <a onclick=checkclick('��ȷ��Ҫɾ���ù��᣿') href="?menu=DelConsortia&id=<%=Rs("id")%>">ɾ������</a></td>
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
<td colspan="2" height="25" align="center" class=a1>�����趨</td>
</tr>
<tr class=a3>
<td>���������ƣ� </td>
<td>
<input size="20" maxlength=7 name="Consortianame" value="<%=Rs("Consortianame")%>"> 
���7���ַ�</td>
</tr>
<tr class=a3>
<td>�����������ƣ� </td>
<td><input size="30" name="FullName" value="<%=Rs("FullName")%>"> </td>
</tr>
<tr class=a3>
<td>�����᳤���ƣ� </td>
<td><input size="30" name="UserName" value="<%=Rs("UserName")%>"> </td>
</tr>
<tr class=a3>
<td>�������ṫ�棺 </td>
<td><input size="60" name="tenet" value="<%=Rs("tenet")%>"> </td>
</tr>
<tr class=a3>
<td colSpan="2">
<div align="center">
<input type="submit" value=" �� �� ">
<input type="reset" value=" �� �� ">
</div>
</td>
</tr>
</table>
</form><p><a href=javascript:history.back()>< �� �� ></a>
<%
end sub


sub editConsortiaok
Consortianame=HTMLEncode(Request("Consortianame"))
FullName=HTMLEncode(Request("FullName"))
tenet=HTMLEncode(Request("tenet"))
UserName=HTMLEncode(Request("UserName"))

if Consortianame="" then error2("������û����д")
if FullName="" then error2("��������û����д")
if UserName="" then error2("�᳤����û����д")

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
error2("�޸ĳɹ�")
end sub

sub DelConsortia
sql="select * from [BBSXP_Consortia] where id="&id&""
Set Rs1=Conn.Execute(sql)
Conn.execute("update [BBSXP_Users] set Consortia='' where Consortia='"&rs1("Consortianame")&"'")
set rs1=nothing
Conn.execute("Delete from [BBSXP_Consortia] where id="&id&"")
error2("ɾ���ɹ�")
end sub


htmlend

%>