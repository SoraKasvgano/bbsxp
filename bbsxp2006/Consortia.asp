<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

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
if Consortia<>"" then error2("���Ѿ����� "&Consortia&" �ˣ������ټ����������ᣡ")
Consortianame=Conn.Execute("Select Consortianame From [BBSXP_Consortia] where id="&id&"")(0)
if Conn.execute("Select count(id) from [BBSXP_Users] where Consortia='"&Consortianame&"'")(0)>99 then error2("�ù����Ѿ�����100����Ա���޷��ټ����»�Ա")
Conn.execute("Delete from [BBSXP_Messages] where id="&int(Request("Messageid"))&" and incept='"&CookieUserName&"'")
Conn.execute("update [BBSXP_Users] set Consortia='"&Consortianame&"' where UserName='"&CookieUserName&"'")
error2("���빫��ɹ�")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="Consortiaout" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Consortia=empty then error("<li>��Ŀǰû�м����κι��ᣡ")
If not Conn.Execute("Select id From [BBSXP_Consortia] where UserName='"&CookieUserName&"'").eof Then error("<li>Ҫ�˳����Ƚ�ɢ����")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&CookieUserName&"'")
Message=Message&"<li>�˳�����ɹ�<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="invite" then
if UserName="" then error("<li>����д�����˵�����")
if CookieUserName=UserName then error("<li>�����Լ������Լ�")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>ֻ�л᳤����Ȩ��ִ�иò���")
if Conn.execute("Select count(id) from [BBSXP_Users] where Consortia='"&Consortia&"'")(0)>99 then error2("�����Ѿ�����100����Ա���޷��ټ����»�Ա")
if Conn.Execute("Select Consortia From [BBSXP_Users] where UserName='"&UserName&"'")(0)<>"" then error("<li>�Է��Ѿ���������������")
Messageid=Conn.execute("select Max(ID)+1 From [BBSXP_Messages]")(0)
Conn.Execute("insert into [BBSXP_Messages] (UserName,incept,content) values ('"&CookieUserName&"','"&UserName&"','<form name=ConsortiaAdd"&Messageid&" method=Post action=Consortia.asp?id="&id&"&Messageid="&Messageid&"><input type=hidden name=menu value=Consortiaadd></form><font color=0000FF>��ϵͳ��Ϣ����"&CookieUserName&" ���������� "&Consortia&" ����<br><br><center><a href=javascript:ConsortiaAdd"&Messageid&".submit()>ͬ��</a>������<a href=Message.asp?menu=Del&id="&Messageid&">�ܾ�</a></font></center>')")
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&UserName&"'")
Message=Message&"<li>�����Ѿ��ɹ�����<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaDel" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>ֻ�л᳤����Ȩ��ִ�иò���")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where Consortia='"&Consortia&"'")
Conn.execute("Delete from [BBSXP_Consortia] where id="&id&"")
Message=Message&"<li>��ɢ����ɹ�<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaUserOut" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if CookieUserName=UserName then error("<li>�����Լ������Լ�")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>ֻ�л᳤����Ȩ��ִ�иò���")
Conn.execute("update [BBSXP_Users] set Consortia='',UserHonor='' where UserName='"&UserName&"' and Consortia='"&Consortia&"'")
Message=Message&"<li>�Ѿ��� "&UserName&" �ӹ����п�����<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="ConsortiaUserUserHonor" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Len(UserHonor)>7 then error("<li>ͷ�γ��Ȳ��ܳ���7���ַ�")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>ֻ�л᳤����Ȩ��ִ�иò���")
Conn.execute("update [BBSXP_Users] set UserHonor='"&UserHonor&"' where UserName='"&UserName&"' and Consortia='"&Consortia&"'")
Message=Message&"<li> "&UserName&" �Ѿ���� "&UserHonor&" ��ͷ��<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="addok" then
FullName=HTMLEncode(Request.Form("FullName"))
tenet=HTMLEncode(Request.Form("tenet"))
if instr(Consortianame,";")>0 then Message=Message&"<li>�������в��ܺ����������"
if Consortia<>empty then Message=Message&"<li>���Ѿ��������������ᣡ"
if experience< 10000 then Message=Message&"<li>���ľ���ֵС�� 10000 ��"
if UserMoney< 10000 then Message=Message&"<li>���Ľ������ 10000 ��"
if Consortianame="" then Message=Message&"<li>������û����д"
if Len(Consortianame)>7 then Message=Message&"<li>���������7���ַ�"
if FullName="" then Message=Message&"<li>����ȫ��û����д"
If not Conn.Execute("Select id From [BBSXP_Consortia] where Consortianame='"&Consortianame&"' or UserName='"&CookieUserName&"'").eof Then  Message=Message&"<li>�������Ѵ���ͬ������<li>���Ѿ�����������"
if Message<>"" then error(""&Message&"")
Conn.Execute("insert into [BBSXP_Consortia] (Consortianame,FullName,tenet,UserName) values ('"&Consortianame&"','"&FullName&"','"&tenet&"','"&CookieUserName&"')")
Conn.execute("update [BBSXP_Users] set Consortia='"&Consortianame&"',[UserMoney]=[UserMoney]-10000 where UserName='"&CookieUserName&"'")
Message=Message&"<li>��������ɹ�<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="xiuok" then
FullName=HTMLEncode(Request.Form("FullName"))
tenet=HTMLEncode(Request.Form("tenet"))
if FullName="" then Message=Message&"<li>����ȫ��û����д"
if Message<>"" then error(""&Message&"")
if Conn.Execute("Select UserName From [BBSXP_Consortia] where id="&id&"")(0)<>CookieUserName then error("<li>ֻ�л᳤����Ȩ��ִ�иò���")
Conn.execute("update [BBSXP_Consortia] set FullName='"&FullName&"',tenet='"&tenet&"' where id="&id&"")
Message=Message&"<li>�޸Ĺ���ɹ�<li><a href=Consortia.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Consortia.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="look" then
sql="select * from [BBSXP_Consortia] where ConsortiaName='"&ConsortiaName&"'"
Set Rs=Conn.Execute(sql)
%>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="Consortia.asp">��������</a></td>
</tr>
</table><br>
<table width="82%" border="0" align="center" cellspacing="1" cellpadding="2"  class=a2 height="150">
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>������:</b></font>
</td>
<td width="82%"><%=Rs("Consortianame")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>����ȫ��:</b></font>
</td>
<td width="82%"><%=Rs("FullName")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>���ṫ��:</b></font>
</td>
<td width="82%"><%=Rs("tenet")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>����ʱ��:</b></font>
</td>
<td width="82%"><%=Rs("DateCreated")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>����᳤:</b></font>
</td>
<td width="82%"><%=Rs("UserName")%></td>
</tr>
<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>���л�Ա:</b></font>
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
%><%=i%>��</td>
</tr>

<tr class=a3>
<td width="15%" align="center">
<font color="000066"><b>��Ա����:</b></font>
</td>
<td width="82%">
<%=list%>
</td>
</tr>

</table>
<br><center><INPUT onclick=history.back(-1) type=button value=" << �� �� ">
<%

htmlend
end if



%>
<center>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="Consortia.asp">��������</a></td>
</tr>
</table><br>
<%
if Request("menu")="add" then

%>
<form method=Post name=form action=Consortia.asp?menu=addok>

<table cellspacing=1 cellpadding=2 width=442 border=0 align="center" class=a2>
<tr class=a1>
<td width=526 colspan="2" align="center" height="25">
��������</td>
</tr>
<tr class=a3>
<td width=187 align="right">
<b><font color="0033CC">�����ƣ�</font></b></td>
<td width=339>
<input maxlength=7 name=Consortianame size="10"> ���7���ַ�</td>
</tr>
<tr class=a3>
<td width=187>
<div align="right"><b><font color="0033CC">����ȫ�ƣ� </font></b></div>
</td>
<td width=339>
<input size=30 name=FullName>
</td>
</tr>
<tr class=a3>
<td width=187 height=15>
<div align="right"><b><font color="0033CC">���ṫ�棺 </font></b></div>
</td>
<td width=339 height=15>
<input size=40 name=tenet>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=8>
<div align=center>
<input type=submit value=" �� �� ">
<input type=reset value=" �� �� ">
</div>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=7>
<ol>
���������ע�����
<li>���ľ���ֵ���� 10000 ����
<li>��Ҫ�۳������� 10000 �����Ϊ������� </li>
<li>�������ֻ������ 100 ����Ա</td>
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
<td colspan="2" height="25" align="center" class=a1>���������趨</td>
</tr>


<tr class=a3>
<td>���������ƣ� </td>
<td>
<%=Rs("Consortianame")%></td>
</tr>
<tr class=a3>
<td>��������ȫ�ƣ� </td>
<td><input size="30" name="FullName" value="<%=Rs("FullName")%>"> </td>
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
</form>
<%
else
%>

		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="?menu=add">��������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="?menu=Consortiaout&sessionid=<%=session.sessionid%>">�˳�����</a>
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
if(UserNames==undefined){alert('��ѡ�����');return;}
if(act=="add"){document.location='Friend.asp?menu=add&UserName='+UserNames+'';}
if(act=="Post"){open('Friend.asp?menu=Post&incept='+UserNames+'','','width=320,height=170')}
if(act=="look"){open('Profile.asp?UserName='+UserNames+'')}
if(act=="out"){document.location='Consortia.asp?menu=ConsortiaUserOut&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+UserNames+'';}
if(act=="UserHonor"){var id=prompt("������û�Ա��ͷ�Σ�","");if(id){document.location='Consortia.asp?menu=ConsortiaUserUserHonor&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+UserNames+'&UserHonor='+id+'';}}
}

function invite(){
var id=prompt("������������뱾����Ļ�Ա���ƣ�","");if(id){document.location='Consortia.asp?menu=invite&id=<%=Rs("id")%>&sessionid=<%=session.sessionid%>&UserName='+id+'';}
}

</SCRIPT>

<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>������</b></font><font color="#000066"><b>:</b></font>
</td>
<td width="28%"><%=Rs("Consortianame")%></td>
<td width="10%" align="center"><font color="000066"><b>����ȫ��:</b></font></td>
<td width="27%"><%=Rs("FullName")%></td>
</tr>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>����᳤:</b></font>
</td>
<td width="28%"><%=Rs("UserName")%></td>
<td width="10%" align="center"><font color="000066"><b>����ʱ��:</b></font></td>
<td width="27%"><%=Rs("DateCreated")%></td>
</tr>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>���ṫ��:</b></font>
</td>
<td width="82%" colspan="3"><%=Rs("tenet")%></td>
</tr>
<form method=Post name=Consortia>
<tr class=a3>
<td width="10%" align="center">
<font color="000066"><b>��Ա����:</b></font><br>
<font size="1">�� <font color=red><%=i%></font> ��</font>
</td>
<td width="65%" colspan="3">

<%=list%>
</td>
</tr>

<tr>
<td width="38%" colspan="2" align="center" class=a1>
��Ա����</td>
<td width="37%" colspan="2" align="center" class=a1>
�᳤����</td>
</tr>
<tr class=a3>
<td width="38%" colspan="2" align="center">
<a onclick="ConsortiaUser(document.Consortia.UserName,'add')" href=#>��Ϊ����</a>����<a onclick="ConsortiaUser(document.Consortia.UserName,'Post')" href=#>������Ϣ</a>����<a onclick="ConsortiaUser(document.Consortia.UserName,'look')" href=#>�鿴����</a></td>
<td width="37%" colspan="2" align="center">


<a onclick="invite()" href=#>��ӻ�Ա</a>����<a onclick="ConsortiaUser(document.Consortia.UserName,'out')" href=#>������Ա</a>����<a onclick="ConsortiaUser(document.Consortia.UserName,'UserHonor')" href=#>����ͷ��</a>����<a href="?menu=xiu&id=<%=Rs("id")%>">�޸�����</a>����<a onclick=checkclick('��ȷ��Ҫ��ɢ�ù��᣿') href="?menu=ConsortiaDel&sessionid=<%=session.sessionid%>&id=<%=Rs("id")%>">��ɢ����</a></td>
</tr>
</form>
<%
Rs.Close


else
%>
<tr class=a3><td align="center">û�д������߼����κι���</td></tr>
<%
end if
%>
</table>






<%
end if


htmlend
%>