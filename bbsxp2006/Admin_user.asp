<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")

UserName=HTMLEncode(Request("UserName"))

response.write "<center>"

select case Request("menu")

case "Activation"
Activation

case "User2"
User2

case "Showall"
Showall

case "Showallok"
Showallok

case "UserDelTopic"
UserDelTopic

case "UserDel"
UserDel

case "Userok"
Userok


case "UserRank"
UserRank


case "UserRankUp"
if Request.Form("RankName")<>"" then
Conn.Execute("insert into [BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('"&Request.Form("RankName")&"','"&Request.Form("PostingCountMin")&"','"&Request.Form("RankIconUrl")&"')")
end if
for each ho in Request.Form("id")
Conn.execute("update [BBSXP_Ranks] set RankName='"&Request.Form("RankName"&ho)&"',PostingCountMin='"&Request.Form("PostingCountMin"&ho)&"',RankIconUrl='"&Request.Form("RankIconUrl"&ho)&"' where ID="&ho&"")
next
response.write "���³ɹ�<br><br><a href=javascript:history.back()>�� ��</a>"

case "UserRankDel"
Conn.execute("Delete from [BBSXP_Ranks] where id="&Request("RankID")&"")
response.write "ɾ���ɹ�<br><br><a href=javascript:history.back()>�� ��</a>"

case "fix"
Conn.execute("update [BBSXP_Users] set Userpass='"&md5("123")&"'  where UserName='"&UserName&"'")
error2("�Ѿ��� "&UserName&" �����뻹ԭ�� 123 ")




case "activationok"
for each ho in request.form("id")
Conn.execute("update [BBSXP_Users] set membercode=1 where id="&ho&"")
next
error2("�Ѿ���������ѡ�û���")

end select




sub Showall
%>
�û����ϣ�<b><font color=red><%=Conn.execute("Select count(id)from [BBSXP_Users]")(0)%></font></b> ��
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 colspan=2 align="center">�û�����</td>
  </tr>


<tr class=a3>
<td colspan="2" align="center"><form method="POST" action="?menu=User2">
�������Ա������: <input size="13" name="UserName">
<input type="submit" value="ȷ��">



</td></form>
</tr>
<form method="POST" action="?menu=Showallok">

  <tr height=25>
    <td class=a1 align=center colspan=2>�߼���ѯ</td>
  </tr>



    <tr height=25 class=a3>
    <td width="30%">�û�����</td>
    <td><select size="1" name="membercode">
<option value="">�����û�</option>
<option value="1">������ͨ��Ա</option>
<option value="0">����δ�����Ա</option>
<option value="2">���й����Ա</option>
<option value="4">���г�������</option>
<option value="5">���й���Ա</option></td>
</tr>




  <tr height=25 class=a3>
    <td>�û�������</td>
    <td><input size="45" name="UserName"></td>
  </tr>




  <tr height=25 class=a3>
    <td>������Ϣ����</td>
    <td><input size="45" name="UserInfo"></td>
  </tr>


  <tr height=25 class=a3>
    <td>��ϵ��Ϣ����</td>
    <td><input size="45" name="UserIM"></td>
  </tr>


  <tr height=25 class=a3>
    <td>Email����</td>
    <td><input size="45" name="UserMail"></td>
  </tr>


  <tr height=25 class=a3>
    <td>��ҳ����</td>
    <td><input size="45" name="Userhome"></td>
  </tr>
  <tr height=25 class=a3>
    <td>ǩ������</td>
    <td><input size="45" name="UserSign"></td>
  </tr>


  <tr height=25 class=a3>
    <td>ע�����ڰ���</td>
    <td><input size="45" name="UserRegTime"></td>
  </tr>
  <tr height=25 class=a3>
    <td>����¼ʱ�����</td>
    <td><input size="45" name="UserLandTime"></td>
  </tr>
    

  
  <tr height=25 class=a3>
    <td>����¼IP����</td>
    <td><input size="45" name="UserLastIP"></td>
  </tr>

  <tr height=25 class=a3>
    <td>�����ʾ��¼��</td>
    <td><input size="45" value="50" name="SearchMax"></td>
  </tr>

  <tr height=25 class=a3>
    <td colspan="2" align="center">
	<input type="submit" value="   ��  ��   "></td>
  </tr>

</form>
  </table>



</div>



<br>
<%
end sub







sub Showallok
%>
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
<TR align=middle class=a1>
<TD>�û���</TD>
<TD height="25">Email</TD>
<TD>��ѶϢ</TD>
<TD>�༭</TD>
<TD>ע��ʱ��</TD>
<TD>����¼ʱ��</TD>
<TD>����¼IP</TD>
<TD>ɾ��</TD>
</TR>
<form method="POST" action="?menu=UserDel">
<%
if Request("membercode")<>"" then item=""&item&" and membercode="&Request("membercode")&""
if Request("UserName")<>"" then item=""&item&" and UserName like '%"&Request("UserName")&"%'"
if Request("UserMail")<>"" then item=""&item&" and UserMail like '%"&Request("UserMail")&"%'"
if Request("Userhome")<>"" then item=""&item&" and Userhome like '%"&Request("Userhome")&"%'"
if Request("UserInfo")<>"" then item=""&item&" and UserInfo like '%"&Request("UserInfo")&"%'"
if Request("UserIM")<>"" then item=""&item&" and UserIM like '%"&Request("UserIM")&"%'"
if Request("UserSign")<>"" then item=""&item&" and UserSign like '%"&Request("UserSign")&"%'"
if Request("UserRegTime")<>"" then item=""&item&" and UserRegTime like '%"&Request("UserRegTime")&"%'"
if Request("UserLandTime")<>"" then item=""&item&" and UserLandTime like '%"&Request("UserLandTime")&"%'"
if Request("UserLastIP")<>"" then item=""&item&" and UserLastIP like '%"&Request("UserLastIP")&"%'"
if item="" or Request("SearchMax")="" then error2("��������Ҫ����������")
item="where"&item&""
item=replace(item,"where and","where")

sql="select top "&Request("SearchMax")&" * from [BBSXP_Users] "&item&""
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
i=i+1


%>
<TR align=middle>
<TD class=a4><a target="_blank" href=Profile.asp?UserName=<%=Rs("UserName")%>><%=Rs("UserName")%></a></TD>
<TD class=a3><%=Rs("UserMail")%></TD>
<TD class=a4><a style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','','width=320,height=170')"><img border="0" src="images/message.gif"></a></TD>
<TD class=a3><a href="Admin_User.asp?menu=User2&UserName=<%=Rs("UserName")%>">�༭</a></TD>
<TD class=a4><%=Rs("UserRegTime")%></TD>
<TD class=a3><%=Rs("UserLandTime")%></TD>
<TD class=a4><%=Rs("UserLastIP")%></TD>
<TD class=a3><input type="checkbox" value="<%=Rs("UserName")%>" name="UserName"></TD></TR>



<%
Rs.MoveNext
loop
Rs.Close
%>

	<TR align=middle class=a3>
<TD colspan="8" align="right">&nbsp;<input onclick="checkclick('��ȷ��Ҫɾ������ѡ�û���ȫ������?');"  type="submit" value="   ȷ ��   "> <input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON">&nbsp;</TD></form>
	</TR>
</TABLE>
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>
<%
end sub








sub User2

sql="select * from [BBSXP_Users] where UserName='"&HTMLEncode(UserName)&"'"
Set Rs=Conn.Execute(sql)
if Rs.eof then error2(""&UserName&" ���û����ϲ�����")

UserInfo=split(Rs("UserInfo"),"\")
realname=UserInfo(0)
country=UserInfo(1)
province=UserInfo(2)
city=UserInfo(3)
Postcode=UserInfo(4)
blood=UserInfo(5)
belief=UserInfo(6)
occupation=UserInfo(7)
marital=UserInfo(8)
education=UserInfo(9)
college=UserInfo(10)
address=UserInfo(11)
phone=UserInfo(12)
character=UserInfo(13)
personal=UserInfo(14)
UserSign=replace(Rs("UserSign"),"<br>", vbCrlf)

%>
<form method="POST" name=form action="?menu=Userok">
<input type=hidden name=UserName value="<%=Rs("UserName")%>">
<div align="center">

<table cellSpacing="1" cellPadding="3" border="0" width="60%" class=a2>
<tr class=a1 id=TableTitleLink>
<td height="25" width="600" colspan="4" align="center">
<font color="000000">
<a target="_blank" href="Profile.asp?UserName=<%=Rs("UserName")%>">
�鿴��<%=Rs("UserName")%>������ϸ����</a></font></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;�û����ƣ�<%=Rs("UserName")%></td>
<td width="300" colspan="2" height="25">&nbsp;�û����룺<a onclick="checkclick('�˲�������Ѹ��û�������ĳɣ�123');" href="?menu=fix&UserName=<%=Rs("UserName")%>">��ԭ����</a></td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;�û�Ȩ�ޣ�<select size="1" name="membercode">
<option value=0 <%if Rs("membercode")=0 then%>selected<%end if%>>��δ����</option>
<option value=1 <%if Rs("membercode")=1 then%>selected<%end if%>>��ͨ��Ա</option>
<option value=2 <%if Rs("membercode")=2 then%>selected<%end if%>>�����Ա</option>
<option value=4 <%if Rs("membercode")=4 then%>selected<%end if%>>��������</option>
<option value=5 <%if Rs("membercode")=5 then%>selected<%end if%>>����Ա</option>
</select>
</td>
<td width="300" colspan="2">&nbsp;�������᣺<input size="10" name="Consortia" value="<%=Rs("Consortia")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;�û�ͷ�Σ�<input size="10" name="UserHonor" value="<%=Rs("UserHonor")%>"></td>
<td width="300" colspan="2">&nbsp;�������£�<input size="10" name="PostTopic" value="<%=Rs("PostTopic")%>"></td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;������ң�<input size="10" name="UserMoney" value="<%=Rs("UserMoney")%>"></td>
<td width="300" colspan="2">&nbsp;�ظ����£�<input size="10" name="Postrevert" value="<%=Rs("Postrevert")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;������<input size="10" name="savemoney" value="<%=Rs("savemoney")%>"></td>
<td width="300" colspan="2">&nbsp;��ɾ���ӣ�<input size="10" name="DelTopic" value="<%=Rs("DelTopic")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;ע�����ڣ�<input size="10" name="UserRegTime" value="<%=Rs("UserRegTime")%>"></td>
<td width="300" colspan="2">&nbsp;�� �� ֵ��<input size="10" name="experience" value="<%=Rs("experience")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;�û�ͷ��<input size="50" name="Userface" value="<%=Rs("Userface")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;�û���Ƭ��<input size="50" name="Userphoto" value="<%=Rs("Userphoto")%>"></td>
</tr>
<tr class=a1 id=TableTitleLink>
<td colspan="4" align="center" height="25">
��������</td>
</tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ʵ������<input type="text" name="realname" size="10" value="<%=realname%>"><td width="50%" height="3" colspan="2">
&nbsp;�ԡ�����<input type="text" name="UserSex" size="10" value="<%=Rs("UserSex")%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;�������ڣ�<input type="text" name="birthday" size="10" value="<%=Rs("birthday")%>"><td width="50%" height="3" colspan="2">
&nbsp;�������ң�<b><input type="text" name="country" size="10" value="<%=country%>"></b></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;ʡ�����ݣ�<input type="text" name="province" size="10" value="<%=province%>"><td width="50%" height="3" colspan="2">
&nbsp;�ǡ����У�<input type="text" name="city" size="10" value="<%=city%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;������ţ�<input type="text" name="Postcode" size="10" value="<%=Postcode%>"><td width="50%" height="3" colspan="2">
&nbsp;Ѫ�����ͣ�<input maxlength=4 size=10 name=blood value="<%=blood%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;�š�������<input type="text" name="belief" size="10" value="<%=belief%>"><td width="50%" colspan="2">
&nbsp;ְ����ҵ��<input type="text" name="occupation" size="10" value="<%=occupation%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;����״����<input maxlength=4 size=10 name=marital value="<%=marital%>"><td width="50%" colspan="2">
&nbsp;���ѧ����<input type="text" name="education" size="10" value="<%=education%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ҵԺУ��<input type="text" name="college" size="20" value="<%=college%>"><td width="50%" colspan="2">
&nbsp;�绰���룺<input type="text" name="phone" size="10" value="<%=phone%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ϵ��ַ��<input type="text" name="address" size="20" value="<%=address%>"><td width="50%" colspan="2">
&nbsp;�ֻ����룺<input type="text" name="UserMobile" size="10" value="<%=Rs("UserMobile")%>"></tr>

<tr class=a4>
<td width="600" colspan="4">&nbsp;�û�ǩ����<textarea name="UserSign" rows="4" cols="50"><%=UserSign%></textarea></td>
</tr>

<tr class=a1 id=TableTitleLink>
<td width="50%" align="center" height="13">
<a onclick="checkclick('��ȷ��Ҫɾ�����û����з����������?');" href="?menu=UserDelTopic&UserName=<%=Rs("UserName")%>">
ɾ�����û�����������</a>


<td width="201" colspan="2" align="center" height="13">
<input type="submit" value=" �� �� "></td>
<td width="50%" align="center" height="13">

<a onclick="checkclick('��ȷ��Ҫɾ�����û�����������?');" href="?menu=UserDel&UserName=<%=Rs("UserName")%>">
ɾ�����û�����������</a></td>

</td>
</tr>


</table>
</form><A href="javascript:history.back()">�� ��</A>
<%
end sub

sub UserDelTopic
Conn.execute("Delete from [BBSXP_Threads] where UserName='"&UserName&"'")
%>
�Ѿ��� <%=UserName%> ���з����������ȫ��ɾ��<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub

sub UserDel

if UserName=CookieUserName then error2("�����Լ�ɾ���Լ�")

for each ho in Request("UserName")
ho=HTMLEncode(ho)
Conn.execute("Delete from [BBSXP_Users] where UserName='"&ho&"'")
next

%>
�Ѿ��ɹ�ɾ�� <%=UserName%> ����������<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub


sub Userok


sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Rs.Open sql,Conn,1,3
Rs("Userface")=Request("Userface")
Rs("Userphoto")=Request("Userphoto")
Rs("membercode")=Request("membercode")
Rs("UserHonor")=Request("UserHonor")
Rs("Consortia")=Request("Consortia")
Rs("PostTopic")=Request("PostTopic")
Rs("Postrevert")=Request("Postrevert")
Rs("experience")=Request("experience")
Rs("UserMoney")=Request("UserMoney")
Rs("savemoney")=Request("savemoney")
Rs("UserRegTime")=Request("UserRegTime")
Rs("DelTopic")=Request("DelTopic")
Rs("birthday")=Request("birthday")
Rs("UserSign")=HTMLEncode(Request.Form("UserSign"))
Rs("UserSex")=HTMLEncode(Request.Form("UserSex"))

Rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("Postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
Rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""


Rs.update
Rs.close
%> ���³ɹ�</b></font></td>
</tr></table><br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub

sub Activation
%>
<form method="POST" action=?menu=activationok>
<table cellspacing="1" cellpadding="1" width="100%" align="center" border="0" class="a2">
<TR height=25 class=a1>
		<td align="center" width="50">
		<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
		<td align="center" width="100">�û���</td>
		<td align="center" width="150">Email</td>
		<TD align="center">����¼IP</TD>
		<td align="center">ע��ʱ��</td></tr>
<%

sql="select * from [BBSXP_Users] where membercode=0 order by id Desc"
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
i=i+1

%>
<TR height=25 class=a3>
      <TD align="center"><INPUT type=checkbox value=<%=Rs("id")%> name=id></TD>
      <TD align="center"><a target="_blank" href=Profile.asp?UserName=<%=Rs("UserName")%>><%=Rs("UserName")%></a></TD>
      <TD align="center"><%=Rs("UserMail")%></TD>
      <TD align=center><%=Rs("UserLastIP")%></TD>
      <TD align=center><%=Rs("UserRegTime")%></TD>
    </TR>
<%

Rs.MoveNext
loop
Rs.Close


%>


</table>
<table border=0 width=100% align=center><tr><td>
<input type=submit  onclick="checkclick('��ȷ��Ҫ������ѡ���û�?');" value=" �� �� "></form>
</td><td valign="top" align="right">
<%ShowPage()%>
</td></tr></table>


<%
end sub


sub UserRank
%><form method="POST" action=?menu=UserRankUp>

<table border="0" cellpadding="5" cellspacing="1" class=a2 width=100%>
	<tr class=a1>
		<td width="50" align="center">ID</td>
		<td width="120">����</td>
		<td width="120">��;�����</td>
		<td>ͼ��·��</td>
		<td width="50" align="center">����</td>
	</tr>
<%
sql="select * from [BBSXP_Ranks] order by PostingCountMin"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
%>

	<tr class=a3>
		<td align="center"><%=Rs("ID")%><input type=hidden name=id value=<%=Rs("ID")%>></td>
		<td><input size="10" name="RankName<%=Rs("ID")%>" value="<%=Rs("RankName")%>"></td>
		<td><input size="10" name="PostingCountMin<%=Rs("ID")%>" value="<%=Rs("PostingCountMin")%>"></td>
		<td><input size="30" name="RankIconUrl<%=Rs("ID")%>" value="<%=Rs("RankIconUrl")%>"> <img src="<%=Rs("RankIconUrl")%>"></td>
		<td align="center"><a href="?menu=UserRankDel&RankID=<%=Rs("ID")%>">ɾ��</a></td>
	</tr>

<%
Rs.Movenext
loop
Set Rs = Nothing
%>

	<tr class=a3>
		<td align="center"><font color="#FF0000">����</font></td>
		<td><input size="10" name="RankName" value=""></td>
		<td><input size="10" name="PostingCountMin" value=""></td>
		<td colspan="2"><input size="30" name="RankIconUrl" value=""></td>
	</tr>


</table>
<input type="submit" value=" �� �� ">


<%
end sub
htmlend

%>