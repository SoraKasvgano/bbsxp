<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->

<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")


select case Request.form("menu")
case "editProfileok"
editProfileok
case "passok"
passok
end select


%>
<SCRIPT src="inc/birthday.js"></SCRIPT>

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
<%


select case Request("menu")
case ""
index
case "pass"
pass
end select


sub index
sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(sql)

UserIM=split(Rs("UserIM"),"\")
qq=UserIM(0)
icq=UserIM(1)
uc=UserIM(2)
aim=UserIM(3)
msn=UserIM(4)
Yahoo=UserIM(5)

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
character=replace(character,"<br>", vbCrlf)
personal=replace(personal,"<br>", vbCrlf)
UserSign=replace(Rs("UserSign"),"<br>",vbCrlf)

%>
<table width=100% cellspacing=1 cellpadding=4 align=center border=0 class=a2>
<form method="POST" name="form" action="EditProfile.asp">
<input type=hidden name=Userface value="<%=Rs("Userface")%>">
<input type=hidden name=menu value="editProfileok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>&nbsp;:::��������:::</b></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%"> ��<b>��ʵ������</b> 
      <input type="text" name="realname" size="17" value="<%=realname%>">
</td>
    <td height="71" align="Left" valign="top" class=a4 rowspan="12" width="64%"> 
      <table width="486" cellspacing="0" cellpadding="5">
<tr>
<td width="476">

<table cellspacing=0 cellpadding=0 border=0>
	<tr>
		<td><b>ͷ���� &nbsp;</b></td>
		<td>
		<table title="ѡ��ͷ��" cellspacing=1 cellpadding=0 border=0 background=Images/FaceBackground.gif width=54 style=cursor:hand onclick="javascript:open('CreateUser.asp?menu=face','','width=500,height=500,resizable,scrollbars')"><tr><td><img src="<%=Rs("Userface")%>" width=40 height=40 name=tus border="0"></td></tr></table>
		</td>
	</tr>
</table>


</td>
</tr>
<tr>
<td width="466"><b>�����գ�</b> <input onfocus="calendar()" value="<%=Rs("birthday")%>" name="birthday"></td>
</tr>
<tr>
<td width="476"><b>�ԡ���</b><br>
<textarea name=character rows=5 cols=65><%=character%></textarea><font color="000000">�� </font>
</td>
</tr>
<tr>
<td width="476"><b>���˼�飺</b><br>
<textarea name=personal rows=5 cols=65><%=personal%></textarea>
</td>
</tr>
<tr>
<td height="10" width="476"><b>ǩ������</b><br>
<textarea name=UserSign rows=5 cols=65><%=UserSign%></textarea>  </td>
</tr>
</table>
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%"><b> ���ԡ�����</b> 
      <select size=1 name=UserSex>
<option value="" selected>[��ѡ��]</option>
<option value=male <%if Rs("UserSex")="male" then%>selected<%end if%>>��</option>
<option value=female <%if Rs("UserSex")="female" then%>selected<%end if%>>Ů</option>
</select>



</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">��<b>�������ң�</b> <b> 
      <input type="text" name="country" size="17" value="<%=country%>">
</b> </td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">��<b>ʡ�����ݣ�</b> 
      <input type="text" name="province" size="17" value="<%=province%>">
</td>
</tr>
<tr>
    <td height="1" valign="top" class=a4 width="34%">��<b>�ǡ����У� </b> 
      <input type="text" name="city" size="17" value="<%=city%>">
</td>
</tr>
<tr>
    <td height="1" valign="top" class=a3 width="34%">��<b>�������룺 </b> 
      <input type="text" name="Postcode" size="17" value="<%=Postcode%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">��<b>Ѫ�����ͣ�</b> 
      <input maxlength=4 size=4
name=blood value="<%=blood%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">��<b>�š�������</b> 
      <input type="text" name="belief" size="17" value="<%=belief%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">��<b>ְ����ҵ��</b> 
      <input type="text" name="occupation" size="17" value="<%=occupation%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">��<b>����״����</b> 
      <input maxlength=4 size=17 name=marital value="<%=marital%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">��<b>���ѧ����</b> 
      <input type="text" name="education" size="17" value="<%=education%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">��<b>��ҵԺУ��</b> 
      <input type="text" name="college" size="17" value="<%=college%>"></td>
</tr>
<tr class=a1>
<td height="25" align="Left" colspan="2"><b>&nbsp;:::��ϵ����:::</b></td>
</tr>

<tr class=a3>
<td valign="middle" align="Left">��<b>�绰���룺</b><input type="text" name="phone" size="20" value="<%=phone%>"></td>
<td valign="middle" align="Left"><b>�ֻ����룺</b><input type="text" name="UserMobile" size="20" value="<%=Rs("UserMobile")%>"></td>
</tr>

<tr class=a4>
<td valign="middle" align="Left" colspan="2">��<b>��ϵ��ַ��</b><input type="text" name="address" size="60" value="<%=address%>"></td>
</tr>

<tr class=a4>
<td valign="middle" align="Left" colspan="2">��<b>������ҳ��</b><input type="text" name="Userhome" size="60" value="<%=Rs("Userhome")%>"></td>
</tr>



<tr class=a1>
<td colspan="2" height="25"><b>&nbsp;:::�ɣ�����:::</b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">��<b>QQ���룺 
<input type="text" name="qq" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=qq%>"></b></td>
<td height="4" valign="middle" align="Left"><b>ICQ��������<input type="text" name="icq" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=icq%>"></b></td>
</tr>
<tr class=a4>
<td height="4" valign="middle" align="Left">��<b>UC���룺 <input type="text" name="uc" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=uc%>"></b></td>
<td height="4" valign="middle" align="Left"><b>AIM��������<input type="text" name="aim" size="20" value="<%=aim%>"></b></td>
</tr>
<tr class=a3>
<td height="2" valign="middle" align="Left">��<b>MSN IM�� <input type="text" name="msn" size="20" value="<%=msn%>"></b>��</td>
<td height="2" valign="middle" align="Left"><b>Yahoo IM�� <input type="text" name="Yahoo" size="20" value="<%=yahoo%>"></b></td>
</tr>

<tr class=a3>
    <td height="1" align="center" width="100%" colspan="2">

<input type="submit" value=" ȷ �� "></td>
</tr>
</table>



</form>

<%
end sub




sub editProfileok


UserSign=HTMLEncode(Request.Form("UserSign"))
Userface=HTMLEncode(Request("Userface"))

if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
UserSign=ReplaceText(UserSign,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


if Len(UserSign)>255 then Message=Message&"<li>ǩ�������ܴ��� 255 ���ֽ�"

if instr(Userface,";")>0 then Message=Message&"<li>ͷ��URL�в��ܺ����������"

if Message<>"" then error(""&Message&"")

sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3

Rs("birthday")=HTMLEncode(Request("birthday"))
Rs("Userface")=Userface
Rs("UserSex")=HTMLEncode(Request("UserSex"))
Rs("UserSign")=UserSign

Rs("UserIM")=""&HTMLEncode(Request("qq"))&"\"&HTMLEncode(Request("icq"))&"\"&HTMLEncode(Request("uc"))&"\"&HTMLEncode(Request("aim"))&"\"&HTMLEncode(Request("msn"))&"\"&HTMLEncode(Request("Yahoo"))&""
Rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("Postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
Rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""
Rs("Userhome")=""&HTMLEncode(Request("Userhome"))&""


Rs.update
Rs.close
Message=Message&"<li>�޸����ϳɹ�<li><a href=UserCp.asp>���������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=UserCp.asp>")
end sub







sub pass
sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(sql)
%>

<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="passok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>�û������޸�</b></td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b> ԭ���룺</b></td>
    <td height="2" align="Left" width="55%"> &nbsp; 
      <input type="password" name="oldUserpass" size="40">
</td>
</tr>
<tr class=a3>
    <td height="2" align="right" width="45%"><b> �����룺</b><br>
      �粻��������˴�������</td>
    <td height="2" align="Left" width="55%"> &nbsp; 
      <input type="password" name="Userpass" size="40" maxlength="16">
</td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b>ȷ�������룺</b><br>
      �������������뱣��һ��</td>
    <td height="2" align="Left" valign="middle" width="55%"> &nbsp; 
      <input type="password" name="Userpass2" size="40" maxlength="16"></td>
</tr>
<tr class=a3>
    <td align="right" width="45%"><b>�ʼ���ַ��</b></td>
    <td align="Left" width="55%" class=a4> &nbsp; 
      <input type="text" name="UserMail" size="40" value="<%=Rs("UserMail")%>"></td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>������ʾ���⣺</b></td>
    <td height="1" align="Left" width="55%"> &nbsp; 
      <input type="text" name="PasswordQuestion" size="40" value="<%=Rs("PasswordQuestion")%>">
</td>
</tr>
<tr class=a3>
    <td height="1" align="right" width="45%"><b>������ʾ�𰸣�</b><br>
	�粻���Ĵ𰸴˴�������</td>
    <td height="1" align="Left" width="55%" class=a4> &nbsp; 
      <input type="text" name="PasswordAnswer" size="40" value=""></td>
</tr>
<tr class=a3>
    <td align="center" width="100%" colspan="2">
<input type="submit" value=" ȷ �� "></td>
</tr>
</table>

</form>
<%
end sub

sub passok
Userpass=Trim(Request("Userpass"))
oldUserpass=Trim(Request("oldUserpass"))
Userpass2=Trim(Request("Userpass2"))
PasswordQuestion=HTMLEncode(Request("PasswordQuestion"))
UserMail=HTMLEncode(Request("UserMail"))


if instr(UserMail,"@")=0 then error("<li>���ĵ����ʼ���ַ��д����")

sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Rs.Open sql,Conn,1,3


if md5(oldUserpass)<>Rs("Userpass") then Message=Message&"<li>����ԭ�������"

if Userpass<>Userpass2 then Message=Message&"<li>�����������ȷ�������벻ͬ"

if Message<>"" then error(""&Message&"")
if Userpass<>empty then Rs("Userpass")=md5(Userpass)
Rs("PasswordQuestion")=PasswordQuestion
if Request("PasswordAnswer")<>empty then Rs("PasswordAnswer")=md5(Request("PasswordAnswer"))
Rs("UserMail")=UserMail
Rs.update
Rs.close
Message=Message&"<li>�޸����ϳɹ�<li><a href=UserCp.asp>���������ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=UserCp.asp>")
end sub


htmlend
%>