<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if Request("menu")="face" then
%>
<script>
function insertface(userface){
self.opener.form.Userface.value=userface;
self.opener.document.images.tus.src=userface;
window.close();
}
</script>
<body topmargin=0>
<title>BBSXP--ͷ���б� - Powered By BBSXP</title>
<table border=0 width=100% cellspacing=1 cellpadding=1>
<tr class=a1><td colspan="10" height="25" align=center>BBSXPͷ��</td></tr>
<tr align=center>
<script>
var ii=1
for (var i=1; i <= 95; i++) {
ii++
document.write("<td class=a3><img src=images/face/"+i+".gif onclick=insertface('images/face/"+i+".gif') style=CURSOR:hand><br>"+i+"</td>");
if(ii >5){document.write("</tr><tr align=center>");ii=1}
}
</script>
</tr>
</table>
<%
CloseDatabase
end if


UserName=HTMLEncode(Request.Form("UserName"))

errorchar=array(" ","��","��","��","	","","","#","`","|","%","&","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error2("�û����в��ܺ����������")
next

if SiteSettings("BannedRegUserName")<>"" then
filtrate=split(""&SiteSettings("BannedRegUserName")&"","|")
for i = 0 to ubound(filtrate)
if instr(UserName,filtrate(i))>0 then error2("����û����а�����ϵͳ��ֹע����ַ�")
next
end if






if SiteSettings("CloseRegUser")= 1 then error("<li>����̳��ʱ���������û�ע�ᣡ")

top
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request("menu")="AddUserName" then

if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>��֤�����"
end if

UserSign=HTMLEncode(Request.Form("UserSign"))

password=Trim(Request.Form("password"))
Userpass2=Trim(Request.Form("Userpass2"))
UserMail=HTMLEncode(Request.Form("UserMail"))
Userhome=HTMLEncode(Request.Form("Userhome"))
PasswordQuestion=HTMLEncode(Request.Form("PasswordQuestion"))
PasswordAnswer=HTMLEncode(Request.Form("PasswordAnswer"))
birthday=HTMLEncode(Request.Form("birthday"))
Userface=HTMLEncode(Request.Form("Userface"))
UserSex=HTMLEncode(Request.Form("UserSex"))


if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
UserSign=ReplaceText(UserSign,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


if UserName="" then Message=Message&"<li>�����û���û����д"

if Len(UserName)>16 then Message=Message&"<li>�����û����в��ܳ���16���ֽ�"

if SiteSettings("EnableUser")=1 then 
Randomize
password=int(rnd*999999)+1
else
if password="" then Message=Message&"<li>��û����������"
if password<>Userpass2 then Message=Message&"<li>��2����������벻��ͬ"
end if

if instr(UserMail,"@")=0 then Message=Message&"<li>���ĵ����ʼ���ַ��д����"

if Len(UserSign)>255 then Message=Message&"<li>ǩ�������ܴ��� 255 ���ֽ�"

if instr(Userface,";")>0 then Message=Message&"<li>ͷ��URL�в��ܺ����������"

If not Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then Message=Message&"<li>���û����Ѿ�������ע����"

if SiteSettings("OnlyMailReg") = 1 then
If not Conn.Execute("Select id From [BBSXP_Users] where UserMail='"&UserMail&"'" ).eof Then Message=Message&"<li>��Email�Ѿ�������ע����"
end if

if Message<>"" then error(""&Message&"")

for each ho in request.form("character")
allcharacter=""&allcharacter&""&ho&""
next

if SiteSettings("EnableUser")=2 then
membercode=0
else
membercode=1
end if

Rs.Open "select * from [BBSXP_Users]",Conn,1,3
Rs.addNew
Rs("UserName")=UserName
Rs("Userpass")=md5(password)
Rs("UserMail")=UserMail
Rs("Userhome")=Userhome
Rs("PasswordQuestion")=PasswordQuestion
Rs("membercode")=membercode
if Request("PasswordAnswer")<>empty then Rs("PasswordAnswer")=md5(Request("PasswordAnswer"))
Rs("birthday")=birthday
Rs("Userface")=Userface
Rs("UserSex")=UserSex
Rs("UserSign")=UserSign
Rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""
Rs("UserFriend")="|"
Rs("UserRegTime")=""&now()&""
Rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
Rs("UserLandTime")=""&now()&""
Rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("Postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
Rs("UserIM")=""&HTMLEncode(Request("qq"))&"\"&HTMLEncode(Request("icq"))&"\"&HTMLEncode(Request("uc"))&"\"&HTMLEncode(Request("aim"))&"\"&HTMLEncode(Request("msn"))&"\"&HTMLEncode(Request("Yahoo"))&""
Rs.update
Rs.close

Conn.execute("update [BBSXP_Statistics_Site] set TotalUser=TotalUser+1,NewUser='"&UserName&"'")


Mailaddress=UserMail
MailTopic="�û���ע��ɹ�"
body=""&vbCrlf&"�װ���"&UserName&", ����!"&vbCrlf&""&vbCrlf&"������ϲ! ���Ѿ��ɹ���ע������������, �ǳ���л��ʹ��"&SiteSettings("CompanyName")&"�ķ���!"&vbCrlf&""&vbCrlf&"��* �����ʺ���:"&UserName&"��������:"&password&""&vbCrlf&""&vbCrlf&"��* "&SiteSettings("SiteName")&"("&SiteSettings("SiteURL")&"Default.asp)"&vbCrlf&""&vbCrlf&"��* ���, �м���ע�����������μ�"&vbCrlf&"1�������ء��������Ϣ�������������ȫ��������취�����һ�й涨��"&vbCrlf&"2��ʹ�����ɶ������Ļ��⣬�����벻Ҫ�漰���Ρ��ڽ̵����л��⡣"&vbCrlf&"3���е�һ����������Ϊ��ֱ�ӻ��ӵ��µ����»����·������Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") �ṩ������������YUZI������(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%

if SiteSettings("EnableUser")=0 then
Response.Cookies("UserName")=escape(UserName)
Response.Cookies("Userpass")=md5(password)
end if


Message=Message&"<li>ע�����û����ϳɹ�<li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")



elseif Request("menu")="write" then


If not Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof or UserName="" Then error2("����ѡ���û��� "&UserName&" �Ѿ�����ʹ�ã�������ѡ��һ���û�����")


%>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <a href="CreateUser.asp">ע��Э��</a> �� 
<a href="CreateUser.asp?menu=Check">����û���</a> �� ��д�û�����</td>
</tr>
</table><br>

<table width=100% align=center cellspacing=0 cellpadding=0 border=0 class=a1>
<form method="POST" name="form" action=?menu=AddUserName onsubmit="return VerifyInput();">
<input type=hidden name=UserName value="<%=UserName%>">
<input type=hidden name=Userface value="images/face/1.gif">
<tr>
<td>
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr class=a1>
<td colspan="2" height="25" valign="middle" align="Left"><b>&nbsp;������������</b>���������ݱ��</td>
</tr>
<tr>
<td class=a3 height="25" align="right" width="46%"><b>�û�����</b></td>
<td class=a3 height="25" align="Left" width="54%">
&nbsp;
<%=UserName%></td>
</tr>

<%if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then%>
<tr>
<td class=a3 height="12" align="right" width="46%"><b>��֤�룺</b></td>
<td class=a3 align="Left">
&nbsp;
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="��֤��,�������?����ˢ����֤��" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></td>
</tr>
<%end if%>


<%if SiteSettings("EnableUser")<>1 then%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>���룺</b></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;
<input type="password" name="password" size="40" maxlength="16">
</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>ȷ�����룺</b><br>�����������뱣��һ��</td>
<td height="2" align="Left" valign="middle" width="54%" class=a3>
&nbsp;
<input type="password" name="Userpass2" size="40">
</td>
</tr>
<%end if%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>����Email��ַ��</b><br>
<%if SiteSettings("EnableUser")=1 then%><font color="FF0000">���뽫ͨ��Email����</font><%end if%></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;
<input type="text" name="UserMail" size="40">
</td>
</tr>
<tr class=a1>
<td height="25" align="Left" valign="middle" colspan="2"><b>&nbsp;������������</b>���������ݽ�����д��</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a3 width="46%"> ��<b>��ʵ������</b>
<input type="text" name="realname" size="18">
</td>
<td height="71" align="Left" valign="top" class=a4 rowspan="16" width="54%">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr>
<td>

<table border=0>
	<tr>
		<td><b>ͷ���� &nbsp;</b></td>
		<td>
		<table title="ѡ��ͷ��" cellspacing=1 cellpadding=0 border=0 background=Images/FaceBackground.gif width=54 style=cursor:hand onclick="javascript:open('CreateUser.asp?menu=face','','width=500,height=500,resizable,scrollbars')"><tr><td><img src="images/face/1.gif" width=40 height=40 name=tus border="0" alt="ѡ��ͷ��"></td></tr></table>
		</td>
	</tr>
</table>


</td>



</tr>
<tr>
<td><b>�����գ�<font color="000000">��</font></b> <input onfocus="calendar()" value="" name="birthday"></td>
</tr>
<tr>
<td><b>������ҳ�� </b> 
<input type="text" name="Userhome" size="40" value="http://"></td>
</tr>
<tr>
<td><b><font color="000000">�ԡ���</font>��</b>
<br>
<SCRIPT>
var moderated="�������� ���ɵ�Ƥ �������� ���ÿɰ� Ƣ������ ������ ������ �ĵ����� ��Ȥ��Ĭ ����˹�� ������ȡ �����ѻ� �������� Բ������ ���Ų�� �����쿪 ����Ƹ� �������� ������� ��С���� ѭ�浸�� �������� ���Կ��� �������� ׷��̼�"
var list= moderated.split (' '); 
for(i=0;i<list.length;i++) {
if (i==5 || i==10 || i==15 || i==20) document.write("<br>")
if (list[i] !=""){document.write("�� <input type=checkbox value=' "+list[i]+"' name=character id=character"+i+"> <label for=character"+i+">"+list[i]+"</label>")}
}
</SCRIPT>

</td>
</tr>
<tr>
<td><b>���˼�飺</b><br>
<textarea name=personal rows=6 cols=65></textarea>
</td>
</tr>
<tr>
<td height="5"><b>ǩ������</b><br>
<textarea name=UserSign rows=6 cols=65></textarea>  </td>
</tr>
</table>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%"><b> ���ԡ�����</b>
<select size=1 name=UserSex>
<option value="" selected>[��ѡ��]</option>
<option value=male>��</option>
<option value=female>Ů</option>
</select>
</td>
</tr>

<tr class=a3>
<td height="2" valign="top" width="46%">��<b>�������ң�</b>
<b>
<input type="text" name="country" size="18">
</b> </td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">��<b>ʡ�����ݣ�</b>
<input type="text" name="province" size="18">
</td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">��<b>�ǡ����У�
</b>
<input type="text" name="city" size="18">
</td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">��<b>������ţ� 
<input type="text" name="Postcode" size="18"></b></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">
��<b>Ѫ�����ͣ�</b>
<select size=1 name=blood>
<option value="" selected>[��ѡ��]</option>
<option
value=A>A</option>
<option value=B>B</option>
<option
value=AB>AB</option>
<option value=O>O</option>
<option
value=����>����</option>
</select>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">��<b>�š�������</b>
<select size=1 name=belief>
<option value="" selected>[��ѡ��]</option>
<option value=���>���</option>
<option
value=����>����</option>
<option value=������>������</option>
<option
value=������>������</option>
<option value=�ؽ�>�ؽ�</option>
<option
value=��������>��������</option>
<option value=����������>����������</option>
<option
value=����>����</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">��<b>ְ����ҵ�� </b>
<select name="occupation">
<option value="" selected>[��ѡ��]</option>
<option value="�ƻ�/����">�ƻ�/����</option>
<option value="����ʦ">����ʦ</option>
<option value="����">����</option>
<option value="����������ҵ">����������ҵ</option>
<option value="��ͥ����">��ͥ����</option>
<option value="����/��ѵ">����/��ѵ</option>
<option value="�ͻ�����/֧��">�ͻ�����/֧��</option>
<option value="������/�ֹ�����">������/�ֹ�����</option>
<option value="����">����</option>
<option value="��ְҵ">��ְҵ</option>
<option value="����/�г�/���">����/�г�/���</option>
<option value="ѧ��">ѧ��</option>
<option value="�о��Ϳ���">�о��Ϳ���</option>
<option value="һ�����/�ල">һ�����/�ල</option>
<option value="����/����">����/����</option>
<option value="ִ�й�/�߼�����">ִ�й�/�߼�����</option>
<option value="����/����/����">����/����/����</option>
<option value="רҵ��Ա">רҵ��Ա</option>
<option value="�Թ�/ҵ��">�Թ�/ҵ��</option>
<option value="����">����</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">��<b>����״����</b>
<select size=1 name=marital>
<option value="" selected>[��ѡ��]</option>
<option value=δ��>δ��</option>
<option
value=�ѻ�>�ѻ�</option>
<option value=����>����</option>
<option
value=ɥż>ɥż</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">��<b>���ѧ����</b>
<select size=1 name=education>
<option value="" selected>[��ѡ��]</option>
<option value=Сѧ>Сѧ</option>
<option
value=����>����</option>
<option value=����>����</option>
<option
value=��ѧ>��ѧ</option>
<option value=˶ʿ>˶ʿ</option>
<option
value=��ʿ>��ʿ</option>
</select></td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">��<b>��ҵԺУ��</b>
<input type="text" name="college" size="18"></td>
</tr>

<tr class=a1>
<td valign="top" height="25"><b>&nbsp;˽����ϵ����</b>���������ϲ�������</td>
</tr>

<tr class=a3>
<td valign="top" width="46%">��<b>��ϵ��ַ�� 
<input type="text" name="address" size="18"></b></td>
</tr>

<tr class=a3>
<td valign="top" width="46%">��<b>�绰���룺 
<input type="text" name="phone" size="18"></b></td>
</tr>

<tr class=a3>
<td valign="top" width="46%">��<b>�ֻ����룺 
<input type="text" name="UserMobile" size="18"></b></td>
</tr>

<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;��ʱͨѶ����</b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">��<b>QQ���룺 
<input type="text" name="qq" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
<td height="4" valign="middle" align="Left"><b>ICQ��������<input type="text" name="icq" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
</tr>
<tr class=a4>
<td height="4" valign="middle" align="Left">��<b>UC���룺 <input type="text" name="uc" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
<td height="4" valign="middle" align="Left"><b>AIM��������<input type="text" name="aim" size="20"></b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">��<b>MSN IM�� 
<input type="text" name="msn" size="20"></b>��</td>
<td height="4" valign="middle" align="Left"><b>Yahoo IM�� <input type="text" name="Yahoo" size="20"></b></td>
</tr>
<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;���뱣������</b></td>
</tr>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>������ʾ���⣺</b></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;


<select name="PasswordQuestion">
<option value="" selected>[��ѡ��]</option>
<option value="��ϲ���ĳ��">��ϲ���ĳ��</option>
<option value="��ϲ���ĵ�Ӱ��">��ϲ���ĵ�Ӱ��</option>
<option value="��������� [��/��/��]��">��������� [��/��/��]��</option>
<option value="���׵����֣�">���׵����֣�</option>
<option value="��ż�����֣�">��ż�����֣�</option>
<option value="��һ�����ӵİ��ƣ�">��һ�����ӵİ��ƣ�</option>
<option value="��ѧ��У����">��ѧ��У����</option>
<option value="���𾴵���ʦ��">���𾴵���ʦ��</option>
<option value="��ϲ�������жӣ�">��ϲ�������жӣ�</option>
</select>

</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>������ʾ�𰸣�</b></td>
<td height="2" align="Left" valign="middle" width="54%"> &nbsp;
<input type="text" name="PasswordAnswer" size="40">
</td>
</tr>

<tr align="center" class=a4>
<td height="2" valign="middle" colspan="2">
<input type="submit" value=" ע �� >>�� һ �� ">
</td>
</tr>
</table>
</td>
</tr></table>



<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="�й�"}
else
if (language=="zh-hk"){country="�й����"}
else
if (language=="zh-tw"){country="�й�̨��"}
else
if (language=="zh-sg"){country="�¼���"}
else
country=""
document.form.country.value=country


function VerifyInput()
{
var Mail = document.form.UserMail.value;
if(Mail.indexOf('@',0) == -1 || Mail.indexOf('.',0) == -1){
alert("�������Email�д���\n�����¼������Email");
document.form.UserMail.focus();
return false;
}

if (document.form.UserMail.value == "")
{
alert("����������Email��ַ");
document.form.UserMail.focus();
return false;
}

return true;
}
</SCRIPT>
<%
elseif Request("menu")="" then
%>
<center>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ע��Э��</td></tr></tbody></table>
<br>
<textarea name="textarea" rows="18" readOnly cols="100"><%=SiteSettings("RegUserAgreement")%></textarea><br><br>
<input type="submit" value=" ͬ �� " onclick="window.location.href='CreateUser.asp?menu=Check';">
<input type="submit" value=" ��ͬ�� " onclick="history.back()">
<%
elseif Request("menu")="Check" then
%>
<center><form method="POST" name="form" action=?menu=write>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� <a href="CreateUser.asp">ע��Э��</a> �� ����û���</td></tr></tbody></table>
<br>
<br>��������Ҫע����û���<BR>
<input type="text" name="UserName" size="28" maxlength="12" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')">
<br><br>
<input type="button" value=" ��һ�� " onclick="history.back()">
<input type="submit" value=" ��һ�� " id='btnadd' disabled></form>
<%
end if

htmlend
%>