<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
top
UserName=HTMLEncode(Request("UserName"))
sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Set Rs1=Conn.Execute(sql)
if Rs1.eof then error("<li>"&UserName&"���û����ϲ�����")

ShowRank()
select case Rs1("UserSex")
case "male"
UserSex="��"
case "female"
UserSex="Ů"
end select
UserInfo=split(Rs1("UserInfo"),"\")
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

UserIM=split(Rs1("UserIM"),"\")
qq=UserIM(0)
icq=UserIM(1)
uc=UserIM(2)
aim=UserIM(3)
msn=UserIM(4)
Yahoo=UserIM(5)


%>
<title><%=UserName%>�û����� - Powered By BBSXP</title>
<script src="inc/birth.js"></script>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �鿴�û���<%=UserName%>������</td>
</tr>
</table><br>

<table width=100% border="0" cellspacing="0" cellpadding="5" align="center">
<tr>
<td><img src="<%=Rs1("Userface")%>" width="32" height="32">��<b><%=Rs1("UserName")%></b></td>
<td align="right" valign="bottom"><b>:::��ز���:::</b>����
<img src="images/finds.gif"> 
<a href=ShowBBS.asp?menu=5&UserName=<%=Rs1("UserName")%>> ��������</a>��

<img src="images/sig.gif">
<a href="Blog.asp?UserName=<%=Rs1("UserName")%>">������־</a>��

<img src="images/Friend.gif"> <a href=Friend.asp?menu=add&UserName=<%=Rs1("UserName")%>> ��Ϊ����</a>��
<img src="images/message.gif"> <a href=# onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs1("UserName")%>','','width=320,height=170')"> ����ѶϢ</a> ��
</tr>
</table>


<table width=100% align=center cellspacing=0 cellpadding=0 border=0>
<tr>
<td class=a2 height="676">
<table width=100% cellspacing=1 cellpadding=6 border=0>
<tr class=a1>
<td colspan="2" height="25" valign="middle" align="Left"><b>&nbsp;:::������Ϣ:::</b></td>
</tr>
<tr>
<td class=a3 height="5" align="Left" valign="middle" width="50%"><b>���ǡ����ƣ�</b><%=Rs1("UserName")%></td>
<td class=a3 height="5" align="Left" valign="middle" width="50%"><b>
���ȼ����ƣ�</b><%=RankName%></td>
</tr>
<tr>
<td class=a4 align="Left" valign="middle" width="50%"><b>���������᣺</b><%
if Rs1("Consortia")=empty then
response.write "��"
else
response.write Rs1("Consortia")
end if
%></td>
<td class=a4 align="Left" valign="middle" width="50%"><b>������ԭ����</b><%=Rs1("PostTopic")%></td>
</tr>
<tr>
<td align="Left" width="50%" class=a3><b>��ͷ�����Σ�</b><%=Rs1("UserHonor")%></td>
<td align="Left" width="50%" class=a3>

<b>�����������</b><%=Rs1("Postrevert")%></td>
</tr>
<tr class=a4>
<td align="Left" width="50%"><b>���䡡��ż��</b><%
if Rs1("Consort")=empty then
response.write "��"
else
response.write "<a href=Profile.asp?UserName="&Rs1("Consort")&">"&Rs1("Consort")&"</a>"
end if
%></td>
<td align="Left" width="50%">

<b>����¼������</b><%=Rs1("goodTopic")%></td>
</tr>
<tr class=a3>
<td align="Left" width="50%"><b>
��ע�����ڣ�</b><%=Rs1("UserRegTime")%></td>
<td align="Left" width="50%">

<b>����ɾ���ӣ�</b><%=Rs1("DelTopic")%></td>
</tr>
<tr class=a4>
<td align="Left" width="50%"><b>
���ϴε�¼��</b><%=Rs1("UserLandTime")%></td>
<td align="Left" width="50%"><b>��������ң�</b><%=Rs1("UserMoney")%></td>
</tr>
<tr class=a3>
<td align="Left" width="50%"><b>
����¼������</b><%=Rs1("UserDegree")%></td>
<td align="Left" width="50%"><b>���� �� ֵ��</b><%=Rs1("experience")%></td>
</tr>

<tr class=a1>
<td height="25" align="Left" valign="middle" colspan="2"><b>&nbsp;:::������Ϣ:::</b></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"> ��<b>��ʵ������</b><%=realname%>
</td>
<td height="100%" align="Left" valign="bottom" class=a3 rowspan="17" width="50%">
<div style="OVERFLOW:auto; HEIGHT:450px">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">

<tr><td valign=top align=center height="100%"><div align=Left><b> ���û���Ƭ��</b></div>
<script>
if("<%=Rs1("Userphoto")%>"!=""){
document.write("<a target=_blank href=<%=Rs1("Userphoto")%>><img src=<%=Rs1("Userphoto")%> border=0 onload='javascript:if(this.width>300)this.width=300'></a>")
}
</script>

</td></tr>

<tr>
<td height="100">

<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>
&nbsp;�ԡ�����</b><br>
��<table border="0">
<tr>
<td width="100%"><%=character%></td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2 height="100">
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;�򡡽飺</b><br>
<br>
<%=personal%> </td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;ǩ������</b><br>
<br><%=Rs1("UserSign")%></td>
</tr>
</table>
</td>
</tr>
</table>


</td>
</tr>


</table></div>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%"><b>���ԡ�����</b><%=UserSex%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"><b> ���������ң�</b><%=country%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>ʡ�����ݣ�</b><%=province%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>�ǡ����У�</b><%=city%></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�������룺</b><%=Postcode%></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>������Ф��</b><script>document.write(getpet("<%=Rs1("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>Ѫ�����ͣ�</b><%=blood%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>�ǡ�������</b><script>document.write(astro("<%=Rs1("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�š�������</b><%=belief%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>ְ����ҵ��</b><%=occupation%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>����״����</b><%=marital%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>���ѧ����</b><%=education%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>��ҵԺУ��</b><%=college%>
</td>
</tr>
<tr class=a1>
<td valign="top" width="50%" height="25"><b>&nbsp;:::��ϵ����:::</b></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�����ʼ���</b><a href="Mailto:<%=Rs1("UserMail")%>"><%=Rs1("UserMail")%></a></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>������ҳ��</b><a target="_blank" href="<%=Rs1("Userhome")%>"><%=Rs1("Userhome")%></a></td>
</tr>
<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;:::��ʱͨѶ:::</b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">��<b>QQ���룺</b><%if qq<>"" then%><a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=qq%>&menu=yes&Site=<%=SiteSettings("SiteName")%>><%=qq%></a><%end if%>
</td>
<td height="4" valign="middle" align="Left">��<b>ICQ��</b><%=icq%></td>
</tr>
<tr class=a4>
<td height="4" valign="middle" align="Left">��<b>UC���룺</b><%=uc%></td>
<td height="4" valign="middle" align="Left">��<b>AIM��</b><%=aim%></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">��<b>MSN IM��</b><%=msn%>��</td>
<td height="4" valign="middle" align="Left">��<b>Yahoo IM��</b><%=Yahoo%></td>
</tr>





</table>

</td>
</tr>
</table>

<center>


<br>
<INPUT onclick=history.back(-1) type=button value=" << �� �� ">
<br><%
Set Rs1 = Nothing
htmlend
%>