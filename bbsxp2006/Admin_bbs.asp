<!-- #include file="Setup.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


id=HTMLEncode(Request("id"))
bbsid=HTMLEncode(Request("bbsid"))
TimeLimit=HTMLEncode(Request("TimeLimit"))
UserName=HTMLEncode(Request("UserName"))


response.write "<center>"



select case Request("menu")
case ""
error2("��ѡ����Ҫ��������Ŀ")

case "ApplyManage"
ApplyManage

case "activation"
activation

case "bbsManage"
bbsManage

case "bbsManagexiu"
bbsManagexiu


case "bbsManagexiuok"
bbsManagexiuok

case "bbsadd"
bbsadd

case "bbsaddok"
bbsaddok

case "classs"
classs


case "upSiteSettings"
upSiteSettings

case "upSiteSettingsok"
upSiteSettingsok


case "bbsManageDel"
Conn.execute("Delete from [BBSXP_Forums] where id="&id&"")
error2("�Ѿ�������̳����������ɾ���ˣ�")



case "Delforumok"
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where lasttime<"&SqlNowString&"-"&TimeLimit&" "&BbsIdList&"")
error2("�Ѿ���"&TimeLimit&"��û�и��¹�������ɾ���ˣ�")


case "DelUserTopicok"
if UserName="" then error2("��û�������û�����")
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where UserName='"&UserName&"' "&BbsIdList&"")
error2("�Ѿ���"&UserName&"���������ɾ���ˣ�")


case "DellikeTopicok"
Topic=HTMLEncode(Request("Topic"))
if Topic="" then error2("��û�������ַ���")
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where Topic like '%"&Topic&"%' "&BbsIdList&" ")
error2("�Ѿ�������������� "&Topic&" ������ȫ��ɾ���ˣ�")


case "DelForumsok"
Conn.execute("Delete from [BBSXP_Forums] where ForumHide=1 and lasttime<"&SqlNowString&"-"&TimeLimit&"")
error2("�Ѿ���"&TimeLimit&"��û�������ӵ���̳ɾ���ˣ�")

case "clean"
Conn.execute("Delete from [BBSXP_Threads] where IsDel=1 and lasttime<"&SqlNowString&"-"&TimeLimit&"")
error2("�Ѿ���� "&TimeLimit&" ����ǰ�����⣡")

case "uniteok"
hbbs=Request("hbbs")
YBBs=Request("YBBs")
if hbbs = YBBs then error2("����ѡ����ͬ��̳��")
if UserName<>"" then UserNamelist="and UserName='"&UserName&"'"
Conn.execute("update [BBSXP_Threads] set ForumID="&int(hbbs)&" where ForumID="&int(YBBs)&" and lasttime<"&SqlNowString&"-"&TimeLimit&" "&UserNamelist&"")
error2("�ƶ���̳���ϳɹ���")

case "BatchCensorship"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsDel=0 where id="&ho&"")
next
error2("�Ѿ��� ����/��ԭ ��ѡ���ӣ�")
case "BatchDel"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("Delete from [BBSXP_Threads] where id="&ho&"")
next
error2("�Ѿ���ɾ����ѡ���ӣ�")

case "Delapplication"
Application.contents.ReMoveAll()
error2("�Ѿ���������������е�application���棡")





end select

sub ApplyManage
%>

<table cellspacing=1 cellpadding=2 width=100% border=0 class=a2 align=center>
<tr class=a1 id=TableTitleLink>
<td align="center" height="25"><a href="?menu=ApplyManage&fashion=id">ID</a></td>
<td width="20%" align="center" height="25">
<a href="?menu=ApplyManage&fashion=ForumName">��̳</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumHide">����</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=moderated">����</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumToday">����</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumThreads">����</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumPosts">����</a></td>
<td align="center" height="25">����</td>
</tr>
<%

if Request("fashion")=empty then
fashion="ForumPosts"
else
fashion=Request("fashion")
end if


sql="select * from [BBSXP_Forums] order by "&fashion&" Desc"
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
<tr class=a3>
<td align="center" height="25"><%=Rs("id")%></td>
<td><a target=_blank href=ShowForum.asp?ForumID=<%=Rs("id")%>><%=Rs("ForumName")%></a></td>
<td align=center><%if ""&Rs("ForumHide")&""="1" then%><font color="#FF0000">����</font><%elseif ""&Rs("followid")&""="0" then%><font color="#0000FF">���</font><%else%>����<%end if%></td>
<td align=center width="200"><%=Rs("moderated")%></td>
<td align=center><b><font color=red><%=Rs("ForumToday")%></font></b></td>
<td align=center><b><font color=red><%=Rs("ForumThreads")%></font></b></td>
<td align=center><b><font color=red><%=Rs("ForumPosts")%></font></b></td>
<td align="center"><a href=?menu=bbsManagexiu&id=<%=Rs("id")%>>�༭��̳</a> | <a onclick=checkclick('��ȷ��Ҫɾ������̳����������?') href=?menu=bbsManageDel&id=<%=Rs("id")%>>ɾ����̳</a>
</tr>
<%
Rs.MoveNext
loop
Rs.Close

%>
</table>
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td>
</tr></table>

<%

end sub


sub classs
%>

<table border="0" width="80%">
	<tr>
		<td align="center">
<form name="form" method="POST" action="?menu=bbsaddok"><input type=hidden name=classid value=0><input type=hidden name=ForumPass value=1><input type=hidden name=ForumHide value=0>
������ƣ������磺�������磩<input name="ForumName" onkeyup="ValidateTextboxAdd(this, 'ForumName1')" onpropertychange="ValidateTextboxAdd(this, 'ForumName1')"><input type="submit" value="����" id='ForumName1' disabled></form>
</td>
		<td align="center">
<form method="POST" action="?menu=bbsManagexiu">
������̳��<INPUT size=2 name=id onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="ID" Helptext="ID">
<input type=submit value="ȷ��" id='btnadd' disabled></form>
</td>
	</tr>
</table>


<table cellspacing=1 cellpadding="2" width="80%" border="0" class="a2" align="center">
<%
sort(0)
%>
</table>


<%

end sub

sub bbsadd

%>


<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">������̳����</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

<form name="form" method="POST" action="?menu=bbsaddok">
<input type=hidden name=classid value=<%=id%>>

��̳����</td>
    <td class=a3>

<input size="20" name="ForumName"></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<input size="30" name=moderated> �����������á�|���������磺yuzi|ԣԣ
</td></tr>

<tr class=a3>
<td height="2" align="center" width="20%">����ר��</td>
<td height="2" align="Left" valign="middle" width="77%">
<input size="30" name="TolSpecialTopic"> 
������á�|���������磺ԭ��|ת��</td></tr>

   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<textarea rows="5" name="ForumIntro" cols="50"></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<textarea rows="5" name="ForumRules" cols="50"></textarea></td></tr>

   <tr height=25>
    <td class=a3 align=middle>

��̳״̬</td>
    <td class=a3>

<select size="1" name="ForumPass">
<option value=0>�ر�</option>
<option value=1 selected>����</option>
<option value=2>�ο�ֹ��</option>
<option value=3>��Ȩ���</option>
<option value=4>��Ȩ����</option>
<option value=5>��鷢��</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

��Ȩ�û�����</td>
    <td class=a3>
<input size="30" name="ForumUserList">  ���á�|���������磺yuzi|ԣԣ
</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

Сͼ��URL</td>
    <td class=a3>

<input size="30" name="ForumIcon"> ��ʾ��������ҳ��̳�����ұ�
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��ͼ��URL</td>
    <td class=a3>

<input size="30" name="ForumLogo"> ��ʾ����̳���Ͻ�
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

ͨ������</td>
    <td class=a3>

<input size="30" name="ForumPassword"> ����ǹ�����̳���˴�������</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

�Ƿ���ʾ����̳�б���������������������
    <td class=a3>

<input type="radio" CHECKED value="0" name="ForumHide">��ʾ 
<input type="radio" value="1" name="ForumHide">����
</tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">

�� <input type="submit" value=" �� �� "><br></td></tr></table>
</form>
<center><br><a href=javascript:history.back()>< �� �� ></a>


<%
end sub
sub bbsaddok
if Request("ForumName")="" then error2("��������̳����")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"���û����ϻ�δע��")
next


ForumUserList=replace(Request("ForumUserList"),vbCrlf,"")

Rs.Open "[BBSXP_Forums]",Conn,1,3
Rs.addNew
Rs("followid")=Request("classid")
Rs("ForumName")=HTMLEncode(Request("ForumName"))
Rs("moderated")=Request("moderated")
Rs("TolSpecialTopic")=Request("TolSpecialTopic")
Rs("ForumPass")=Request("ForumPass")
Rs("ForumPassword")=Request("ForumPassword")
Rs("ForumUserList")=ForumUserList
Rs("ForumIntro")=HTMLEncode(Request.Form("ForumIntro"))
Rs("ForumRules")=HTMLEncode(Request.Form("ForumRules"))
Rs("ForumIcon")=Request("ForumIcon")
Rs("ForumLogo")=Request("ForumLogo")
Rs("ForumHide")=Request("ForumHide")
Rs.update
id=Rs("id")

Rs.close

classs

end sub




sub bbsManagexiuok

if Request("ForumName")="" then error2("��������̳����")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"���û����ϻ�δע��")
next

ForumUserList=replace(Request("ForumUserList"),vbCrlf,"")

sql="select * from [BBSXP_Forums] where id="&id&""
Rs.Open sql,Conn,1,3


Rs("followid")=Request("classid")
Rs("SortNum")=int(Request("SortNum"))
Rs("ForumName")=HTMLEncode(Request("ForumName"))
Rs("moderated")=Request("moderated")
Rs("TolSpecialTopic")=Request("TolSpecialTopic")
Rs("ForumPass")=Request("ForumPass")
Rs("ForumPassword")=Request("ForumPassword")
Rs("ForumUserList")=ForumUserList
Rs("ForumIntro")=HTMLEncode(Request.Form("ForumIntro"))
Rs("ForumRules")=HTMLEncode(Request.Form("ForumRules"))
Rs("Forumicon")=Request("Forumicon")
Rs("ForumLogo")=Request("ForumLogo")
Rs("ForumHide")=Request("ForumHide")
Rs.update

Rs.close
%>
�༭�ɹ�<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub


sub bbsManagexiu

sql="select * from [BBSXP_Forums] where id="&id&""
Set Rs=Conn.Execute(sql)
if Rs.EOF then error2("ϵͳ�����ڸ���̳������")
ForumIntro=replace(""&Rs("ForumIntro")&"","<br>",vbCrlf)
ForumRules=replace(""&Rs("ForumRules")&"","<br>",vbCrlf)
%>


<form method="POST" action="?menu=bbsManagexiuok" name=form><input type=hidden name=id value=<%=Rs("id")%>>
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">�༭��̳����</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>


<input size="15" name="ForumName" value="<%=Rs("ForumName")%>"> &nbsp; ���� <input size="2" name="SortNum" value="<%=Rs("SortNum")%>">
��С��������</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳���</td>
    <td class=a3>


<select name="classid">
<option value="<%=Rs("followid")%>">Ĭ��</option>
<option value="0">һ������</option>
<%BBSList(0)%><%=ForumsList%>
</select> һ�����ཫ�����������</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��̳����</td>
    <td class=a3>


<input size="30" name=moderated value="<%=Rs("moderated")%>"> �����������á�|���������磺yuzi|ԣԣ
</td></tr>

<tr class=a3>
<td height="2" align="center" width="20%">����ר��</td>
<td height="2" align="Left" valign="middle" width="77%">
<input size="30" name="TolSpecialTopic" value="<%=Rs("TolSpecialTopic")%>"> 
������á�|���������磺ԭ��|ת��</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��̳����</td>
    <td class=a3>


<textarea rows="5" name="ForumIntro" cols="50"><%=ForumIntro%></textarea></td></tr>


   <tr height=25>
    <td class=a3 align=middle>

��̳����</td>
    <td class=a3>

<textarea rows="5" name="ForumRules" cols="50"><%=ForumRules%></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

��̳״̬</td>
    <td class=a3>

<select size="1" name="ForumPass">
<option value=0 <%if Rs("ForumPass")=0 then%>selected<%end if%>>�ر�</option>
<option value=1 <%if Rs("ForumPass")=1 then%>selected<%end if%>>����</option>
<option value=2 <%if Rs("ForumPass")=2 then%>selected<%end if%>>�ο�ֹ��</option>
<option value=3 <%if Rs("ForumPass")=3 then%>selected<%end if%>>��Ȩ���</option>
<option value=4 <%if Rs("ForumPass")=4 then%>selected<%end if%>>��Ȩ����</option>
<option value=5 <%if Rs("ForumPass")=5 then%>selected<%end if%>>��鷢��</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

��Ȩ�û�����</td>
    <td class=a3>
<input size="30" name="ForumUserList" value="<%=Rs("ForumUserList")%>"> ������á�|���������磺yuzi|ԣԣ
</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


Сͼ��URL</td>
    <td class=a3>


<input size="30" name="ForumIcon" value="<%=Rs("ForumIcon")%>"> ��ʾ��������ҳ��̳�����ұ�
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


��ͼ��URL</td>
    <td class=a3>


<input size="30" name="ForumLogo" value="<%=Rs("ForumLogo")%>"> ��ʾ����̳���Ͻ�</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

ͨ������</td>
    <td class=a3>

<input size="30" name="ForumPassword" value="<%=Rs("ForumPassword")%>"> ����ǹ�����̳���˴�������</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


�Ƿ���ʾ����̳�б�</td>
    <td class=a3>


<input type="radio" <%if Rs("ForumHide")=0 then%>CHECKED <%end if%>value="0" name="ForumHide" value="0">��ʾ 
<input type="radio" <%if Rs("ForumHide")=1 then%>CHECKED <%end if%>value="1" name="ForumHide" value="1">���� </td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">


<input type="submit" value=" �� �� "></td></tr></table><br></form>
<center><br><a href=javascript:history.back()>< �� �� ></a>

<%
end sub



sub bbsManage
BBSList(0)
%>


��̳����<b><font color=red><%=Conn.execute("Select count(id)from [BBSXP_Forums]")(0)%></font></b>������������<b><font color=red><%=Conn.execute("Select count(id)from [BBSXP_Threads]")(0)%></font></b><br>



<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="3">����ɾ������</td>
  </tr>


<form method="POST" action="?menu=Delforumok">
<tr height=25><td class=a3 align=middle width="50%">
ɾ�� <INPUT size=3 name=TimeLimit value="180">
��û�и��µ�����</td>
    <td class=a3 align=middle>
<select name="bbsid">
<option value="">������̳</option>
<%=ForumsList%>
</select></td>
    <td class=a3 align=middle>
 <input type="submit" value=" ȷ �� "></td></form></tr>

   <tr>
    <td class=a4 align=middle>
    <form method="POST" action="?menu=DelUserTopicok">
ɾ�� <input size="10" name="UserName"> �������������
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">������̳</option>
<%=ForumsList%>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" ȷ �� "></tr>
	</tr></form>

   <tr height=25>
    <td class=a4 align=middle>
    <form method="POST" action="?menu=DellikeTopicok">
ɾ������������� <input size="10" name="Topic"> ����������
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">������̳</option>
<%=ForumsList%>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" ȷ �� "></tr></form>

</table>

<br>




<form method="POST" action="?menu=DelForumsok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>����ɾ����̳</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>
ɾ�� <INPUT size=3 name=TimeLimit value="90">
��û�������ӵ���̳

<input type="submit" value=" ȷ �� "></td>
    </tr>
    </table></form>



<form method="POST" action="?menu=uniteok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=3>�ƶ���̳����</td>
  </tr>
   <tr>
    <td class=a3 align=middle>

�� <select name="YBBs">
<%=ForumsList%>
</select> �ƶ���
</td>
    <td class=a3 align=middle>

<select name="hbbs">
<%=ForumsList%>
</select>
</td>
    <td class=a3 align=middle rowspan="2">

<INPUT type=submit value=" ȷ �� "></td>
	</tr>
   <tr>
    <td class=a3 align=middle width="49%">

���ƶ�

<input size="2" name="TimeLimit" value="0"> ��ǰ������
</td>
    <td class=a3 align=middle>

���ƶ�

<input size="8" name="UserName"> ���������</td>
    </tr>
   </table></form>






<%
end sub

sub upSiteSettings
%>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�������ϸ���</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
�˲�����������̳���ϣ��޸���̳ͳ�Ƶ���Ϣ<br>
<a href="?menu=upSiteSettingsok">������������̳ͳ������</a><br>


<a href="?menu=Delapplication">����������ϵ�application����</a><br>



</td></tr></table><br>

<%
end sub

sub upSiteSettingsok


Rs.Open "[BBSXP_Forums]",Conn
do while not Rs.eof

allarticle=Conn.execute("Select count(ID) from [BBSXP_Threads] where IsDel=0 and ForumID="&Rs("id")&"")(0)
if allarticle>0 then
allrearticle=Conn.execute("Select sum(replies) from [BBSXP_Threads] where IsDel=0 and ForumID="&Rs("id")&"")(0)
else
allrearticle=0
end if

Conn.execute("update [BBSXP_Forums] set ForumThreads="&allarticle&",ForumPosts="&allarticle+allrearticle&" where ID="&Rs("id")&"")


Rs.Movenext
loop
Rs.close

%>
�����ɹ���<br>
<br>
�Ѿ���������������̳��ͳ������<br>

<%
end sub

htmlend

sub activation
if Request("type")="Recycle" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime<>lasttime order by lasttime Desc"
response.write "�� �� վ"
elseif Request("type")="Censorship" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime=lasttime order by lasttime Desc"
response.write "�� �� ��"
end if
%>
<form method="POST" action=?>
<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
<%
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
ShowThread()


Rs.MoveNext
loop
Rs.Close


%>
    


</table>
<table border=0 width=100% align=center><tr><td valign="top">
<input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">ȫѡ
<%
if Request("type")="Recycle" then
%><input type="radio" value="BatchCensorship" name=menu>��ԭ����<%
elseif Request("type")="Censorship" then
%><input type="radio" value="BatchCensorship" name=menu>ͨ�����<%
end if
%>
 <input type="radio" value="BatchDel" name=menu>����ɾ��	
<input onclick="checkclick('��ȷ��ִ�б��β���?');" type="submit" value=" ִ �� ">

</td></form>
	<td>
	<form method="POST" action="?menu=clean">��� <input maxlength="1" size="1" value="7" name="TimeLimit"> ����ǰ������ <input type="submit" onclick="checkclick('ִ�б���������ջ���վ�������������?');" value="ȷ��"></form>
	</td>
<td align="right" valign="top">
<%ShowPage()%></td></tr></table>

<%
end sub

dim ii
ii=0
sub sort(selec)
	sql="Select * From [BBSXP_Forums] where followid="&selec&" and ForumHide=0 order by SortNum"
	Set Rs1=Conn.Execute(sql)


	do while not rs1.eof

if selec=0 then
%>
  <tr class=a1 id=TableTitleLink height=25>
<td>��<a target=_blank href=ShowForum.asp?ForumID=<%=rs1("id")%>><%=rs1("ForumName")%></a></td>
<td align="right" width="190">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>������̳</a> | <a href=?menu=bbsManagexiu&id=<%=rs1("id")%>>�༭��̳</a> | 
<a onclick=checkclick('��ȷ��Ҫɾ������̳����������?') href=?menu=bbsManageDel&id=<%=rs1("id")%>>ɾ����̳</a>
</tr>

<%
else
%>
<tr class=a3 height=25>
<td>��<%=string(ii*2,"��")%><a target=_blank href=ShowForum.asp?ForumID=<%=rs1("id")%>><%=rs1("ForumName")%></a></td>
<td align="right">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>������̳</a> | <a href=?menu=bbsManagexiu&id=<%=rs1("id")%>>�༭��̳</a> | 
<a onclick=checkclick('��ȷ��Ҫɾ������̳����������?') href=?menu=bbsManageDel&id=<%=rs1("id")%>>ɾ����̳</a>
</tr>
<%
end if
ii=ii+1
	sort rs1("id")
ii=ii-1
	rs1.Movenext
	loop
	Set Rs1 = Nothing
end sub

%>