<!-- #include file="Setup.asp" -->


<%
if CookieUserName=empty then error2("���¼�����ʹ�ô˹��ܣ�")
incept=HTMLEncode(Request("incept"))
UserName=HTMLEncode(Request("UserName"))


select case Request("menu")
case "add"
add
case "bad"
bad
case "Del"
Del
case "Post"
Post
case "look"
look
case "loadLog"
loadLog
case "addPost"
addPost
case ""
index
end select



sub add
if UserName="" then error2("��������Ҫ���ӵĺ������֣�")

if UserName=CookieUserName then error2("���������Լ���")

If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then error2("���ݿⲻ���ڴ��û������ϣ�")

sql="select UserFriend from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Rs.Open sql,Conn,1,3
master=split(Rs("UserFriend"),"|")
if ubound(master) > 20 then error2("ϵͳ�����������20�����ѣ����Ѿ��ﵽ����")
if instr(Rs("UserFriend"),"|"&UserName&"|")>0 then error2("�˺����Ѿ����ӣ�")
Rs("UserFriend")=""&Rs("UserFriend")&""&UserName&"|"
Rs.update
Rs.close
error2("�Ѿ��ɹ����Ӻ���!")
end sub


sub Del
sql="select UserFriend from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Rs.Open sql,Conn,1,3
Rs("UserFriend")=replace(Rs("UserFriend"),"|"&UserName&"|","|")
Rs.update
Rs.close
index
end sub

sub look

Conn.execute("update [BBSXP_Users] set NewMessage=0 where UserName='"&CookieUserName&"'")

Page=Request("Page")
if Page<1 then
disabled="disabled=true"
Page=0
end if
count=Conn.execute("Select count(id)from [BBSXP_Messages] where incept='"&CookieUserName&"'")(0)
sql="select UserName,content from [BBSXP_Messages] where incept='"&CookieUserName&"' order by id Desc"
Set Rs=Conn.Execute(sql)
if Count-Page<2 then disabled2="disabled=true"

if Rs.eof then error2("��û�ж�ѶϢ��")

Rs.Move Page
%>


<title>�鿴��Ϣ - Powered By BBSXP</title>
<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-Left: 1px; BORDER-BOTTOM: 1px;}
</style>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><TD>&nbsp;�ǳƣ�<input readOnly type="text" value="<%=Rs("UserName")%>" size="8"> Email��<input  readOnly type="text" value="<%=Conn.Execute("Select UserMail From [BBSXP_Users] where UserName='"&Rs("UserName")&"'")(0)%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?UserName=<%=Rs("UserName")%>><img border="0" src="<%=Conn.Execute("Select Userface From [BBSXP_Users] where UserName='"&Rs("UserName")&"'")(0)%>" width="32" height="32" alt=�û���ϸ����>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
<script>
document.write ('<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>')
frames.HtmlEditor.document.write ('<%=Rs("content")%>');
frames.HtmlEditor.document.body.style.fontSize="10pt";
frames.HtmlEditor.document.close();
</script>

</TD></TR></TABLE>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="�ظ�ѶϢ" onclick=javascript:open('Friend.asp?menu=Post&incept=<%=Rs("UserName")%>','_top','width=320,height=170')>
</td><TD><input type="button" value="<<��һ��" <%=disabled%> onclick=javascript:open('Friend.asp?menu=look&Page=<%=Page-1%>','_top','')> </td><TD><input type="button" value="��һ��>>" <%=disabled2%> onclick=javascript:open('Friend.asp?menu=look&Page=<%=Page+1%>','_top','')>
</td>
</TR></TABLE>
<%
end sub

sub loadLog
sql="select * from [BBSXP_Messages] where (UserName='"&CookieUserName&"' and incept='"&incept&"') or (UserName='"&incept&"' and incept='"&CookieUserName&"') order by id Desc"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
content=content&"<font color=red>("&Rs("DateCreated")&")����"&Rs("UserName")&"</font><br><font color=0000FF>"&Rs("content")&"</font><br>"
Rs.Movenext
loop
Rs.close
if content="" then content="(��)"
%>
<script>
parent.Log.innerHTML='<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>'
parent.frames.HtmlEditor.document.write ('<%=content%>');
parent.frames.HtmlEditor.document.body.style.fontSize="10pt";
parent.frames.HtmlEditor.document.close();
</script>
<%
end sub

sub Post

if incept="" then error2("�Բ�����û�������û����ƣ�")

sql="select UserName,Userface,UserMail from [BBSXP_Users] where UserName='"&incept&"'"
Set Rs=Conn.Execute(sql)
if Rs.eof then error2("ϵͳ�����ڸ��û�������")


%>

<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-Left: 1px; BORDER-BOTTOM: 1px;}
</style><title>������Ϣ - Powered By BBSXP</title>

<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="Friend.asp" method="POST">
<input type="hidden" name="menu" value="addPost">
<input type="hidden" name="incept" value="<%=Rs("UserName")%>">
<TD>
&nbsp;�ǳƣ�<input readOnly type="text" value="<%=Rs("UserName")%>" size="8"> Email��<input  readOnly type="text" value="<%=Rs("UserMail")%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?UserName=<%=Rs("UserName")%>><img border="0" src="<%=Rs("Userface")%>" width="32" height="32" alt=�û���ϸ����>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey() onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')"><%=HTMLEncode(Request("content"))%></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="�����¼" onclick=lookLog(); name=LookLogButton>
</td><TD><input type="reset" value="ȡ������" OnClick="window.close();"> </td><TD><input type="submit" value="����ѶϢ" name=okButton onclick="return check(this.form)" id='btnadd' disabled></td>
</TR></form>
</TABLE>

<span id=Log></span>
<script>
function lookLog(){
resizeTo(330,330)
document.frames["hiddenframe"].location.replace("?menu=loadLog&incept=<%=incept%>");
document.form.LookLogButton.disabled = true;
}
function check(theForm) {
if (theForm.content.value.lengtd > 200){
alert("�Բ�������ѶϢ���ܳ��� 200 ���ֽڣ�");
return false;
}
}
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}
<%if Request("Log")=1 then%>lookLog()<%end if%>
</script><iframe height=0 width=0 name=hiddenframe></iframe>
<%
Rs.close


end sub

sub addPost
if Len(Request.Form("content"))<2 then error2("���ݲ���С�� 2 �ַ�")
if incept=CookieUserName then error2("���ܸ��Լ�����ѶϢ��")
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&incept&"'" ).eof Then error2("ϵͳ������"&incept&"������")

sql="insert into [BBSXP_Messages](UserName,incept,content) values ('"&CookieUserName&"','"&incept&"','"&HTMLEncode(Request.Form("content"))&"')"
Conn.Execute(SQL)
Conn.execute("update [BBSXP_Users] set NewMessage=NewMessage+1 where UserName='"&incept&"'")
%>
���ͳɹ���<script>close();</script>
<%
end sub


sub index

top


sql="select UserName From [BBSXP_UsersOnline] where Eremite=0"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF
OnlineUserList=OnlineUserList&""&Rs("UserName")&"|"
Rs.MoveNext
loop
Rs.close


%>

<SCRIPT>
var Onlinelist= "<%=Onlinelist%>";

function add(){
var id=prompt("��������Ҫ���ӵĺ������ƣ�","");
if(id){
document.location='Friend.asp?menu=add&UserName='+id+'';
}
}
</SCRIPT>


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
<form method="POST">
<input type=hidden name="menu" value="Del">

<table width="100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
  <tr>
    <td class=a1 align="center">
    �ǳ� </th>
    <td class=a1 align="center">
    �ʼ� </th>
    <td class=a1 align="center">
    ��ҳ </th>
    <td class=a1 align="center">
    ״̬ </th>
    <td class=a1 align="center">
    ������ </th>
    <td class=a1 align="center">
    ���� </th>
  </tr>
<%

on error resume next '�Ҳ�����������ʱ����Դ���

sql="select UserFriend,Userface from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs=Conn.Execute(sql)
master=split(Rs("UserFriend"),"|")
for i = 1 to ubound(master)-1


Set Rs1=Conn.Execute("[BBSXP_Users] where UserName='"&master(i)&"'")
UserMail=Rs1("UserMail")
Userhome=Rs1("Userhome")
Set Rs1 = Nothing


%>
  <tr class=a3>
    <td vAlign=center align=middle><a href=Profile.asp?UserName=<%=master(i)%>><%=master(i)%></a>
    <td align=middle><a href=Mailto:<%=UserMail%>><%=UserMail%></a>
    <td><a href=<%=Userhome%> target=_blank><%=Userhome%></a>
    <td align=middle><%if instr("|"&OnlineUserList&"","|"&master(i)&"|")>0 then%> �� �� <%else%> �� �� <%end if%></td>
    <td align=middle><a style=cursor:hand onclick="javascript:open('Friend.asp?menu=Post&incept=<%=master(i)%>','','width=320,height=170')">����</a></td>
    <td align=middle><INPUT type=radio value=<%=master(i)%> name=UserName></td>
  </tr>
  
  
  

<%
next
Rs.close
%>


  
  

  
  <tr class=a3>
    <td vAlign="center" align="right" colSpan="6">
    <input onclick="javascript:add();" type="button" value="���Ӻ���" name="action">&nbsp;<input onclick="checkclick('ȷ��ɾ��ѡ���ĺ�����?');" type="submit" value="ɾ��"></td>
  </tr></form>
</table>

<%

htmlend
end sub

CloseDatabase
%>