<!-- #include file="Setup.asp" -->

<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

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
<TABLE cellSpacing=0 cellPadding=0 width=100% align=center>

  <TR>
    <TD vAlign=top width="28%">
      <TABLE style="WIDTH: 95%" height="100%" cellSpacing=1 cellPadding=3 border=0 class=a2>
        
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>�û�ͷ��</TH></b></TR>
        <TR align=middle>
          <TD colspan="2" class=a4><img src="<%=Userface%>" onload='javascript:if(this.width>200)this.width=200'></TD></TR>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>��������</TH></b></TR>
<TR class=a3><TD>

       <%
sql="select * from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Set Rs1=Conn.Execute(sql)
ShowRank()
response.write "�û��ȼ��� "&RankName&"<br>�������᣺ "&Rs1("Consortia")&"<br>ͷ�����Σ� "&Rs1("UserHonor")&"<br>�䡡��ż�� "&Rs1("Consort")&"<br>����ԭ���� "&Rs1("PostTopic")&"<br>���������� "&Rs1("Postrevert")&"<br>���������� "&Rs1("goodTopic")&"<br>��ɾ���ӣ� "&Rs1("DelTopic")&"<br>������ң� "&Rs1("UserMoney")&"<br>�û����飺 "&Rs1("experience")&"<br>��¼������ "&Rs1("UserDegree")&"<br>ע��ʱ�䣺 "&Rs1("UserRegTime")&"<br>�ϴε�¼�� "&Rs1("UserLandTime")&""
Rs1.close
%>   
</TD></TR>

</TABLE></TD>
    <TD vAlign=top>
      <TABLE style="WIDTH: 100%" height=29 cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=Left height=25 class=a1>
          <b>-=&gt; ��̳��ѶϢ</b></TD></TR>
        <TR class=a3>
          <TD align="center"><a href="Message.asp"><FONT 
            color=ff0000>�ռ���</FONT></a><font color="FF0000">&nbsp; </font>�� <B>[<%=Conn.execute("Select count(id)from [BBSXP_Messages] where incept='"&SqlString(CookieUserName)&"'")(0)%>]</B> ��ѶϢ��<a href="Message.asp?send=1"><FONT 
            color=ff0000>�ѷ���ѶϢ</FONT></a><B> [<%=Conn.execute("Select count(id)from [BBSXP_Messages] where UserName='"&SqlString(CookieUserName)&"'")(0)%>]</B> ����<BR></TD></TR></TABLE><BR>
      <TABLE  style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
        
        <TR>
          <TD align=Left colSpan=5 height=25 class=a1>
          <b>-=&gt; �����յ��Ķ�Ѷ</TH></b></TR>
        <TR class=a3>
          <TD vAlign=center align=middle width=30>��</TD>
          <TD vAlign=center align=middle width=100><B>������</B></TD>
          <TD vAlign=center align=middle width=300><B>����</B></TD>
          <TD align=middle width=160><B>����</B></TD>
          <TD vAlign=center align=middle width=50><B>��С</B></TD>
         </TR>
<%

sql="select top 5 * from [BBSXP_Messages] where incept='"&SqlString(CookieUserName)&"' order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not Rs.EOF 
size=Len(""&Rs("content")&"")
if size>20 then
content=Left(""&Rs("content")&"",16)&"..."
else
content=Rs("content")
end if
%>
   
        <TR class=a3>
          <TD vAlign=center align=middle><IMG src="images/f_norm.gif"></TD>
          <TD vAlign=center align=middle><A href="Profile.asp?UserName=<%=Rs("UserName")%>" ><%=Rs("UserName")%></A></TD>
          <TD><A href=Message.asp><%=content%></A></TD>
          <TD align="center"><%=Rs("DateCreated")%></FONT></TD>
          <TD align="center"><%=size%></TD>
        </TR>
          
<%          
Rs.MoveNext
loop
Rs.Close      
%>          
          
          
          
          </TD></TR></TABLE><BR>
      <TABLE style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=Left height=25 class=a1>
          <b>-=&gt; 
  �������������</TH></b></TR>
    <%
sql="select top 5 * from [BBSXP_Threads] where IsDel=0 and UserName='"&SqlString(CookieUserName)&"' order by id Desc"
Set Rs=Conn.Execute(sql)

Do While Not Rs.EOF 

%>

  <tr class=a3>
    <td align="Left">&nbsp;����<a href=ShowPost.asp?ThreadID=<%=Rs("id")%>><%=Rs("Topic")%></a> -- <%=Rs("lasttime")%></td>
  </tr>


<%


Rs.MoveNext
loop
Rs.Close

  %>
</table></TD></TR></TABLE>
  
<%
htmlend
%>