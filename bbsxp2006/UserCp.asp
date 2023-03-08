<!-- #include file="Setup.asp" -->

<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 控制面板</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">控制面板</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">密码修改</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">个性设置</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">附件管理</a></b></td>
      <TD align="center"><b><a href="Message.asp">短信服务</a></b></td>
      <TD align="center"><b><a href="Friend.asp">好友列表</a></b></td>
      </TR></TABLE>
<br>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=center>

  <TR>
    <TD vAlign=top width="28%">
      <TABLE style="WIDTH: 95%" height="100%" cellSpacing=1 cellPadding=3 border=0 class=a2>
        
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>用户头像</TH></b></TR>
        <TR align=middle>
          <TD colspan="2" class=a4><img src="<%=Userface%>" onload='javascript:if(this.width>200)this.width=200'></TD></TR>
        <TR>
          <TD height=25 class=a1 align="center" colspan="2">
          <b>基本资料</TH></b></TR>
<TR class=a3><TD>

       <%
sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
Set Rs1=Conn.Execute(sql)
ShowRank()
response.write "用户等级： "&RankName&"<br>公　　会： "&Rs1("Consortia")&"<br>头　　衔： "&Rs1("UserHonor")&"<br>配　　偶： "&Rs1("Consort")&"<br>发表原贴： "&Rs1("PostTopic")&"<br>发表回贴： "&Rs1("Postrevert")&"<br>精华帖数： "&Rs1("goodTopic")&"<br>被删帖子： "&Rs1("DelTopic")&"<br>社区金币： "&Rs1("UserMoney")&"<br>用户经验： "&Rs1("experience")&"<br>登录次数： "&Rs1("UserDegree")&"<br>注册时间： "&Rs1("UserRegTime")&"<br>上次登录： "&Rs1("UserLandTime")&""
Rs1.close
%>   
</TD></TR>

</TABLE></TD>
    <TD vAlign=top>
      <TABLE style="WIDTH: 100%" height=29 cellSpacing=1 cellPadding=3 align=center border=0 class=a2>

        <TR>
          <TD align=Left height=25 class=a1>
          <b>-=&gt; 论坛短讯息</b></TD></TR>
        <TR class=a3>
          <TD align="center"><a href="Message.asp"><FONT 
            color=ff0000>收件箱</FONT></a><font color="FF0000">&nbsp; </font>共 <B>[<%=Conn.execute("Select count(id)from [BBSXP_Messages] where incept='"&CookieUserName&"'")(0)%>]</B> 条讯息，<a href="Message.asp?send=1"><FONT 
            color=ff0000>已发送讯息</FONT></a><B> [<%=Conn.execute("Select count(id)from [BBSXP_Messages] where UserName='"&CookieUserName&"'")(0)%>]</B> 条。<BR></TD></TR></TABLE><BR>
      <TABLE  style="WIDTH: 100%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
        
        <TR>
          <TD align=Left colSpan=5 height=25 class=a1>
          <b>-=&gt; 最新收到的短讯</TH></b></TR>
        <TR class=a3>
          <TD vAlign=center align=middle width=30>　</TD>
          <TD vAlign=center align=middle width=100><B>发件人</B></TD>
          <TD vAlign=center align=middle width=300><B>内容</B></TD>
          <TD align=middle width=160><B>日期</B></TD>
          <TD vAlign=center align=middle width=50><B>大小</B></TD>
         </TR>
<%

sql="select top 5 * from [BBSXP_Messages] where incept='"&CookieUserName&"' order by id Desc"
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
  最近发表的主题</TH></b></TR>
    <%
sql="select top 5 * from [BBSXP_Threads] where IsDel=0 and UserName='"&CookieUserName&"' order by id Desc"
Set Rs=Conn.Execute(sql)

Do While Not Rs.EOF 

%>

  <tr class=a3>
    <td align="Left">&nbsp;□　<a href=ShowPost.asp?ThreadID=<%=Rs("id")%>><%=Rs("Topic")%></a> -- <%=Rs("lasttime")%></td>
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