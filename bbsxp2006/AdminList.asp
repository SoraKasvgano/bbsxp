<!-- #include file="Setup.asp" -->
<%
top
%>

<table border=0 align=center cellspacing=1 cellpadding=4 width=100% class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 管理团队</td>
</tr>
</table><br>



<table cellpadding=3 cellspacing=1 border=0 height="25" align=center width=100% class=a2><tr>
<td align=center style="font-size: 9pt" class=a1 height="25" width="30%">
<b>管 理 人 员</b></td>
<td align=center style="font-size: 9pt" class=a1>
<b>详 细 信 息</b></td></tr>
<%
sql="select * from [BBSXP_Users] where membercode > 3 order by membercode Desc,UserLandTime Desc"
Set Rs1=Conn.Execute(sql)

TotalCount=conn.Execute("Select count(ID) From [BBSXP_Users] where membercode > 3")(0) '获取数据数量
PageSetup=10 '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = cint(Request.QueryString("PageIndex")) '获取当前页
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>1 then RS1.Move (PageCount-1) * pagesetup


i=0
Do While Not Rs1.EOF and i<PageSetup
i=i+1
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
ShowRank()
%>
<tr class=a3><td width="30%" align="center" valign="top">
<script>
if("<%=Rs1("Userphoto")%>"!=""){
document.write("<a target=_blank href=<%=Rs1("Userphoto")%>><img src=<%=Rs1("Userphoto")%> border=0 onload='javascript:if(this.width>200)this.width=200'></a>")
}
</script>

</td><td valign=top>
<table cellSpacing=0 cellPadding=5 border=0 style="border-collapse: collapse" style="TABLE-LAYOUT: fixed">
<tr><td align=middle width="50">
<img src=<%=Rs1("Userface")%> width="32" height="32"></td>
<td width="80"><b>昵 称</b></td>
<td align=Left width="120"><a href=Profile.asp?UserName=<%=Rs1("UserName")%>><%=Rs1("UserName")%></a>
<td align=Left width="30">
<img src=images/Friend.gif align=abscenter border=0 width="16" height="16"></td>
<td align=Left width="80"><a href=Friend.asp?menu=add&UserName=<%=Rs1("UserName")%>> 加为好友</a></td>
<td align=Left>
<img src=images/message.gif width="16" height="16"> <a href=# onclick=javascript:open('Friend.asp?menu=Post&incept=<%=Rs1("UserName")%>','','width=320,height=170')> 
发送讯息</a></td>
</tr><tr><td align=middle width="50">
<img src=images/level.gif border=0 width="16" height="16"></td>
<td width="80"><b>用户等级</b></td>
<td align=Left width="120"><%=RankName%></td>
<td align=Left width="30">
<img src=images/Mail.gif border=0 width="16" height="16"></td>
<td align=Left width="80"><b>Email</b></td><td align=Left><a href=Mailto:<%=Rs1("UserMail")%>><%=Rs1("UserMail")%></a>
</tr><tr><td align=middle width="50">
<img src=images/Registered.gif border=0 width="16" height="16"></td>
<td width="80"><b>注册日期</b></td><td align=Left width="120"><%=Rs1("UserRegTime")%>
<td align=Left width="30">
<img src=images/home.gif border=0 width="16" height="16"></td>
<td align=Left width="80"><b>主页地址</b></td>
<td align=Left><a target="_blank" href="<%=Rs1("Userhome")%>"><%=Rs1("Userhome")%></a></tr><tr>
<td align=middle width="50">
<img src=images/Posts.gif border=0 width="16" height="16"></td>
<td width="80"><b>总发贴数</b></td>
<td align=Left width="120"><%=Rs1("PostTopic")+Rs1("Postrevert")%>
<td align=Left width="110" colspan="2">
<b>最后登录时间</b></td>
<td align=Left><%=Rs1("UserLandTime")%>
</td>
</tr><tr>
<td>
</td>
<td colspan="5"><%=personal%>
　</td>
</tr>
</table></td></tr>
<%
Rs1.MoveNext
loop
Rs1.Close
%></table>

<table border=0 width=100% align=center><tr><td>
<%ShowPage()%>
</tr></td></table>
<%
htmlend
%>