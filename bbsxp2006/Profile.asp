<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
top
UserName=HTMLEncode(Request("UserName"))
sql="select * from [BBSXP_Users] where UserName='"&UserName&"'"
Set Rs1=Conn.Execute(sql)
if Rs1.eof then error("<li>"&UserName&"的用户资料不存在")

ShowRank()
select case Rs1("UserSex")
case "male"
UserSex="男"
case "female"
UserSex="女"
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
<title><%=UserName%>用户资料 - Powered By BBSXP</title>
<script src="inc/birth.js"></script>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 查看用户“<%=UserName%>”资料</td>
</tr>
</table><br>

<table width=100% border="0" cellspacing="0" cellpadding="5" align="center">
<tr>
<td><img src="<%=Rs1("Userface")%>" width="32" height="32">　<b><%=Rs1("UserName")%></b></td>
<td align="right" valign="bottom"><b>:::相关操作:::</b>　｛
<img src="images/finds.gif"> 
<a href=ShowBBS.asp?menu=5&UserName=<%=Rs1("UserName")%>> 个人帖子</a>　

<img src="images/sig.gif">
<a href="Blog.asp?UserName=<%=Rs1("UserName")%>">个人日志</a>　

<img src="images/Friend.gif"> <a href=Friend.asp?menu=add&UserName=<%=Rs1("UserName")%>> 加为好友</a>　
<img src="images/message.gif"> <a href=# onclick="javascript:open('Friend.asp?menu=Post&incept=<%=Rs1("UserName")%>','','width=320,height=170')"> 发送讯息</a> ｝
</tr>
</table>


<table width=100% align=center cellspacing=0 cellpadding=0 border=0>
<tr>
<td class=a2 height="676">
<table width=100% cellspacing=1 cellpadding=6 border=0>
<tr class=a1>
<td colspan="2" height="25" valign="middle" align="Left"><b>&nbsp;:::社区信息:::</b></td>
</tr>
<tr>
<td class=a3 height="5" align="Left" valign="middle" width="50%"><b>　昵　　称：</b><%=Rs1("UserName")%></td>
<td class=a3 height="5" align="Left" valign="middle" width="50%"><b>
　等级名称：</b><%=RankName%></td>
</tr>
<tr>
<td class=a4 align="Left" valign="middle" width="50%"><b>　公　　会：</b><%
if Rs1("Consortia")=empty then
response.write "无"
else
response.write Rs1("Consortia")
end if
%></td>
<td class=a4 align="Left" valign="middle" width="50%"><b>　发表原贴：</b><%=Rs1("PostTopic")%></td>
</tr>
<tr>
<td align="Left" width="50%" class=a3><b>　头　　衔：</b><%=Rs1("UserHonor")%></td>
<td align="Left" width="50%" class=a3>

<b>　发表回贴：</b><%=Rs1("Postrevert")%></td>
</tr>
<tr class=a4>
<td align="Left" width="50%"><b>　配　　偶：</b><%
if Rs1("Consort")=empty then
response.write "无"
else
response.write "<a href=Profile.asp?UserName="&Rs1("Consort")&">"&Rs1("Consort")&"</a>"
end if
%></td>
<td align="Left" width="50%">

<b>　收录精华：</b><%=Rs1("goodTopic")%></td>
</tr>
<tr class=a3>
<td align="Left" width="50%"><b>
　注册日期：</b><%=Rs1("UserRegTime")%></td>
<td align="Left" width="50%">

<b>　被删帖子：</b><%=Rs1("DelTopic")%></td>
</tr>
<tr class=a4>
<td align="Left" width="50%"><b>
　上次登录：</b><%=Rs1("UserLandTime")%></td>
<td align="Left" width="50%"><b>　社区金币：</b><%=Rs1("UserMoney")%></td>
</tr>
<tr class=a3>
<td align="Left" width="50%"><b>
　登录次数：</b><%=Rs1("UserDegree")%></td>
<td align="Left" width="50%"><b>　经 验 值：</b><%=Rs1("experience")%></td>
</tr>

<tr class=a1>
<td height="25" align="Left" valign="middle" colspan="2"><b>&nbsp;:::生活信息:::</b></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"> 　<b>真实姓名：</b><%=realname%>
</td>
<td height="100%" align="Left" valign="bottom" class=a3 rowspan="17" width="50%">
<div style="OVERFLOW:auto; HEIGHT:450px">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">

<tr><td valign=top align=center height="100%"><div align=Left><b> 　用户照片：</b></div>
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
&nbsp;性　　格：</b><br>
　<table border="0">
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
<td class=a4 height="100" valign="top"><b>&nbsp;简　介：</b><br>
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
<td class=a4 height="100" valign="top"><b>&nbsp;签名档：</b><br>
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
<td valign="top" class=a4 width="50%"><b>　性　　别：</b><%=UserSex%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"><b> 　国　　家：</b><%=country%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>省　　份：</b><%=province%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>城　　市：</b><%=city%></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>邮政编码：</b><%=Postcode%></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>生　　肖：</b><script>document.write(getpet("<%=Rs1("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>血　　型：</b><%=blood%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>星　　座：</b><script>document.write(astro("<%=Rs1("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>信　　仰：</b><%=belief%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>职　　业：</b><%=occupation%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>婚姻状况：</b><%=marital%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>最高学历：</b><%=education%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>毕业院校：</b><%=college%>
</td>
</tr>
<tr class=a1>
<td valign="top" width="50%" height="25"><b>&nbsp;:::联系方法:::</b></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">　<b>电子邮件：</b><a href="Mailto:<%=Rs1("UserMail")%>"><%=Rs1("UserMail")%></a></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">　<b>个人主页：</b><a target="_blank" href="<%=Rs1("Userhome")%>"><%=Rs1("Userhome")%></a></td>
</tr>
<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;:::即时通讯:::</b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">　<b>QQ号码：</b><%if qq<>"" then%><a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=qq%>&menu=yes&Site=<%=SiteSettings("SiteName")%>><%=qq%></a><%end if%>
</td>
<td height="4" valign="middle" align="Left">　<b>ICQ：</b><%=icq%></td>
</tr>
<tr class=a4>
<td height="4" valign="middle" align="Left">　<b>UC号码：</b><%=uc%></td>
<td height="4" valign="middle" align="Left">　<b>AIM：</b><%=aim%></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">　<b>MSN IM：</b><%=msn%>　</td>
<td height="4" valign="middle" align="Left">　<b>Yahoo IM：</b><%=Yahoo%></td>
</tr>





</table>

</td>
</tr>
</table>

<center>


<br>
<INPUT onclick=history.back(-1) type=button value=" << 返 回 ">
<br><%
Set Rs1 = Nothing
htmlend
%>