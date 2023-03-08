<!-- #include file="setup.asp" --><%top

Subject=HTMLEncode(Request.Form("Subject"))
content=ContentEncode(Request.Form("content"))
lookdate=HTMLEncode(Request("lookdate"))
adddate=HTMLEncode(Request("adddate"))
hide=int(Request("hide"))
id=int(Request("id"))


if Request("menu")="add" then
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
if Subject=empty then error("<li>您没有输入日志主题")
if content=empty then error("<li>您没有输入日志内容")
if adddate="" then adddate=""&year(now)&"-"&month(now)&"-"&day(now)&""



if id=0 then
sql="insert into [BBSXP_Calendar](title,UserName,content,hide,adddate) values ('"&Subject&"','"&CookieUserName&"','"&content&"','"&hide&"','"&adddate&"')"
conn.Execute(SQL)
else
conn.execute("update [BBSXP_Calendar] set title='"&Subject&"',content='"&content&"',hide='"&hide&"' where id="&id&" and UserName='"&CookieUserName&"'")
end if


message="<li>发表日志成功<li><a href=blog.asp?UserName="&CookieUserName&">返回个人日志</a><li><a href=calendar.asp?menu=show&lookdate="&adddate&">返回"&adddate&"日志</a><li><a href=calendar.asp>返回日历</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=blog.asp?UserName="&CookieUserName&">")

elseif Request("menu")="del" then
if membercode > 3 then
conn.execute("delete from [BBSXP_Calendar] where id="&id&"")
else
conn.execute("delete from [BBSXP_Calendar] where id="&id&" and UserName='"&CookieUserName&"'")
end if
error2("删除成功！")

end if

%>
<table class="a2" cellSpacing="1" cellPadding="4" width="100%" align="center" border="0">
	<tr class="a3">
		<td colSpan="2">
		<table cellSpacing="0" cellPadding="0" width="100%" border="0" id="table2">
			<tr>
				<td height="18">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> → <a href="calendar.asp">社区日志</a></td>
				<td align="right" height="18"> <img src="images/jt.gif"> <a href=blog.asp?UserName=<%=CookieUserName%>>我的日志</a> 
				<img src="images/jt.gif"> <a href=calendar.asp?menu=NewCalendar&lookdate=<%=lookdate%>>发表日志</a></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<br>
<%

select case Request("menu")
case ""
Default
case "show"
show
case "NewCalendar"
NewCalendar

end select

sub NewCalendar
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
if Request("id")<>empty then
sql="select * from [BBSXP_Calendar] where id="&id&""
Set Rs=Conn.Execute(sql)
UserName=rs("UserName")
content=rs("content")
title=rs("title")
adddate=rs("adddate")
hide=rs("hide")
Set Rs = Nothing
if UserName<>""&CookieUserName&"" then error2("只有原作者才能编辑该日志！")
end if


%>
	
	<table cellspacing="1" cellpadding="4" width="100%" align="center" border="0" class="a2">
		<form method="POST" name="yuziform" action="calendar.asp?menu=add" onsubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=content%>'>
			<input type="hidden" name="id" value="<%=id%>">
			<input type="hidden" name="adddate" value="<%=adddate%>">
			<tr class="a1">
				<td width="100%" colspan="2" align="center"><b>发表日志</b></td>
			</tr>
			<tr class="a3">
				<td width="14%"><b>日志标题：</b></td>
				<td width="83%">
				<input maxlength="30" name="Subject" style="width:80%" value="<%=title%>"> <select name="hide" size="1">
				<option value="0" <%if hide=0 then%>selected<%end if%>>公开</option>
				<option value="1" <%if hide=1 then%>selected<%end if%>>隐藏</option>
				</select></td>
			</tr>
			<tr class="a3">
				<td width="14%"><b>日志内容：</b></td>
				<td width="83%" height="200">
				<script src="inc/post.js"></script>
				</td>
			</tr>
			<tr class="a3">
				<td colspan="2" align="center">　<input type="submit" value=" 发 送 " name=EditSubmit>&nbsp;
				<input type="reset" value=" 重 置 "></td>
			</tr>
		</form>
	</table>
	
<%
end sub


sub show
%> <br>

<div align="center">
	<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
		<tr class="a1">
			<td align="center" width="75%"><b><%=lookdate%> 社区日志</b></td>
		</tr>
		<tr class="a3">
			<td align="center" valign="top">
			<%
			
if lookdate<>"" then lookdateSQL="and adddate='"&lookdate&"'"
sql="select * from [BBSXP_Calendar] where (hide=0 or UserName='"&CookieUserName&"') "&lookdateSQL&" order by id Desc"

rs.Open sql,Conn,1
pagesetup=10 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
UserName=rs("UserName")
content=ReplaceText(rs("content"),"<[^>]*>","")
if len(content)>200 then content=left(""&content&"",200)&"..."


%>
				<table border="0" width="100%" cellspacing="10">
					<tr>
						<td>
						<b><font size="4"><%=rs("title")%></font></b><br>
						<br>
<%=content%>
<%if rs("hide")=1 then%><br><br>注：<font color="#FF0000">本篇日志为隐藏状态</font><%end if%>
<hr>
<a href="blog.asp?id=<%=rs("id")%>">阅读全文</a>
 | 
<%if CookieUserName=UserName then%>
<a href="calendar.asp?menu=NewCalendar&id=<%=rs("id")%>">编辑</a>
<%else%>
<font color="#C0C0C0">编辑</font>
<%end if%>
 | 
<%if UserName=CookieUserName or membercode > 3 then%>
<a href="calendar.asp?menu=del&id=<%=rs("id")%>" onclick="checkclick('您确定要删除此条日志?')">删除</a>
<%else%>
<font color="#C0C0C0">删除</font>
<%end if%>
 | <%=rs("DateCreated")%> by <a href="Profile.asp?UserName=<%=UserName%>"><%=UserName%></a></td>
					</tr>
				</table>
<%
RS.MoveNext
loop

RS.Close  
%>
</td>
		</tr>
	</table>
	
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td>
</tr></table>

	<%
end sub


sub Default
%>
	<script src="inc/calendar.js"></script>
	<script language="VBScript">
<!--
'===== 算世界时间
Function TimeAdd(UTC,T)
   Dim PlusMinus, DST, y
   If Left(T,1)="-" Then PlusMinus = -1 Else PlusMinus = 1
   UTC=Right(UTC,Len(UTC)-5)
   UTC=Left(UTC,Len(UTC)-4)
   y = Year(UTC)
   TimeAdd=DateAdd("n", (Cint(Mid(T,2,2))*60 + Cint(Mid(T,4,2))) * PlusMinus, UTC)
   '美国日光节约期间: 4月第一个星日00:00 至 10月最後一个星期日00:00
   If Mid(T,6,1)="*" And DateSerial(y,4,(9 - Weekday(DateSerial(y,4,1)) mod 7) ) <= TimeAdd And DateSerial(y,10,31 - Weekday(DateSerial(y,10,31))) >= TimeAdd Then
      TimeAdd=CStr(DateAdd("h", 1, TimeAdd))
      tSave.innerHTML = "R"
   Else
      tSave.innerHTML = ""
   End If
   TimeAdd = CStr(TimeAdd)
End Function
'-->
</script>
	<body onload="initial()">
<div id="detail" style="LEFT: 12px; WIDTH: 200px; POSITION: absolute; TOP: 0px; HEIGHT: 19px"></div>
	<form name="CLD">
		
		<table cellspacing="1" cellpadding="0" width="100%" border="0" class="a2" align="center">
			<tr>
				<td align="middle" width="30%" class="a4">
				<script language="JavaScript">
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "星期日"
if(today.getDay()==1) day = "星期一"
if(today.getDay()==2) day = "星期二"
if(today.getDay()==3) day = "星期三"
if(today.getDay()==4) day = "星期四"
if(today.getDay()==5) day = "星期五"
if(today.getDay()==6) day = "星期六"
document.fgColor = "black";
date = " 公元 " + (today.getYear()) + " 年 " +
(today.getMonth() + 1 ) + "月 " + today.getDate() + "日 " +
day +"";
document.write(date)
    </script>
				</font></td>
				<td width="400" class="a3" rowspan="2" align="center" valign="top">
				<table border="0">
					<tr>
						<td class="a1" colspan="7" align="center" height="25">公元<select style="FONT-SIZE: 9pt" onchange="changeCld()" name="SY">
						<script language="JavaScript"><!--        
            for(i=1900;i<2050;i++) document.write('<option>'+i)        
            //--></script>
						</select>年<select style="FONT-SIZE: 9pt" onchange="changeCld()" name="SM">
						<script language="JavaScript"><!--        
            for(i=1;i<13;i++) document.write('<option>'+i)        
            //--></script>
						</select>月</font>&nbsp;&nbsp; <font id="GZ"></font><br>
						</td>
					</tr>
					<tr align="middle" class="a4">
						<td width="54" height="20">日</td>
						<td width="54" height="20">一</td>
						<td width="54" height="20">二</td>
						<td width="50" height="20">三</td>
						<td width="54" height="20">四</td>
						<td width="54" height="20">五</td>
						<td width="54" height="20">六</td>
					</tr>
					<script language="JavaScript">    
            var gNum        
            for(i=0;i<6;i++) {        
               document.write('<tr align=center>')        
               for(j=0;j<7;j++) {        
                  gNum = i*7+j        
                  document.write('<td ><font id="SD' + gNum +'" size=5 face="Arial Black"></font><br><font id="LD' + gNum + '"></font></td>')
               }        
               document.write('</tr>')        
            }        
</script>
				</table>
				</td>
				<td align="middle" class="a4" height="25">公开日志</td>
			</tr>
			<tr>
				<td valign="top" align="middle" width="30%" class="a3"><br>
				<font id="Clock" face="Arial" color="000080" size="4" align="center">
				</font>
				<p>
				<!--时区 *表示自动调整为日光节约时间--><font style="FONT-SIZE: 9pt" size="2">
				<select style="FONT-SIZE: 9pt" onchange="changeTZ()" name="TZ">
				<option value="+0800 北京、香港" selected>北京</option>
				<option value="-1200 安尼威土克、瓜甲兰">国际换日线</option>
				<option value="-1100 萨摩亚">萨摩亚</option>
				<option value="-1000 檀香山">檀香山</option>
				<option value="-0900 安克雷奇">安克雷奇</option>
				<option value="-0800 洛杉矶">洛杉矶</option>
				<option value="-0700 丹佛">丹佛</option>
				<option value="-0600 芝加哥">芝加哥</option>
				<option value="-0500 纽约">纽约</option>
				<option value="-0400 加拉加斯">加拉加斯</option>
				<option value="-0300 里约热内卢">里约热内卢</option>
				<option value="-0200 大西洋中部">大西洋中部</option>
				<option value="-0100 亚速尔群岛、维德角群岛">亚速尔</option>
				<option value="+0000 格林威治时">格林威治时</option>
				<option value="+0000 伦敦">伦敦</option>
				<option value="+0100 巴黎">巴黎</option>
				<option value="+0200 开罗">开罗</option>
				<option value="+0300 莫斯科">莫斯科</option>
				<option value="+0400 迪拜">迪拜</option>
				<option value="+0500 卡拉奇">卡拉奇</option>
				<option value="+0530 德里">德里</option>
				<option value="+0600 达卡">达卡</option>
				<option value="+0700 曼谷">曼谷</option>
				<option value="+0900 汉城">汉城</option>
				<option value="+1000 悉尼">悉尼</option>
				<option value="+1100 努美阿">努美阿</option>
				<option value="+1200 威灵顿">威灵顿</option>
				</select> 时间</font><font id="tSave" style="FONT-SIZE: 18pt; COLOR: red; FONT-FAMILY: Wingdings">
				</font><br>
				<br>
				<font style="FONT-SIZE: 120pt; COLOR: #13b0f4; FONT-FAMILY: Webdings">
				&ucirc;</font><br>
				<br>
				<font id="CITY"></font></p>
				</td>
				<td class="a3" height="30%" align="center" valign="top">
				<table border="0" width="95%">
				

<tr><td>
<%
sql="select top 20 * from [BBSXP_Calendar] where hide=0 order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
%><li><a href="blog.asp?id=<%=rs("id")%>"><%=rs("title")%></a></li><%
RS.MoveNext
loop
RS.Close   
%></td></tr>
				</table>
				</td>
			</tr>
		</table>
		
	</form>
	</div>
<%
end sub

htmlend
%>