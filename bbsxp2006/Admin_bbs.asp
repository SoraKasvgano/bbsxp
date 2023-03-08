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
error2("请选择您要操作的项目")

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
error2("已经将该论坛的所有数据删除了！")



case "Delforumok"
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where lasttime<"&SqlNowString&"-"&TimeLimit&" "&BbsIdList&"")
error2("已经将"&TimeLimit&"天没有更新过的主题删除了！")


case "DelUserTopicok"
if UserName="" then error2("您没有输入用户名！")
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where UserName='"&UserName&"' "&BbsIdList&"")
error2("已经将"&UserName&"发表的主题删除了！")


case "DellikeTopicok"
Topic=HTMLEncode(Request("Topic"))
if Topic="" then error2("您没有输入字符！")
if bbsid<>"" then BbsIdList="and ForumID="&bbsid&""
Conn.execute("Delete from [BBSXP_Threads] where Topic like '%"&Topic&"%' "&BbsIdList&" ")
error2("已经将标题里包含有 "&Topic&" 的帖子全部删除了！")


case "DelForumsok"
Conn.execute("Delete from [BBSXP_Forums] where ForumHide=1 and lasttime<"&SqlNowString&"-"&TimeLimit&"")
error2("已经将"&TimeLimit&"天没有新帖子的论坛删除了！")

case "clean"
Conn.execute("Delete from [BBSXP_Threads] where IsDel=1 and lasttime<"&SqlNowString&"-"&TimeLimit&"")
error2("已经清空 "&TimeLimit&" 天以前的主题！")

case "uniteok"
hbbs=Request("hbbs")
YBBs=Request("YBBs")
if hbbs = YBBs then error2("不能选择相同论坛！")
if UserName<>"" then UserNamelist="and UserName='"&UserName&"'"
Conn.execute("update [BBSXP_Threads] set ForumID="&int(hbbs)&" where ForumID="&int(YBBs)&" and lasttime<"&SqlNowString&"-"&TimeLimit&" "&UserNamelist&"")
error2("移动论坛资料成功！")

case "BatchCensorship"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("update [BBSXP_Threads] set IsDel=0 where id="&ho&"")
next
error2("已经将 激活/还原 所选帖子！")
case "BatchDel"
for each ho in request.form("ThreadID")
ho=int(ho)
Conn.execute("Delete from [BBSXP_Threads] where id="&ho&"")
next
error2("已经将删除所选帖子！")

case "Delapplication"
Application.contents.ReMoveAll()
error2("已经清除服务器上所有的application缓存！")





end select

sub ApplyManage
%>

<table cellspacing=1 cellpadding=2 width=100% border=0 class=a2 align=center>
<tr class=a1 id=TableTitleLink>
<td align="center" height="25"><a href="?menu=ApplyManage&fashion=id">ID</a></td>
<td width="20%" align="center" height="25">
<a href="?menu=ApplyManage&fashion=ForumName">论坛</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumHide">属性</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=moderated">版主</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumToday">今日</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumThreads">主题</a></td>
<td align=center height="25"><a href="?menu=ApplyManage&fashion=ForumPosts">帖子</a></td>
<td align="center" height="25">操作</td>
</tr>
<%

if Request("fashion")=empty then
fashion="ForumPosts"
else
fashion=Request("fashion")
end if


sql="select * from [BBSXP_Forums] order by "&fashion&" Desc"
Rs.Open sql,Conn,1

PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

%>
<tr class=a3>
<td align="center" height="25"><%=Rs("id")%></td>
<td><a target=_blank href=ShowForum.asp?ForumID=<%=Rs("id")%>><%=Rs("ForumName")%></a></td>
<td align=center><%if ""&Rs("ForumHide")&""="1" then%><font color="#FF0000">隐藏</font><%elseif ""&Rs("followid")&""="0" then%><font color="#0000FF">类别</font><%else%>正常<%end if%></td>
<td align=center width="200"><%=Rs("moderated")%></td>
<td align=center><b><font color=red><%=Rs("ForumToday")%></font></b></td>
<td align=center><b><font color=red><%=Rs("ForumThreads")%></font></b></td>
<td align=center><b><font color=red><%=Rs("ForumPosts")%></font></b></td>
<td align="center"><a href=?menu=bbsManagexiu&id=<%=Rs("id")%>>编辑论坛</a> | <a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsManageDel&id=<%=Rs("id")%>>删除论坛</a>
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
类别名称：（例如：电脑网络）<input name="ForumName" onkeyup="ValidateTextboxAdd(this, 'ForumName1')" onpropertychange="ValidateTextboxAdd(this, 'ForumName1')"><input type="submit" value="建立" id='ForumName1' disabled></form>
</td>
		<td align="center">
<form method="POST" action="?menu=bbsManagexiu">
查找论坛：<INPUT size=2 name=id onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="ID" Helptext="ID">
<input type=submit value="确定" id='btnadd' disabled></form>
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
    <td class=a1 align=middle colspan="2">建立论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

<form name="form" method="POST" action="?menu=bbsaddok">
<input type=hidden name=classid value=<%=id%>>

论坛名称</td>
    <td class=a3>

<input size="20" name="ForumName"></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛版主</td>
    <td class=a3>

<input size="30" name=moderated> 多版主添加请用“|”隔开，如：yuzi|裕裕
</td></tr>

<tr class=a3>
<td height="2" align="center" width="20%">帖子专题</td>
<td height="2" align="Left" valign="middle" width="77%">
<input size="30" name="TolSpecialTopic"> 
添加请用“|”隔开，如：原创|转贴</td></tr>

   <tr height=25>
    <td class=a3 align=middle>

论坛介绍</td>
    <td class=a3>

<textarea rows="5" name="ForumIntro" cols="50"></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛规则</td>
    <td class=a3>

<textarea rows="5" name="ForumRules" cols="50"></textarea></td></tr>

   <tr height=25>
    <td class=a3 align=middle>

论坛状态</td>
    <td class=a3>

<select size="1" name="ForumPass">
<option value=0>关闭</option>
<option value=1 selected>正常</option>
<option value=2>游客止步</option>
<option value=3>授权浏览</option>
<option value=4>授权发帖</option>
<option value=5>审查发帖</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

授权用户名单</td>
    <td class=a3>
<input size="30" name="ForumUserList">  请用“|”隔开，如：yuzi|裕裕
</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

小图标URL</td>
    <td class=a3>

<input size="30" name="ForumIcon"> 显示在社区首页论坛介绍右边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

大图标URL</td>
    <td class=a3>

<input size="30" name="ForumLogo"> 显示在论坛左上角
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

通行密码</td>
    <td class=a3>

<input size="30" name="ForumPassword"> 如果是公开论坛，此处请留空</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

是否显示在论坛列表　　　　　　　　　　　
    <td class=a3>

<input type="radio" CHECKED value="0" name="ForumHide">显示 
<input type="radio" value="1" name="ForumHide">隐藏
</tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">

　 <input type="submit" value=" 建 立 "><br></td></tr></table>
</form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>


<%
end sub
sub bbsaddok
if Request("ForumName")="" then error2("请输入论坛名称")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
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

if Request("ForumName")="" then error2("请输入论坛名称")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
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
编辑成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub


sub bbsManagexiu

sql="select * from [BBSXP_Forums] where id="&id&""
Set Rs=Conn.Execute(sql)
if Rs.EOF then error2("系统不存在该论坛的资料")
ForumIntro=replace(""&Rs("ForumIntro")&"","<br>",vbCrlf)
ForumRules=replace(""&Rs("ForumRules")&"","<br>",vbCrlf)
%>


<form method="POST" action="?menu=bbsManagexiuok" name=form><input type=hidden name=id value=<%=Rs("id")%>>
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">编辑论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛名称</td>
    <td class=a3>


<input size="15" name="ForumName" value="<%=Rs("ForumName")%>"> &nbsp; 排序 <input size="2" name="SortNum" value="<%=Rs("SortNum")%>">
从小到大排序</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛类别</td>
    <td class=a3>


<select name="classid">
<option value="<%=Rs("followid")%>">默认</option>
<option value="0">一级分类</option>
<%BBSList(0)%><%=ForumsList%>
</select> 一级分类将当作类别名称</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛版主</td>
    <td class=a3>


<input size="30" name=moderated value="<%=Rs("moderated")%>"> 多版主添加请用“|”隔开，如：yuzi|裕裕
</td></tr>

<tr class=a3>
<td height="2" align="center" width="20%">帖子专题</td>
<td height="2" align="Left" valign="middle" width="77%">
<input size="30" name="TolSpecialTopic" value="<%=Rs("TolSpecialTopic")%>"> 
添加请用“|”隔开，如：原创|转贴</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛介绍</td>
    <td class=a3>


<textarea rows="5" name="ForumIntro" cols="50"><%=ForumIntro%></textarea></td></tr>


   <tr height=25>
    <td class=a3 align=middle>

论坛规则</td>
    <td class=a3>

<textarea rows="5" name="ForumRules" cols="50"><%=ForumRules%></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛状态</td>
    <td class=a3>

<select size="1" name="ForumPass">
<option value=0 <%if Rs("ForumPass")=0 then%>selected<%end if%>>关闭</option>
<option value=1 <%if Rs("ForumPass")=1 then%>selected<%end if%>>正常</option>
<option value=2 <%if Rs("ForumPass")=2 then%>selected<%end if%>>游客止步</option>
<option value=3 <%if Rs("ForumPass")=3 then%>selected<%end if%>>授权浏览</option>
<option value=4 <%if Rs("ForumPass")=4 then%>selected<%end if%>>授权发帖</option>
<option value=5 <%if Rs("ForumPass")=5 then%>selected<%end if%>>审查发帖</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

授权用户名单</td>
    <td class=a3>
<input size="30" name="ForumUserList" value="<%=Rs("ForumUserList")%>"> 添加请用“|”隔开，如：yuzi|裕裕
</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


小图标URL</td>
    <td class=a3>


<input size="30" name="ForumIcon" value="<%=Rs("ForumIcon")%>"> 显示在社区首页论坛介绍右边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


大图标URL</td>
    <td class=a3>


<input size="30" name="ForumLogo" value="<%=Rs("ForumLogo")%>"> 显示在论坛左上角</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

通行密码</td>
    <td class=a3>

<input size="30" name="ForumPassword" value="<%=Rs("ForumPassword")%>"> 如果是公开论坛，此处请留空</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


是否显示在论坛列表</td>
    <td class=a3>


<input type="radio" <%if Rs("ForumHide")=0 then%>CHECKED <%end if%>value="0" name="ForumHide" value="0">显示 
<input type="radio" <%if Rs("ForumHide")=1 then%>CHECKED <%end if%>value="1" name="ForumHide" value="1">隐藏 </td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">


<input type="submit" value=" 编 辑 "></td></tr></table><br></form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>

<%
end sub



sub bbsManage
BBSList(0)
%>


论坛数：<b><font color=red><%=Conn.execute("Select count(id)from [BBSXP_Forums]")(0)%></font></b>　　主题数：<b><font color=red><%=Conn.execute("Select count(id)from [BBSXP_Threads]")(0)%></font></b><br>



<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="3">批量删除主题</td>
  </tr>


<form method="POST" action="?menu=Delforumok">
<tr height=25><td class=a3 align=middle width="50%">
删除 <INPUT size=3 name=TimeLimit value="180">
天没有更新的主题</td>
    <td class=a3 align=middle>
<select name="bbsid">
<option value="">所有论坛</option>
<%=ForumsList%>
</select></td>
    <td class=a3 align=middle>
 <input type="submit" value=" 确 定 "></td></form></tr>

   <tr>
    <td class=a4 align=middle>
    <form method="POST" action="?menu=DelUserTopicok">
删除 <input size="10" name="UserName"> 发表的所有主题
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">所有论坛</option>
<%=ForumsList%>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" 确 定 "></tr>
	</tr></form>

   <tr height=25>
    <td class=a4 align=middle>
    <form method="POST" action="?menu=DellikeTopicok">
删除标题里包含有 <input size="10" name="Topic"> 的所有主题
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">所有论坛</option>
<%=ForumsList%>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" 确 定 "></tr></form>

</table>

<br>




<form method="POST" action="?menu=DelForumsok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>批量删除论坛</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>
删除 <INPUT size=3 name=TimeLimit value="90">
天没有新帖子的论坛

<input type="submit" value=" 确 定 "></td>
    </tr>
    </table></form>



<form method="POST" action="?menu=uniteok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=3>移动论坛帖子</td>
  </tr>
   <tr>
    <td class=a3 align=middle>

将 <select name="YBBs">
<%=ForumsList%>
</select> 移动到
</td>
    <td class=a3 align=middle>

<select name="hbbs">
<%=ForumsList%>
</select>
</td>
    <td class=a3 align=middle rowspan="2">

<INPUT type=submit value=" 确 定 "></td>
	</tr>
   <tr>
    <td class=a3 align=middle width="49%">

仅移动

<input size="2" name="TimeLimit" value="0"> 天前的帖子
</td>
    <td class=a3 align=middle>

仅移动

<input size="8" name="UserName"> 发表的帖子</td>
    </tr>
   </table></form>






<%
end sub

sub upSiteSettings
%>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>社区资料更新</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
此操作将更新论坛资料，修复论坛统计的信息<br>
<a href="?menu=upSiteSettingsok">点击这里更新论坛统计数据</a><br>


<a href="?menu=Delapplication">清除服务器上的application缓存</a><br>



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
操作成功！<br>
<br>
已经更新社区所有论坛的统计数据<br>

<%
end sub

htmlend

sub activation
if Request("type")="Recycle" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime<>lasttime order by lasttime Desc"
response.write "回 收 站"
elseif Request("type")="Censorship" then
sql="select * from [BBSXP_Threads] where IsDel=1 and PostTime=lasttime order by lasttime Desc"
response.write "审 查 区"
end if
%>
<form method="POST" action=?>
<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">主题</td>
<td align="center" width="10%">作者</td>
<td align="center" width="6%">回复</td>
<td align="center" width="6%">点击</td>
<td align="center" width="25%">最后更新</td>
</tr>
<%
Rs.Open sql,Conn,1

PageSetup=20 '设定每页的显示数量
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '总页数
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
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
<input type="checkbox" name="chkall" onclick="ThreadIDCheckAll(this.form)" value="ON">全选
<%
if Request("type")="Recycle" then
%><input type="radio" value="BatchCensorship" name=menu>还原主题<%
elseif Request("type")="Censorship" then
%><input type="radio" value="BatchCensorship" name=menu>通过审查<%
end if
%>
 <input type="radio" value="BatchDel" name=menu>彻底删除	
<input onclick="checkclick('您确定执行本次操作?');" type="submit" value=" 执 行 ">

</td></form>
	<td>
	<form method="POST" action="?menu=clean">清空 <input maxlength="1" size="1" value="7" name="TimeLimit"> 天以前的主题 <input type="submit" onclick="checkclick('执行本操作将清空回收站和审查区的主题?');" value="确定"></form>
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
<td>　<a target=_blank href=ShowForum.asp?ForumID=<%=rs1("id")%>><%=rs1("ForumName")%></a></td>
<td align="right" width="190">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>建立论坛</a> | <a href=?menu=bbsManagexiu&id=<%=rs1("id")%>>编辑论坛</a> | 
<a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsManageDel&id=<%=rs1("id")%>>删除论坛</a>
</tr>

<%
else
%>
<tr class=a3 height=25>
<td>　<%=string(ii*2,"　")%><a target=_blank href=ShowForum.asp?ForumID=<%=rs1("id")%>><%=rs1("ForumName")%></a></td>
<td align="right">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>建立论坛</a> | <a href=?menu=bbsManagexiu&id=<%=rs1("id")%>>编辑论坛</a> | 
<a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsManageDel&id=<%=rs1("id")%>>删除论坛</a>
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