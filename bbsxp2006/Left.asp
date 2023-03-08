<!-- #include file="setup.asp" -->

<SCRIPT language=javascript src="inc/Menu.js"></SCRIPT>
<body class="left" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div style=width:100%;height:100%;OVERFLOW:auto>

<embed width=140  height=60  src='images/bbsxplogo.swf' wmode=transparent>
<SCRIPT language=javascript>
var yuzimenu=new menutree('yuzimenu','main')
yuzimenu.addNode('<%=SiteSettings("SiteName")%>','Default.asp')


<%
dim LeftMenuList
TreeList(0)
Response.Write LeftMenuList
%>
yuzimenu.addFolder('我的工具箱')
	yuzimenu.addFolder('控制面板','usercp.asp',false)
		yuzimenu.addNode('资料修改','EditProfile.asp')
		yuzimenu.addNode('密码修改','EditProfile.asp?menu=pass')
		yuzimenu.addNode('个性设置','MySettings.asp')
		yuzimenu.addNode('附件管理','MyAttachment.asp')
		yuzimenu.addNode('短信服务','message.asp')
		yuzimenu.addNode('好友列表','Friend.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('个人服务')
		yuzimenu.addNode('我的资料','Profile.asp?UserName=<%=CookieUserName%>')
		yuzimenu.addNode('上传头像','upface.asp')
		yuzimenu.addNode('上传照片','upphoto.asp')
		yuzimenu.addNode('社区日志','calendar.asp')
		yuzimenu.addNode('我的日志','blog.asp?UserName=<%=CookieUserName%>')
		yuzimenu.addNode('更改用户','login.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('收 藏 夹')
		yuzimenu.addNode('帖子收藏夹','MyFavorites.asp?menu=Topic')
		yuzimenu.addNode('论坛收藏夹','MyFavorites.asp?menu=Forum')
		yuzimenu.addNode('网站收藏夹','MyFavorites.asp')
	yuzimenu.endFolder()
yuzimenu.endFolder()

<%clubmenu()%>
	yuzimenu.addFolder('帖子状态')
		yuzimenu.addNode('最新帖子','ShowBBS.asp')
		yuzimenu.addNode('人气帖子','ShowBBS.asp?menu=1')
		yuzimenu.addNode('热门帖子','ShowBBS.asp?menu=2')
		yuzimenu.addNode('精华帖子','ShowBBS.asp?menu=3')
		yuzimenu.addNode('投票帖子','ShowBBS.asp?menu=4')
		yuzimenu.addNode('我的帖子','ShowBBS.asp?menu=5&UserName=<%=CookieUserName%>')
	yuzimenu.endFolder()
	yuzimenu.addFolder('论坛状态')
		yuzimenu.addNode('在线情况','online.asp')
		yuzimenu.addNode('在线图例','online.asp?menu=cutline')
		yuzimenu.addNode('性别图例','online.asp?menu=UserSex')
		yuzimenu.addNode('今日图例','online.asp?menu=TodayPage')
		yuzimenu.addNode('主题图例','online.asp?menu=board')
		yuzimenu.addNode('帖子图例','online.asp?menu=ForumPosts')
		yuzimenu.addNode('会员列表','usertop.asp')
		yuzimenu.addNode('管理团队','adminlist.asp')
		yuzimenu.addNode('申请论坛','ApplyForum.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('个人状态')
		yuzimenu.addNode('在线','cookies.asp?menu=online')
		yuzimenu.addNode('隐身','cookies.asp?menu=eremite')
	yuzimenu.endFolder()

<%if CookieUserName=empty then%>
yuzimenu.addNode('登录论坛','login.asp')
<%else%>
yuzimenu.addNode('退出社区','login.asp?menu=out')
<%end if%>

yuzimenu.init()
top.document.title="<%=SiteSettings("SiteName")%> - Powered By BBSXP"; 
</SCRIPT>
</div>
</body>


<%



sub clubmenu()

Set Rs = Conn.ExeCute("Select * From [BBSXP_Menu] Where followid=0")
Do While Not Rs.Eof
	Response.Write vbTab & "yuzimenu.addFolder('"&Rs("name")&"')" & vbCrlf
	Set Rs1 = Conn.ExeCute("Select * From [BBSXP_Menu] Where followid="&Rs("ID"))
	Do While Not Rs1.Eof
		Response.Write String(2,vbTab) & "yuzimenu.addNode('"&Rs1("name")&"','"&Rs1("url")&"')" & vbCrlf
		Rs1.MoveNext()
	Loop
	Response.Write vbTab & "yuzimenu.endFolder()" & vbCrlf
	Rs1.Close
	Rs.MoveNext()
Loop
Rs.Close
Set Rs = Nothing
end sub



sub TreeList(selec)
sql="Select * From [BBSXP_Forums] where followid="&selec&" and ForumHide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
	ForumName=rs1("ForumName")
	bEof = conn.execute("Select * From [BBSXP_Forums] Where followid="&rs1("id")&" and ForumHide=0").Eof
	if bEof Then
LeftMenuList=LeftMenuList&vbCrlf & string(ii,vbTab) & "yuzimenu.addNode('"&ForumName&"','ShowForum.asp?ForumID="&rs1("id")&"')"
	else
LeftMenuList=LeftMenuList&vbCrlf & string(ii,vbTab) & "yuzimenu.addFolder('"&ForumName&"','ShowForum.asp?ForumID="&rs1("id")&"',false)"
	end if
	ii=ii+1
	TreeList rs1("id")
	ii=ii-1
if Not bEof Then LeftMenuList=LeftMenuList&vbCrlf & string(ii,vbTab) & "yuzimenu.endFolder()"
rs1.movenext
loop
Set Rs1 = Nothing
end sub


CloseDatabase
%>