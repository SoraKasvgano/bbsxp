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
yuzimenu.addFolder('�ҵĹ�����')
	yuzimenu.addFolder('�������','usercp.asp',false)
		yuzimenu.addNode('�����޸�','EditProfile.asp')
		yuzimenu.addNode('�����޸�','EditProfile.asp?menu=pass')
		yuzimenu.addNode('��������','MySettings.asp')
		yuzimenu.addNode('��������','MyAttachment.asp')
		yuzimenu.addNode('���ŷ���','message.asp')
		yuzimenu.addNode('�����б�','Friend.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('���˷���')
		yuzimenu.addNode('�ҵ�����','Profile.asp?UserName=<%=CookieUserName%>')
		yuzimenu.addNode('�ϴ�ͷ��','upface.asp')
		yuzimenu.addNode('�ϴ���Ƭ','upphoto.asp')
		yuzimenu.addNode('������־','calendar.asp')
		yuzimenu.addNode('�ҵ���־','blog.asp?UserName=<%=CookieUserName%>')
		yuzimenu.addNode('�����û�','login.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('�� �� ��')
		yuzimenu.addNode('�����ղؼ�','MyFavorites.asp?menu=Topic')
		yuzimenu.addNode('��̳�ղؼ�','MyFavorites.asp?menu=Forum')
		yuzimenu.addNode('��վ�ղؼ�','MyFavorites.asp')
	yuzimenu.endFolder()
yuzimenu.endFolder()

<%clubmenu()%>
	yuzimenu.addFolder('����״̬')
		yuzimenu.addNode('��������','ShowBBS.asp')
		yuzimenu.addNode('��������','ShowBBS.asp?menu=1')
		yuzimenu.addNode('��������','ShowBBS.asp?menu=2')
		yuzimenu.addNode('��������','ShowBBS.asp?menu=3')
		yuzimenu.addNode('ͶƱ����','ShowBBS.asp?menu=4')
		yuzimenu.addNode('�ҵ�����','ShowBBS.asp?menu=5&UserName=<%=CookieUserName%>')
	yuzimenu.endFolder()
	yuzimenu.addFolder('��̳״̬')
		yuzimenu.addNode('�������','online.asp')
		yuzimenu.addNode('����ͼ��','online.asp?menu=cutline')
		yuzimenu.addNode('�Ա�ͼ��','online.asp?menu=UserSex')
		yuzimenu.addNode('����ͼ��','online.asp?menu=TodayPage')
		yuzimenu.addNode('����ͼ��','online.asp?menu=board')
		yuzimenu.addNode('����ͼ��','online.asp?menu=ForumPosts')
		yuzimenu.addNode('��Ա�б�','usertop.asp')
		yuzimenu.addNode('�����Ŷ�','adminlist.asp')
		yuzimenu.addNode('������̳','ApplyForum.asp')
	yuzimenu.endFolder()
	yuzimenu.addFolder('����״̬')
		yuzimenu.addNode('����','cookies.asp?menu=online')
		yuzimenu.addNode('����','cookies.asp?menu=eremite')
	yuzimenu.endFolder()

<%if CookieUserName=empty then%>
yuzimenu.addNode('��¼��̳','login.asp')
<%else%>
yuzimenu.addNode('�˳�����','login.asp?menu=out')
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