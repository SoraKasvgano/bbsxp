<!-- #include file="Setup.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


dir=Request("dir")

response.write "<center>"
select case Request("menu")

case "Posts"
Posts

case "PostAttachment"
PostAttachment

case "DelPostAttachment"
for each ho in Request.form("ID")
ho=int(ho)
Conn.execute("Delete from [BBSXP_PostAttachments] where id="&ho&"")
next
error2("删除成功")


case "changeobjectowner"
SQL="sp_changeobjectowner '[BBSXP_Posts"&Request("tablename")&"]','dbo'"
Conn.execute(SQL)
error2("已经成功将该表的所有者更改为dbo")

case "CreatTable"

'建立数据库
Sql="CREATE TABLE [BBSXP_Posts"&Request("tablename")&"] ("&_
	"ID int IDENTITY (1, 1) NOT NULL ,"&_
	"ThreadID int NOT NULL ,"&_
	"IsTopic int Default 0 NOT NULL ,"&_
	"Subject varchar(255) NOT NULL ,"&_
	"UserName varchar(255) NOT NULL ,"&_
	"Content text NOT NULL ,"&_
	"Posttime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"PostIP varchar(255) NOT NULL "&_
")"
Conn.execute(sql)

'建立关联
Sql="ALTER TABLE [BBSXP_Posts"&Request("tablename")&"] ADD CONSTRAINT [FK_BBSXP_Posts"&Request("tablename")&"_BBSXP_Threads] FOREIGN KEY ([ThreadID]) REFERENCES [BBSXP_Threads] ([Id]) ON DELETE CASCADE  ON UPDATE CASCADE"
Conn.execute(sql)

Conn.execute("update [BBSXP_SiteSettings] set DefaultPostsName="&Request("tablename"))

error2("建立成功！")


case "update"
if Request("tablename") = "" Then ThisValue = "Null" Else ThisValue = Request("tablename")
Conn.execute("update [BBSXP_SiteSettings] set DefaultPostsName="&ThisValue)
error2("设置成功！") 



case "Del"
if request("tablename")="" then error2("系统默认的表不能删除！")
If not Conn.Execute("Select PostsTableName From [BBSXP_Threads] where PostsTableName='"&request("tablename")&"'" ).eof Then error2("该表中有对应的主题，不能删除！")
if ""&SiteSettings("DefaultPostsName")&""=""&request("tablename")&"" then error2("当前正在使用中的数据表不能删除！")
Conn.Execute("drop table [BBSXP_Posts"&request("tablename")&"]")
error2("删除成功！")


case "files"
files

case "Delfiles"

if dir<>"" then dir=""&dir&"/"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
for each ho in request.form("files")
MyFileObject.DeleteFile""&Server.MapPath("./UpFile/"&dir&""&ho&"")&""
next
error2("已经成功删除所选的文件！")


case "bak"
bak

case "bakbf"

set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(DB)&"",""&Server.MapPath(Request.Form("BakDbPath"))&""
error2("备份成功！")

case "bakhf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request.Form("BakDbPath"))&"",""&Server.MapPath(DB)&""
error2("恢复成功！")


case "statroom"
statroom

end select

sub bak


If IsSqlDataBase<>0 Then error2("SQL版本无法进行ACCESS数据库管理")

%>


<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>备份数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
<form method="POST" action="?menu=bakbf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%" height="25">数据库路径： </td>
<td width="70%"><%=Db%></td>
</tr>
<tr>
<td width="30%" height="25">备份的数据库路径：</td>
<td width="70%"><input size="30" value="<%=replace(""&Db&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 备 份 "><br></td>
</tr>
</table>
</td></tr></table></form>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>恢复数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<form method="POST" action="?menu=bakhf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%" height="25">备份的数据库路径： </td>
<td width="70%">
<input size="30" value="<%=replace(""&Db&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath"></td>
</tr>
<tr>
<td width="30%" height="25">数据库路径：</td>
<td width="70%"><%=Db%></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 恢 复 "><br></td>
</tr>
</table></td></tr></table>
</form>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>压缩数据库</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>

<form action=Compact.asp method=Post>
<input type=hidden name=sessionid value=<%=session.sessionid%>>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="70%">


<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>

<td width="30%">数据库路径： </td>
<td width="70%">
<input size="30" value="<%=Db%>" name="dbpath"></td>
</tr>
<tr>
<td width="30%">数据库格式：</td>
<td width="70%"><input type="radio" value="True" name="boolIs97" id=boolIs97><label for=boolIs97>Access 97</label>　<input type="radio" value="" name="boolIs97" checked id=boolIs97_1><label for=boolIs97_1>Access 2000、2002、2003</label></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" 压 缩 "></td>
</tr>



</table>
</td>
</table>
</td></tr></table>
</form>

<%
end sub



sub statroom
set fso=server.createobject("Scripting.FileSystemObject")
UpFacedir=server.mappath("./UpFile/UpFace")
set d=fso.getfolder(UpFacedir)
UpFacesize=d.size

UpPhotodir=server.mappath("./UpFile/UpPhoto")
set d=fso.getfolder(UpPhotodir)
UpPhotosize=d.size

UpFiledir=server.mappath("./UpFile/UpAttachment")
set d=fso.getfolder(UpFiledir)
UpFilesize=d.size




toldir=server.mappath(".")
set d=fso.getfolder(toldir)
tolsize=d.size

totalBytes=UpFacesize+UpPhotosize+UpFilesize+tolsize



%>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center" height="130">
  <tr height=25>
    <td class=a1 align=middle colspan=3>统计占用空间</td>
  </tr>
   <tr height=25>
    <td class=a3 width="29%">
    

上传头像占用空间<br>




</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/0.gif" width=<%=Int(UpFacesize/totalBytes*100)%>%> </td>
    <td class=a3><%=CheckSize(UpFacesize)%></td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

上传照片占用空间</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/1.gif" width=<%=Int(UpPhotosize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(UpPhotosize)%>

　</td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

上传附件占用空间</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/2.gif" width=<%=Int(UpFilesize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(UpFilesize)%>

　</td></tr>



   <tr height=25>
    <td class=a3 width="29%" height="24">
    

BBSXP目录总共占用空间</td>
    <td class=a3 width="60%" height="24">
    

<IMG height=10 src="images/bar/4.gif" width=<%=Int(tolsize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(tolsize)%>

　</td></tr></table>


<%
end sub


sub files
thisdir=server.mappath("./UpFile/"&dir&"")
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fdir=fs.GetFolder(thisdir)

count=fdir.Files.count

%>

<a href="?menu=files&dir=<%=dir%>"><%=dir%>目录下</a>共 <font color="FF0000"><b><%=count%></b></font> 个文件

<table border="0" width="100%"><tr><td valign=top>
<img src=images/folder.gif> <a href="?menu=files">..</a><br>

<%

set gFolder=fs.GetFolder(thisdir)

for each theFolder in gFolder.SubFolders
%>
<img src=images/folder.gif> <a href="?menu=files&dir=<%=theFolder.Name%>"><%=theFolder.Name%></a><br>
<%
next

%>

</td>
<td valign=top align="center">
<%

PageSetup=10 '设定每页的显示数量
If Count/PageSetup > (Count\PageSetup) then
TotalPage=(Count\PageSetup)+1
else TotalPage=(Count\PageSetup)
End If
if Request.QueryString("PageIndex")<>"" then PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


%>
<form method=Post name=form action="?menu=Delfiles">
<input type=hidden name=dir value=<%=dir%>>

<table cellspacing="1" cellpadding="2" border="0" width=88% class="a2">
<tr class=a1 ><td width=50 align="center">

<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
	<td width=208 align=center>名称</td>
  <td width=128 align=center>
大小</td><td width=140 align="center">类型</td><td align="center">
修改时间</td></tr>
<%
Pagesize=20
Page=request.querystring("Page")
if Page="" or not isnumeric(Page) then
	Page=1
else
	Page=int(Page)
end if
filenum=fdir.Files.count
Pagenum=int(filenum/Pagesize)
if filenum mod Pagesize>0 then
	Pagenum=Pagenum+1
end if
if Page> Pagenum then
	Page=1
end if
i=0


genre=Request("genre")

For each thing in fdir.Files


if instr(thing.name,genre) > 0 then

	i=i+1
	if i>(Page-1)*Pagesize and i<=Page*Pagesize then

response.write "<tr class=a3><td align=center><input type='CheckBox' value='"&thing.name&"' name=files></td><td align=center><a target=_blank href='UpFile/"&dir&"/"&thing.Name&"'>"&thing.Name&"</a></td><td align=center>" &CheckSize(thing.size)& "</td><td align=center>" & thing.type & "</td><td align=center>" & cstr(thing.datelastmodified) & "</td></tr>"

	elseif i>Page*Pagesize then
		exit for
	end if
end if

Next

%></table>


</div>


<table width="88%" cellpadding=2 cellspacing=2>
<tr><td>
<INPUT type=submit onclick="checkclick('您确定要删除您所选的文件?');" value=" 删 除 ">

</td></form>

	<td align="center">
	
<form method=Post name=form action="?menu=files">
<input type=hidden name=dir value=<%=dir%>>
搜索文件名包含
<input size="20" name="genre" value="<%=genre%>">
<INPUT type=submit value=" 搜 索 "></td>

	<td align="right">

<%
if Page>1 then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page=1>首页</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Page-1 &">上一页</a>&nbsp;&nbsp;"
else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
end if
if Page<i/Pagesize then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Page+1 &">下一页</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Pagenum &">尾页</a>"
else
	response.write "下一页&nbsp;&nbsp;尾页"
end if
%>
</td>

</tr></table>


</td>
	</tr>
</table>
<%

end sub

sub Posts
set Rs=conn.openSchema(20)

%>
<form method=Post name=form action="?menu=update">
<table  cellspacing="1" cellpadding="2" border="0" width="100%" class="a2">
	<tr class=a1>
		<td align="center" colspan="8" height="25"><%=rs(0)%>帖子数据表管理</td>
	</tr>
	<tr class=a3>
		<td align="center">表名</td>
		<td align="center">帖子数</td>
		<td align="center">默认</td>
		<td align="center">创建时间</td>
		<td align="center">修改时间</td>
		<td align="center">说明</td>
		<td align="center">所有者</td>
		<td align="center">管理</td>
	</tr>

<%
Do Until Rs.EOF
if instr(Rs("TABLE_NAME"),"BBSXP_Posts")>0 and Rs("TABLE_TYPE")="TABLE" then
TableNameNO=replace(Rs("TABLE_NAME"),"BBSXP_Posts","")
%>
	<tr class=a3>
		<td align="center"><%=Rs("TABLE_NAME")%></td>
		<td align="center"><%=Conn.execute("Select count(id)from ["&Rs("TABLE_NAME")&"]",0,1)(0)%></td>
		<td align="center"><input type="radio" <%if "BBSXP_Posts"&SiteSettings("DefaultPostsName")&""=""&Rs("TABLE_NAME")&"" then%>CHECKED<%end if%> value="<%=TableNameNO%>" name="TableName"></td>
		<td align="center"><%=Rs(7)%></td>
		<td align="center"><%=Rs(8)%></td>
		<td align="center"><%=Rs(5)%></td>
		<td align="center"><%=Rs(1)%></td>
		<td align="center"><a onclick="checkclick('您确定将表的所有者更改为dbo?');" href="?menu=changeobjectowner&tablename=<%=TableNameNO%>">更改所有者</a> <a href="?menu=Del&tablename=<%=TableNameNO%>">删除数据表</a></td>
	</tr>
<%
end if
Rs.movenext
Loop

%>
	
	<tr class=a3>
		<td align="center" colspan="8">
		<input type="submit" value="提 交"></td>
	</tr>
	
	<tr class=a3>
		<td colspan="8">
		<i>注释：</i>上面数据表中选中的为当前论坛所使用来保存帖子数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您下列单个数据表中的数据有超过几十万的帖子时不妨新添一个数据表来保存帖子数据，您会发现论坛速度快很多。 
		您也可以将当前所使用的数据表在数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认保存帖子的数据表.<br>
		<i>建议：</i>ACCESS 版本超过 200,000 条数据建立新表<br>
		　　　 MSSQL 版本超过 500,000 条数据建立新表</td>
	</tr>
</table></form>
<br>
<form method=Post name=form action="?menu=CreatTable">
<table  cellspacing="1" cellpadding="2" border="0" width="100%" class="a2">
	<tr class=a1>
		<td align="center" colspan="2" height="25">
		添加数据表</td>
	</tr>
	<tr class=a3>
		<td align="center" width="70%">
		添加的表名:<input value="1" name="tablename" size="2" onkeyup=if(isNaN(this.value))this.value='' maxLength=2>&nbsp;&nbsp; 只需要填写数字( 1 - 9 )即可</td>
		<td align="center" width="30%">
		<input type="submit" value="提 交"></td>
	</tr>
	</table></form>
<%
end sub

sub PostAttachment

if Request("Search")<>"" then SearchSql="where "&Request("Search")&" like '%"&Request("Key")&"%'"

if Request("order")<>"" then
order=HTMLEncode(Request("order"))
else
order="id"
end if

sql="[BBSXP_PostAttachments] "&SearchSql&""
rs.Open ""&sql&" order by "&order&" Desc",Conn,1
TotalCount=conn.Execute("Select count(ID) From "&sql&"")(0) '获取数据数量
PageSetup=20 '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = cint(Request.QueryString("PageIndex")) '获取当前页
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
%>
共搜索到 <font color="#FF0000"><b><%=TotalCount%></b></font> 个附件
<table cellspacing="1" cellpadding="2" border="0" width=100% class="a2"><form method=Post name=form action="?menu=DelPostAttachment">
<tr class=a1 id="TableTitleLink"><td width=50 align="center">
<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
<td align=center width="15%"><a href="?menu=PostAttachment&order=FileName">附件名称</a></td>
<td align=center width="80"><a href="?menu=PostAttachment&order=ContentSize">大小</a></td>
<td align="center"><a href="?menu=PostAttachment&order=ContentType">类型</a></td>
<td align="center"><a href="?menu=PostAttachment&order=Created">上传时间</a></td>
<td align="center"><a href="?menu=PostAttachment&order=UserName">上传用户</a></td>
<td align="center" width="50"><a href="?menu=PostAttachment&order=TotalDownloads">下载次数</a></td>
<td align="center" width="25%"><a href="?menu=PostAttachment&order=Description">关联帖子</a></td></tr>

<%


i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1
%>


<tr class=a3><td align="center"><input type=checkbox name=ID value="<%=RS("id")%>"></td>
	<td align=center style=word-break:break-all>
	<a target="_blank" href="PostAttachment.asp?AttachmentID=<%=RS("id")%>"><%=RS("FileName")%></a></td>
	<td align=center><%=CheckSize(RS("ContentSize"))%></td>
	<td align=center><%=RS("ContentType")%></td>
	<td align=center><%=RS("Created")%></td>
	<td align=center><a target="_blank" href="Profile.asp?UserName=<%=RS("UserName")%>"><%=RS("UserName")%></a></td>
	<td align=center><%=RS("TotalDownloads")%></td>
	<td align=center style=word-break:break-all><a target="_blank" href="ShowPost.asp?ThreadID=<%=RS("ThreadID")%>"><%=RS("Description")%></a></td>
</tr>

<%

Rs.MoveNext
loop
Rs.Close
%>
</table>
<table border="0" width="100%">
	<tr>
		<td><input type="submit" value="删 除"></form></td>
		<td>
		
			<form name="form" method="post" action="?menu=PostAttachment">
			搜索 <select size="1" name="Search">
<option value="FileName" selected>附件名称</option>
<option value="ContentType">附件类型</option>
<option value="UserName">上传用户</option>
<option value="Description">关联帖子</option>
</select> 包含
				<input onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" name="Key">
				<input id="btnadd" disabled type="submit" value="搜索"> 
			</form>		
		</td>
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>


<%
end sub



htmlend
%>