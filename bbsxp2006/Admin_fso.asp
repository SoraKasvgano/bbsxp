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
error2("ɾ���ɹ�")


case "changeobjectowner"
SQL="sp_changeobjectowner '[BBSXP_Posts"&Request("tablename")&"]','dbo'"
Conn.execute(SQL)
error2("�Ѿ��ɹ����ñ�������߸���Ϊdbo")

case "CreatTable"

'�������ݿ�
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

'��������
Sql="ALTER TABLE [BBSXP_Posts"&Request("tablename")&"] ADD CONSTRAINT [FK_BBSXP_Posts"&Request("tablename")&"_BBSXP_Threads] FOREIGN KEY ([ThreadID]) REFERENCES [BBSXP_Threads] ([Id]) ON DELETE CASCADE  ON UPDATE CASCADE"
Conn.execute(sql)

Conn.execute("update [BBSXP_SiteSettings] set DefaultPostsName="&Request("tablename"))

error2("�����ɹ���")


case "update"
if Request("tablename") = "" Then ThisValue = "Null" Else ThisValue = Request("tablename")
Conn.execute("update [BBSXP_SiteSettings] set DefaultPostsName="&ThisValue)
error2("���óɹ���") 



case "Del"
if request("tablename")="" then error2("ϵͳĬ�ϵı���ɾ����")
If not Conn.Execute("Select PostsTableName From [BBSXP_Threads] where PostsTableName='"&request("tablename")&"'" ).eof Then error2("�ñ����ж�Ӧ�����⣬����ɾ����")
if ""&SiteSettings("DefaultPostsName")&""=""&request("tablename")&"" then error2("��ǰ����ʹ���е����ݱ���ɾ����")
Conn.Execute("drop table [BBSXP_Posts"&request("tablename")&"]")
error2("ɾ���ɹ���")


case "files"
files

case "Delfiles"

if dir<>"" then dir=""&dir&"/"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
for each ho in request.form("files")
MyFileObject.DeleteFile""&Server.MapPath("./UpFile/"&dir&""&ho&"")&""
next
error2("�Ѿ��ɹ�ɾ����ѡ���ļ���")


case "bak"
bak

case "bakbf"

set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(DB)&"",""&Server.MapPath(Request.Form("BakDbPath"))&""
error2("���ݳɹ���")

case "bakhf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request.Form("BakDbPath"))&"",""&Server.MapPath(DB)&""
error2("�ָ��ɹ���")


case "statroom"
statroom

end select

sub bak


If IsSqlDataBase<>0 Then error2("SQL�汾�޷�����ACCESS���ݿ����")

%>


<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�������ݿ�</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
<form method="POST" action="?menu=bakbf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%" height="25">���ݿ�·���� </td>
<td width="70%"><%=Db%></td>
</tr>
<tr>
<td width="30%" height="25">���ݵ����ݿ�·����</td>
<td width="70%"><input size="30" value="<%=replace(""&Db&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath"></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" �� �� "><br></td>
</tr>
</table>
</td></tr></table></form>
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>�ָ����ݿ�</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
<form method="POST" action="?menu=bakhf">
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="90%">
<tr>
<td width="30%" height="25">���ݵ����ݿ�·���� </td>
<td width="70%">
<input size="30" value="<%=replace(""&Db&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath"></td>
</tr>
<tr>
<td width="30%" height="25">���ݿ�·����</td>
<td width="70%"><%=Db%></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" �� �� "><br></td>
</tr>
</table></td></tr></table>
</form>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>ѹ�����ݿ�</td>
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

<td width="30%">���ݿ�·���� </td>
<td width="70%">
<input size="30" value="<%=Db%>" name="dbpath"></td>
</tr>
<tr>
<td width="30%">���ݿ��ʽ��</td>
<td width="70%"><input type="radio" value="True" name="boolIs97" id=boolIs97><label for=boolIs97>Access 97</label>��<input type="radio" value="" name="boolIs97" checked id=boolIs97_1><label for=boolIs97_1>Access 2000��2002��2003</label></td>
</tr>
<tr>
<td width="100%" align="center" colspan="2">
<input type="submit" value=" ѹ �� "></td>
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
    <td class=a1 align=middle colspan=3>ͳ��ռ�ÿռ�</td>
  </tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ�ͷ��ռ�ÿռ�<br>




</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/0.gif" width=<%=Int(UpFacesize/totalBytes*100)%>%> </td>
    <td class=a3><%=CheckSize(UpFacesize)%></td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ���Ƭռ�ÿռ�</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/1.gif" width=<%=Int(UpPhotosize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(UpPhotosize)%>

��</td></tr>
   <tr height=25>
    <td class=a3 width="29%">
    

�ϴ�����ռ�ÿռ�</td>
    <td class=a3 width="60%">
    

<IMG height=10 src="images/bar/2.gif" width=<%=Int(UpFilesize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(UpFilesize)%>

��</td></tr>



   <tr height=25>
    <td class=a3 width="29%" height="24">
    

BBSXPĿ¼�ܹ�ռ�ÿռ�</td>
    <td class=a3 width="60%" height="24">
    

<IMG height=10 src="images/bar/4.gif" width=<%=Int(tolsize/totalBytes*100)%>%> </td>
    <td class=a3>
    <%=CheckSize(tolsize)%>

��</td></tr></table>


<%
end sub


sub files
thisdir=server.mappath("./UpFile/"&dir&"")
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set fdir=fs.GetFolder(thisdir)

count=fdir.Files.count

%>

<a href="?menu=files&dir=<%=dir%>"><%=dir%>Ŀ¼��</a>�� <font color="FF0000"><b><%=count%></b></font> ���ļ�

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

PageSetup=10 '�趨ÿҳ����ʾ����
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
	<td width=208 align=center>����</td>
  <td width=128 align=center>
��С</td><td width=140 align="center">����</td><td align="center">
�޸�ʱ��</td></tr>
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
<INPUT type=submit onclick="checkclick('��ȷ��Ҫɾ������ѡ���ļ�?');" value=" ɾ �� ">

</td></form>

	<td align="center">
	
<form method=Post name=form action="?menu=files">
<input type=hidden name=dir value=<%=dir%>>
�����ļ�������
<input size="20" name="genre" value="<%=genre%>">
<INPUT type=submit value=" �� �� "></td>

	<td align="right">

<%
if Page>1 then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page=1>��ҳ</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Page-1 &">��һҳ</a>&nbsp;&nbsp;"
else
	response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
end if
if Page<i/Pagesize then
	response.write "<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Page+1 &">��һҳ</a>&nbsp;&nbsp;<a href=?menu=files&dir="&dir&"&genre="&genre&"&Page="& Pagenum &">βҳ</a>"
else
	response.write "��һҳ&nbsp;&nbsp;βҳ"
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
		<td align="center" colspan="8" height="25"><%=rs(0)%>�������ݱ����</td>
	</tr>
	<tr class=a3>
		<td align="center">����</td>
		<td align="center">������</td>
		<td align="center">Ĭ��</td>
		<td align="center">����ʱ��</td>
		<td align="center">�޸�ʱ��</td>
		<td align="center">˵��</td>
		<td align="center">������</td>
		<td align="center">����</td>
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
		<td align="center"><a onclick="checkclick('��ȷ������������߸���Ϊdbo?');" href="?menu=changeobjectowner&tablename=<%=TableNameNO%>">����������</a> <a href="?menu=Del&tablename=<%=TableNameNO%>">ɾ�����ݱ�</a></td>
	</tr>
<%
end if
Rs.movenext
Loop

%>
	
	<tr class=a3>
		<td align="center" colspan="8">
		<input type="submit" value="�� ��"></td>
	</tr>
	
	<tr class=a3>
		<td colspan="8">
		<i>ע�ͣ�</i>�������ݱ���ѡ�е�Ϊ��ǰ��̳��ʹ���������������ݵı�һ�������ÿ�����е�����Խ����̳������ʾ�ٶ�Խ�죬�������е������ݱ��е������г�����ʮ�������ʱ��������һ�����ݱ��������������ݣ����ᷢ����̳�ٶȿ�ܶࡣ 
		��Ҳ���Խ���ǰ��ʹ�õ����ݱ������ݱ����л�����ǰ��ʹ�õ��������ݱ���ǰ��̳�û�����ʱĬ�ϱ������ӵ����ݱ�.<br>
		<i>���飺</i>ACCESS �汾���� 200,000 �����ݽ����±�<br>
		������ MSSQL �汾���� 500,000 �����ݽ����±�</td>
	</tr>
</table></form>
<br>
<form method=Post name=form action="?menu=CreatTable">
<table  cellspacing="1" cellpadding="2" border="0" width="100%" class="a2">
	<tr class=a1>
		<td align="center" colspan="2" height="25">
		������ݱ�</td>
	</tr>
	<tr class=a3>
		<td align="center" width="70%">
		��ӵı���:<input value="1" name="tablename" size="2" onkeyup=if(isNaN(this.value))this.value='' maxLength=2>&nbsp;&nbsp; ֻ��Ҫ��д����( 1 - 9 )����</td>
		<td align="center" width="30%">
		<input type="submit" value="�� ��"></td>
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
TotalCount=conn.Execute("Select count(ID) From "&sql&"")(0) '��ȡ��������
PageSetup=20 '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = cint(Request.QueryString("PageIndex")) '��ȡ��ǰҳ
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
%>
�������� <font color="#FF0000"><b><%=TotalCount%></b></font> ������
<table cellspacing="1" cellpadding="2" border="0" width=100% class="a2"><form method=Post name=form action="?menu=DelPostAttachment">
<tr class=a1 id="TableTitleLink"><td width=50 align="center">
<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
<td align=center width="15%"><a href="?menu=PostAttachment&order=FileName">��������</a></td>
<td align=center width="80"><a href="?menu=PostAttachment&order=ContentSize">��С</a></td>
<td align="center"><a href="?menu=PostAttachment&order=ContentType">����</a></td>
<td align="center"><a href="?menu=PostAttachment&order=Created">�ϴ�ʱ��</a></td>
<td align="center"><a href="?menu=PostAttachment&order=UserName">�ϴ��û�</a></td>
<td align="center" width="50"><a href="?menu=PostAttachment&order=TotalDownloads">���ش���</a></td>
<td align="center" width="25%"><a href="?menu=PostAttachment&order=Description">��������</a></td></tr>

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
		<td><input type="submit" value="ɾ ��"></form></td>
		<td>
		
			<form name="form" method="post" action="?menu=PostAttachment">
			���� <select size="1" name="Search">
<option value="FileName" selected>��������</option>
<option value="ContentType">��������</option>
<option value="UserName">�ϴ��û�</option>
<option value="Description">��������</option>
</select> ����
				<input onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" name="Key">
				<input id="btnadd" disabled type="submit" value="����"> 
			</form>		
		</td>
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>


<%
end sub



htmlend
%>