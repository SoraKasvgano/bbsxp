<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")

select case Request("menu")
	case "PostAttachment"
		PostAttachment
	case "DelPostAttachment"
		for each ho in Request.Form("UpFileID")
			ho=int(ho)
			sql="Select * from ["&TablePrefix&"PostAttachments] where UpFileID="&ho&""
			Rs.Open sql,Conn,1
				if Not Rs.Eof then
					if ""&Rs("FilePath")&""<>"" then IsDelFile=DelFile(""&Rs("FilePath")&"")
					if ""&Rs("FilePath")&""="" or (""&Rs("FilePath")&""<>"" and IsDelFile=True) then
						Execute("Delete from ["&TablePrefix&"PostAttachments] where UpFileID="&Rs("UpFileID")&"")
					end if
				End if
			Rs.close
		next
		Alert("ɾ���ɹ���")
	
	case "ClearAttachment"
		Rs.Open "Select * from ["&TablePrefix&"PostAttachments]",Conn,1,3
			if Not Rs.eof then
				do while not Rs.eof
					if Execute("select PostID from ["&TablePrefix&"Posts] where PostID="&Rs("PostID")&"").eof or Execute("select UserName from ["&TablePrefix&"Users] where UserName='"&Rs("UserName")&"'").eof then
						if ""&Rs("FilePath")&""<>"" then IsDelFile=DelFile(""&Rs("FilePath")&"")
						if ""&Rs("FilePath")&""="" or (""&Rs("FilePath")&""<>"" and IsDelFile=True) then
							Rs.Delete
							Rs.update
						end if
					end if
					Rs.movenext
				loop
			end if
		Rs.close
		Alert("�ѳɹ�����ռ������õĸ�����")

	case "bak"
		bak
	case "bakbf"
		Set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
		MyFileObject.CopyFile ""&Server.MapPath(SqlDataBase)&"",""&Server.MapPath(Request.Form("BakDbPath"))&""
		Alert("���ݳɹ���")
	case "bakhf"
		Set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
		MyFileObject.CopyFile ""&Server.MapPath(Request.Form("BakDbPath"))&"",""&Server.MapPath(SqlDataBase)&""
		Alert("�ָ��ɹ���")
	case "statroom"
		statroom
end select

Sub bak
	If IsSqlDataBase<>0 Then Alert("SQL�汾�޷�����ACCESS���ݿ����")
%>

<table cellspacing="1" cellpadding="5" width="70%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan=2>�������ݿ�</td>
	</tr>
	<tr class="CommonListCell">
		<td align=center colspan=2>
		<form method="POST" action="?menu=bakbf">
			<table cellpadding="0" cellspacing="0" width="90%">
				<tr>
					<td width="30%">���ݿ�·���� </td>
					<td width="70%"><%=SqlDataBase%></td>
				</tr>
				<tr>
					<td width="30%">���ݵ����ݿ�·����</td>
						<td width="70%"><input size="30" value="<%=replace(""&SqlDataBase&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath" /></td>
				</tr>
				<tr>
						<td width="100%" align="center" colspan="2"><input type="submit" value=" �� �� " /><br /></td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="5" width="70%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan=2>�ָ����ݿ�</td>
	</tr>
	<tr class="CommonListCell">
		<td align=center colspan=2>
		<form method="POST" action="?menu=bakhf">
			<table cellpadding="0" cellspacing="0" width="90%">
				<tr>
					<td width="30%">���ݵ����ݿ�·���� </td>
					<td width="70%"><input size="30" value="<%=replace(""&SqlDataBase&"",".mdb","("&Date()&").mdb")%>" name="BakDbPath" /></td>
				</tr>
				<tr>
					<td width="30%">���ݿ�·����</td>
					<td width="70%"><%=SqlDataBase%></td>
				</tr>
				<tr>
					<td width="100%" align="center" colspan="2"><input type="submit" value=" �� �� " /><br /></td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="5" width="70%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan=2>ѹ�����ݿ�</td>
	</tr>
	<tr class="CommonListCell">
		<td align=center colspan=2>
		<form action=Compact.asp method=Post>
		<input type=hidden name=sessionid value=<%=session.sessionid%> />
			<table cellpadding="0" cellspacing="0" width="90%">
				<tr>
					<td width="70%">
						<table cellpadding="0" cellspacing="0" width="90%">
							<tr>
								<td width="30%">���ݿ�·���� </td>
								<td width="70%"><input size="30" value="<%=SqlDataBase%>" name="dbpath" /></td>
							</tr>
							<tr>
								<td width="30%">���ݿ��ʽ��</td>
								<td width="70%"><input type="radio" value="True" name="boolIs97" id=boolIs97 /><label for=boolIs97>Access 97</label>��<input type="radio" value="" name="boolIs97" checked="checked" id=boolIs97_1 /><label for=boolIs97_1>Access 2000��2002��2003</label></td>
							</tr>
							<tr>
								<td width="100%" align="center" colspan="2"><input type="submit" value=" ѹ �� " /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
<%
End Sub

Sub statroom
	Set fso=server.createobject("Scripting.FileSystemObject")
	
	UpFacedir=server.mappath("./UpFile/UpFace")
	Set d=fso.getfolder(UpFacedir)
	UpFacesize=d.size
	UpFaceCounts=d.Files.count
	Set d = Nothing
	
	UpFiledir=server.mappath("./UpFile/UpAttachment")
	Set d=fso.getfolder(UpFiledir)
	UpFilesize=d.size
	Set d = Nothing

	toldir=server.mappath(".")
	Set d=fso.getfolder(toldir)
	tolsize=d.size
	Set d = Nothing

	Set fso = Nothing

	totalBytes=UpFacesize+UpFilesize+tolsize
%>
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
    	<td align=center colspan=3>ͳ��ռ�ÿռ�</td>
	</tr>
	<tr class="CommonListCell">
		<td width="29%">�ϴ�ͷ��ռ�ÿռ�(ͷ������<%=UpFaceCounts%>)<br /></td>
		<td width="60%">
<div class="percent"><div style="width:<%=Int(UpFacesize/totalBytes*100)%>%"></div></div></td>
		<td><%=CheckSize(UpFacesize)%></td></tr>
	<tr class="CommonListCell">
		<td width="29%">�ϴ�����ռ�ÿռ�</td>
		<td width="60%"><div class="percent"><div style="width:<%=Int(UpFilesize/totalBytes*100)%>%"></div></div> </td>
		<td><%=CheckSize(UpFilesize)%></td>
	</tr>
	<tr class="CommonListCell">
		<td width="29%">BBSXPĿ¼�ܹ�ռ�ÿռ�</td>
		<td width="60%"><div class="percent"><div style="width:<%=Int(tolsize/totalBytes*100)%>%"></div></div> </td>
		<td><%=CheckSize(tolsize)%></td>
	</tr>
</table>
<%
End Sub


Sub PostAttachment

	Search=HTMLEncode(Request("Search"))
	Key=HTMLEncode(Request("Key"))
	order=HTMLEncode(Request("order"))

	if Search<>"" then SearchSql="where "&Search&" like '%"&Key&"%'"
	if order="" then order="UpFileID"
	sql="["&TablePrefix&"PostAttachments] "&SearchSql&""
	rs.Open ""&sql&" order by "&order&" Desc",Conn,1
		TotalCount=Execute("Select count(UpFileID) From "&sql&"")(0) '��ȡ��������
		PageSetup=20 '�趨ÿҳ����ʾ����
		TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
		PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
%>

<table border="0" width="100%">
	<tr>
		<td>
			<form name="form" method="post" action="?menu=PostAttachment" style="margin:0px;padding:0px">
			���� <select size="1" name="Search">
<option value="FileName" selected="selected">�ļ�����</option>
<option value="ContentType">�ļ�����</option>
<option value="UserName">�ϴ��û�</option>
<option value="Description">��������</option>
</select> ����
				<input onkeyup="ValidateTextboxAdd(this, 'btnadd')" name="Key" />
				<input id="btnadd" disabled="disabled" type="submit" value="����" /> 
			</form>		
		</td>
		<td align="right"><a href="?menu=ClearAttachment" onclick="return window.confirm('ȷ�����û���������ĸ�����')">�����������</a></td>
		<td align="right" width=130>�������� <font color="#FF0000"><b><%=TotalCount%></b></font> ���ļ�</td>
	</tr>
</table>
<table cellspacing="1" cellpadding="2" border="0"  width=100% class=CommonListArea>
<form method=Post name=form action="?menu=DelPostAttachment" style="margin:0px;padding:0px">
	<tr class=CommonListTitle>
		<td width=50 align="center"><input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON" /></td>
		<td align=center width="15%"><a href="?menu=PostAttachment&order=FileName">�ļ�����</a></td>
		<td align=center width="80"><a href="?menu=PostAttachment&order=ContentSize">��С</a></td>
		<td align="center"><a href="?menu=PostAttachment&order=ContentType">����</a></td>
		<td align="center" width="15%"><a href="?menu=PostAttachment&order=Created">�ϴ�ʱ��</a></td>
		<td align="center"><a href="?menu=PostAttachment&order=UserName">�ϴ��û�</a></td>
		<td align="center"><a href="?menu=PostAttachment&order=Description">����</a></td>
	</tr>
<%
		i=0
		Do While Not Rs.EOF and i<PageSetup
			i=i+1
			FilePath=RS("FilePath")
			if ""&FilePath&""="" then FilePath="GetAttachment.asp?AttachmentID="&Rs("UpFileID")&""
%>
	<tr class="CommonListCell" id="UpFile<%=Rs("UpFileID")%>">
		<td align="center"><input type=checkbox name=UpFileID value="<%=RS("UpFileID")%>" onclick="CheckSelected(this.form,this.checked,'UpFile<%=Rs("UpFileID")%>')" /></td>
		<td align=center><a target="_blank" href="<%=FilePath%>"><%=RS("FileName")%></a></td>
		<td align=center><%=CheckSize(RS("ContentSize"))%></td>
		<td align=center><%=RS("ContentType")%></td>
		<td align=center><%=RS("Created")%></td>
		<td align=center><a target="_blank" href="Profile.asp?UserName=<%=RS("UserName")%>"><%=RS("UserName")%></a></td>
		<td align=center><%if Rs("PostID")>0 then response.Write("<a href=ShowPost.asp?PostID="&Rs("PostID")&" target=_blank>"&RS("Description")&"</a>") else response.Write(RS("Description")) end if%></td>
	</tr>
<%
			Rs.MoveNext
		loop
	Rs.Close
%>
</table>
<table border="0" width="100%">
	<tr>
		<td><input type="submit" value="ɾ ��" /></form></td>
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>
<%
End Sub
AdminFooter
%>