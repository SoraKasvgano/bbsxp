<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")

if Request("menu")="DelPostAttachment" then
	if Request.Form("UpFileID")="" then error("��û��ѡ��Ҫɾ�����ļ���")
	for each ho in Request.Form("UpFileID")
		ho=int(ho)
		sql="Select * from ["&TablePrefix&"PostAttachments] where UpFileID="&ho&" and UserName='"&CookieUserName&"'"
		Rs.Open sql,Conn,1
		if Not Rs.Eof then
			if ""&Rs("FilePath")&""<>"" then IsDelFile=DelFile(""&Rs("FilePath")&"")
			if ""&Rs("FilePath")&""="" or (""&Rs("FilePath")&""<>"" and IsDelFile=True) then
				Execute("Delete from ["&TablePrefix&"PostAttachments] where UpFileID="&Rs("UpFileID")&"")
			end if
		end if
		Rs.close
	next
	succeed "ɾ���ɹ�",""
end if



RoleMaxPostAttachmentsSize=Execute("select RoleMaxPostAttachmentsSize from ["&TablePrefix&"Roles] where RoleID="&CookieUserRoleID&"")(0)
UpMaxPostAttachmentsSize=SiteConfig("MaxPostAttachmentsSize")*1024		'����������

if RoleMaxPostAttachmentsSize>0 then UpMaxPostAttachmentsSize = RoleMaxPostAttachmentsSize*1024


sql="["&TablePrefix&"PostAttachments] where UserName='"&CookieUserName&"'"
rs.Open ""&sql&" order by UpFileID Desc",Conn,1
	TotalCount=Execute("Select count(UpFileID) From "&sql&"")(0) '��ȡ��������
	PageSetup=20 '�趨ÿҳ����ʾ����
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
	PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	TotalUserPostAttachments=Execute("Select sum(ContentSize) from "&sql&"")(0)
	if ""&TotalUserPostAttachments&""="" then TotalUserPostAttachments=0
	
	BytesUsed=int(TotalUserPostAttachments/(UpMaxPostAttachmentsSize)*100)
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �ϴ�����</div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center"><a href="EditProfile.asp">�����޸�</a></td>
		<td align="center"><a href="EditProfile.asp?menu=pass">�����޸�</a></td>
		<td align="center"><a href="MyUpFiles.asp">�ϴ�����</a></td>
		<td align="center"><a href="MyFavorites.asp">�� �� ��</a></td>
		<td align="center"><a href="MyMessage.asp">���ŷ���</a></td>
	</tr>
</table>
<br />
<table cellspacing=1 cellpadding=5 border=0 width=100%>
	<tr>
		<td align="center">�����ϴ��ռ�ʹ����Ϣ����ʹ�� <%=CheckSize(""&TotalUserPostAttachments&"")%>��ʣ�� <%=CheckSize((UpMaxPostAttachmentsSize-TotalUserPostAttachments))%></td>
	</tr>
	<tr>
		<td>
		<table cellspacing="0" cellpadding="0" width="100%" border="0">
			<tr>
				<td width="30"><font color="#ff0000"><%=BytesUsed%>%</font></td>
				<td>
				<table cellspacing="0" cellpadding="0" width="100%" border="0">
					<tr>
						<td>
							<div class="percent"><div style="width:<%=BytesUsed%>%"></div></div>
						</td>
					</tr>
				</table>
				</td>
				<td align="center" width="30"><font color="#ff0000">100%</font></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<form method=Post name=form action="?menu=DelPostAttachment">
<table cellspacing=1 cellpadding=3 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center" width=30><input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON" /></td>
		<td align=center width="20%">�ļ�����</td>
		<td align=center width="60">�ļ���С</td>
		<td align="center">�ļ�����</td>
		<td align="center">�ϴ�ʱ��</td>
		<td align="center" width="25%">����</td>
	</tr>
<%
i=0
Do While Not Rs.EOF and i<PageSetup
	i=i+1
	FilePath=RS("FilePath")
	if ""&FilePath&""="" then FilePath="GetAttachment.asp?AttachmentID="&Rs("UpFileID")&""
%>
	<tr class="CommonListCell" id="UpFile<%=Rs("UpFileID")%>"><td align="center"><input type=checkbox name=UpFileID value="<%=Rs("UpFileID")%>" onclick="CheckSelected(this.form,this.checked,'UpFile<%=Rs("UpFileID")%>')" /></td>
		<td align="center"><a target="_blank" href="<%=FilePath%>"><%=RS("FileName")%></a></td>
		<td align=center><%=CheckSize(RS("ContentSize"))%></td>
		<td align=center><%=RS("ContentType")%></td>
		<td align=center><%=RS("Created")%></td>
		<td align=center><%if Rs("PostID")>0 then response.Write("<a href=ShowPost.asp?PostID="&Rs("PostID")&" target=_blank>"&RS("Description")&"</a>") else response.Write(RS("Description")) end if%></td>
	</tr>
<%
	Rs.MoveNext
loop
Rs.Close
%>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td><input type="submit" value="ɾ ��" /></form></td>
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>

<%
HtmlBottom
%>