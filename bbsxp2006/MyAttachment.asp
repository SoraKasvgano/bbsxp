<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

if Request("menu")="DelPostAttachment" then
for each ho in Request.form("ID")
ho=RequestInt(ho)
Conn.execute("Delete from [BBSXP_PostAttachments] where id="&ho&" and UserName='"&SqlString(CookieUserName)&"'")
next
error2("ɾ���ɹ�")
end if

sql="[BBSXP_PostAttachments] where UserName='"&SqlString(CookieUserName)&"'"
rs.Open ""&sql&" order by id Desc",Conn,1
TotalCount=conn.Execute("Select count(ID) From "&sql&"")(0) '��ȡ��������
PageSetup=20 '�趨ÿҳ����ʾ����
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
PageCount = cint(Request.QueryString("PageIndex")) '��ȡ��ǰҳ
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
TotalUserPostAttachments=conn.execute("Select sum(ContentSize) from "&sql&"")(0)
if ""&TotalUserPostAttachments&""="" then TotalUserPostAttachments=0
BytesUsed=int(TotalUserPostAttachments/SiteSettings("MaxPostAttachmentsSize")*100)

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �������</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">�������</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">�����޸�</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">�����޸�</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">��������</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">��������</a></b></td>
      <TD align="center"><b><a href="Message.asp">���ŷ���</a></b></td>
      <TD align="center"><b><a href="Friend.asp">�����б�</a></b></td>
      </TR></TABLE>
<br>
<%if TotalUserPostAttachments=0 and SiteSettings("AttachmentsSaveOption")=0 then%>
<center>������ֻ�ܹ����洢�����ݿ��еĸ�����<br>��̳�ϴ��ĸ������Ǵ洢�����ݿ��У�</center>
<%
htmlend

end if
%>
<table cellSpacing="0" cellPadding="4" width="100%" align="center" border="0">
	<tr>
		<td align="center">���ĸ����ռ�ʹ����Ϣ����ʹ�� <%=CheckSize(""&TotalUserPostAttachments&"")%>��ʣ�� <%=CheckSize((SiteSettings("MaxPostAttachmentsSize")-TotalUserPostAttachments))%></td>
	</tr>
	<tr>
		<td>
		<table cellSpacing="0" cellPadding="0" width="100%" border="0">
			<tr>
				<td width="30"><font color="#ff0000"><%=BytesUsed%>%</font></td>
				<td>
				<table cellSpacing="1" cellPadding="0" width="100%" border="0" class=a1>
					<tr>
						<td class=a3 style="BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-BOTTOM: #ffffff 1px solid">
						<table cellSpacing="0" cellPadding="0" width=<%=BytesUsed%>% border="0" class=a1>
							<tr>
								<td>��</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
				<td align="middle" width="30"><font color="#ff0000">100%</font></td>
			</tr>
		</table>
		</td>
	</tr>
</table>


<table cellspacing="1" cellpadding="2" border="0" width=100% class="a2"><form method=Post name=form action="?menu=DelPostAttachment">
<tr class=a1 id="TableTitleLink">
<td align="center" width=30><input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></td>
<td align=center width="20%">��������</td>
<td align=center width="60">��С</td>
<td align="center">����</td>
<td align="center">�ϴ�ʱ��</td>
<td align="center" width="50">���ش���</td>
<td align="center" width="25%">��������</td></tr>

<%


i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1
%>


<tr class=a3><td align="center"><input type=checkbox name=ID value="<%=RS("id")%>"></td>
	<td align=center>
	<a target="_blank" href="PostAttachment.asp?AttachmentID=<%=RS("id")%>" style=word-break:break-all><%=RS("FileName")%></a></td>
	<td align=center><%=CheckSize(RS("ContentSize"))%></td>
	<td align=center><%=RS("ContentType")%></td>
	<td align=center><%=RS("Created")%></td>
	<td align=center><%=RS("TotalDownloads")%></td>
	<td align=center style=word-break:break-all><a href="ShowPost.asp?ThreadID=<%=RS("ThreadID")%>"><%=RS("Description")%></a></td>
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
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>


<%
htmlend
%>