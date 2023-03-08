<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

if Request("menu")="DelPostAttachment" then
for each ho in Request.form("ID")
ho=int(ho)
Conn.execute("Delete from [BBSXP_PostAttachments] where id="&ho&" and UserName='"&CookieUserName&"'")
next
error2("删除成功")
end if

sql="[BBSXP_PostAttachments] where UserName='"&CookieUserName&"'"
rs.Open ""&sql&" order by id Desc",Conn,1
TotalCount=conn.Execute("Select count(ID) From "&sql&"")(0) '获取数据数量
PageSetup=20 '设定每页的显示数量
TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数
PageCount = cint(Request.QueryString("PageIndex")) '获取当前页
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
TotalUserPostAttachments=conn.execute("Select sum(ContentSize) from "&sql&"")(0)
if ""&TotalUserPostAttachments&""="" then TotalUserPostAttachments=0
BytesUsed=int(TotalUserPostAttachments/SiteSettings("MaxPostAttachmentsSize")*100)

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 控制面板</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=100% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="UserCp.asp">控制面板</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">密码修改</a></b></td>
      <TD align="center"><b><a href="MySettings.asp">个性设置</a></b></td>
      <TD align="center"><b><a href="MyAttachment.asp">附件管理</a></b></td>
      <TD align="center"><b><a href="Message.asp">短信服务</a></b></td>
      <TD align="center"><b><a href="Friend.asp">好友列表</a></b></td>
      </TR></TABLE>
<br>
<%if TotalUserPostAttachments=0 and SiteSettings("AttachmentsSaveOption")=0 then%>
<center>本功能只能管理存储于数据库中的附件！<br>论坛上传的附件并非存储于数据库中！</center>
<%
htmlend

end if
%>
<table cellSpacing="0" cellPadding="4" width="100%" align="center" border="0">
	<tr>
		<td align="center">您的附件空间使用信息：已使用 <%=CheckSize(""&TotalUserPostAttachments&"")%>，剩余 <%=CheckSize((SiteSettings("MaxPostAttachmentsSize")-TotalUserPostAttachments))%></td>
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
								<td>　</td>
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
<td align=center width="20%">附件名称</td>
<td align=center width="60">大小</td>
<td align="center">类型</td>
<td align="center">上传时间</td>
<td align="center" width="50">下载次数</td>
<td align="center" width="25%">关联帖子</td></tr>

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
		<td><input type="submit" value="删 除"></form></td>
		<td align="right" valign="top"><%ShowPage()%></td>
	</tr>
</table>


<%
htmlend
%>