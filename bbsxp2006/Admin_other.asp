<!-- #include file="Setup.asp" -->
<%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"


response.write "<center>"

select case Request("menu")

case "discreteness"
discreteness

case "Log"
Log

case "circumstance"
circumstance

end select





sub circumstance

%>

<table border="0" cellpadding="2" cellspacing="1" width=90% class=a2 align=center>
<tr class=a1>
<td height="16" align="center">��Ŀ</td>
<td width="66%" align="center" height="25">ֵ</td>
</tr>
<tr class=a3>
<td height="25">������������</td>
<td width="66%" height="25"><%=Request.ServerVariables("server_name")%></td>
</tr>
<tr class=a4>
<td height="25">��������IP��ַ</td>
<td width="66%" height="25"><%=Request.ServerVariables("LOCAL_ADDR")%>
</tr>
<tr class=a3>
<td height="25">����������ϵͳ</td>
<td width="66%" height="25"><%=Request.ServerVariables("OS")%>
</tr>
<tr class=a4>
<td height="25">��������������</td>
<td width="66%" height="25"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
</tr>
<tr class=a3>
<td height="25">��������������Ƽ��汾</td>
<td width="66%" height="25"><%=Request.ServerVariables("SERVER_SOFTWARE")%>
</tr>
<tr class=a4>
<td height="25">�������������еĶ˿�</td>
<td width="66%" height="25"><%=Request.ServerVariables("server_port")%>
</tr>
<tr class=a3>
<td height="25">������CPU����</td>
<td width="66%" height="25"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></td>
</tr>

<tr class=a4>
<td height="25">������Application����</td>
<td width="66%" height="25"><%=Application.contents.Count%></td>
</tr>

<tr class=a3>
<td height="25">������Session����</td>
<td width="66%" height="25"><%=Session.contents.Count%></td>
</tr>

<tr class=a4>
<td height="25">���������·��</td>
<td width="66%" height="25"><%=Request.ServerVariables("path_translated")%>
</tr>
<tr class=a3>
<td height="25">�����URL</td>
<td width="66%" height="25">http://<%=Request.ServerVariables("server_name")%><%=Request.ServerVariables("script_name")%></td>
</tr>
<tr class=a4>
<td height="25">��������ǰʱ��</td>
<td width="66%" height="25"><%=now()%>
</tr>
<tr class=a4>
<td height="25">�ű����ӳ�ʱʱ��</td>
<td width="66%" height="25"><%=Server.ScriptTimeout%> ��</td>
</tr>
</table>
<%
end sub

sub Log
%>
����¼ֻ����һ������
<table cellspacing=1 width=100% cellpadding=2 class=a2 align=center>
<tr class=a1 id=TableTitleLink height="25">
<td align="center"><a href="?menu=Log&order=UserName">�û���</a></td>
<td align="center"><a href="?menu=Log&order=IPAddress">IP</a></td>
<td align="center"><a href="?menu=Log&order=HttpVerb">�ύ����</a></td>
<td align=center width="25%">������Ŀ</td>
<td align="center"><a href="?menu=Log&order=DateCreated">����ʱ��</a></td>
<td align="center" width="20%"><a href="?menu=Log&order=UserAgent">����ϵͳ�������</a></td>
</tr>
<%


if Request("order")<>"" then
order=HTMLEncode(Request("order"))
else
order="DateCreated"
end if

Search=HTMLEncode(Request("Search"))
if Search<>"" then Searchlist="where UserName like '%"&Search&"%' or IPAddress like '%"&Search&"%' or PathAndQuery like '%"&Search&"%'"

sql="select * from [BBSXP_Log] "&Searchlist&" order by "&order&" Desc"
Rs.Open sql,Conn,1

PageSetup=20 '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��


i=0
Do While Not Rs.EOF and i<PageSetup
i=i+1

response.write "<tr class=a3><td align=center>"&Rs("UserName")&"</td><td align=center>"&Rs("IPAddress")&"</td><td align=center>"&Rs("HttpVerb")&"</td><td style=word-break:break-all>"&Rs("PathAndQuery")&"</td><td align=center>"&Rs("DateCreated")&"</td><td align=center style=word-break:break-all><font size=1>"&Rs("UserAgent")&"</font></td></tr>"



Rs.MoveNext
loop
Rs.close
%>
</table>
<table cellspacing=1 width=100% align=center>

	<tr>
		<td width="50%">
<%ShowPage()%>
</td>

		<td align="right"><form action="Admin_other.asp?menu=Log" method="POST">
			������¼��<input size="15" value="<%=Search%>" name="Search"> <input type="submit" value=" ȷ�� "></td></form>
	</tr>
</table>



<%
end sub

sub discreteness
%>
<table border=0 width="90%" cellspacing=1 cellpadding=3 class=a2 align=center>
<tr class=a1>
<td width="57%" height="25">&nbsp;�������</td><td width="41%" height="25">֧�ּ��汾</td>
</tr>
<%
Dim theInstalledObjects(19)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "MSWC.PermissionChecker"
theInstalledObjects(7) = "WScript.Shell"
theInstalledObjects(8) = "ADODB.Stream"
theInstalledObjects(9) = "Scripting.Dictionary"
theInstalledObjects(10) = "Adodb.Connection"
theInstalledObjects(11) = "Scripting.FileSystemObject"
theInstalledObjects(12) = "SoftArtisans.FileUp"
theInstalledObjects(13) = "SoftArtisans.FileManager"
theInstalledObjects(14) = "JMail.Message"
theInstalledObjects(15) = "CDONTS.NewMail"
theInstalledObjects(16) = "Persits.MailSender"
theInstalledObjects(17) = "Persits.Jpeg"
theInstalledObjects(18) = "Persits.Upload.1"
theInstalledObjects(19) = "w3.upload"


For i=0 to 19
Response.Write "<TR class=a3><TD>&nbsp;" & theInstalledObjects(i) & "<font color=888888>&nbsp;"
select case i
case 10
Response.Write "(ACCESS ���ݿ�)"
case 11
Response.Write "(FSO �ı��ļ���д)"
case 12
Response.Write "(SA-FileUp �ļ��ϴ�)"
case 13
Response.Write "(SA-FM �ļ�����)"
case 14
Response.Write "(JMail �ʼ�����)"
case 15
Response.Write "(WIN����SMTP ����)"
case 16
Response.Write "(ASPEmail �ʼ�����)"
case 17
Response.Write "(ASPJPEG ͼ����)"
case 18
Response.Write "(ASPUpload �ļ��ϴ�)"
case 19
Response.Write "(w3 upload �ļ��ϴ�)"
end select
Response.Write "</font></td><td height=25>"
If Not IsObjInstalled(theInstalledObjects(i)) Then
Response.Write "<font color=red><b>��</b></font>"
Else
Response.Write "<b>��</b> " & getver(theInstalledObjects(i)) & ""
End If
Response.Write "</td></TR>" & vbCrLf
Next
%>
</table>
<%
end sub


''''''''''''''''''''''''''''''
Function getver(Classstr)
On Error Resume Next
Set xTestObj = Server.CreateObject(Classstr)
If 0 = Err Then getver=xtestobj.version
Set xTestObj = Nothing
On Error GoTo 0
End Function
''''''''''''''''''''''''''''''
htmlend
%>