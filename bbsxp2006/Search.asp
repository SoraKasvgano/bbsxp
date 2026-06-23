<!-- #include file="Setup.asp" --><%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")


if Request("menu")="ok" then

if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")

Search=""&Request("Search")&""
ForumID=RequestInt("ForumID")
TimeLimit=RequestInt("TimeLimit")
content=HTMLEncode(Request("content"))
Searchxm=HTMLEncode(Request("Searchxm"))


if content="" then error("<li>��û������ؼ���")


if ForumID>0 then ForumIDor="ForumID="&ForumID&" and"

if Search="author" then
if Len(Searchxm)>9 then error("<li>�Ƿ�����")
if Searchxm<>"UserName" and Searchxm<>"LastUserName" then error("<li>非法操作")
item=""&Searchxm&"='"&SqlString(content)&"'"
elseif Search="key" then
item="Topic like '%"&SqlLikeString(content)&"%'"
else
error("<li>非法操作")
end if

if TimeLimit>0 then TimeLimitList="and lasttime>"&SqlNowString&"-"&TimeLimit&""


sql="select * from [BBSXP_Threads] where IsDel=0 and "&ForumIDor&" "&item&" "&TimeLimitList&" order by lasttime Desc"
Rs.Open sql,Conn,1

count=Rs.recordcount    '����������
if Count=0 then error("<li>�Բ���û���ҵ���Ҫ��ѯ������")

PageSetup=SiteSettings("ThreadsPerPage") '�趨ÿҳ����ʾ����
Rs.Pagesize=PageSetup
TotalPage=Rs.Pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��


%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� �������</td>
	</tr>
</table>
<br>
<table width="100%" align="center">
	<tr>
		<td width="100%" align="right">�����ؼ��ʣ�<font color="FF0000"><%=content%></font>�������ҵ��� 
		<b><font color="FF0000"><%=Count%></font></b> ƪ�������</td>
	</tr>
</table>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
<tr height="25" id="TableTitleLink" class="a1">
<td align="center" colspan="3">����</td>
<td align="center" width="10%">����</td>
<td align="center" width="6%">�ظ�</td>
<td align="center" width="6%">���</td>
<td align="center" width="25%">������</td>
</tr>
	<%
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
ShowThread()
Rs.MoveNext
loop
Rs.Close
%>
</table>

<table cellspacing="0" cellpadding="1" width="100%" align="center" border="0">
	<tr height="25">
		<td width="100%" height="2">
		<table cellspacing="0" cellpadding="3" width="100%" border="0">
			<tr>
				<td height="2">
				<%
ShowPage()%>
</td>
				<form name="form" action="Search.asp?menu=ok&ForumID=<%=ForumID%>&Search=key" method="POST"><input type=hidden name=sessionid value=<%=session.sessionid%>>
					<td height="2" align="right">����������<input name="content" size="20" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="�ؼ���" Helptext="�ؼ���">
					<input type="submit" value="����" id="btnadd" disabled>
					</td>
				</form>
			</tr>
		</table>
		</td>
	</tr>
</table>

<%
htmlend


end if
%>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		�� ��������</td>
	</tr>
</table>
<br>

<table height="207" cellspacing="1" cellpadding="3" width="100%" class="a2" border="0" align="center">
	<form method="POST" action="Search.asp?menu=ok" name="form"><input type=hidden name=sessionid value=<%=session.sessionid%>>
		<tr>
			<td colspan="2" height="25" class="a1" align="center">������Ҫ�����Ĺؼ���</td>
		</tr>
		<tr class=a3>
			<td valign="top" colspan="2" height="8">
			<p align="center">
			<input size="40" name="content" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')"></p>
			</td>
		</tr>
		<tr>
			<td class="a1" colspan="2" height="25" align="center">����ѡ��</td>
		</tr>
		<tr class=a3>
			<td width="41%" height="24" align="right">�ؼ�������<input type="radio" value="key" name="Search" checked></td>
			<td height="25" width="58%">&nbsp;<select size=1>
			<option>�������������ؼ���</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="21" align="right">��������<input type="radio" value="author" name="Search" id="Search"></td>
			<td height="25" width="58%">&nbsp;<select size="1" name="Searchxm">
			<option selected value="UserName">������������</option>
			<option value="lastname">�������ظ�����</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="23" align="right">���ڷ�Χ</td>
			<td height="25" width="58%">&nbsp;<select size="1" name="TimeLimit">
			<option value="">��������</option>
			<option value="1">��������</option>
			<option value="5" selected>5������</option>
			<option value="10">10������</option>
			<option value="30">30������</option>
			</select></td>
		</tr>
		<tr class=a3>
			<td width="41%" height="26" align="right">��ѡ��Ҫ��������̳</td>
			<td height="26" width="58%">&nbsp;<select name="ForumID" size="1">
			<option value="" selected>ȫ����̳</option>
			<%BBSList(0)%><%=ForumsList%></select>
			<input type="submit" value="��ʼ����" id="btnadd" disabled></td>
		</tr>
	</form>
</table>

</center><%
htmlend
%>