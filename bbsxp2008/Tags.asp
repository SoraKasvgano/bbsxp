<!-- #include file="Setup.asp" -->
<%
if Request("menu")="UpSelectTags" then
	if CookieUserName=empty then AlertForModal("����δ��¼��̳")
	for each Tag in request.form("SelectTags")
		SelectedTags=""&SelectedTags&","&Tag&""
	next
%>
<script type="text/javascript" language="javascript">
	TargetTag = parent.$('Tags');
	TargetTag.value += TargetTag.value =='' ? "<%=Mid(SelectedTags,2)%>" : "<%=SelectedTags%>";
	parent.BBSXP_Modal.Close();
</script>
<%
elseif Request("menu")="SelectTags" then
%>
<title>ѡ���ǩ</title>
<style type="text/css">body,table{FONT-SIZE:9pt;}</style>
<%
	if CookieUserName=empty then AlertForModal("����δ��¼��̳")
	sql="select top 10 * From ["&TablePrefix&"PostTags] where IsEnabled=1 and DateDiff("&SqlChar&"m"&SqlChar&",MostRecentPostDate,"&SqlNowString&")<1 order by TotalPosts Desc"
	Set Rs=Execute(Sql)
	if Rs.eof then AlertForModal("��ʱû�б�ǩ��ѡ��")
	Response.write("<table border=0 cellpadding=2 width='100%'><form name=form action='?' method=post><input type=hidden name=menu value='UpSelectTags' /><tr><td colspan=2><b>���ű�ǩ��</b></td></tr><tr>")
	i=0
	do while not Rs.eof and i<10
		if i mod 2=0 then Response.write("</tr><tr>")
		Response.write("<td width='50%'><label for='Tag_"&i&"' style='cursor:pointer'><input type=checkbox name='SelectTags' id='Tag_"&i&"' value='"&Rs("TagName")&"' />"&Rs("TagName")&"</label></td>")
		i=i+1
		Rs.movenext
	loop
	Rs.close
	Response.write("</tr>")
	CookiePostTags=RequestCookies("CookiePostTags")
	if CookiePostTags<>"" then
		Response.write("<tr><td colspan=2><hr width='100%'></td></tr><tr><td colspan=2><b>���ʹ�õı�ǩ��</b></td></tr>")
		TagsArray=split(CookiePostTags,",")
		i=0
		for j=0 to UBound(TagsArray)
			if TagsArray(j)<>"" and j<10 then
				if i mod 2=0 then Response.write("</tr><tr>")
				Response.write("<td width='50%'><label for='Tag2_"&i&"' style='cursor:pointer'><input type=checkbox name='SelectTags' id='Tag2_"&i&"' value='"&TagsArray(j)&"' />"&TagsArray(j)&"</label></td>")
				i=i+1
			end if
		next
		
		
	end if
	Response.write("<tr><td colspan=2 align=center><input type=submit value=' ȷ�� ' />��<input type=button onclick='parent.BBSXP_Modal.Close();' value=' ȡ�� ' /></td></tr></form></table>")
else

	HtmlHeadTitle="��ǩ�б�"
	HtmlTop
	
	TagID=RequestInt("TagID")
	if TagID>0 then
		ShowTag
	else
		ShowAllTags
	end if


	Sub ShowAllTags
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <a href="Tags.asp">��ǩ�б�</a></div>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class=CommonListTitle>
		<td align=center>��ǩ�б�</td>
	</tr>
	<tr class="CommonListCell">
		<td class="TagList">
<%
		sql="["&TablePrefix&"PostTags] where IsEnabled=1"
		TotalCount=Execute("Select count(TagID) From "&sql&" ")(0) '��ȡ��������
		PageSetup=210 '�趨ÿҳ����ʾ����
		TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
		PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		if PageCount<11 then
			Set Rs=Execute("select * from "&sql&" order by TotalPosts Desc")
		else
			rs.Open "select * from "&sql&" order by TotalPosts Desc",Conn,1
		end if
		if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
		i=0
		do while not Rs.eof and i<PageSetup
			i=i+1
			Response.Write "<li><a href='Tags.asp?TagID="&Rs("TagID")&"'>"&Rs("TagName")&"</a>("&Rs("TotalPosts")&")</li>"
			Rs.movenext
		loop
		Rs.close
%>
		</td>
	</tr>
</table>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="right"><%ShowPage()%></td>
	</tr>
</table>
<%
	End Sub

	Sub ShowTag
		Rs.open "select TagName from ["&TablePrefix&"PostTags] where IsEnabled=1 and TagID="&TagID&"",Conn,1
		if Rs.eof then error("�Ҳ�����Ҫ�鿴�ı�ǩ")
		TagName=Rs("TagName")
		Rs.Close
		
		if TagsStatic(TagID)=0 then Execute("Delete from ["&TablePrefix&"PostTags] where TagID="&TagID&"")

		PageSetup=SiteConfig("ThreadsPerPage") '�趨ÿҳ����ʾ����
		TotalCount=Execute("Select count(TagID) From ["&TablePrefix&"PostInTags] where TagID="&TagID&"")(0) '��ȡ��������
		TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
		PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		Sql="select TagID,PostID from ["&TablePrefix&"PostInTags] where TagID="&TagID&""
		Set Rs=Execute(sql)
		If Rs.Eof then Execute("Delete from ["&TablePrefix&"PostTags] where TagID="&TagID&"")
		If TotalPage>1 then RS.Move (PageCount-1) * pagesetup
		IF Not Rs.Eof then TagGetRow=Rs.GetRows(pagesetup)
		Rs.close
	
		if IsArray(TagGetRow) then
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <a href="Tags.asp">��ǩ�б�</a> �� <a href="?TagID=<%=TagID%>"><%=TagName%></a></div>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="3"><div style="float:left">����</div></td>
	</tr>
	<tr class=CommonListHeader align="center">
		<td width="40%">����</td>
		<td width="10%">����</td>
		<td width="10%">����ʱ��</td>
	</tr>
<%
			for i=0 to Ubound(TagGetRow,2)
			
				Set Rs=Execute("Select * from ["&TablePrefix&"Posts] where PostID="&TagGetRow(1,i)&" and Visible=1")
				if Rs.eof then
					Execute("delete from ["&TablePrefix&"PostInTags] where PostID="&TagGetRow(1,i)&"")
				else
					Subject=Rs("Subject")
					if ""&Subject&""="" then Subject=Left(ReplaceText(""&BBCode(Rs("Body"))&"","<[^>]*>",""),10)
%>
	<tr class="CommonListCell">
		<td><a href="ShowPost.asp?PostID=<%=Rs("PostID")%>" target=_blank><%=Subject%></a></td>
		<td align=center><a href="Profile.asp?UserName=<%=Rs("PostAuthor")%>" target="_blank"><%=Rs("PostAuthor")%></a></td>
		<td align=center><%=Rs("PostDate")%></td>
	</tr>
<%
					Set Rs = Nothing
				end if
			next
%>
</table>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="right"><%ShowPage()%></td>
	</tr>
</table>
<%
		End if
		TagGetRow=null
	End Sub

	HtmlBottom
end if
%>