<!-- #include file="Setup.asp" --><%
HtmlTop

ForumID=RequestInt("ForumID")
DateComparer=RequestInt("DateComparer")

if ForumID>0 then
	Set Rs=Execute("Select Moderated from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")
	If Not Rs.eof Then Moderated=Rs("Moderated")
	Rs.Close
	%><!-- #include file="Utility/ForumPermissions.asp" --><%
elseif BestRole=1 then
	PermissionManage=1
end if


if Request("menu")="Result" then
	Keywords=HTMLEncode(Request("Keywords"))

	SortBy=HTMLEncode(Request("SortBy"))
	Item=HTMLEncode(Request("Item"))

	if Keywords="" then error("��û�������κβ�ѯ������")
	if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("��֤�����")

	SQLSearch="Visible=1 and "&Item&" like '%"&Keywords&"%' "
	
	if DateComparer > 0 then SQLSearch=SQLSearch&" and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&DateComparer&" "

	if ForumID > 0 then SQLSearch=SQLSearch&" and ForumID="&ForumID&" "


	sql="Select * from ["&TablePrefix&"Threads] where "&SQLSearch&" order by ThreadID "&SortBy&""
	Rs.Open sql,Conn,1
		count=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where "&SQLSearch&"")(0)    '����������
		if Count=0 then error("�Բ���û���ҵ���Ҫ��ѯ������")
		PageSetup=SiteConfig("ThreadsPerPage") '�趨ÿҳ����ʾ����
		Rs.Pagesize=PageSetup
		TotalPage=Rs.Pagecount  '��ҳ��
		PageCount = RequestInt("PageIndex")
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �������</div>

<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td width="100%" align="right">���ҵ��� <b><font color="FF0000"><%=Count%></font></b> ƪ�������</td>
	</tr>
</table>

<%if PermissionManage=1 then response.write("<form style='margin:0px' method=POST action=Manage.asp><input type='hidden' name='ForumID' value='"&ForumID&"' />")%>

<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="5">����</td>
        <%if PermissionManage=1 then response.write("<td width=50 align=center><input type=checkbox name=chkall onclick='CheckAll(this.form)' value=ON /></td>")%>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>����</td>
		<td>����</td>
		<td>�ظ�</td>
		<td>�鿴</td>
		<td>������</td>
        <%if PermissionManage=1 then response.write("<td width=50 align=center></td>")%>
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

<%if PermissionManage=1 then%>
<table cellspacing=0 cellpadding=5 border=0 width=100%>
	<tr>
		<td align=right>
		����ѡ�<select name=menu size=1>
        	<%if BestRole = 1 then%>
			<optgroup label="����">
				<option value="Top">��Ϊ����</option>
				<option value="UnTop">ȡ������</option>
			</optgroup>
            <%end if%>
			<optgroup label="����">
				<option value="IsLocked">��������</option>
				<option value="DelIsLocked">��������</option>
			</optgroup>
			<optgroup label="����">
				<option value="IsGood">��������</option>
				<option value="DelIsGood">ȡ������</option>
			</optgroup>
			<optgroup label="���">
				<option value="Visible">���ͨ��</option>
				<option value="InVisible">���ʧ��</option>
			</optgroup>
			<optgroup label="�ö�">
				<option value="ThreadTop">�ö�����</option>
				<option value="DelTop">ȡ���ö�</option>
			</optgroup>
			<optgroup label="����">
				<option value="DelThread">ɾ������</option>
				<option value="Fix">�޸�����</option>
				<option value="MoveNew">��ǰ����</option>
				<option value="MoveThread">�ƶ�����</option>
			</optgroup>
		</select>��<input type="submit" value=" ִ �� " onclick="return VerifyRadio('Item');" />
		</td>
</table>
<%
	response.write("</form>")
end if
%>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td><%ShowPage()%></td>
	</tr>
</table>

<script language="JavaScript" type="text/javascript">
function VerifyRadio() {
	objYN=false;
	if (window.confirm('��ȷ��ִ�б��β���?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
		if (objYN==false) {alert ('��ѡ����Ҫ���������⣡');return false;}
	}
	return objYN;
}
</script>
<%
	HtmlBottom
end if
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ��������</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<form method="POST" action="Search.asp?menu=Result" name="form" onsubmit="return doSearch();">
		<tr class=CommonListTitle>
			<td colspan="2">����ѡ��</td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">�� �� �ʣ�</td>
			<td><input size="40" name="Keywords" onkeyup="ValidateTextboxAdd(this, 'btnadd')" /></td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">������ʽ��</td>
			<td> 
			<input type="radio" name="searchMethod" value="BBS" id="BBS" checked="checked" /><label for="BBS">վ��</label> 
			<input type="radio" name="searchMethod" value="DuoCi" id="DuoCi" /><label for="DuoCi">DuoCi����</label> </td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">�������ݣ�</td>
			<td> 
			<input type="radio" name="Item" value="Topic" id="Topic" checked="checked" /><label for="Topic">����</label> 
			<input type="radio" name="Item" value="Category" id="Category" /><label for="Category">���</label> 
			<input type="radio" name="Item" value="PostAuthor" id="PostAuthor" /><label for="PostAuthor">����</label> 
			<input type="radio" name="Item" value="LastName" id="LastName" /><label for="LastName">�����µ�����</label> 
		</tr>
		<tr class="CommonListCell">
			<td align="right">���ڷ�Χ��</td>
			<td>
			<select size="1" name="DateComparer">
				<option value="1" >��������</option>
				<option value="2" >2 ������</option>
				<option value="7" >1 ������</option>
				<option value="10" >10 ������</option>
				<option value="14" >2 ������</option>
				<option value="30" >1 ��������</option>
				<option value="45" >45 ������</option>
				<option value="60" >2 ��������</option>
				<option value="75" >75 ������</option>
				<option value="100" >100 ������</option>
				<option value="365" selected="selected">1 ������</option>
				<option value="-1">�κ�ʱ��</option>
				</select>
</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">����ʽ��</td>
			<td>
				<select  name="SortBy">
					<option value="">�Ӿɵ���</option>
					<option value="Desc" selected="">���µ���</option>
				</select>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">������̳��</td>
			<td>
				<select name="ForumID">
				<option value="0">ȫ����̳</option>
				<%=GroupList(ForumID)%>
				</select>
			</td>
		</tr>
		
	<tr class="CommonListCell">
		<td align="right">�� ֤ �룺</td>
		<td><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">���������ȡ��֤��</span> <span id="CheckVerifyCode" style="color:red"></span></td>
	</tr>
		<tr class="CommonListCell">
			<td colspan="2" align="center"><input type="submit" value="��ʼ����" /></td>
		</tr>
	</form>
</table>
<script language="javascript" type="text/javascript">
function doSearch() {
	searchMethod = document.form.searchMethod[1].checked;
	Keyword = document.form.Keywords.value;
	if (Keyword == "") {
		alert(�����������ؼ���);
		document.form.Keywords.focus();
		return false;
	}
	Keyword = "site:<%=Server_Name%> "+Keyword;
	if (searchMethod == true) {
		window.open("http://www.duoci.com/Search/?Charset=<%=BBSxpCharset%>&word="+Keyword);
		return false
	}
	else {
		return true;
	}
}

function VerifyRadio() {
	objYN=false;
	if (window.confirm('��ȷ��ִ�б��β���?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
	
		if (objYN==false) {alert ('��ѡ����Ҫ���������⣡');return false;}
	}
}
</script>
<%
HtmlBottom
%>