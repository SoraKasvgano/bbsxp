<!-- #include file="Setup.asp" --><%

ForumID=RequestInt("ForumID")
SortOrder=RequestInt("SortOrder")
TimeLimit=RequestInt("TimeLimit")
GoodTopic=RequestInt("GoodTopic")
VoteTopic=RequestInt("VoteTopic")
order=HTMLEncode(Request("order"))
Category=HTMLEncode(Request("Category"))
if ""&order&""<>"" then
	if Len(order)>12 then error("�Ƿ�����")
	if instr("|threadid|lasttime|posttime|lastvieweddate|IsGood|totalreplies|","|"&lcase(order)&"|")<1 then error("��Ч�Ĳ���ֵ")
end if
if Len(Category)>25 then error("�������̫��")


%><!-- #include file="Utility/ForumPermissions.asp" --><%
HtmlHeadTitle=ForumName

HtmlTop




if TotalCategorys<>"" then
	filtrate=split(TotalCategorys,"|")
	CategorysList="[<a href='ShowForum.asp?ForumID="&ForumID&"'>ȫ��</a>] "
	for i = 0 to ubound(filtrate)
		CategorysList=CategorysList&"[<a href='ShowForum.asp?ForumID="&ForumID&"&Category="&filtrate(i)&"'>"&filtrate(i)&"</a>] "
	next
end if

%>

<meta http-equiv="refresh" content="300">


<div class="CommonBreadCrumbArea">
	<div style="float:left"><%=ClubTree%> �� <%=ForumTree(ParentID)%><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></div>
<%
if Moderated<>empty then
	filtrate=split(Moderated,"|")
	for i = 0 to ubound(filtrate)
		ModeratedList=ModeratedList&"<div class=menuitems><a href=Profile.asp?UserName="&filtrate(i)&">"&filtrate(i)&"</a></div>"
	next
%>
	<div style="float:right"><img src="images/team.gif" /> <a onmouseover="showmenu(event,'<%=ModeratedList%>')" style="cursor:default">��̳����</a></div>
<%end if%>

</div>
<%
sql="Select * From ["&TablePrefix&"Forums] where ParentID="&ForumID&" and SortOrder>0 and IsActive=1 order by SortOrder"
Set Rs=Execute(sql)
	if not Rs.eof then
	%><table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="7"><a href="ShowForum.asp?ForumID=<%=ForumID%>"><%=ForumName%></a></td>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td width="30"></td>
		<td>��̳</td>
		<td width="50">����</td>
		<td width="50">����</td>
		<td width="150">��󷢱�</td>
		<td width="100">����</td>
	</tr>
	<%
	do while not Rs.eof
		ShowForum()
		Rs.Movenext
	loop
	response.write "</table><br />"
	end if
Rs.close


if ForumRules<>"" then
%>

<table cellpadding="8" class="CommonListCell" width="100%" style=" BORDER-RIGHT:#ccc 1px dotted; BORDER-TOP:#ccc 1px dotted; BORDER-LEFT:#ccc 1px dotted; BORDER-BOTTOM:#ccc 1px dotted;">
	<tr>
		<td><strong><font color="#ff0000">�� �� �� �� ��</font></strong><br />
		<%=ForumRules%></td>
	</tr>
</table><br />
<%end if%>

<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="Left"><%if PermissionPost=1 then%>
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/NewPost.gif)" href="AddTopic.asp?ForumID=<%=ForumID%>">
		��������</a> <%end if%> <%if PermissionCreatePoll=1 then%>
		<a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/Poll.gif)" href="AddTopic.asp?ForumID=<%=ForumID%>&Poll=1">
		����ͶƱ</a> <%end if%> </td>
		<td align="right"><%if ModerateNewThread=1 then%>
		<a href="?ForumID=<%=ForumID%>&Order=Visible&SortOrder=1">
		�������(<%=Execute("Select count(ThreadID) From ["&TablePrefix&"Threads] where Visible=0 and ForumID="&ForumID&"")(0)%>)</a> 
		| <%end if%> <a href="?ForumID=<%=ForumID%>&GoodTopic=1">��������</a> | 
        <a href="?ForumID=<%=ForumID%>&VoteTopic=1">ͶƱ����</a> |
		<a href="ShowForumPermissions.asp?ForumID=<%=ForumID%>">��̳Ȩ��</a> |
		<a href="ForumManage.asp?menu=ForumData&ForumID=<%=ForumID%>">������</a> |
<script language="JavaScript" type="text/javascript">
var AllForumNameList=getCookie("ForumNameList");
var NowForumName="<option value='ShowForum.asp?ForumID=<%=ForumID%>'><%=ForumName%></option>";
if(AllForumNameList.indexOf(NowForumName)==-1){document.cookie= "ForumNameList" + "=" + escape(NowForumName+AllForumNameList);}
document.write("<select onchange='location=this.options[this.selectedIndex].value;'><option>������ʵİ��...</option>" + AllForumNameList+ "</select>");
</script>
		</td>
	</tr>
</table>


<%if PermissionManage=1 then%>
<form style="margin:0px;" method="POST" action="Manage.asp">
<input type="hidden" name="ForumID" value="<%=ForumID%>" />
<%end if%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="5">
		<div style="float:left">����</div>
		<div style="float:right"><%=CategorysList%></div>
		</td>
		<%
		if PermissionManage=1 then
			Visible=" and (Visible=1 or Visible=0)"
			response.write("<td width=50 align=center><input type=checkbox name=chkall onclick='CheckAll(this.form)' value=ON></td>")
		else
			Visible=" and Visible=1"
		end if
		%>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>����</td>
		<td>����</td>
		<td>�ظ�</td>
		<td>�鿴</td>
		<td>������</td>
		<%if PermissionManage=1 then response.write("<td width=50>&nbsp;</td>")%>
	</tr>
	<%
	if RequestInt("PageIndex") < 2 then
		sql="["&TablePrefix&"Threads] where (ThreadTop=2 or ThreadTop=1 and ForumID="&ForumID&")"&Visible&" order by ThreadTop DESC"
		Set Rs=Execute(sql)
	
		if Not RS.EOF then
			Do While Not RS.EOF
				ShowThread()
				Rs.MoveNext
			loop
			if PermissionManage=1 then
				response.write("<tr class=CommonListHeader><td colspan=5>�������</td><td width=50>&nbsp;</td></tr>")
			else
				response.write "<tr class=CommonListHeader><td colspan=5>�������</td></tr>"
			end if
		end if

		Rs.Close
	end if

	if Order=empty then Order="lasttime"
	if Category<>"" then SQLCategory="and Category='"&Category&"'"
	if GoodTopic > 0 then SQLGoodTopic="and IsGood=1"
	if VoteTopic > 0 then SQLVoteTopic="and IsVote=1"
			
	if TimeLimit > 0 then SQLTimeLimit="and DateDiff("&SqlChar&"d"&SqlChar&",lasttime,"&SqlNowString&") < "&TimeLimit&""
		
	if SortOrder="1" then
		SqlSortOrder=""
	else
		SqlSortOrder="Desc"
	end if

	topsql="["&TablePrefix&"Threads] where ForumID="&ForumID&Visible&" and ThreadTop=0 "&SQLCategory&" "&SQLTimeLimit&" "&SQLGoodTopic&" "&SQLVoteTopic&""
	TotalCount=Execute("Select count(ThreadID) From "&topsql&" ")(0) '��ȡ��������
	PageSetup=SiteConfig("ThreadsPerPage") '�趨ÿҳ����ʾ����
	TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '��ҳ��
	PageCount = RequestInt("PageIndex") '��ȡ��ǰҳ
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	
	sql="Select * from "&topsql&" order by "&order&" "&SqlSortOrder&""

	if PageCount<11 then
		Set Rs=Execute(sql)
	else
		rs.Open sql,Conn,1
	end if
	if TotalPage>1 then RS.Move (PageCount-1) * pagesetup
	i=0
	Do While Not RS.EOF and i<PageSetup
		i=i+1
		ShowThread()
		Rs.MoveNext
	loop
	Rs.Close
%>
</table>
<%if PermissionManage=1 then%>
<table cellspacing="0" cellpadding="5" border="0" width="100%">
	<tr>
		<td align="right">
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
			<optgroup label="ɾ��">
				<option value="DelThread">ɾ������</option>
				<option value="UnDelThread">ȡ��ɾ��</option>
			</optgroup>
			<optgroup label="����">
				<option value="Fix">�޸�����</option>
				<option value="MoveNew">��ǰ����</option>
				<option value="MoveThread">�ƶ�����</option>
			</optgroup>
		</select>��<input type="submit" value=" ִ �� " onclick="return VerifyRadio('Item');" />
</td></form>
	</tr>
</table>
<%end if%>
<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td>
		<a onmousedown="ToggleMenuOnOff('ForumOption')" class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/ForumSettings.gif)" href="#ForumOption">
		ѡ��</a>
		</td>
		<td align="right"><%ShowPage()%></td>
	</tr>
</table>
<div id="ForumOption" style="display:none;">
	<fieldset>
	<legend>ѡ��</legend>
	<table border="0" width="100%">
		<tr>
			<td valign="top">
			<form name="form" action="ShowForum.asp?ForumID=<%=ForumID%>" method="POST">
				�������<select name="order">
				<option value="">������ʱ��</option>
				<option value="ThreadID">���ⷢ��ʱ��</option>
				<option value="IsGood">��������</option>
				<option value="IsVote">ͶƱ����</option>
				<option value="Topic">����</option>
				<option value="PostAuthor">����</option>
				<option value="TotalViews">�����</option>
				<option value="TotalReplies">�ظ���</option>
				</select> ���� <select name="SortOrder">
				<option value="0" selected="selected">����</option>
				<option value="1">����</option>
				</select> ����<br />
				���ڹ��ˣ�<select name="TimeLimit">
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
				<option value="365">1 ������</option>
				<option value="-1">�κ�ʱ��</option>
				</select><br />
				<input type="submit" value=" Ӧ�� " /></form>
			</td>
			<td valign="top" align="right">�� <b><%if PermissionPost=0 then%>��<%end if%>��</b> 
			�ڴ˰淢��������<br />
			�� <b><%if PermissionReply=0 then%>��<%end if%>��</b> �ڴ˰�ظ�����<br />
			�� <b><%if PermissionEdit=0 then%>��<%end if%>��</b> �ڴ˰��޸������������<br />
			�� <b><%if PermissionDelete=0 then%>��<%end if%>��</b> �ڴ˰�ɾ�������������<br />
			�� <b><%if PermissionCreatePoll=0 then%>��<%end if%>��</b> �ڴ˰淢��ͶƱ<br />
			�� <b><%if PermissionVote=0 then%>��<%end if%>��</b> �ڴ˰����ͶƱ<br />
			�� <b><%if PermissionAttachment=0 then%>��<%end if%>��</b> �ڴ˰��ϴ�����<br />
			�ð�����<b><%if ModerateNewThread=0 then%>��<%end if%>��Ҫ���</b><br />
			�ð�����<b><%if ModerateNewPost=0 then%>��<%end if%>��Ҫ���</b></td>
		</tr>
	</table>
	</fieldset> </div>
	
	
<!-- #include file="Utility/OnLine.asp" -->
<%
if SiteConfig("DisplayForumUsers")=1 then
	ForumIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ForumID="&ForumID&"")(0)
	regForumIDOnline=Execute("Select count(sessionid) from ["&TablePrefix&"UserOnline] where ForumID="&ForumID&" and UserName<>''")(0)
%>
	
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>�û�������Ϣ</td>
	</tr>
	<tr class="CommonListCell">
		<td>
		<img src="images/plus.gif" id="followImg" style="cursor:pointer;" onclick="loadThreadFollow('ForumID=<%=ForumID%>')" /> 
		Ŀǰ��̳������ <b><%=Onlinemany%></b> �ˣ�������̳���� <b><%=ForumIDOnline%></b> �����ߡ�����ע���û� 
		<b><%=regForumIDOnline%></b> �ˣ��ÿ� <b><%=ForumIDOnline-regForumIDOnline%></b> 
		�ˡ�<div style="display:none" id="follow">
			<hr width="90%" size="1" align="left"><span id="followTd" class="UserList">
			<img src="images/loading.gif" />���ڼ���...</span></div>
		</td>
	</tr>
</table>
<br />
<%end if%>
<table border="0" width="90%" align="center">
	<tr>
		<td><img src="images/topic-announce.gif" border="0" align="absmiddle" /> ��������</td>
		<td><img src="images/topic-pinned.gif" border="0" align="absmiddle" /> �ö�����</td>
		<td><img src="images/topic-popular.gif" align="absmiddle" /> ��������</td>
		<td><img src="images/topic-locked.gif" border="0" align="absmiddle" /> ��������</td>
		<td><img src="images/topic-poll.gif" border="0" align="absmiddle" /> ͶƱ����</td>
		<td><img src="images/topic-hot.gif" border="0" title="�ظ����ﵽ <%=SiteConfig("PopularPostThresholdPosts")%> ���ߵ�����ﵽ <%=SiteConfig("PopularPostThresholdViews")%>" align="absmiddle" /> ��������</td>
		<td><img src="images/topic.gif" border="0" align="absmiddle" /> ��ͨ����</td>
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
%>