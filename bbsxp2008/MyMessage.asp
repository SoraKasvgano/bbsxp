<!-- #include file="Setup.asp" --><%

RecipientUserName=HTMLEncode(Request("RecipientUserName"))
UserName=HTMLEncode(Request("UserName"))
Subject=HTMLEncode(Request("Subject"))
Body=HTMLEncode(Request("Body"))
box=HTMLEncode(Request("box"))
MessageID=RequestInt("MessageID")
ForumID=RequestInt("ForumID")

MaxPrivateMessageSize=SiteConfig("MaxPrivateMessageSize")

if Request("menu")="Post" then
	if CookieUserName=empty then AlertForModal("����δ��¼��̳")

	if MessageID > 0 then
		Rs.open "Select top 1 * from ["&TablePrefix&"PrivateMessages] where MessageID="&MessageID&"",conn,1
		if rs.eof then AlertForModal("����/�ظ��Ķ���Ϣ������")
		Subject=Rs("Subject")
		Body=replace(Rs("Body"),"<br>",vbcrlf)

		if RecipientUserName<>empty then	'�ظ�
			Subject="RE: "&Subject
			Body=""
		else					'ת��
			Body=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"------------------------------------"&vbcrlf&"����: "&Rs("CreateTime")&vbcrlf&"������: "&Rs("SenderUserName")&vbcrlf&"�ռ���: "&Rs("RecipientUserName")&vbcrlf&"����: "&Subject&vbcrlf&vbcrlf&Body&""
			Subject="FW: "&Subject
		end if
		Rs.close
	end if
Response.clear
%>
<title>����ѶϢ</title>
<style type="text/css">body{FONT-SIZE:9pt;}</style>
<form name="form" action="MyMessage.asp" method="POST" onsubmit="return CheckMessage()">
<input type="hidden" name="menu" value="addPost" />
<input type="hidden" name="ForumID" value="<%=ForumID%>" />
�����ˣ�<input type="text" name="SendUserName" value="<%=CookieUserName%>" readonly="readonly" /><br />
�ռ��ˣ�<input id="RecipientUserName" name="RecipientUserName" type="text" value="<%=RecipientUserName%>" />

<%
if ForumID > 0 then
	if RequestCookies("LastReportTime")<>"" then
		ReportInterval=SiteConfig("PostInterval")*3
		StopPostTime=int(DateDiff("s",RequestCookies("LastReportTime"),Now()))
		if StopPostTime < ReportInterval then AlertForModal("��̳����һ���������������α������������ "&ReportInterval&" �룡�������ٵȴ� "&ReportInterval-StopPostTime&" �룡")
	end if
%>
<select onchange="Javascript:document.getElementById('RecipientUserName').value=this.value">
<%
	ModeratedList=Execute("Select Moderated from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")(0)
	if ModeratedList<>"" then
		Response.Write("<option value=''>�����б�</option>")
		Moderated=split(ModeratedList,"|")
		for i=0 to Ubound(Moderated)
			Response.Write("<option value="&Moderated(i)&">"&Moderated(i)&"</option>")
		next
	else
		Response.Write("<option value=''>����Ա�б�</option>")
		ModeratedList=(0)
		Set Rs=Execute("Select UserName from ["&TablePrefix&"Users] where (UserRoleID=1 or UserRoleID=2) and UserAccountStatus=1")
		do while not Rs.eof
			Response.Write("<option value="&Rs("UserName")&">"&Rs("UserName")&"</option>")
			Rs.movenext
		loop
	end if
%>
</select>
<%
end if
%><br />
�����⣺<input size="60" name="Subject" value="<%=Subject%>" /><br />
<textarea name="Body" style="width:100%;height:200px"><%=Body%></textarea><br />


<input type="checkbox" name="IsSaveOutbox" id="IsSaveOutbox" value="1" /><label for=IsSaveOutbox>���浽�ѷ���</label>

<div style="float:right"><input type="submit" value="����ѶϢ" />��</div>
</form>

<script language="JavaScript" type="text/javascript">
function CheckMessage(){
	var FormObject = document.form;
	if (FormObject.elements["RecipientUserName"].value == "")
	{
		alert("�������ռ������֣�");
		FormObject.elements["RecipientUserName"].focus()
		return false;
	}
	if ( FormObject.elements["Subject"].value.length < 2)
	{
		alert("���������ⲻ������2���ֽڣ�");
		FormObject.elements["Subject"].focus()
		return false;
	}
	if (FormObject.elements["Body"].value.length < 2)
	{
		alert("���������ݲ�������2���ֽڣ�");
		FormObject.elements["Body"].focus()
		return false;
	}
	if (FormObject.elements["RecipientUserName"].value == FormObject.elements["SendUserName"].value)
	{
		alert("�����Լ����Լ�����ѶϢ��");
		FormObject.elements["RecipientUserName"].focus()
		return false;
	}
	return true;
}
</script>
<%


elseif Request("menu")="addPost" then
	if CookieUserName=empty then AlertForModal("����δ��¼��̳")

	IsSaveOutbox=RequestInt("IsSaveOutbox")
	IsSenderDelete=0
	if IsSaveOutbox=0 then IsSenderDelete=1
	if Len(Subject)<2 then Alert("���ⲻ��С�� 2 �ַ�")
	if Len(Body)<2 then Alert("���ݲ���С�� 2 �ַ�")
	if RecipientUserName="" or Lcase(RecipientUserName)=Lcase(CookieUserName) then Alert("��������յĶ����Ҳ������Լ�")
	if Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&RecipientUserName&"'").eof Then Alert("ϵͳ������"&RecipientUserName&"������")
	
	if IsSaveOutbox=1 and Execute("Select count(MessageID) From ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 or SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0) > MaxPrivateMessageSize Then
		AddApplication "Message_"&CookieUserName,"��ϵͳѶϢ�����Ķ�Ѷ�ļ��пռ��������������ļ��У������������ܱ����κζ�Ѷ��"
		Alert("���Ķ������������ܱ����Ѷ�����η���ʧ�ܣ�")
	end if

	if Execute("Select count(MessageID) From ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&RecipientUserName&"' and IsRecipientDelete=0 or SenderUserName='"&RecipientUserName&"' and IsSenderDelete=0")(0)>MaxPrivateMessageSize Then
		AddApplication "Message_"&RecipientUserName,"��ϵͳѶϢ�����Ķ�Ѷ�ļ��пռ��������������ļ��У����������ղ����κζ�Ѷ��"
		Alert("�û�"&RecipientUserName&"�Ķ��������������η���ʧ�ܣ�")
	end if
	
	sql="insert into ["&TablePrefix&"PrivateMessages](SenderUserName,RecipientUserName,Subject,Body,IsSenderDelete) values ('"&CookieUserName&"','"&RecipientUserName&"','"&Subject&"','"&Body&"',"&IsSenderDelete&")"
	Execute(sql)
	Execute("update ["&TablePrefix&"Users] Set NewMessage=NewMessage+1 where UserName='"&RecipientUserName&"'")
	if ForumID>0 then ResponseCookies "LastReportTime",""&now()&"","999"
%>
<script language="JavaScript" type="text/javascript">parent.BBSXP_Modal.Close();</script>
<%


else
	if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
	HtmlTop
	RecipientCount=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 or SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0)

UsedMessageSize=RecipientCount/MaxPrivateMessageSize*100

%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� ���ŷ���</div>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td align="center"><a href="EditProfile.asp">�����޸�</a></td>
		<td align="center"><a href="EditProfile.asp?menu=pass">�����޸�</a></td>
		<td align="center"><a href="MyUpFiles.asp">�ϴ�����</a></td>
		<td align="center"><a href="MyFavorites.asp">�� �� ��</a></td>
		<td align="center"><a href="MyMessage.asp">���ŷ���</a></td>
	</tr>
</table>
<br />
<table border="0" width="100%" id="table14" cellspacing="0" cellpadding="0">
	<tr>
		<td width="150" valign="top">
		<table width="100%" cellspacing="1" cellpadding="5" border="0" class="CommonListArea">
			<tr class="CommonListTitle">
				<td><b>�ļ���</b></td>
			</tr>
			<tr class="CommonListCell">
				<td><img src="images/i_inbox.gif" align="absmiddle" />
				<a href="?menu=Inbox">�ռ���</a>[<b><%=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0")(0)%></b>]</td>
			</tr>
			<tr class="CommonListCell">
				<td><img src="images/i_sent.gif" align="absmiddle" />
				<a href="?menu=Outbox">�ѷ���</a>[<b><%=Execute("Select count(messageID) from ["&TablePrefix&"PrivateMessages] where SenderUserName='"&CookieUserName&"' and IsSenderDelete=0")(0)%></b>]</td>
			</tr>
		</table>
		<br />
		<table cellspacing="1" cellpadding="1" width="100%">
			<tr>
				<td>
				<div class="percent">
					<div style="width:<%=UsedMessageSize%>%">
					</div>
				</div>
				<%=UsedMessageSize%>% ��<%=RecipientCount%>/<%=MaxPrivateMessageSize%>��</td>
			</tr>
		</table>
		</td>
		<td width="1%"></td>
		<td valign="top"><%
select case Request("menu")

	case "Read"
		ReadMessage

	case "Delete"
		for each ho in Request("MessageID")
			ho=int(ho)
			rs.open "Select * from ["&TablePrefix&"PrivateMessages] where ( RecipientUserName='"&CookieUserName&"' or SenderUserName='"&CookieUserName&"' ) and MessageID="&ho&"",Conn,1,3
				if LCASE(rs("RecipientUserName"))=LCASE(CookieUserName) then rs("IsRecipientDelete")=1
				if LCASE(rs("SenderUserName"))=LCASE(CookieUserName) then rs("IsSenderDelete")=1
			rs.update
			rs.close
		next
		Execute("Delete from ["&TablePrefix&"PrivateMessages] where IsRecipientDelete=1 and IsSenderDelete=1")

		response.redirect("?")

	case else
		InOutbox
end select
%> </td>
	</tr>
</table>
<br />
<%
Sub InOutbox
MessageMenu=Request("Menu")
if MessageMenu=empty then MessageMenu="Inbox"

	if MessageMenu="Inbox" then
		tdstr="<td align=center width='15%'>������</td>"
		sql="Select * from ["&TablePrefix&"PrivateMessages] where RecipientUserName='"&CookieUserName&"' and IsRecipientDelete=0 order by MessageID Desc"
	else
		tdstr="<td align=center width='15%'>�ռ���</td>"
		sql="Select * from ["&TablePrefix&"PrivateMessages] where SenderUserName='"&CookieUserName&"' and IsSenderDelete=0 order by MessageID Desc"
	end if
%>
<form name="showMessage" action="?Menu=Delete" method="post">
	<table cellspacing="1" cellpadding="3" width="100%" class="CommonListArea">
		<tr class="CommonListTitle">
			<td colspan="5">
			<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?Menu=Post',600,350);">
			<img src="images/i_new_msg.gif" border="0" align="absmiddle" />�½�</a>��<a onclick="if(window.confirm('��ȷ��ִ�б��β���?')){document.showMessage.submit();}"><img src="images/i_delete.gif" border="0" align="absmiddle" />ɾ��</a></td>
		</tr>
		<tr class="CommonListHeader">
			<td align="center" width="6%">
			<input type="checkbox" name="chkall" onclick="CheckAll(this.form)" value="ON" /></td>
			<%=tdstr%>
			<td align="center" width="46%">����</td>
			<td align="center" width="22%">����</td>
			<td align="center" width="10%">��С</td>
		</tr>
		<%
	Rs.Open sql,Conn,1
	
	PageSetup=20 '�趨ÿҳ����ʾ����
	Rs.Pagesize=PageSetup
	TotalPage=Rs.Pagecount  '��ҳ��
	PageCount = RequestInt("PageIndex")
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ�� 
	i=0
	Do While Not Rs.EOF and i<PageSetup
		i=i+1
%>
		<tr class="CommonListCell" id="PrivateMessage<%=Rs("MessageID")%>" <%if rs("isread")=0 then response.write" style='font-weight:bold'"%>>
			<td valign="center" align="center" width="5%">
			<input type="checkbox" value="<%=Rs("MessageID")%>" name="MessageID" onclick="CheckSelected(this.form,this.checked,'PrivateMessage<%=Rs("MessageID")%>')" /></td>
			<%if MessageMenu="Inbox" then%>
			<td align="center">
			<a href="Profile.asp?UserName=<%=Rs("SenderUserName")%>" target="_blank">
			<%=Rs("SenderUserName")%></a></td>
			<%else%>
			<td align="center">
			<a href="Profile.asp?UserName=<%=Rs("RecipientUserName")%>" target="_blank">
			<%=Rs("RecipientUserName")%></a></td>
			<%end if%>
			<td width="10%">
			<a href="?Menu=Read&MessageID=<%=Rs("MessageID")%>&box=<%=MessageMenu%>">
			<%=Rs("Subject")%></a></td>
			<td align="center" width="15%"><%=Rs("CreateTime")%></td>
			<td align="center" width="10%"><%=CheckSize(Len(""&Rs("Body")&""))%></td>
		</tr>
		<%
		Rs.MoveNext
	loop
	Rs.Close
%>
	</table>
	<table cellspacing="1" cellpadding="1" width="100%">
		<tr>
			<td colspan="5"><%ShowPage()%></td>
		</tr>
	</table>
</form>
<%
End Sub

Sub ReadMessage

if box="Inbox" then
	FieldName="RecipientUserName"
	DeleteFieldName="IsRecipientDelete"
else
	FieldName="SenderUserName"
	DeleteFieldName="IsSenderDelete"
end if

if Request("pageIndex")="previous" then
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where "&FieldName&"='"&CookieUserName&"' and "&DeleteFieldName&"=0 and MessageID>"&MessageID&" order by messageID"

elseif Request("pageIndex")="next" then
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where "&FieldName&"='"&CookieUserName&"' and "&DeleteFieldName&"=0 and MessageID<"&MessageID&" order by messageID desc"

else
	SQL="Select top 1 * from ["&TablePrefix&"PrivateMessages] where (RecipientUserName='"&CookieUserName&"' or SenderUserName='"&CookieUserName&"') and MessageID="&MessageID&""
end if


Rs.Open SQL,Conn,1,3

if Http_Referer="" then Http_Referer="Default.asp"
if Rs.eof then response.redirect Http_Referer

%>
<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td>
		<div style="float:left">
			<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post', 600, 350);">
			<img src="images/i_new_msg.gif" border="0" align="absmiddle" />�½�</a>��<a href="?menu=Delete&MessageID=<%=Rs("MessageID")%>" onclick="return window.confirm('��ȷ��ִ�б��β���?');"><img src="images/i_delete.gif" border="0" align="absmiddle" />ɾ��</a><%if Rs("SenderUserName")<>CookieUserName then%>��<img src="images/i_reply.gif" border="0" align="absmiddle" /><a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&MessageID=<%=Rs("MessageID")%>&RecipientUserName=<%=Rs("SenderUserName")%>',600,350);">�ظ�</a><%end if%>��<a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&MessageID=<%=Rs("MessageID")%>',600,350);"><img src="images/i_forward.gif" border="0" align="absmiddle" />ת��</a></div>
		<div style="float:right">
			<a href="?menu=Read&MessageID=<%=Rs("messageID")%>&pageIndex=previous&box=<%=box%>" title="��һ��">
			<img src="images/i_previous.gif" border="0" align="absmiddle" /></a>
			<a href="?menu=Read&MessageID=<%=Rs("messageID")%>&pageIndex=next&box=<%=box%>" title="��һ��">
			<img src="images/i_next.gif" border="0" align="absmiddle" /></a></div>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td width="10%" style="line-height:150%">
		<div style="float:left">
			�����ˣ�<a target="_blank" href="Profile.asp?UserName=<%=Rs("SenderUserName")%>"><%=Rs("SenderUserName")%></a><%if Rs("SenderUserName")<>CookieUserName then%>��<a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("SenderUserName")%>',true);">+��Ӻ���</a><%end if%></div>
		<div style="float:right">
			<%=Rs("CreateTime")%></div>
		<br />
		�ռ��ˣ�<a target="_blank" href="Profile.asp?UserName=<%=Rs("RecipientUserName")%>"><%=Rs("RecipientUserName")%></a><%if Rs("RecipientUserName")<>CookieUserName then%>��<a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("RecipientUserName")%>',true);">+��Ӻ���</a><%end if%>
		<br />
		�����⣺<%=Rs("Subject")%></td>
	</tr>
	<tr class="CommonListCell">
		<td><%=Rs("Body")%></td>
	</tr>
</table>
<%
	if LCASE(Rs("RecipientUserName"))=LCASE(CookieUserName) and Rs("IsRead")=0 then Rs("IsRead")=1
	Rs.update
	Rs.close


End Sub

HtmlBottom

end if
%>