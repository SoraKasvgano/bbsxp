<!-- #include file="Setup.asp" -->
<%

HtmlTop
if CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")
ThreadID=RequestInt("ThreadID")
PostID=RequestInt("PostID")
Category=HTMLEncode(Request("Category"))

sql="Select * From ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
Rs.Open sql,Conn,1
	if Rs.eof or Rs.bof then error"<li>ϵͳ�����ڸ����ӵ�����"
	ForumID=Rs("ForumID")
	ThreadStyle=Rs("ThreadStyle")
	Topic=Rs("Topic")
	Category=Rs("Category")
	ThreadEmoticonID=Rs("ThreadEmoticonID")
	ThreadTop=Rs("ThreadTop")
	StickyDate=Rs("StickyDate")
	IsVote=Rs("IsVote")
	if ThreadTop>0 then StickyDay=DateDiff("d",now(),StickyDate)
Rs.close


%><!-- #include file="Utility/ForumPermissions.asp" --><%

sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
Set Rs=Execute(sql)
	if Rs.eof or Rs.bof then error"<li>ϵͳ�����ڸ����ӵ�����"
	if Rs("PostAuthor")<>CookieUserName and PermissionManage=0 then error("�Բ�������Ȩ�޲�����")	
	if SiteConfig("PostEditBodyAgeInMinutes")>0 and DateDiff("n", Rs("PostDate"), now())> SiteConfig("PostEditBodyAgeInMinutes") then error"<li>���ӷ��� "&SiteConfig("PostEditBodyAgeInMinutes")&" ����֮�������ٴα��༭"
	Subject=ReplaceText(""&Rs("Subject")&"","<[^>]*>","")
	Body=replace(Rs("Body"),"<br />",vbcrlf)
	PostParentID=Rs("ParentID")
Rs.close



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request_Method = "POST" then
	Subject=HTMLEncode(Request.Form("PostSubject"))
	Body=BodyEncode(Request.Form("Body"))
	Description=Left(HTMLEncode(Request.Form("Description")),200)

	Category=HTMLEncode(Request.Form("Category"))
	Tags=Request.Form("Tags")
	EditNotes=HTMLEncode(Request("EditNotes"))
	ThreadEmoticonID=RequestInt("ThreadEmoticonID")
	if PermissionManage=1 then ThreadStyle=HTMLEncode(Request.Form("ThreadStyle"))


	if Request.Form("DisableBBCode")=1 then Body=Replace(Body,CHR(91),"&#91")
	if PostParentID=0 then
		if Len(Subject)<2 then Message=Message&"<li>�������ⲻ��С�� 2 �ַ�"
		
		'vote submit start
		if PermissionCreatePoll=1 and RequestInt("IsVote")=1 then
			j=0
			for each formElement in request.form
				if instr(formElement,"Bbsxp_Input_")>0 then
					if trim(request.form(""&formElement&""))<>"" then
						allpollTopic=""&allpollTopic&""&request.form(""&formElement&"")&"|"
						allpollResult=""&allpollResult&""&RequestInt(""&replace(formElement,"Bbsxp_Input_","Bbsxp_Result_")&"")&"|"
						j=j+1
					end if
				end if
			next
			if j<SiteConfig("MinVoteOptions") then Message=Message&"<li>ͶƱѡ������� "&SiteConfig("MinVoteOptions")&" ��"
			if j>SiteConfig("MaxVoteOptions") then Message=Message&"<li>ͶƱѡ��ܳ��� "&SiteConfig("MaxVoteOptions")&" ��"
		end if

		IsMultiplePoll=RequestInt("IsMultiplePoll")
		VoteItems=HTMLEncode(allpollTopic)
		VoteExpiry=now()+RequestInt("VoteExpiry")
		'vote submit end
		
		if Not Isdate(VoteExpiry) then Message=Message&"<li>ͶƱ����ʱ�����"
		
	end if


	if Len(Body)<2 then Message=Message&"<li>�������ݲ���С�� 2 �ַ�"
	if SiteConfig("RequireEditNotes")=1 and ""&EditNotes&""="" then Message=Message&"<li>���������ӱ༭ԭ��"
	
	TagArray=split(Tags,",")
	if Ubound(TagArray)>5 then Message=Message&"<li>��ǩ���ܳ���5��"

	if Message<>"" then error(""&Message&"")
	
	sql="Select * from ["&TablePrefix&"Posts] where PostID="&PostID&""
	Rs.Open sql,Conn,1,3
		if Rs("ParentID")=0 then 
			Execute("update ["&TablePrefix&"Threads] Set ThreadStyle='"&ThreadStyle&"',Topic='"&Subject&"',Description='"&Description&"',Category='"&Category&"',ThreadEmoticonID="&ThreadEmoticonID&" where ThreadID="&Rs("ThreadID")&"")
			if PermissionCreatePoll=1 and RequestInt("IsVote")=1 then
				Execute("Update ["&TablePrefix&"Votes] set IsMultiplePoll="&IsMultiplePoll&",Items='"&VoteItems&"',Result='"&allpollResult&"',Expiry='"&VoteExpiry&"' where ThreadID="&Rs("ThreadID")&"")
			end if
		end if
		Rs("Subject")=Subject
		Rs("Body")=Body
	Rs.update
	Rs.close
	
	
	SQL="Select * from ["&TablePrefix&"PostEditNotes] where PostID="&PostID&""
	Rs.Open sql,Conn,1,3
	if Rs.eof then Rs.addNew
	Rs("PostID")=PostID
	Rs("EditNotes")="��"&CookieUserName&" �� "&now()&" �༭����"&vbCrlf&""&EditNotes&""
	Rs.update
	Rs.close
	
	
	
	if SiteConfig("DisplayPostTags")=1 then AddTags()
	
	
	
	if Request.Form("UpFileID")<>"" then
		UpFileID=split(Request.Form("UpFileID"),",")
		for i = 0 to ubound(UpFileID)-1
			Execute("update ["&TablePrefix&"PostAttachments] Set Description='"&Subject&"',PostID='"&PostID&"' where UpFileID="&int(UpFileID(i))&" and UserName='"&CookieUserName&"'")
		next
	end if
	Message="<li>�޸����ӳɹ�<li><a href=ShowPost.asp?ThreadID="&ThreadID&">��������</a><li><a href=ShowForum.asp?ForumID="&	ForumID&">������̳</a>"
	succeed Message,"ShowForum.asp?ForumID="&ForumID&""
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <%=ForumTree(ParentID)%> <a href=ShowForum.asp?ForumID=<%=ForumID%>><%=ForumName%></a> �� <a href="ShowPost.asp?ThreadID=<%=ThreadID%>"><%=Topic%></a> �� �༭����</div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
<form name="form" method="post" onSubmit="return CheckForm(this);">
<%if SiteConfig("RequireEditNotes")=1 then%>
<input name="RequireEditNotes" type="hidden" value='<%=SiteConfig("RequireEditNotes")%>' />
<%end if%>
<input name="Body" type="hidden" value='<%=server.htmlencode(Body)%>' />
<%	if PostParentID=0 then%><input name="Description" type="hidden" /><%end if%>
<input name="UpFileID" type="hidden" />
<input name="ThreadStyle" id="ThreadStyle" type="hidden" value='<%=ThreadStyle%>' />

	<tr class=CommonListTitle>
		<td colspan=2>�༭����</td>
	</tr>
	<tr class="CommonListCell">
		<td width="180"><b>����</b><%if PermissionManage=1 and PostParentID=0 then%>��<a href="javascript:BBSXP_Modal.Open('Utility/SelectStyle.htm',500,420);">����</a>��<%end if%></td>
		<td><input type="text" size="60" name="PostSubject" value='<%=Subject%>' <%if PostParentID=0 then%>id="Subject" style="<%=ThreadStyle%>"<%end if%> /></td>
	</tr>

<%	if PostParentID=0 then%>
	<tr class="CommonListCell">
		<td><b>���</b></td>
		<td>
		<select name=Category size=1>
		<option value="" selected="selected">��</option>
		<option value="<%=Category%>" selected="selected"><%=Category%></option>
<%
		if TotalCategorys<>empty then
			filtrate=split(TotalCategorys,"|")
			for i = 0 to ubound(filtrate)
				response.write "<OPTION value='"&filtrate(i)&"'>"&filtrate(i)&"</OPTION>"
			next
		end if
%>
		</select> <a href="javascript:BBSXP_Modal.Open('Utility/AddCategory.htm', 500, 100);">������</a>
		</td>
	</tr>

	<tr class="CommonListCell">
		<td valign=top align=Left><b>����</b></td>
		<td><input type=radio value=0 name=ThreadEmoticonID checked="checked" />��<br />
		<script language="JavaScript" type="text/javascript">
		for(i=1;i<=30;i++) {
			if (i ==<%=ThreadEmoticonID%>){IsChecked="checked"}else{IsChecked=""}
			document.write("<input type=radio value="+i+" name=ThreadEmoticonID "+IsChecked+"  /><img src=images/Emoticons/"+i+".gif width=23 height=23 />��")
			if (i ==10 || i ==20){document.write("<br />")}
		}
		</script></td>
	</tr>
	<%
	if IsVote=1 then
		Set Rs=Execute("Select * from ["&TablePrefix&"Votes] where ThreadID="&ThreadID&"")
		if Rs.eof then
			Execute("Update ["&TablePrefix&"Threads] set IsVote=0 where ThreadID="&ThreadID&"")
		else
			IsMultiplePoll=Rs("IsMultiplePoll")
			Expiry=Rs("Expiry")
			Items=Rs("Items")
			Result=Rs("Result")
			
			ExpiryDay=DateDiff("d",now(),Expiry)
			ItemsArray=split(Items,"|")
			ResultArray=split(Result,"|")
		end if
		Rs.close
	%>
	<tr class="CommonListCell">
		<td valign=top align=Left><b>ͶƱ</b><br />ÿ��һ��ͶƱ��Ŀ<br /><input type=radio<%if IsMultiplePoll=0 then%> checked="checked"<%end if%> value=0 name=IsMultiplePoll id=IsMultiplePoll /><label for=IsMultiplePoll>��ѡͶƱ</label><br /><input type=radio<%if IsMultiplePoll=1 then%> checked="checked"<%end if%> value=1 name=IsMultiplePoll id=IsMultiplePoll_1 /><label for=IsMultiplePoll_1>��ѡͶƱ</label></font> <br />����ʱ�� <input type="text" size="2" name="VoteExpiry" value="<%=ExpiryDay%>" onkeyup=if(isNaN(this.value))this.value='<%=ExpiryDay%>' /> ���
		</td>
		<td valign="top">
		<script language="javascript" type="text/javascript" src="Utility/LabelDom.js"></script>
		<div id="VoteOptionList">
		<%
		if IsArray(ItemsArray) then
			Response.Write("<input type=hidden name=IsVote value=1 />")
			j=0
			for i=0 to Ubound(ItemsArray)
				if ItemsArray(i)<>"" then ItemsStr=ItemsStr&"<div id=""Bbsxp_Div_"&i&"""><input type=""text"" size=""50"" name=""Bbsxp_Input_"&i&""" id=""Bbsxp_Input_"&i&""" value="""&ItemsArray(i)&""" />��Ʊ��:<input type=""text"" size=""5"" name=""Bbsxp_Result_"&i&""" value="""&ResultArray(i)&""" /></div>":j=j+1
			next
		end if
		response.Write(ItemsStr)
		%>
		</div>
		<a class="CommonTextButton" href="javascript:AddLabel('VoteOptionList')">�����Ŀ</a>
		<script language="javascript" type="text/javascript">
			var MinCount=parseInt('<%=SiteConfig("MinVoteOptions")%>');
			var MaxCount=parseInt('<%=SiteConfig("MaxVoteOptions")%>');
			var CurrentCount=parseInt('<%=j%>');
			var RealCount=parseInt('<%=j%>');
		</script>
		</td>
	</tr>
<%
	end if
end if
if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
	<tr class="CommonListCell">
		<td valign=top><b>�ϴ�����</b></td>
		<td><iframe id="UpLoadIframe" name="UpLoadIframe" src="UploadAttachment.asp" frameborder="0" width="100%" height="20" scrolling="no"></iframe></td>
	</tr>
<%end if%>

	<tr class="CommonListCell">
		<td valign=top>
			<br /><b>����</b><br />��<a href="javascript:CheckLength();">�鿴���ݳ���</a>��<br /><br /><input id=DisableBBCode name=DisableBBCode type=checkbox value=1 /><label for=DisableBBCode> ���� BB ����</label>
            
            
		<%if PostParentID=0 and PermissionManage=1 then%><br /><br />
		<select name=StickyDate size=1>
			<option value="0"<%if StickyDay=0 then response.write("selected")%>>�ö�ѡ��</option>
			<option value="1"<%if StickyDay=1 then response.write("selected")%>>1 ��</option>
			<option value="3"<%if StickyDay>1 and StickyDay<=3 then response.write("selected")%>>3 ��</option>
			<option value="7"<%if StickyDay>3 and StickyDay<=7 then response.write("selected")%>>1 ��</option>
			<option value="14"<%if StickyDay>7 and StickyDay<=14 then response.write("selected")%>>2 ��</option>
			<option value="30"<%if StickyDay>14 and StickyDay<=30 then response.write("selected")%>>1 ��</option>
			<option value="90"<%if StickyDay>30 and StickyDay<=90 then response.write("selected")%>>3 ��</option>
			<option value="180"<%if StickyDay>90 and StickyDay<=180 then response.write("selected")%>>6 ��</option>
			<option value="366"<%if StickyDay>180 and StickyDay<=366 then response.write("selected")%>>1 ��</option>
			<%if BestRole=1 then%><option value="999"<%if StickyDay>366 then response.write("selected")%>>����</option><%end if%>
		</select>
		<%end if%>
            
			<%if SiteConfig("UpFileOption")<>empty and PermissionAttachment=1 then%>
						<br /><br /><span id=UpFile></span>
			<%end if%>
		</td>
		<td height=250><script type="text/javascript" src="Editor/Post.js"></script></td>
	</tr>
<%
if SiteConfig("DisplayPostTags")=1 then
%>
	<tr class="CommonListCell">
		<td><b>��ǩ<br /></b>�Զ��š�,���ָ�</td>
		<td><input type="text" name="Tags" size="80" value="<%=ShowTags(PostID)%>" /> <a href="javascript:BBSXP_Modal.Open('Tags.asp?menu=SelectTags',500,420);" class="CommonTextButton">ѡ���ǩ</a></td>
	</tr>
<%end if



%>
	<tr class="CommonListCell">
		<td><b>�༭ԭ��<br /></b></td>
		<td><input type="text" name="EditNotes" size="80" /></td>
	</tr>
<%
	Set Rs=Execute("Select * from ["&TablePrefix&"PostEditNotes] where PostID="&PostID&" order by EditNoteID desc")
	If Not Rs.eof Then EditNotesRecordset=Rs("EditNotes")
	Rs.close
%>
	<tr class="CommonListCell">
		<td><b>�༭��¼��<br /></b></td>
		<td><textarea name="EditNotesRecordset" cols="78" rows="5" readonly="readonly"><%=EditNotesRecordset%></textarea></td>
	</tr>
<%


%>
	<tr class="CommonListCell">
		<td align=center colspan=2><input type=submit value=" ���� " name=EditSubmit />��<input type="Button" value=" Ԥ�� " onClick="Preview()" />��<input onClick="history.back()" type="button" value=" ȡ�� " />��<input type="button" id="recoverdata" onclick="RestoreData()" title="�ָ��ϴ��Զ����������" value="�ָ�����" /></td>
	</tr>
</form>
</table>
<div name="Preview" id="Preview"></div>
<%
HtmlBottom
%>