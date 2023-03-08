<!-- #include file="conn.asp" -->
<%
if Request_Method <> "POST" then Response.End

select case Request("menu")

	case "CloseMsg"
		RemoveApplication("Message_"&CookieUserName)
		Response.write("临时讯息（已阅读）")

	case "Reputation"
		CommentFor=HTMLEncode(unescape(Request.QueryString("CommentFor")))
		CommentBy=HTMLEncode(unescape(Request.QueryString("CommentBy")))
		
		if ""&CommentFor&""="" and ""&CommentBy&""="" then response.End()
		
		if CommentFor<>empty then
			Sql=" from ["&TablePrefix&"Reputation] where CommentFor='"&CommentFor&"'"
			TitleStr="评价人"
			PageUrl="CommentFor="&CommentFor
		elseif CommentBy<>empty then
			Sql=" from ["&TablePrefix&"Reputation] where CommentBy='"&CommentBy&"'"
			TitleStr="被评价人"
			PageUrl="CommentBy="&CommentBy
		end if
		
		PageSetup=5 '设定每页的显示数量
		TotalCount=Execute("Select count(ReputationID)"&Sql)(0)
		TotalPage=Abs(Int(TotalCount/PageSetup*(-1))) '总页数	
		PageCount = RequestInt("PageIndex") '获取当前页
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		TopCount=PageCount*PageSetup


		if TotalCount<1 then
			CommentAreaStr="<table cellspacing=0 cellpadding=0 width='100%' class='PannelBody'><tr><td>暂时没有相关评价</td></tr></table>"
		else
			Set Rs=Execute("select top "&TopCount&" *"&Sql&" order by DateCreated desc")
			If TotalPage>1 then Rs.Move (PageCount-1) * PageSetup
			i=0
			CommentAreaStr="<table cellspacing=0 cellpadding=0 width='100%' class='PannelBody' style='table-layout: fixed'><tr align=center><td width='15%'>&nbsp;</td><td width='60%'>评价内容</td><td width='25%'>"&TitleStr&"</td></tr>"
			do while not Rs.eof and i<PageSetup
				i=i+1
				if Rs("Reputation")>0 then
					ReputationTitle="好评"
					ImgUrl="Reputation_Excellent.gif"
					FontColor="#FF0000"
					ReputationValue="+"&Rs("Reputation")&""
					
				elseif Rs("Reputation")=0 then
					ReputationTitle="中评"
					ImgUrl="Reputation_Average.gif"
					FontColor="#007700"
					ReputationValue="不计分"
				else
					ReputationTitle="差评"
					ImgUrl="Reputation_Poor.gif"
					FontColor=""
					ReputationValue=Rs("Reputation")
				end if

				CommentAreaStr=CommentAreaStr&"<tr>"
				CommentAreaStr=CommentAreaStr&"<td align=center width='15%'><img src='images/"&ImgUrl&"'> <font color='"&FontColor&"'>"&ReputationTitle&"</font><br />( "&ReputationValue&" )</td>"
				CommentAreaStr=CommentAreaStr&"<td width='60%'>"&Rs("Comment")&"　<em><font color=#C0C0C0>"&Rs("DateCreated")&"</em></font></td>"
			if CommentFor<>empty then
				CommentAreaStr=CommentAreaStr&"<td width='25%' align=center><a href='?UserName="&Rs("CommentBy")&"'>"&Rs("CommentBy")&"</a></td>"
			else
				CommentAreaStr=CommentAreaStr&"<td width='25%' align=center><a href='?UserName="&Rs("CommentFor")&"'>"&Rs("CommentFor")&"</a></td>"
			end if
				CommentAreaStr=CommentAreaStr&"</tr>"
				Rs.movenext
			loop
			
			CommentAreaStr=CommentAreaStr&"<tr><td colspan=3 align=right>"&AjaxShowPage(TotalPage,PageCount,"Loading.asp?menu=Reputation&"&PageUrl)&"</td></tr>"
			CommentAreaStr=CommentAreaStr&"</table>"
			Rs.close
			Set Rs = Nothing
		end if
		response.Write(CommentAreaStr)
		
	case "ForumTree"
		GroupID=RequestInt("GroupID")
		ParentID=RequestInt("ParentID")
		QueryStr="ParentID="&ParentID&""
		if GroupID>0 then QueryStr="GroupID="&GroupID&" and ParentID=0"
		sql="Select * From ["&TablePrefix&"Forums] where "&QueryStr&" and IsActive=1 and  SortOrder>0 order by SortOrder"
		Set Rs=Execute(sql)
		do while not rs.eof
			alltree=""&alltree&"<div class=menuitems><a href=ShowForum.asp?ForumID="&rs("ForumID")&">"&rs("ForumName")&"</a></div>"
			rs.Movenext
		loop
		Set Rs = Nothing
		if GroupID>0 then
%>
<a onmouseover="showmenu(event,'<%=alltree%>')" href=Default.asp?GroupID=<%=GroupID%>><%=Execute("Select GroupName From ["&TablePrefix&"Groups] where GroupID="&GroupID&"")(0)%></a><%else%>
<a onmouseover="showmenu(event,'<%=alltree%>')" href=ShowForum.asp?ForumID=<%=ParentID%>><%=Execute("Select ForumName From ["&TablePrefix&"Forums] where ForumID="&ParentID&"")(0)%></a><%
		end if

	case "ForumList"
		ii=0
		ForumsList="<select name='ParentID'><option value='0'>--</option>"
		GroupID=RequestInt("GroupID")
		ForumListLoad GroupID,0
		ForumsList=ForumsList&"</select>"
		response.write(ForumsList)

	case "UsersOnline"
		ForumID=RequestInt("ForumID")
		ThreadID=RequestInt("ThreadID")
		if ThreadID>0 then
			sql="Select * from ["&TablePrefix&"UserOnline] where UserName<>'' and IsInvisible<>1 and ThreadID="&ThreadID&""
		elseif ForumID>0 then
			sql="Select * from ["&TablePrefix&"UserOnline] where UserName<>'' and IsInvisible<>1 and ForumID="&ForumID&""
		else
			sql="Select * from ["&TablePrefix&"UserOnline] where UserName<>'' and IsInvisible<>1"
		end if

		Set Rs=Execute(sql)
		do while not Rs.eof
			content=""&content&"<li><a href=Profile.asp?UserName="&Rs("UserName")&">"&Rs("UserName")&"</a></li>"
			Rs.Movenext
		loop
		Rs.close
		response.write content
		
		
	case "ThreadStatus"
		ThreadID=RequestInt("ThreadID")
		ThreadStatus=RequestInt("ThreadStatus")
		
		if ThreadID<=0 then Response.End()
		
		Execute("update ["&TablePrefix&"Threads] Set ThreadStatus="&ThreadStatus&" where ThreadID="&ThreadID&"")
	case "Subscription"
		ThreadID=RequestInt("ThreadID")
		
		if ThreadID<=0 or ""&CookieUserName&""="" then Response.End()
		
		sql="Select * from ["&TablePrefix&"Subscriptions] where UserName='"&CookieUserName&"' and ThreadID="&ThreadID&""
		Rs.open sql,conn,1,3
		if Rs.eof then
			Rs.addnew
			Rs("UserName")=CookieUserName
			Rs("ThreadID")=ThreadID
			Rs("Email")=CookieUserEmail
			Rs.update
			BgImage="tracktopic-on.gif"
			ButtonText="取消订阅"
		else
			Execute("Delete from ["&TablePrefix&"Subscriptions] where UserName='"&CookieUserName&"' and ThreadID="&ThreadID&"")
			BgImage="tracktopic.gif"
			ButtonText="订阅主题"
		end if
		Rs.close
		%><a class="CommonImageTextButton" style="BACKGROUND-IMAGE: url(images/<%=BgImage%>)" href="javascript:Ajax_CallBack(false,'ThreadSubscription','Loading.asp?Menu=Subscription&amp;ThreadID=<%=ThreadID%>')"><%=ButtonText%></a><%
	case "Threadstar"
		ThreadID=RequestInt("ThreadID")
		VoteValue=RequestInt("Rate")
		
		if ThreadID<=0 or VoteValue<0 or ""&CookieUserName&""="" then Response.End()
		
		sql="Select * From ["&TablePrefix&"ThreadRating] where UserName='"&CookieUserName&"' and ThreadID="&ThreadID&""
		Rs.Open SQL,Conn,1,3
		if Rs.eof then Rs.addNew
			Rs("ThreadID")=ThreadID
			Rs("Rating")=VoteValue
			Rs("UserName")=CookieUserName
			Rs("DateCreated")=Now()
		Rs.update
		Rs.close
			
		RatingSum=Execute("Select sum(Rating) from ["&TablePrefix&"ThreadRating] where ThreadID="&ThreadID&"")(0)
		TotalRatings=Execute("Select count(Rating) from ["&TablePrefix&"ThreadRating] where ThreadID="&ThreadID&"")(0)

		sql="Select * from ["&TablePrefix&"Threads] where ThreadID="&ThreadID&""
		Rs.Open SQL,Conn,1,3
			Rs("RatingSum")=RatingSum
			Rs("TotalRatings")=TotalRatings
		Rs.update
			RatingSum=Rs("RatingSum")
			TotalRatings=Rs("TotalRatings")
		Rs.close

		Threadstar=RatingSum/TotalRatings

		Response.write(Threadstar)

	case "CheckUserName"
		UserName=HTMLEncode(unescape(Request.QueryString("UserName")))
		UserNameLength=RequestInt("UserNameLength")
		if len(UserName)<>UserNameLength then
			Response.write("<img src='images/check_error.gif' />&nbsp;您输入的用户名“"&UserName&"”含有URL所不能传送的字符")
			Response.end
		end if

		ErrMsg=CheckUser(UserName)
		if ErrMsg<>"" then
			Response.write("<img src='images/check_error.gif' />&nbsp;"&ErrMsg&"")
			Response.end
		end if
		
		if Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&UserName&"'" ).eof Then
			Response.write("<img src='images/check_right.gif' />")
		else
			Response.write("<img src='images/check_error.gif' align=absmiddle />&nbsp;"&UserName&" 已经有人使用，请另选一个!")
		end if
	case "CheckMail"
		UserEmail=HTMLEncode(unescape(Request.QueryString("Mail")))
		If Execute("Select UserID From ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"'" ).eof Then
			Response.write("<img src='images/check_right.gif' />")
		else
			Response.write("<img src='images/check_error.gif' />&nbsp;"&UserEmail&" 已经有人使用，请另选一个。")
		end if
	case "CheckVerifyCode"
		VerifyCode=Request.QueryString("VerifyCode")
		If VerifyCode=Session("VerifyCode") Then
			Response.write("<img src='images/check_right.gif' />")
		else
			Response.write("<img src='images/check_error.gif' />&nbsp;您的验证码输入错误，请重新输入或点击刷新验证码图片")
		end if
	case "IsObjInstalled"
		strClassString=Request.QueryString("Object")
		On Error Resume Next
		IsInstalled = "<img src='images/check_error.gif' /> 您的服务器不支持 "&Request.QueryString("Object")&" 组件"
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsInstalled = "<img src='images/check_right.gif' />"
		Set xTestObj = Nothing
		On Error GoTo 0
		Response.write(IsInstalled)
	case "Preview"
		Subject=unescape(Request.Form("Subject")&Request.Form("PostSubject"))
		Body=unescape(Request.Form("Body"))
%>
<table class=CommonListArea cellspacing="1" cellpadding="5" width="100%" align="center" border="0">
	<tr class=CommonListTitle><td><div style="float:left">预览</div><div style="float:right"><a href="javascript:ToggleMenuOnOff('Preview')">关闭</a></div></td></tr>
	<tr class="CommonListHeader"><td><div style="float:left"><b><%=HTMLEncode(Subject)%></b></div><div style="float:right"><%=now()%></div></td></tr>
	<tr class="CommonListCell"><td><%=BBCode(BodyEncode(Body))%></td></tr>
	<tr class="CommonListCell"><td align="right"><input type="Button" value=" 发表 " onclick="javascript:document.form.EditSubmit.click()"></td></tr>
</table>
<%
	case "ShowTag"
		TagID=RequestInt("TagID")

		if TagID<0 then Response.End()

		TagName=Execute("Select TagName from ["&TablePrefix&"PostTags] where TagID="&TagID&"")(0)
		ShowTagString="<div class='PostTag'><h5>"&TagName&"</h5><ul>"
		i=0
		'if TagID>0 then
			Sql="Select Top 5 * from ["&TablePrefix&"PostInTags] where TagID="&TagID&""
			Set Rs=Execute(Sql)
			do while not Rs.eof and i<5
				Set Rs1=Execute("Select Top 5 * from ["&TablePrefix&"Posts] where PostID="&Rs("PostID")&"")
					Subject=Rs1("Subject")
					if ""&Subject&""="" then Subject=Left(ReplaceText(""&Rs("Body")&"","<[^>]*>",""),10)
					ShowTagString=ShowTagString&"<li><a href='ShowPost.asp?PostID="&Rs("PostID")&"' target='_blank'>"&Subject&"</a></li>"
				Rs1.close
				i=i+1
				Rs.movenext
			loop
			Rs.Close
		'end if
		ShowTagString=ShowTagString&"<li class=""more""><a href=""Tags.asp?TagID="&TagID&""" target=""_blank"">查看更多&raquo;</a></li>"
		ShowTagString=ShowTagString&"</ul></div>"
		Response.Write(ShowTagString)


	case "LoadThemes"
		Set XMLDOM=Server.CreateObject("msxml2.FreeThreadedDOMDocument.3.0")
		XMLDOM.async = False
		XMLDOM.load(Server.MapPath("Xml/Themes.xml"))
		Response.Write(XMLDOM.xml)
end select

Sub ForumListLoad(GroupID,ParentID)
	sql="Select * From ["&TablePrefix&"Forums] where GroupID="&GroupID&" and ParentID="&ParentID&" and IsActive=1 order by SortOrder"
	Set Rs1=Execute(sql)
	do while not rs1.eof
		ForumsList=ForumsList&"<option value='"&rs1("ForumID")&"'>"&string(ii,"　")&"&gt; "&rs1("ForumName")&"</option>"
		ii=ii+1
		ForumListLoad GroupID,rs1("ForumID")
		ii=ii-1
		rs1.Movenext
	loop
	Set Rs1 = Nothing
End Sub
%>