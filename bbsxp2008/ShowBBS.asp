<!-- #include file="Setup.asp" -->
<%
UserName=HTMLEncode(Request("UserName"))
TempStr="MyTopic|MyPost|MySubscriptions"
if instr("|"&TempStr&"|","|"&Request("menu")&"|")>0 and CookieUserName=empty then error("����δ<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>��̳")

HtmlTop
select case Request("menu")
	case ""
		ForumTitle="��������"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 order by ThreadID Desc"
	case "HotViews"
		ForumTitle="��������"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&SiteConfig("PopularPostThresholdDays")&" order by TotalViews Desc"
	case "HotReplies"
		ForumTitle="��������"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&SiteConfig("PopularPostThresholdDays")&" order by TotalReplies Desc"
	case "GoodTopic"
		ForumTitle="��������"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and IsGood=1 order by ThreadID Desc"
	case "VoteTopic"
		ForumTitle="ͶƱ����"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and IsVote=1 order by ThreadID Desc"
	case "NoReplies"
		ForumTitle="δ�ظ���"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and TotalReplies=0 order by ThreadID Desc"
	case "MyTopic"
		ForumTitle="�ҵ�����"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and PostAuthor='"&UserName&"' order by ThreadID Desc"
	case "MyPost"
		ForumTitle="�Ҳ��������"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and ThreadID in (Select distinct ThreadID from ["&TablePrefix&"Posts] where PostAuthor='"&UserName&"' and Visible=1 and ParentID>0) order by ThreadID Desc"
	case "MySubscriptions"
		ForumTitle="�ҵĶ���"
		sql="Select top 200 * from ["&TablePrefix&"Threads] where Visible=1 and ThreadID in (Select distinct ThreadID from ["&TablePrefix&"Subscriptions] where UserName='"&UserName&"') order by ThreadID Desc"
	case else
		error("���������")
end select

Rs.Open sql,Conn,1
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> �� <%=ForumTitle%></div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
      <td align="center"><a href="?">��������</a></td>
      <td align="center"><a href="?menu=HotViews">��������</a></td>
      <td align="center"><a href="?menu=HotReplies">��������</a></td>
      <td align="center"><a href="?menu=NoReplies">δ�ظ���</a></td>
      <td align="center"><a href="?menu=GoodTopic">��������</a></td>
      <td align="center"><a href="?menu=VoteTopic">ͶƱ����</a></td>
      <td align="center"><a href="?menu=MyTopic&UserName=<%=CookieUserName%>">�ҵ�����</a></td>
      <td align="center"><a href="?menu=MyPost&UserName=<%=CookieUserName%>">�ҵĻ���</a></td>
      <td align="center"><a href="?menu=MySubscriptions&UserName=<%=CookieUserName%>">�ҵĶ���</a></td>
	</tr>
</table>
<br />
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td colspan="5">����</td>
	</tr>
	<tr class=CommonListHeader align=center>
		<td>����</td>
		<td>����</td>
		<td>�ظ�</td>
		<td>�鿴</td>
		<td>������</td>
	</tr>
<%

	PageSetup=SiteConfig("ThreadsPerPage") '�趨ÿҳ����ʾ����
	Rs.Pagesize=PageSetup
	TotalPage=Rs.Pagecount  '��ҳ��
	PageCount = RequestInt("PageIndex")
	if PageCount <1 then PageCount = 1
	if PageCount > TotalPage then PageCount = TotalPage
	if TotalPage>0 then Rs.absolutePage=PageCount '��ת��ָ��ҳ��
	i=0
	Do While Not RS.EOF and i<pagesetup
		i=i+1
		ShowThread()
		Rs.MoveNext
	loop
Rs.Close
%>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%><tr><td><%ShowPage()%></td></tr></table>
<%
HtmlBottom
%>