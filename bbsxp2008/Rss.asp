<!--#include file="conn.asp"--><%

if SiteConfig("EnableForumsRSS")=0 then
	response.write "ϵͳ�ر��� RSS ����Դ����"
	Response.End
end if

ForumID=RequestInt("ForumID")


if ForumID > 0 then

	Rs.Open "["&TablePrefix&"ForumPermissions] where ForumID="&ForumID&" and RoleID=0",Conn,1
		
	if Rs.Eof then
		response.write "�Ҳ���ָ���İ���Ȩ�ޱ�"
		Response.End
	end if
	
	PermissionRead=Rs("PermissionRead")
	PermissionView=Rs("PermissionView")
	Rs.Close

	if PermissionView=0 or PermissionRead=0 then
		response.write "�ð��û�п����ο�������Ķ�Ȩ�ޣ��޷�ʹ�� RSS ����Դ��"
		Response.End
	end if


	ChannelTitle=Execute("Select ForumName from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")(0)
	ChannelLink=""&SiteConfig("SiteUrl")&"/ShowForum.asp?ForumID="&ForumID&""
end if


Response.contentType="application/xml"

%><?xml version="1.0" encoding="<%=BBSxpCharset%>"?>
<rss version="2.0">
<channel>

	<title><%=XmlEncode(ChannelTitle)%></title>
	<link><%=XmlEncode(ChannelLink)%></link>
	<description>Latest <%=SiteConfig("RSSDefaultThreadsPerFeed")%> threads of all forums</description>

	<image>
		<title><%=XmlEncode(SiteConfig("SiteName"))%></title>
		<url><%=XmlEncode(SiteConfig("SiteUrl"))%>/images/logo.gif</url>
		<link><%=XmlEncode(SiteConfig("SiteUrl"))%>/Default.asp</link>
	</image>

<%
	sql="Select top "&SiteConfig("RSSDefaultThreadsPerFeed")&" * from ["&TablePrefix&"Threads] where Visible=1 and ForumID="&ForumID&" order by ThreadID desc"
	Set Rs=Execute(sql)
		do while Not Rs.Eof
%>
	<item>
		<title><%=XmlEncode(Rs("Topic"))%></title>
		<link><%=XmlEncode(SiteConfig("SiteUrl"))%>/ShowPost.asp?ThreadID=<%=Rs("ThreadID")%></link><category><%=XmlEncode(Rs("Category"))%></category><author><%=XmlEncode(Rs("PostAuthor"))%></author><pubDate><%=FormatTime(Rs("PostTime"))%></pubDate><description><![CDATA[<%=XmlEncode(Rs("Description"))%>]]></description></item><%
		Rs.MoveNext
		Loop
	Set Rs = Nothing
%></channel></rss>