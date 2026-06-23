<!-- #include file="conn.asp" -->
<%
SiteSettings=Conn.Execute("[BBSXP_SiteSettings]")
Response.contentType="application/xml;charset=gbk"%><?xml version="1.0" encoding="gbk"?><rss version="2.0"><channel>
<%
ForumID=RequestInt("ForumID")
if ForumID<=0 then Response.End()
sql="select * from [BBSXP_Forums] where id="&ForumID&""
Set Rs=Conn.Execute(sql)
if Rs.eof then Response.End()
ForumPass=Rs("ForumPass")
ForumPassword=Rs("ForumPassword")
ForumUserList=Rs("ForumUserList")
Rs.close
%>
<!-- #include file="inc/Validate.asp" -->
<title><%=XmlEncode(Conn.Execute("Select ForumName From [BBSXP_Forums] where id="&ForumID&"")(0))%></title>
<link><%=XmlEncode(SiteSettings("SiteURL"))%>ShowForum.asp?ForumID=<%=ForumID%></link>
<%
sql="select top 20 * from [BBSXP_Threads] where IsDel=0 and ForumID="&ForumID&" order by id desc"
Set Rs=Conn.Execute(sql)
do while Not Rs.Eof
Topic=XmlEncode(ReplaceText(Rs("Topic"),"<[^>]*>",""))
TableName=SafeTableSuffix(Rs("PostsTableName"))
if TableName="" then TableName="0"
content=XmlEncode(Conn.Execute("Select content From [BBSXP_Posts"&TableName&"] where ThreadID="&Rs("id")&" and IsTopic=1")(0))
if len(content)>200 then content=left(""&content&"",200)&"..."
%>
<item>
<title><%=Topic%></title>
<category><%=XmlEncode(Rs("SpecialTopic"))%></category>
<link><%=XmlEncode(SiteSettings("SiteURL"))%>ShowPost.asp?ThreadID=<%=Rs("id")%></link><author><%=XmlEncode(Rs("UserName"))%></author><description><![CDATA[<%=content%>]]></description><pubDate><%=Rs("PostTime")%></pubDate></item><%
Rs.MoveNext
Loop
Rs.close
%></channel></rss><%CloseDatabase%>