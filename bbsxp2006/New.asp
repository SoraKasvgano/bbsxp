<!-- #include file="conn.asp" -->
<%
SiteSettings=Conn.Execute("[BBSXP_SiteSettings]")
'=========================================================
' ���ô���
'<script src=http://www.***.com/New.asp?TopicCount=10&TitleLen=15&Showtime=1&icon=��></script>
'---------------------------------------------------------
' ForumID:    ָ����̳��ID
' TopicCount: ��ʾ���������⣨���100����
' TimeLimit:  ��ʾ�೤ʱ���ڵ����ӣ���λ���죩
' Sort:       ����ʽ ID Views Replies IsGood
' icon:       ����ǰ�ķ���
' TitleLen: ��ʾ����ĳ���
' Showtime:   ��ʾ����ʱ�� 1=�� 0=��
'=========================================================
ForumID=int(Request("ForumID"))
TopicCount=int(Request("TopicCount"))
TitleLen=int(Request("TitleLen"))
TimeLimit=int(Request("TimeLimit"))
Sort=Server.HTMLEncode(Request("Sort"))

if Sort = empty then
SqlSort="ID"
else
SqlSort=Sort
end if

if TitleLen=empty then
SqlTitleLen=100
else
SqlTitleLen=TitleLen
end if


if TopicCount=empty then
SqlTopicCount=10
else
SqlTopicCount=TopicCount
end if


if ForumID<>empty then SqlForumID=" and ForumID="&ForumID&""
if TimeLimit<>empty then SqlTimeLimit=" and PostTime>"&SqlNowString&"-"&TimeLimit&""

if Len(SqlSort)>10 or TopicCount>100 then Response.End

	sql="select top "&SqlTopicCount&" * from [BBSXP_Threads] where IsDel=0 "&SqlForumID&" "&SqlTimeLimit&" order by "&SqlSort&" desc"
	Set Rs=Conn.Execute(sql)
	do while Not Rs.Eof

	Topic=ReplaceText(Rs("Topic"),"<[^>]*>","")
	if Request("Showtime")=1 then Showtime=" ["&Rs("PostTime")&"]"
	list=list&""&Request("icon")&" <a href="&SiteSettings("SiteURL")&"ShowPost.asp?ThreadID="&Rs("id")&" target=_blank>"&Left(Topic,SqlTitleLen)&"</a>"&Showtime&"<br>"

	Rs.MoveNext
	Loop
	Rs.close
%>
document.write("<%=list%>")
<%CloseDatabase%>