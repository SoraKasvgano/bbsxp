<!-- #include file="conn.asp" -->
<%

'=========================================================
' 调用代码
'<script type="text/javascript" src="http://bbs.yuzi.net/New.asp?TopicCount=10&TitleLen=15&Showtime=1&icon=◎"></script>
'---------------------------------------------------------
'调用的参数及说明：
' GroupID:    指定论坛组的ID（若指定了此参数，则不需要再指定 ForumID 这个参数。）
' ForumID:    指定论坛的ID（若不指定，调用整个论坛）
' TopicCount: 显示多少条主题（最高100条）
' TimeLimit:  显示多长时间内的帖子（单位：天）
' Sort:       排序方式 ThreadID[新帖] 、TotalViews[人气帖] 、TotalReplies[热门帖] 、IsGood[精华帖]
' icon:       标题前的符号（如："◎"）
' TitleLen:   显示标题的长度
' Showtime:   是否显示发表时间 1=是 0=否
'=========================================================
GroupID=RequestInt("GroupID")
ForumID=RequestInt("ForumID")
TopicCount=RequestInt("TopicCount")
TitleLen=RequestInt("TitleLen")
TimeLimit=RequestInt("TimeLimit")
icon=HTMLEncode(Request("icon"))
Sort=HTMLEncode(Request("Sort"))

if Sort = empty then
	SqlSort="ThreadID"
elseif len(Sort)<20 then
	SqlSort=Sort
end if

if TitleLen=0 then
	SqlTitleLen=100
else
	SqlTitleLen=TitleLen
end if


if TopicCount=0 then
	SqlTopicCount=10
elseif TopicCount>100 then
	SqlTopicCount=100
else
	SqlTopicCount=TopicCount
end if


if GroupID>0 then
	ForumIDSet=""
	Set Rs=Execute("Select ForumID from ["&TablePrefix&"Forums] where GroupID="&GroupID&"")
    if not Rs.eof then
    do while not Rs.eof
    	ForumIDSet=ForumIDSet&","&Rs("ForumID")
    	Rs.movenext
    loop
    Rs.close
    SqlForumID=" and ForumID in ("&Mid(ForumIDSet,2)&")"
    end if
elseif ForumID > 0 then
	SqlForumID=" and ForumID="&ForumID&""
end if

if TimeLimit > 0 then SqlTimeLimit=" and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&TimeLimit&""


	sql="Select top "&SqlTopicCount&" * from ["&TablePrefix&"Threads] where Visible=1 "&SqlForumID&" "&SqlTimeLimit&" order by "&SqlSort&" desc"
    
	Set Rs=Execute(sql)
    i=0
	do while Not Rs.Eof and i<SqlTopicCount
		Topic=ReplaceText(Rs("Topic"),"<[^>]*>","")
		if Request("Showtime")=1 then Showtime=" ["&Rs("PostTime")&"]"
		list=list&""&icon&" <a href="&SiteConfig("SiteUrl")&"/ShowPost.asp?ThreadID="&Rs("ThreadID")&" target=_blank>"&Left(Topic,SqlTitleLen)&"</a>"&Showtime&"<br />"
		i=i+1
		Rs.MoveNext
	Loop
	Rs.close
    
%>
document.write("<%=list%>")