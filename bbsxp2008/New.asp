<!-- #include file="conn.asp" -->
<%

'=========================================================
' ���ô���
'<script type="text/javascript" src="http://bbs.yuzi.net/New.asp?TopicCount=10&TitleLen=15&Showtime=1&icon=��"></script>
'---------------------------------------------------------
'���õĲ�����˵����
' GroupID:    ָ����̳���ID����ָ���˴˲���������Ҫ��ָ�� ForumID �����������
' ForumID:    ָ����̳��ID������ָ��������������̳��
' TopicCount: ��ʾ���������⣨���100����
' TimeLimit:  ��ʾ�೤ʱ���ڵ����ӣ���λ���죩
' Sort:       ����ʽ ThreadID[����] ��TotalViews[������] ��TotalReplies[������] ��IsGood[������]
' icon:       ����ǰ�ķ��ţ��磺"��"��
' TitleLen:   ��ʾ����ĳ���
' Showtime:   �Ƿ���ʾ����ʱ�� 1=�� 0=��
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