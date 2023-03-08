<!-- #include file="Setup.asp" -->
<%
ThreadID=RequestInt("ThreadID")
HtmlHead
%>
<body topmargin=0>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>参与者</td>
		<td>评级</td>
		<td>参与时间</td>
	</tr>
<%
sql="Select * from ["&TablePrefix&"ThreadRating] where ThreadID="&ThreadID&""
Set Rs=Execute(sql)
Do While Not Rs.EOF 
%>
	<tr class="CommonListCell">
		<td> <a href="javascript:opener.location=('Profile.asp?UserName=<%=RS("UserName")%>'); self.close();"><%=RS("UserName")%></a></td>
		<td><img border=0 src=Images/Star/<%=RS("Rating")%>.gif /></td>
		<td><%=RS("DateCreated")%></td>
	</tr><%
Rs.MoveNext
loop
Rs.Close
%>
</table><input onclick="javascript:self.close();" type="button" value="关闭" />
</body>