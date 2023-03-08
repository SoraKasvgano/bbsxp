<!-- #include file="Setup.asp" -->
<%
id=int(Request("id"))
sql="Select * From [BBSXP_Forums] where followid="&id&" and ForumHide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
do while not rs1.eof
alltree=""&alltree&"<div class=menuitems><a href=ShowForum.asp?ForumID="&rs1("id")&">"&rs1("ForumName")&"</a></div>"
rs1.Movenext
loop
Set Rs1 = Nothing
%>
<SCRIPT>
parent.temp<%=id%>.innerHTML="<a onmouseover=\"showmenu(event,'<%=alltree%>')\" href=ShowForum.asp?ForumID=<%=id%>><%=Conn.Execute("Select ForumName From [BBSXP_Forums] where id="&id&"")(0)%></a>"
</SCRIPT>
<%CloseDatabase%>