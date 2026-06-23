<!-- #include file="setup.asp" --><%top

Subject=HTMLEncode(Request.Form("Subject"))
content=ContentEncode(Request.Form("content"))
lookdate=HTMLEncode(Request("lookdate"))
adddate=HTMLEncode(Request("adddate"))
hide=RequestInt("hide")
id=RequestInt("id")


if Request("menu")="add" then
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
if Subject=empty then error("<li>��û��������־����")
if content=empty then error("<li>��û��������־����")
if adddate="" then adddate=""&year(now)&"-"&month(now)&"-"&day(now)&""



if id=0 then
sql="insert into [BBSXP_Calendar](title,UserName,content,hide,adddate) values ('"&SqlString(Subject)&"','"&SqlString(CookieUserName)&"','"&SqlString(content)&"','"&hide&"','"&SqlString(adddate)&"')"
conn.Execute(SQL)
else
conn.execute("update [BBSXP_Calendar] set title='"&SqlString(Subject)&"',content='"&SqlString(content)&"',hide='"&hide&"' where id="&id&" and UserName='"&SqlString(CookieUserName)&"'")
end if


message="<li>������־�ɹ�<li><a href=blog.asp?UserName="&CookieUserName&">���ظ�����־</a><li><a href=calendar.asp?menu=show&lookdate="&adddate&">����"&adddate&"��־</a><li><a href=calendar.asp>��������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=blog.asp?UserName="&CookieUserName&">")

elseif Request("menu")="del" then
if membercode > 3 then
conn.execute("delete from [BBSXP_Calendar] where id="&id&"")
else
conn.execute("delete from [BBSXP_Calendar] where id="&id&" and UserName='"&SqlString(CookieUserName)&"'")
end if
error2("ɾ���ɹ���")

end if

%>
<table class="a2" cellSpacing="1" cellPadding="4" width="100%" align="center" border="0">
	<tr class="a3">
		<td colSpan="2">
		<table cellSpacing="0" cellPadding="0" width="100%" border="0" id="table2">
			<tr>
				<td height="18">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> �� <a href="calendar.asp">������־</a></td>
				<td align="right" height="18"> <img src="images/jt.gif"> <a href=blog.asp?UserName=<%=CookieUserName%>>�ҵ���־</a> 
				<img src="images/jt.gif"> <a href=calendar.asp?menu=NewCalendar&lookdate=<%=lookdate%>>������־</a></td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<br>
<%

select case Request("menu")
case ""
Default
case "show"
show
case "NewCalendar"
NewCalendar

end select

sub NewCalendar
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")
if Request("id")<>empty then
sql="select * from [BBSXP_Calendar] where id="&id&""
Set Rs=Conn.Execute(sql)
UserName=rs("UserName")
content=rs("content")
title=rs("title")
adddate=rs("adddate")
hide=rs("hide")
Set Rs = Nothing
if UserName<>""&CookieUserName&"" then error2("ֻ��ԭ���߲��ܱ༭����־��")
end if


%>
	
	<table cellspacing="1" cellpadding="4" width="100%" align="center" border="0" class="a2">
		<form method="POST" name="yuziform" action="calendar.asp?menu=add" onsubmit="return CheckForm(this);">
<input name="content" type="hidden" value='<%=content%>'>
			<input type="hidden" name="id" value="<%=id%>">
			<input type="hidden" name="adddate" value="<%=adddate%>">
			<tr class="a1">
				<td width="100%" colspan="2" align="center"><b>������־</b></td>
			</tr>
			<tr class="a3">
				<td width="14%"><b>��־���⣺</b></td>
				<td width="83%">
				<input maxlength="30" name="Subject" style="width:80%" value="<%=title%>"> <select name="hide" size="1">
				<option value="0" <%if hide=0 then%>selected<%end if%>>����</option>
				<option value="1" <%if hide=1 then%>selected<%end if%>>����</option>
				</select></td>
			</tr>
			<tr class="a3">
				<td width="14%"><b>��־���ݣ�</b></td>
				<td width="83%" height="200">
				<script src="inc/post.js"></script>
				</td>
			</tr>
			<tr class="a3">
				<td colspan="2" align="center">��<input type="submit" value=" �� �� " name=EditSubmit>&nbsp;
				<input type="reset" value=" �� �� "></td>
			</tr>
		</form>
	</table>
	
<%
end sub


sub show
%> <br>

<div align="center">
	<table cellspacing="1" cellpadding="6" width="100%" border="0" class="a2">
		<tr class="a1">
			<td align="center" width="75%"><b><%=lookdate%> ������־</b></td>
		</tr>
		<tr class="a3">
			<td align="center" valign="top">
			<%
			
if lookdate<>"" then lookdateSQL="and adddate='"&lookdate&"'"
sql="select * from [BBSXP_Calendar] where (hide=0 or UserName='"&CookieUserName&"') "&lookdateSQL&" order by id Desc"

rs.Open sql,Conn,1
pagesetup=10 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("PageIndex"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
UserName=rs("UserName")
content=ReplaceText(rs("content"),"<[^>]*>","")
if len(content)>200 then content=left(""&content&"",200)&"..."


%>
				<table border="0" width="100%" cellspacing="10">
					<tr>
						<td>
						<b><font size="4"><%=rs("title")%></font></b><br>
						<br>
<%=content%>
<%if rs("hide")=1 then%><br><br>ע��<font color="#FF0000">��ƪ��־Ϊ����״̬</font><%end if%>
<hr>
<a href="blog.asp?id=<%=rs("id")%>">�Ķ�ȫ��</a>
 | 
<%if CookieUserName=UserName then%>
<a href="calendar.asp?menu=NewCalendar&id=<%=rs("id")%>">�༭</a>
<%else%>
<font color="#C0C0C0">�༭</font>
<%end if%>
 | 
<%if UserName=CookieUserName or membercode > 3 then%>
<a href="calendar.asp?menu=del&id=<%=rs("id")%>" onclick="checkclick('��ȷ��Ҫɾ��������־?')">ɾ��</a>
<%else%>
<font color="#C0C0C0">ɾ��</font>
<%end if%>
 | <%=rs("DateCreated")%> by <a href="Profile.asp?UserName=<%=UserName%>"><%=UserName%></a></td>
					</tr>
				</table>
<%
RS.MoveNext
loop

RS.Close  
%>
</td>
		</tr>
	</table>
	
<table border=0 width=100% align=center><tr><td>
<%ShowPage()%></td>
</tr></table>

	<%
end sub


sub Default
%>
	<script src="inc/calendar.js"></script>
	<script language="VBScript">
<!--
'===== ������ʱ��
Function TimeAdd(UTC,T)
   Dim PlusMinus, DST, y
   If Left(T,1)="-" Then PlusMinus = -1 Else PlusMinus = 1
   UTC=Right(UTC,Len(UTC)-5)
   UTC=Left(UTC,Len(UTC)-4)
   y = Year(UTC)
   TimeAdd=DateAdd("n", (Cint(Mid(T,2,2))*60 + Cint(Mid(T,4,2))) * PlusMinus, UTC)
   '�����չ��Լ�ڼ�: 4�µ�һ������00:00 �� 10������һ��������00:00
   If Mid(T,6,1)="*" And DateSerial(y,4,(9 - Weekday(DateSerial(y,4,1)) mod 7) ) <= TimeAdd And DateSerial(y,10,31 - Weekday(DateSerial(y,10,31))) >= TimeAdd Then
      TimeAdd=CStr(DateAdd("h", 1, TimeAdd))
      tSave.innerHTML = "R"
   Else
      tSave.innerHTML = ""
   End If
   TimeAdd = CStr(TimeAdd)
End Function
'-->
</script>
	<body onload="initial()">
<div id="detail" style="LEFT: 12px; WIDTH: 200px; POSITION: absolute; TOP: 0px; HEIGHT: 19px"></div>
	<form name="CLD">
		
		<table cellspacing="1" cellpadding="0" width="100%" border="0" class="a2" align="center">
			<tr>
				<td align="middle" width="30%" class="a4">
				<script language="JavaScript">
var enabled = 0; today = new Date();
var day; var date;
if(today.getDay()==0) day = "������"
if(today.getDay()==1) day = "����һ"
if(today.getDay()==2) day = "���ڶ�"
if(today.getDay()==3) day = "������"
if(today.getDay()==4) day = "������"
if(today.getDay()==5) day = "������"
if(today.getDay()==6) day = "������"
document.fgColor = "black";
date = " ��Ԫ " + (today.getYear()) + " �� " +
(today.getMonth() + 1 ) + "�� " + today.getDate() + "�� " +
day +"";
document.write(date)
    </script>
				</font></td>
				<td width="400" class="a3" rowspan="2" align="center" valign="top">
				<table border="0">
					<tr>
						<td class="a1" colspan="7" align="center" height="25">��Ԫ<select style="FONT-SIZE: 9pt" onchange="changeCld()" name="SY">
						<script language="JavaScript"><!--        
            for(i=1900;i<2050;i++) document.write('<option>'+i)        
            //--></script>
						</select>��<select style="FONT-SIZE: 9pt" onchange="changeCld()" name="SM">
						<script language="JavaScript"><!--        
            for(i=1;i<13;i++) document.write('<option>'+i)        
            //--></script>
						</select>��</font>&nbsp;&nbsp; <font id="GZ"></font><br>
						</td>
					</tr>
					<tr align="middle" class="a4">
						<td width="54" height="20">��</td>
						<td width="54" height="20">һ</td>
						<td width="54" height="20">��</td>
						<td width="50" height="20">��</td>
						<td width="54" height="20">��</td>
						<td width="54" height="20">��</td>
						<td width="54" height="20">��</td>
					</tr>
					<script language="JavaScript">    
            var gNum        
            for(i=0;i<6;i++) {        
               document.write('<tr align=center>')        
               for(j=0;j<7;j++) {        
                  gNum = i*7+j        
                  document.write('<td ><font id="SD' + gNum +'" size=5 face="Arial Black"></font><br><font id="LD' + gNum + '"></font></td>')
               }        
               document.write('</tr>')        
            }        
</script>
				</table>
				</td>
				<td align="middle" class="a4" height="25">������־</td>
			</tr>
			<tr>
				<td valign="top" align="middle" width="30%" class="a3"><br>
				<font id="Clock" face="Arial" color="000080" size="4" align="center">
				</font>
				<p>
				<!--ʱ�� *��ʾ�Զ�����Ϊ�չ��Լʱ��--><font style="FONT-SIZE: 9pt" size="2">
				<select style="FONT-SIZE: 9pt" onchange="changeTZ()" name="TZ">
				<option value="+0800 ���������" selected>����</option>
				<option value="-1200 ���������ˡ��ϼ���">���ʻ�����</option>
				<option value="-1100 ��Ħ��">��Ħ��</option>
				<option value="-1000 ̴��ɽ">̴��ɽ</option>
				<option value="-0900 ��������">��������</option>
				<option value="-0800 ��ɼ��">��ɼ��</option>
				<option value="-0700 ����">����</option>
				<option value="-0600 ֥�Ӹ�">֥�Ӹ�</option>
				<option value="-0500 ŦԼ">ŦԼ</option>
				<option value="-0400 ������˹">������˹</option>
				<option value="-0300 ��Լ����¬">��Լ����¬</option>
				<option value="-0200 �������в�">�������в�</option>
				<option value="-0100 ���ٶ�Ⱥ����ά�½�Ⱥ��">���ٶ�</option>
				<option value="+0000 ��������ʱ">��������ʱ</option>
				<option value="+0000 �׶�">�׶�</option>
				<option value="+0100 ����">����</option>
				<option value="+0200 ����">����</option>
				<option value="+0300 Ī˹��">Ī˹��</option>
				<option value="+0400 �ϰ�">�ϰ�</option>
				<option value="+0500 ������">������</option>
				<option value="+0530 ����">����</option>
				<option value="+0600 �￨">�￨</option>
				<option value="+0700 ����">����</option>
				<option value="+0900 ����">����</option>
				<option value="+1000 Ϥ��">Ϥ��</option>
				<option value="+1100 Ŭ����">Ŭ����</option>
				<option value="+1200 �����">�����</option>
				</select> ʱ��</font><font id="tSave" style="FONT-SIZE: 18pt; COLOR: red; FONT-FAMILY: Wingdings">
				</font><br>
				<br>
				<font style="FONT-SIZE: 120pt; COLOR: #13b0f4; FONT-FAMILY: Webdings">
				&ucirc;</font><br>
				<br>
				<font id="CITY"></font></p>
				</td>
				<td class="a3" height="30%" align="center" valign="top">
				<table border="0" width="95%">
				

<tr><td>
<%
sql="select top 20 * from [BBSXP_Calendar] where hide=0 order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
%><li><a href="blog.asp?id=<%=rs("id")%>"><%=rs("title")%></a></li><%
RS.MoveNext
loop
RS.Close   
%></td></tr>
				</table>
				</td>
			</tr>
		</table>
		
	</form>
	</div>
<%
end sub

htmlend
%>