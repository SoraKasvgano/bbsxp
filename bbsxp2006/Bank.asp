<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>����δ<a href=Login.asp>��¼</a>��̳")

sql="select * from [BBSXP_Users] where UserName='"&SqlString(CookieUserName)&"'"
Rs.Open sql,Conn,1,3

accrual=fix(Rs("savemoney")/1000*(now-Rs("SaveMoneyTime")))


select case Request.Form("menu")
case "save"
save
case "draw"
draw
case "virement"
virement
end select

%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="Bank.asp">��������</a></td>
</tr>
</table><br>




<center>
<table border="0" width="100%">
	<tr>
		<td width="50%" align="center" valign="top">

<img src="images/Bank.gif"><br>
��<table border="0" cellpadding="4" cellspacing="1" width="377" class=a2>
<tr>
<td width="50%" colspan="4" class=a1 align="center">���������˺�</td>
</tr>
<tr class=a4>
<td width="16%" align="center">�ֽ�</td>
<td width="28%"><b><font color="aa0000"><%=Rs("UserMoney")%> </font>
���</b></td>
<td width="16%" align="center">��Ϣ��</td>
<td width="31%"><b><font color="aa0000"><%=accrual%></font> ���</b></td>
</tr>
<tr class=a4>
<td width="16%" align="center">��</td>
<td width="28%"><b><font color="aa0000"><%=Rs("savemoney")%></font> ���</b></td>
<td width="16%" align="center">�ܹ���</td>
<td width="31%"><b><font color="aa0000"><%=Rs("savemoney")+Rs("UserMoney")+accrual%></font>
���</b></td>
</tr>
<tr class=a4>
<td width="23%" align="center">���ʱ�䣺</td>
<td width="68%" colspan="3"><%=Rs("SaveMoneyTime")%></td>
</tr>
<tr class=a4>
<td width="91%" colspan="4">�����е���ϢΪÿ�� <font color="#FF0000"><b>0.1%</b></font>��ÿ�δ�ȡ���Զ�������Ϣ��</td>
</tr>
</table>
		</td>
		<td align="center">
		
<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫ���</b>&nbsp;</td>
<td class=a1 height="25" align="center">�����ֽ� <b><%=Rs("UserMoney")%></b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="save">&nbsp; ��Ҫ��
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b>
</td>
<td height="25" align="center">
<input type="submit" value=" �� �� " name="B2"></td></tr></table></form>

<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫȡ��</b>&nbsp;
</td>
<td class=a1 height="25" align="center">���д�� <b><%=Rs("savemoney")%></b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="draw">&nbsp; ��Ҫȡ
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b>
</td>
<td height="25" align="center">
<input type="submit" value=" ȡ �� " name="B2"></td></tr></table></form>
		
		
<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>��Ҫת��</b>&nbsp;
</td>
<td class=a1 height="25" align="right">���ת�˽��Ϊ <b>1000</b> <b>���</b></td></tr><tr class=a4>
<td height="25" align="center" colspan="2">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="virement">&nbsp; ��Ҫ��
<input size="5" value="1000" name="qmoney" MAXSIZE="32"><b> ���</b> ת��
<input size="10" name="dxname" MAXSIZE="32"> ���˻�
<input type="submit" value=" ȷ �� " name="B2"></td>
</tr></table></form>
		
		</td>
	</tr>
	</table>






<%
Rs.close
htmlend

sub save
qmoney=RequestInt("qmoney")
if qmoney > Rs("UserMoney") then error("<li>�����ֽ�û����ô��ɣ�")
if qmoney<1 then error("<li>����Ϊ�㣡")

Rs("savemoney")=Rs("savemoney")+qmoney+accrual
Rs("UserMoney")=Rs("UserMoney")-qmoney
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close
Message="<li>���ɹ�<li><a href=Bank.asp>��������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub


sub draw
qmoney=RequestInt("qmoney")
if qmoney>Rs("savemoney") then error("<li>���Ĵ�����")
if qmoney<1 then error("<li>ȡ���Ϊ�㣡")

Rs("savemoney")=Rs("savemoney")-qmoney+accrual
Rs("UserMoney")=Rs("UserMoney")+qmoney
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close
Message="<li>ȡ��ɹ�<li><a href=Bank.asp>��������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub

sub virement
dxname=HTMLEncode(Request.form("dxname"))

if dxname=CookieUserName then error("<li>����������Լ����˺ţ�")


qmoney=RequestInt("qmoney")
if qmoney>Rs("savemoney") then error"<li>�����ʻ���������"
if qmoney<1000 then error"<li>ת�ʲ��ܵ���1000��"
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&dxname&"'" ).eof Then error("<li>����"&dxname&"���˺�")

Rs("savemoney")=Rs("savemoney")-qmoney+accrual
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close

Conn.execute("update [BBSXP_Users] set [UserMoney]=[UserMoney]+"&qmoney&" where UserName='"&dxname&"'")

Log(""&CookieUserName&" ͨ������ת�� "&qmoney&" ��Ҹ� "&dxname&"")

Message="<li>ת�˳ɹ�<li><a href=Bank.asp>��������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub


%>