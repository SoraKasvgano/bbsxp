<!-- #include file="Setup.asp" -->
<%
top
if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")

sql="select * from [BBSXP_Users] where UserName='"&CookieUserName&"'"
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
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="Bank.asp">社区银行</a></td>
</tr>
</table><br>




<center>
<table border="0" width="100%">
	<tr>
		<td width="50%" align="center" valign="top">

<img src="images/Bank.gif"><br>
　<table border="0" cellpadding="4" cellspacing="1" width="377" class=a2>
<tr>
<td width="50%" colspan="4" class=a1 align="center">您的银行账号</td>
</tr>
<tr class=a4>
<td width="16%" align="center">现金：</td>
<td width="28%"><b><font color="aa0000"><%=Rs("UserMoney")%> </font>
金币</b></td>
<td width="16%" align="center">利息：</td>
<td width="31%"><b><font color="aa0000"><%=accrual%></font> 金币</b></td>
</tr>
<tr class=a4>
<td width="16%" align="center">存款：</td>
<td width="28%"><b><font color="aa0000"><%=Rs("savemoney")%></font> 金币</b></td>
<td width="16%" align="center">总共：</td>
<td width="31%"><b><font color="aa0000"><%=Rs("savemoney")+Rs("UserMoney")+accrual%></font>
金币</b></td>
</tr>
<tr class=a4>
<td width="23%" align="center">存款时间：</td>
<td width="68%" colspan="3"><%=Rs("SaveMoneyTime")%></td>
</tr>
<tr class=a4>
<td width="91%" colspan="4">本银行的利息为每天 <font color="#FF0000"><b>0.1%</b></font>，每次存款、取款自动结算利息。</td>
</tr>
</table>
		</td>
		<td align="center">
		
<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要存款</b>&nbsp;</td>
<td class=a1 height="25" align="center">您有现金 <b><%=Rs("UserMoney")%></b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="save">&nbsp; 我要存
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b>
</td>
<td height="25" align="center">
<input type="submit" value=" 存 了 " name="B2"></td></tr></table></form>

<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要取款</b>&nbsp;
</td>
<td class=a1 height="25" align="center">您有存款 <b><%=Rs("savemoney")%></b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="draw">&nbsp; 我要取
<input size="10" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b>
</td>
<td height="25" align="center">
<input type="submit" value=" 取 了 " name="B2"></td></tr></table></form>
		
		
<table cellSpacing="1" cellPadding="3" border="0" width="377" height="47" class=a2><tr>
<td class=a1 height="25">&nbsp; <b>我要转帐</b>&nbsp;
</td>
<td class=a1 height="25" align="right">最低转账金额为 <b>1000</b> <b>金币</b></td></tr><tr class=a4>
<td height="25" align="center" colspan="2">
<form action="Bank.asp" method="POST"><input type=hidden name=menu value="virement">&nbsp; 我要将
<input size="5" value="1000" name="qmoney" MAXSIZE="32"><b> 金币</b> 转到
<input size="10" name="dxname" MAXSIZE="32"> 的账户
<input type="submit" value=" 确 定 " name="B2"></td>
</tr></table></form>
		
		</td>
	</tr>
	</table>






<%
Rs.close
htmlend

sub save
qmoney=int(Request("qmoney"))
if qmoney > Rs("UserMoney") then error("<li>您的现金没有这么多吧！")
if qmoney<1 then error("<li>存款不能为零！")

Rs("savemoney")=Rs("savemoney")+qmoney+accrual
Rs("UserMoney")=Rs("UserMoney")-qmoney
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close
Message="<li>存款成功<li><a href=Bank.asp>返回银行</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub


sub draw
qmoney=int(Request("qmoney"))
if qmoney>Rs("savemoney") then error("<li>您的存款不够！")
if qmoney<1 then error("<li>取款不能为零！")

Rs("savemoney")=Rs("savemoney")-qmoney+accrual
Rs("UserMoney")=Rs("UserMoney")+qmoney
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close
Message="<li>取款成功<li><a href=Bank.asp>返回银行</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub

sub virement
dxname=HTMLEncode(Request.form("dxname"))

if dxname=CookieUserName then error("<li>您输入的是自己的账号？")


qmoney=int(Request("qmoney"))
if qmoney>Rs("savemoney") then error"<li>您的帐户余额不够！！"
if qmoney<1000 then error"<li>转帐不能低于1000！"
If Conn.Execute("Select id From [BBSXP_Users] where UserName='"&dxname&"'" ).eof Then error("<li>查无"&dxname&"的账号")

Rs("savemoney")=Rs("savemoney")-qmoney+accrual
Rs("SaveMoneyTime")=now()
Rs.update
Rs.close

Conn.execute("update [BBSXP_Users] set [UserMoney]=[UserMoney]+"&qmoney&" where UserName='"&dxname&"'")

Log(""&CookieUserName&" 通过银行转帐 "&qmoney&" 金币给 "&dxname&"")

Message="<li>转账成功<li><a href=Bank.asp>返回银行</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(Message&"<meta http-equiv=refresh content=3;url='Bank.asp'>")
end sub


%>