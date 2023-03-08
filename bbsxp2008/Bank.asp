<!-- #include file="Setup.asp" -->
<%
HtmlTop
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

sql="select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&""
Rs.Open sql,Conn,1,3

SaveBankDay=int(now-Rs("BankDate"))

accrual=fix(Rs("BankMoney")*0.001)*SaveBankDay

select case Request.Form("menu")
	case "save"
		save
	case "draw"
		draw
	case "virement"
		virement
end select
%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> → <a href="Bank.asp">银行</a></div>
<br />
<table border=0 width=100%>
	<tr>
		<td width="50%" align="center" valign="top"><img src="images/Plus/Bank.gif" /><br /><br /><br />
		<table border="0" cellpadding="5" cellspacing="1" width="377" class=CommonListArea>
			<tr class=CommonListTitle>
				<td width="50%" colspan="4" align="center"><%=CookieUserName%>的银行帐号</td>
			</tr>
			<tr class="CommonListCell">
				<td width="16%" align="center">现金：</td>
				<td width="28%"><b><font color="aa0000">￥<%=Rs("UserMoney")%> </font></b></td>
				<td width="16%" align="center">存款：</td>
				<td width="31%"><b><font color="aa0000">￥<%=Rs("BankMoney")%></font></b></td>
			</tr>
			<tr class="CommonListCell">
				<td width="16%" align="center">利息：</td>
				<td width="28%"><b><font color="aa0000">￥<%=accrual%></font></b> （<%=SaveBankDay%>天）</td>
				<td width="16%" align="center">总共：</td>
				<td width="31%"><b><font color="aa0000">￥<%=Rs("BankMoney")+Rs("UserMoney")+accrual%></font></b></td>
			</tr>
			<tr class="CommonListCell">
				<td width="23%" align="center">存款时间：</td>
				<td width="68%" colspan="3"><%=Rs("BankDate")%></td>
			</tr>
			<tr class="CommonListCell">
				<td width="91%" colspan="4">本银行的利息为每天 <font color="#FF0000"><b>0.001%</b></font>，每次存款、取款自动结算利息。</td>
			</tr>
		</table>
		</td>
		<td align="center">
		<form action="Bank.asp" method="POST">
		<table cellspacing="1" cellpadding="5" border="0" width="377" height="47" class=CommonListArea>
			<tr class=CommonListTitle>
				<td width="60%">&nbsp; <b>存款</b></td>
				<td align="center">现金 <b>￥<%=Rs("UserMoney")%></b></td>
			</tr>
			<tr class="CommonListCell">
				<td align="center"><input type=hidden name=menu value="save" />&nbsp; 我要存 ￥<input size="10" value="<%=Rs("UserMoney")%>" name="qmoney" maxlength="32" /></td>
				<td align="center"><input type="submit" value=" 存 了 " name="B2" /></td>
			</tr>
		</table>
		</form>
		<form action="Bank.asp" method="POST">
		<table cellspacing="1" cellpadding="5" border="0" width="377" height="47" class=CommonListArea>
			<tr class=CommonListTitle>
				<td width="60%">&nbsp; <b>取款</b></td>
				<td align="center">存款 <b>￥<%=Rs("BankMoney")%></b></td>
			</tr>
			<tr class="CommonListCell">
				<td align="center"><input type=hidden name=menu value="draw" />&nbsp; 我要取 ￥<input size="10" value="<%=Rs("BankMoney")%>" name="qmoney" maxlength="32" /></td>
				<td align="center"><input type="submit" value=" 取 了 " name="B2" /></td>
			</tr>
		</table>
		</form>
		<form action="Bank.asp" method="POST">
		<table cellspacing="1" cellpadding="5" border="0" width="377"  class=CommonListArea>
			<tr class=CommonListTitle>
				<td width="60%">&nbsp; <b>现金转帐</b>（手续费 10 %）</td>
				<td align="center">现金 <b>￥<%=Rs("UserMoney")%></b></td>
			</tr>
			<tr class="CommonListCell">
				<td align="center" colspan="2">
				<input type=hidden name=menu value="virement" />&nbsp; 我要将 ￥<input size="5" value="1000" name="qmoney"  maxlength="32" /> 转到 <input size="10" name="dxname"  maxlength="32" /> 的帐户<br /><input type="submit" value=" 确 定 " name="B2" />
				</td>
			</tr>
		</table>
		</form>
		</td>
	</tr>
</table>
<%
Rs.close
HtmlBottom

sub save
	qmoney=RequestInt("qmoney")
	if qmoney > Rs("UserMoney") then error("<li>您的现金不够！")
	if qmoney<1 then error("<li>存款不能为零！")

	Rs("BankMoney")=Rs("BankMoney")+qmoney+accrual
	Rs("UserMoney")=Rs("UserMoney")-qmoney
	Rs("BankDate")=now()
	Rs.update
	Rs.close
	Message="<li>存款成功<li><a href=Bank.asp>返回银行</a>"
	succeed Message,"Bank.asp"
end sub


sub draw
	qmoney=RequestInt("qmoney")
	if qmoney>Rs("BankMoney") then error("<li>您的存款不够！")
	if qmoney<1 then error("<li>取款不能为零！")

	Rs("BankMoney")=Rs("BankMoney")-qmoney+accrual
	Rs("UserMoney")=Rs("UserMoney")+qmoney
	Rs("BankDate")=now()
	Rs.update
	Rs.close
	Message="<li>取款成功<li><a href=Bank.asp>返回银行</a>"
	succeed Message,"Bank.asp"
end sub

sub virement
	dxname=HTMLEncode(Request.form("dxname"))
	if dxname=CookieUserName then error("<li>您输入的是自己的帐号？")
	If Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&dxname&"'" ).eof Then error("<li>系统不存在"&dxname&"的帐号")

	qmoney=RequestInt("qmoney")
	if qmoney<10 then error"<li>转帐不能低于10！"

	if qmoney*1.1>Rs("UserMoney") then error"<li>您的现金不够！目前现金 ￥"&Rs("UserMoney")&"<li>本次汇款 ￥"&qmoney&"，手续费 ￥"&qmoney*0.1&"，共需要 ￥"&qmoney*1.1&""

	Rs("UserMoney")=Rs("UserMoney")-qmoney*1.1
	Rs("BankDate")=now()
	Rs.update
	Rs.close

	Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&qmoney&" where UserName='"&dxname&"'")

	AddApplication "Message_"&dxname,"【系统讯息】"&CookieUserName&" 已经通过银行转帐 ￥"&qmoney&" 到您的现金帐户中！"


	Log(""&CookieUserName&" 通过银行转帐 ￥"&qmoney&" 给 "&dxname&"")

	Message="<li>转帐成功<li><a href=Bank.asp>返回银行</a>"
	succeed Message,"Bank.asp"
end sub
%>