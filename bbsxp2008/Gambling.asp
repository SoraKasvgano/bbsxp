<!-- #include file="Setup.asp" -->
<%
HtmlHeadTitle="赌骰子"
HtmlTop

if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

PlayMoney=RequestInt("PlayMoney")


%>
<div class="CommonBreadCrumbArea">	<%=ClubTree%> → <a href="?">赌场</a></div>

<%
Randomize

select case Request("menu")
	case ""
		index
		
	case "VsBigSmall"
		VsBigSmall
		
	case "VsBigSmallResult"
		VsBigSmallResult
		
	case "GuessNumber"
		GuessNumber
		
	case "GuessNumberResult"
		GuessNumberResult
		
	case "GuessNumber3"
		GuessNumber3
		
	case "GuessNumber3Result"
		GuessNumber3Result
		
	case "GuessNumber50"
		GuessNumber50
		
	case "GuessNumber50Result"
		GuessNumber50Result
		
		end select

sub CheckMoney
	if Request_Method <> "POST" then Alert("提交方式错误！您本次使用的是"&Request_Method&"提交方式！")
	if PlayMoney <10 or PlayMoney>100 then Alert("赌注不能少于10金币或超过100金币")
	if PlayMoney > CookieUserMoney then Alert("你的金币不够！")
end sub

sub index
%>
<table border="0" width="100%" cellspacing="0" cellpadding="6">
	<tr>
		<td width="40%" align="center">
		<p>
		<img src="images/plus/Gambling.gif" width="303" height="237" border="0" /></p>
		</td>
		<td width="60%">
		<ul>
			<li><b><a href="?menu=VsBigSmall">比大小</a></b> (赔率１：１，赌注范围１０－１００金币)<p>庄家随机出一个骰子，玩家也随机出一个骰子， 
			点数必须大于庄家才能赢。</p>
			</li>
		</ul>
		<p align="right">　</p>
		<ul>
			<li><b><a href="?menu=GuessNumber">猜点数</a></b> (赔率１：５，赌注范围１０－１００金币)<p>电脑随机出一个骰子，您选择骰子的点数，如果骰子的点数和您选的点数一样，您就赢了。</p>
			</li>
		</ul>
		<p align="right">　</p>
		<ul>
			<li><b><a href="?menu=GuessNumber3">赌点数</a></b> (赔率１：
			１－３，赌注范围１０－５００金币)<p>电脑随机出三个骰子，您可以一次压许多个骰子，如果电脑出的任何一个骰子的点数和您压的点数一样，您就赢了。</p>
			</li>
			
		</ul>
		<p align="right">　</p>
		<ul>
			
			<li><b><a href="?menu=GuessNumber50">赌数字</a></b> (赔率１：４０
			，赌注范围１０－５０００金币)<p>电脑随机出一个数字，您可以一次压许多个数字，如果电脑出的的数字和您压的数字一样，您就赢了。</p>
			</li>
		</ul>
		</td>
	</tr>
</table>
<%
end sub







sub VsBigSmall
%>
<form method="POST" action="?">
	<input type="hidden" name="menu" value="VsBigSmallResult" />
	<table border="0" cellspacing="1" cellpadding="5" width="500" class="CommonListArea" height="115">
		<tr class="CommonListTitle">
			<td height="9" align="center">比 大 小</b></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center" height="12">最大下注是 <b><font color="CC0000">100</font> 
			金币</b></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center" height="20">我要下注：<input type="text" name="PlayMoney" size="10" value="10" />
			<b>金币</b></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><input type="submit" value="确 定" /></td>
		</tr>
	</table>
</form>
<%
end sub

sub VsBigSmallResult
	CheckMoney

	
	Computer=fix(rnd*6)+1
	player=fix(rnd*6)+1
%> <br />
<table border="0" cellspacing="1" cellpadding="5" width="500" class="CommonListArea">
	<tr align="center" class="CommonListTitle">
		<td colspan="3" height="15">比 大 小</font></td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td height="8" width="192"><font color="CC0000">您的骰子</font></td>
		<td height="15" width="86" rowspan="2"><font size="6"><b>VS</b></font></td>
		<td height="8" width="190"><font color="0000CC">庄家骰子</font></td>
	</tr>
	<tr align="center" class="CommonListCell">
		<td height="7" width="192"><img src="images/plus/dice<%=player%>.gif" /></td>
		<td height="7" width="190"><img src="images/plus/dice<%=Computer%>.gif" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" height="50" colspan="3"><%
if Computer < player then
	%><font size="3"> 恭喜，你赢了 <b><font color="CC0000"><%=PlayMoney%></font> 金币</b></font><%
	PlayMoney=PlayMoney
else
	%><font size="3"> 很遗憾，你输了 <b><font color="CC0000"><%=PlayMoney%></font> 金币</b></font><%
	PlayMoney=-PlayMoney
end if


Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&PlayMoney&" where UserID="&CookieUserID&"")

%> </td>
	</tr>
</table>
<%
end sub



sub GuessNumber
%> <br />
<form method="POST" action="?">
	<input type="hidden" name="menu" value="GuessNumberResult" />
	<table border="0" cellspacing="1" cellpadding="5" width="500" class="CommonListArea">
		<tr align="center" class="CommonListTitle">
			<td width="33%">猜 点 数</td>
		</tr>
		<tr align="center" class="CommonListCell">
			<td width="33%"><img src="images/plus/random.gif" /></td>
		</tr>
		<tr class="CommonListTitle">
			<td width="960" align="center">请 选 择</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">
			<table border="0" cellspacing="0" cellpadding="0" align="center">
				<tr align="center">
					<td width="17%"><img src="images/plus/dice1.gif" /></td>
					<td width="17%"><img src="images/plus/dice2.gif" /></td>
					<td width="16%"><img src="images/plus/dice3.gif" /></td>
					<td width="17%"><img src="images/plus/dice4.gif" /></td>
					<td width="17%"><img src="images/plus/dice5.gif" /></td>
					<td width="16%"><img src="images/plus/dice6.gif" /></td>
				</tr>
				<tr align="center">
					<td width="17%">
					<input type="radio" name="dddimg" value="1" checked="checked" /></td>
					<td width="17%"><input type="radio" name="dddimg" value="2" /></td>
					<td width="16%"><input type="radio" name="dddimg" value="3" /></td>
					<td width="17%"><input type="radio" name="dddimg" value="4" /></td>
					<td width="17%"><input type="radio" name="dddimg" value="5" /></td>
					<td width="16%"><input type="radio" name="dddimg" value="6" /></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">我要下注：<input type="text" name="PlayMoney" size="10" value="10" />
			<b>金币</b></td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><input type="submit" value="确 定" /></td>
		</tr>
	</table>
</form>
<%
end sub




sub GuessNumberResult
	CheckMoney
	Player=cint(request("dddimg"))
	Computer=fix(rnd*6)+1
%> <br />
<table border="0" cellpadding="3" width="500" cellspacing="1" class="CommonListArea">
	<tr class="CommonListTitle">
		<td align="center" width="50%">你的下注</td>
		<td align="center">结 果</td>
	</tr>
	<tr>
		<td align="center" class="CommonListCell">
		<img src="images/plus/dice<%=Player%>.gif" /></td>
		<td align="center" class="CommonListCell">
		<img src="images/plus/dice<%=Computer%>.gif" /></td>
	</tr>
	<tr class="CommonListCell">
		<td colspan="2" align="center" height="50"><%
if Computer=Player then
	PlayMoney=PlayMoney*5
	%><font size="3">恭喜，你赢了  <b><font color="CC0000"><%=PlayMoney%></font> 金币</b></font><%
else
	%><font size="3">很遗憾，你输了  <b><font color="CC0000"><%=PlayMoney%> </font>金币</b></font><%	
	PlayMoney=-PlayMoney
end if

Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&PlayMoney&" where UserID="&CookieUserID&"")

%> </td>
	</tr>
</table>
<%
end sub






sub GuessNumber3
%> <br />
<form method="POST" action="?">
	<input type="hidden" name="menu" value="GuessNumber3Result" />
	<table border="0" cellspacing="1" cellpadding="5" width="600" class="CommonListArea">
		<tr align="center" class="CommonListTitle">
			<td width="33%">赌 点 数</td>
		</tr>
		<tr align="center" class="CommonListCell">
			<td width="33%"><img src="images/plus/random.gif" />
			<img src="images/plus/random.gif" /> <img src="images/plus/random.gif" /></td>
		</tr>
		<tr class="CommonListTitle">
			<td width="960" align="center">请 下 注</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">
				<table border="0" cellspacing="0" cellpadding="5" width="100%">
					<tr align="center">
						<td>点数</td>
						<td><img src="images/plus/dice1.gif" /></td>
						<td><img src="images/plus/dice2.gif" /></td>
						<td><img src="images/plus/dice3.gif" /></td>
						<td><img src="images/plus/dice4.gif" /></td>
						<td><img src="images/plus/dice5.gif" /></td>
						<td><img src="images/plus/dice6.gif" /></td>
					</tr>
					<tr align="center">
						<td>金额</td>
						<td><input type="text" name="PlayMoney1" size="5" value="0" /></td>
						<td><input type="text" name="PlayMoney2" size="5" value="0" /></td>
						<td><input type="text" name="PlayMoney3" size="5" value="0" /></td>
						<td><input type="text" name="PlayMoney4" size="5" value="0" /></td>
						<td><input type="text" name="PlayMoney5" size="5" value="0" /></td>
						<td><input type="text" name="PlayMoney6" size="5" value="0" /></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><input type="submit" value="确 定" /></td>
		</tr>
	</table>
</form>
<%
end sub



sub GuessNumber3Result
	Computer1=fix(rnd*6)+1
	Computer2=fix(rnd*6)+1
	Computer3=fix(rnd*6)+1
	
			
	PlayMoney=0
	i=0
	Do While i < 6
		i=i+1
		TempMoney=RequestInt("PlayMoney"&i)
		if TempMoney>0 then
			TotlePlayMoney=TotlePlayMoney+TempMoney
			PlaySelectList=PlaySelectList&"<Img src=images/plus/dice"&i&".gif title='"&TempMoney&" 金币' > "
			
			StartMoney=PlayMoney
			if i=Computer1 then PlayMoney=PlayMoney+TempMoney
			if i=Computer2 then PlayMoney=PlayMoney+TempMoney
			if i=Computer3 then PlayMoney=PlayMoney+TempMoney
			if StartMoney=PlayMoney then PlayMoney=PlayMoney-TempMoney '如果没中就扣钱
			
		end if
	loop
	
	
	if Request_Method <> "POST" then Alert("提交方式错误！您本次使用的是"&Request_Method&"提交方式！")
	if TotlePlayMoney < 10 then Alert("最少赌注不能少于10金币")
	if TotlePlayMoney > 500 then Alert("最大赌注不能超过500金币")
	if TotlePlayMoney > CookieUserMoney then Alert("你的金币不够！")

%> <br />
<table border="0" cellpadding="3" width="500" cellspacing="1" class="CommonListArea">
	<tr class="CommonListTitle">
		<td align="center">你的下注</td>
	</tr>
	<tr>
		<td align="center" class="CommonListCell"><%=PlaySelectList%> </td>
	</tr>
	<tr class="CommonListTitle">
		<td align="center">结果</td>
	</tr>
	<tr>
		<td align="center" class="CommonListCell">
		<img src="images/plus/dice<%=Computer1%>.gif" />
		<img src="images/plus/dice<%=Computer2%>.gif" />
		<img src="images/plus/dice<%=Computer3%>.gif" /> <br />
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" height="50"><%
	if PlayMoney=0 then
		%><font size="3">本局持平</font><%
	elseif PlayMoney>0 then

		%><font size="3">恭喜，你赢了 <b><font color="CC0000"><%=PlayMoney%></font> 金币</b></font><%	
	else
		%><font size="3">很遗憾，你输了 <b><font color="CC0000"><%=Abs(PlayMoney)%> </font> 金币</b></font><%	
end if

Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&PlayMoney&" where UserID="&CookieUserID&"")

%> </td>
	</tr>
</table>
<%
end sub




sub GuessNumber50
%> <br />
<form method="POST" action="?" name="form" >
	<input type="hidden" name="menu" value="GuessNumber50Result" />
	<table border="0" cellspacing="1" cellpadding="5" width="600" class="CommonListArea">
		<tr align="center" class="CommonListTitle">
			<td width="33%">赌 数 字</td>
		</tr>
		<tr align="center" class="CommonListCell">
			<td width="33%"><div style="float:left">快速选择：<input type="radio" value="Big" id="FastSelectType"  name="FastSelectType" />大　<input type="radio" value="Small" id="FastSelectType" name="FastSelectType" />小　<input type="radio" value="One" id="FastSelectType" name="FastSelectType" />单　<input type="radio" id="FastSelectType" value="Two" name="FastSelectType" />双</div>
			
			<div style="float:right"><input type=text id=PlayMoney name=PlayMoney size=5 value=500 /> 
				<input type="button" value="快速压注" onclick="FastSelect()" /></div></td>
		</tr>
		<tr class="CommonListTitle">
			<td width="960" align="center">请 下 注</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center">
			<div align="center">
				<table border="0" cellspacing="0" cellpadding="5" width="100%">
<%
	i=0
	Do While i < 50

		if i mod 10 = 0 then response.write "<tr align=center>"
		i=i+1
response.write "<td><b>"&i&"</b><br /><input type=text id=PlayMoney"&i&" name=PlayMoney"&i&" size=5 value=0 /></td>"

		if i mod 10 = 0 then response.write "</tr>"
			loop
%>


				</table>
			</div>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="center"><input type="submit" value="确 定" /></td>
		</tr>
	</table>
</form>



<script language="JavaScript" type="text/javascript">
function FastSelect(){

PlayMoney=$("PlayMoney").value

var FastSelectType;
var FormObject = document.forms["form"];
for (var i=0;i<FormObject.elements["FastSelectType"].length;i++) {
	if(FormObject.elements["FastSelectType"][i].checked){   
		FastSelectType=FormObject.elements["FastSelectType"][i].value;
		break;
	}
}

if (!FastSelectType){alert("你没有选择快速压注的类型")}

if (FastSelectType=="One"){
	for (var i=1; i <= 50; i++) {
	if(i % 2  == 1){
			$("PlayMoney"+i).value= PlayMoney/25;
		}else{
			$("PlayMoney"+i).value= 0;
		}
	}
}

if (FastSelectType=="Two"){
	for (var i=1; i <= 50; i++) {
	if(i % 2  == 0){
			$("PlayMoney"+i).value= PlayMoney/25;
		}else{
			$("PlayMoney"+i).value= 0;
		}
	}
}


if (FastSelectType=="Big"){

	for (var i=1; i <= 25; i++) {
	$("PlayMoney"+i).value= 0;
	}
	
	for (var i=26; i <= 50; i++) {
	$("PlayMoney"+i).value= PlayMoney/25;
	}

}



if (FastSelectType=="Small"){
	for (var i=1; i <= 25; i++) {
	$("PlayMoney"+i).value= PlayMoney/25;
	}
	for (var i=26; i <= 50; i++) {
	$("PlayMoney"+i).value= 0;
	}
}

}
</script>

<%
end sub



sub GuessNumber50Result
	Computer=fix(rnd*50)+1

	PlayMoney=0
	i=0
	Do While i < 50
		i=i+1
		TempMoney=RequestInt("PlayMoney"&i)
		if TempMoney>0 then

			TotlePlayMoney=TotlePlayMoney+TempMoney
			PlaySelectList=PlaySelectList&"<b><font size=5 title='"&TempMoney&" 金币'>"&i&"</font></b>　"
			
			StartMoney=PlayMoney
			if i=Computer then PlayMoney=PlayMoney+TempMoney*40
			if StartMoney=PlayMoney then PlayMoney=PlayMoney-TempMoney '如果没中就扣钱
			
		end if
	loop
	
	
	if Request_Method <> "POST" then Alert("提交方式错误！您本次使用的是"&Request_Method&"提交方式！")
	if TotlePlayMoney<10 then Alert("最少赌注不能少于10金币")
	if TotlePlayMoney>5000 then Alert("最大赌注不能超过5000金币")
	if TotlePlayMoney > CookieUserMoney then Alert("你的金币不够！")

%> <br />
<table border="0" cellpadding="3" width="500" cellspacing="1" class="CommonListArea">
	<tr class="CommonListTitle">
		<td align="center">你的下注</td>
	</tr>
	<tr>
		<td align="center" class="CommonListCell"><%=PlaySelectList%> </td>
	</tr>
	<tr class="CommonListTitle">
		<td align="center">结果</td>
	</tr>
	<tr>
		<td align="center" class="CommonListCell"><b><font size=7><%=Computer%></font></b>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" height="50"><%
	if PlayMoney=0 then
		%><font size="3">本局持平</font><%
	elseif PlayMoney>0 then

		%><font size="3">恭喜，你赢了 <b><font color="CC0000"><%=PlayMoney%></font> 金币</b></font><%	
	else
		%><font size="3">很遗憾，你输了 <b><font color="CC0000"><%=Abs(PlayMoney)%> </font> 金币</b></font><%	
end if

Execute("update ["&TablePrefix&"Users] Set UserMoney=UserMoney+"&PlayMoney&" where UserID="&CookieUserID&"")

%> </td>
	</tr>
</table>
<%
end sub


%> 

<hr width="250" size="1" />您现在身上共有 <b><font color="CC0000"><%=Execute("Select UserMoney from ["&TablePrefix&"Users] where UserName='"&CookieUserName&"'")(0)%></font> 
金币</b><br />
<hr width="250" size="1" /><br />
<input onclick="history.back()" type="button" value=" &lt;&lt; 返 回 " /> <br />

<%
HtmlBottom
%>