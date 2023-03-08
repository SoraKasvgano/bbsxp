document.writeln("<div id=\"eMeng\" style=\"BORDER-RIGHT:#455690 1px solid; BORDER-TOP:#a6b4cf 1px solid; Z-INDEX:99999; LEFT:0px; VISIBILITY:hidden; BORDER-LEFT:#a6b4cf 1px solid; WIDTH:180px; BORDER-BOTTOM:#455690 1px solid; POSITION:absolute; TOP:0px; BACKGROUND-COLOR:#c9d3f3\">");
document.writeln("	<table width=\"100%\" border=0 cellpadding=0 cellspacing=0>");
document.writeln("		<tr class=a1>");
document.writeln("			<td height=\"23\" valign=\"middle\"><span id=\"MsgTitle\"></span></td>");
document.writeln("			<td valign=\"middle\" align=\"right\">");
document.writeln("			<img src=\"images/msgClose.jpg\" hspace=\"3\" style=\"CURSOR:pointer\" onclick=\"closeDiv()\" title=\"关闭\"></td>");
document.writeln("		</tr>");
document.writeln("		<tr>");
document.writeln("			<td colspan=\"2\" height=60 class=a3>");
document.writeln("			<div>");
document.writeln("				<span id=\"MsgCenter\"></span></div>");
document.writeln("			</td>");
document.writeln("		</tr>");
document.writeln("	</table>");
document.writeln("</div>");

		var divTop,divLeft,divWidth,divHeight,docHeight,docWidth,objTimer,i = 0;
		function getMsg(MM1,MM2)
		{
			try{
			MsgTitle.innerHTML=MM1;//小窗口标题。
			MsgCenter.innerHTML=MM2;//小窗口内容。
		
			divTop = parseInt(document.getElementById("eMeng").style.top,10)
			divLeft = parseInt(document.getElementById("eMeng").style.left,10)
			divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10)
			divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10)
			docWidth = document.body.clientWidth;
			docHeight = document.body.clientHeight;
			document.getElementById("eMeng").style.top = parseInt(document.body.scrollTop,10) + docHeight + 10;//  divHeight
			document.getElementById("eMeng").style.left = parseInt(document.body.scrollLeft,10) + docWidth - divWidth
			document.getElementById("eMeng").style.visibility="visible"
			objTimer = window.setInterval("moveDiv()",10)
			}
			catch(e){}
		}

		function resizeDiv()
		{
			i+=1
			if(i>1288) closeDiv()
			try{
			divHeight = parseInt(document.getElementById("eMeng").offsetHeight,10)
			divWidth = parseInt(document.getElementById("eMeng").offsetWidth,10)
			docWidth = document.body.clientWidth;
			docHeight = document.body.clientHeight;
			document.getElementById("eMeng").style.top = docHeight - divHeight + parseInt(document.body.scrollTop,10)
			document.getElementById("eMeng").style.left = docWidth - divWidth + parseInt(document.body.scrollLeft,10)
			}
			catch(e){}
		}

		function moveDiv()
		{
			try
			{
			if(parseInt(document.getElementById("eMeng").style.top,10) <= (docHeight - divHeight + parseInt(document.body.scrollTop,10)))
			{
			window.clearInterval(objTimer)
			objTimer = window.setInterval("resizeDiv()",1)
			}
			divTop = parseInt(document.getElementById("eMeng").style.top,10)
			document.getElementById("eMeng").style.top = divTop - 1
			}
			catch(e){}
		}
		function closeDiv()
		{
			document.getElementById('eMeng').style.visibility='hidden';
			if(objTimer) window.clearInterval(objTimer)
		}
		
		setTimeout('closeDiv()',10000);//几秒后自动关闭