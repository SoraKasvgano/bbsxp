<!-- #include file="Setup.asp" -->
<%top%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
	<a href="help.asp">使用帮助</a></td></tr></table>
<br>

<%if Request("menu")="Ranks"then%>
<table border="0" cellpadding="3" cellspacing="1" class=a2 width=100%>
	<tr class=a1>
		<td width="30%" align="center">等级名称</td>
		<td width="30%">
		<p align="center">经验值</p>
		</td>
		<td width="30%">级别图标</td>
	</tr>
<%
sql="select * from [BBSXP_Ranks] order by PostingCountMin"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
%>
	<tr class=a3>
		<td width="30%" align="center"><%=Rs("RankName")%></td>
		<td width="30%" align="center">≥<%=Rs("PostingCountMin")%></td>
		<td width="30%">
		<img src="<%=Rs("RankIconUrl")%>">
		</td>
	</tr>
<%
Rs.Movenext
loop
Set Rs = Nothing
%>
	<tr class=a3>
		<td width="30%" align="center">尚未激活</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode0.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">贵宾会员</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode2.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">版主</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode3.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">超级版主</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode4.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">管理员</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode5.gif" /></td>
	</tr>
</table>

<%elseif Request("menu")="popedom"then%>

<table cellspacing="1" cellpadding="0" border="0" class=a2 width=100%>
	<tr height="25" class=a1>
		<td align="middle" width="20%">操作权限</td>
		<td align="middle" width="10%">访客</td>
		<td align="middle" width="10%">注册会员</td>
		<td align="middle" width="10%">贵宾会员</td>
		<td align="middle" width="10%">论坛版主</td>
		<td align="middle" width="10%">超级版主</td>
		<td align="center" width="10%">管理员</td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">浏览正规论坛</td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">浏览会员论坛</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">发送短讯</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">参与投票</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">编辑帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">搜索帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">发表帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">发表彩色标题</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">拉前帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">关闭帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">移动帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">删除帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">固顶帖子</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">添加精华帖</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">帖子总固顶</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">管理社区监狱</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">删除公开日志</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">查看在线会员IP</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b>√</b></td>
		<td align="center"><b>√</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">后台管理</td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="middle"><b><font color="red">×</font></b></td>
		<td align="center"><b>√</b></td>
	</tr>
	</table>
<%elseif Request("menu")="ybb"then%>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
	<tr>
		<td class="a1" height="25"><b>什么是YBB(Yuzi Bulletin Board)代码？ </b></td>
	</tr>
	<tr class=a3>
		<td>YBB代码是HTML的一个变种。您也许已经对它很熟悉了。一般情况下，如果允许您用HTM，也就可以使用YBB代码。即使您的讨论区不能让您使用HTML，您也可以使用YBB代码。 
		由于要求使用的编码很少，即使可以使用HTML，您可能也想使用YBB代码，因为代码错误的可能性大大减小了。
		</td>
	</tr>
	<tr>
		<td class="a1" height="25"><b>URL超级链接</b></td>
	</tr>
	<tr class=a3>
		<td>在您的信息里加入超级链接，只要按下列方式套入就可以了<br><br>
		<font color="red">[url]</font><a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a><font color="red">[/url]</font>
		<br>或者 <br>
		<font color="red">[url=http://www.yuzi.net]</font><a target="_blank" href="http://www.yuzi.net">YUZI工作室</a><font color="red">[/url]</font><br><br>
		按上例套入，YBB代码会自动对URL产生链接，并保证当用户点击新的窗口时这个链接是打开着的。注意URL的&quot;http://&quot;这一部分是随意的。</td>
	</tr>
	<tr class="a1">
		<td height="25"><b>电子邮件链接</b></td>
	</tr>
	<tr class=a3>
		<td>在您的信息里加入电子邮件的超级链接，只要按照下例套入就可以了<br>
<br>
		<font color="red">[email]</font><a href="mailto:yuzi@yuzi.net">yuzi@yuzi.net</a><font color="red">[/email]</font><br>
<br>
按上例套入，YBB代码会对电子邮件自动产生链接。 </td>
	</tr>
	<tr class="a1">
		<td height="25"><b>粗体与斜体</b></td>
	</tr>
	<tr class=a3>
		<td>您可以使用 [b] [i] 等这些标志以达到在帖子中使用相应的效果<br>
		<br>
		您好,<font color="red">[b]</font><strong>管理员</strong><font color="red">[/b]</font><br>
		您好,<font color="red">[i]</font><em>管理员</em><font color="red">[/i]</font><br>
		您好,<font color="red">[u]</font><u>管理员</u><font color="red">[/u]</font><br>
		您好,<font color="red">[strike]</font><strike>管理员</strike><font color="red">[/strike]</font></td>
	</tr>
	<tr class="a1">
		<td><b>移动文字 </b></td>
	</tr>
	<tr class=a3>
		<td height="42">在您的信息里加入移动文字，只要按下例套入就可以了<br><br>
		<font color="red">[marquee]</font>移动文字<font color="red">[/marquee]</font><br>
		
		<marquee>移动文字</marquee></td>
	</tr>
	<tr class="a1">
		<td><b>引用其他信息 </b></td>
	</tr>
	<tr class=a3>
		<td>引用一些人的帖子，只要剪切和拷贝然后按下例套入就可以了<br>
		<br><font color="red">
		[QUOTE]</font>别问您的国家能为您作什么......<br>
		问您能为您的国家作什么？<font color="red">[/QUOTE]</font>
		<BLOCKQUOTE><strong>引用</strong>：<HR Size=1>别问您的国家能为您作什么......<br>
		问您能为您的国家作什么？ <HR SIZE=1></BLOCKQUOTE>
		</td>
	</tr>
	<tr class="a1">
		<td><b>彩色文字</b></td>
	</tr>
	<tr class=a3>
		<td><font color="red">[COLOR=green]</font><font COLOR="#008000">彩色文字</font><font COLOR="red">[/COLOR]</font>　　“green”代表文字的颜色<font color="red"><br>
		[FONT=黑体]</font><font face="黑体">彩色文字</font><font color="red">[/FONT]</font>　　“黑体”代表文字的字体<font color="red"><br>
		[SIZE=5]</font><font size="5">彩色文字</font><font color="red">[/SIZE]</font>　　“5”代表文字的大小</td>
	</tr>
	<tr>
		<td height="14" class="a1"><b>特别注意</b></td>
	</tr>
	<tr>
		<td height="76" bgcolor="FFFFFF">
		<p>您不可以同时使用<font face="Verdana, Arial">HTML</font>和YBB代码的同一种功能。并且注意YBB代码不对大小写敏感。所以您可以用<font color="red">[URL]</font> 
		或 <font color="red">[url]</font> <font color="800000">
		<br>
		不正确的</font><font color="800000" face="Verdana, Arial">YBB</font><font color="800000">代码使用：</font><font face="Verdana, Arial"><font color="red"><br>
		[url]</font> www.yuzi.net <font color="red">[/url]</font> </font>不要在括号和您输入的文字之间有空格<font face="Verdana, Arial"><font color="red"><br>
		[email]</font>yuzi@yuzi.net<font color="red">[email]</font>
		</font>在结束时，不要忘了在括号内加入斜杠<font color="red">[/email]</font>
		</p>
		</td>
	</tr>
</tbody>
</table>


<%else%>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">常见问题解答 </td>
	</tr>
	<tr>
		<td class="a3"><b>注册和登录</b> <br>
		<a href="#A1">我为什么要注册？ </a><br>
		<a href="#A2">怎样才能注册？ </a><br>
		<a href="#A3">我已经注册了用户名和密码，怎么登录？ </a><br>
		<a href="#A4">我已经登录，为什么还会自动注销？</a> <br>
		<a href="#A5">我忘记了用户名/密码.</a> <br>
		<a href="#A6">为什么我已注册但仍不能登录？</a> <br>
		<a href="#A7">以前注册，但是现在不能登录？</a>
		<p><b>用户个人资料 &amp; 设置 </b><br>
		<a href="#B1">什么是个人资料？</a> <br>
		<a href="#B2">如何在帖子上添加签名？</a> <br>
		<a href="#B3">什么是头像？</a> <br>
		<a href="#B4">如何设置我的头像？ </a><br>
		<a href="#B5">为什么我需要登录才能发贴、浏览会员资料？</a> <br>
		<a href="?menu=Ranks">我想知道会员的等级名称？</a></p>
		<p><b>隐私 &amp; 安全</b> <br>
		<a href="#C1">如何更改密码？ </a><br>
		<a href="#C2">如何更改e-mail地址？</a> <br>
		<a href="#C3">要求设置的个人资料？</a> </p>
		<p><b>导航</b> <br>
		<a href="#D1">什么是论坛组？</a> <br>
		<a href="#D2">什么是论坛？</a> <br>
		<a href="#D3">什么是主题？ </a><br>
		<a href="#D4">主题图标是什么意思？</a> <br>
		<a href="#D5">我浏览论坛时看不见任何主题/帖子？</a> <br>
		<a href="#D6">为什么看不见我刚发的帖子？</a> <br>
		<a href="#D7">什么是固顶贴？ </a><br>
		<a href="#D8">什么是锁定贴</a> <br>
		<a href="#D9">当浏览论坛时我能将主题排序吗？ </a><br>
		<a href="#D10">论坛的RSS图标是什么？</a> <br>
		<a href="#D11">我不能登录以前访问过的论坛？</a> </p>
		<p><b>发贴 </b><br>
		<a href="#E1">我能使用HTML吗？</a> <br>
		<a href="#E2">什么是YBB代码？</a> <br>
		<a href="#E3">我可以在我的贴子里面添加附件吗？</a> <br>
		<a href="#E4">什么是表情图标？</a> <br>
		<a href="#E5">我怎样才能在板块中发表新贴子？</a> <br>
		<a href="#E6">怎么回复一个贴子？</a> <br>
		<a href="#E7">怎么样修改我的贴子？</a> <br>
		<a href="#E8">怎么删除我的贴子？</a> <br>
		<a href="#E9">为什么我的贴子中有些词被替换成了“***”？</a> <br>
		<a href="#E10">我怎么在我的贴子中加签名？</a> <br>
		<a href="#E11">怎么把头像添加到贴子中？</a> <br>
		<a href="#E12">我想知道论坛的发帖规则？</a></p>
		<p><b>用户权限</b> <br>
		<a href="#F1">什么是管理员？</a> <br>
		<a href="#F2">什么是版主？</a> <br>
		<a href="?menu=popedom">我想知道论坛所有人的权限？</a> </p>
		<p><b>关于 BBSXP</b> <br>
		<a href="#G1">什么是 BBSXP？</a> <br>
		<a href="#G2">谁在使用 BBSXP？</a> <br>
		<a href="#G3">谁创建了 BBSXP？ </a><br>
		<a href="#G4">在哪获取 BBSXP 的拷贝？</a> </p>
		</td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">注册 &amp; 登录 </td>
	</tr>
	<tr>
		<td class="a3"><a name="A1"></a><b>我为什么要注册？ </b><br>
		发帖、阅读其他人的帖子前是否需要登录，这取决于论坛管理员是否设置了论坛的游客权限。大多数设置允许不需要注册登录就可以阅读信息。通过注册您可以享受到所有的附加功能，例如：设置您的头像、
		收藏感兴趣的帖子、给用户发送邮件、私人留言、访问一些受保护的论坛等等，注册步骤也非常简单，一分钟就能完成。<br><a href="#top">返回顶部</a></a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A2"></a><b>怎样才能注册？ </b><br>
		要注册一个新帐号，你需要访问注册并且按照要求填写注册表单。您必须提供一个用户名和一个有效的电子信箱地址。管理员可能要求你指定密码，如果不需要指定密码，那么在注册完成后您将收到确认邮件。 
		<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A3"></a><b>我已经注册了用户名和密码，怎么登录？ </b><br>
		注册成功后，你将拥有一个用户名和密码，你可以访问登录页面并输入您的用户名和密码登录到论坛。 <br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A4"></a><b>我已经登录，为什么还会自动注销？</b> <br>
		当你登录的时候，如果没有选中 &quot;自动登录&quot; 复选框，您会在离开时自动注销。假如您想一直保持登录状态，请在登录的时候选中 &quot;自动登录&quot; 复选框。 
		<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A5"></a><b>我忘记了密码.</b> <br>
		假如您忘记了密码，您可以访问 &quot;找回密码&quot; 
		页输入您注册时的电子信箱地址，会有一份有新密码的邮件发送到您的注册信箱。因为密码是不可逆加密存储，所以我们无法找回您的原始密码。一旦您收到您的新密码，您可以登录并修改您的密码。 
		<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A6"></a><b>已注册但是仍然不能登录？ </b><br>
		如果你已注册但是仍然不能登录,请确认你是否有一个合法的用户名和密码。 
		如果你确认你的用户名和密码是正确的，但仍然无法登入，则可能你的帐号需要激活或你的帐号在停用中。 如果是这个原因，你最好联系版主或是管理员。 <br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A7"></a><b>已注册，但是现在不能登录？ </b><br>
		请先检查你的用户名和密码是否正确。如果仍然不能登录，你的帐号可能由于长期未登录已被删除。请与论坛管理员或者版主联系。 <br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">用户资料 &amp; 设置 </td>
	</tr>
	<tr>
		<td class="a3"><a name="B1"></a><b>什么是个人资料？</b> <br>
		个人资料就是你的帐号信息，帐号控制着你如何查看论坛上的帖子。它包含着你所发表过帖子的详细资料，你乐意与他人共享的私人信息如你的网址、博客地址，以及一些控制着你如何与论坛相互作用的设置
		。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B2"></a><b>如何在帖子里添加签名？ </b><br>
		签名就是附加在你发表于论坛上的每一篇帖子后面的信息. 你可以在个人资料页面编辑你的签名. 此签名将显示在你所发表的任何信息的尾部。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B3"></a><b>什么是头像？</b> <br>
		头像是一个帖子附带的形象。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B4"></a><b>我如何设置头像？ </b><br>
		如果管理员已启用头像，当你查看个人资料时可以看见头像区。在这里你可以选择一个你喜欢的头像或是上传或是输入一个图片的网址来作为你的头像。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B5"></a><b>为什么我需要登录才能发贴、浏览会员资料？ </b>
		<br>
		根据管理员如何配置论坛，你可能需要登录才能浏览/使用论坛的某些区域，这主要是为了保护用户的隐私。<br><a href="#top">返回顶部</a> 
		<br>
		 </td>
	</tr>

	</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">私人信息 &amp; 安全 </td>
	</tr>
	<tr>
		<td class="a3"><a name="C1"></a><b>如何更改密码？ </b><br>
		登录以后，你能在个人资料页面里更改密码。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="C2"></a><b>如何更改email地址？</b> <br>
		你可以在控制面板－〉密码修改里更改Email地址。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="C3"></a><b>要求设置的个人资料？ </b><br>
		只有电子邮件地址是必须填的。 这个邮件地址用来向你发送你预定的论坛信息，及向你发送忘记的用户名或密码。其它的资料都是不是必须填的。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">导航 </td>
	</tr>
	<tr>
		<td class="a3"><a name="D1"></a><b>什么是论坛组？</b> <br>
		论坛组是个包含几个相关版面的大类。一个论坛组包括一个或多个子版面。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D2"></a><b>什么是版块？ </b><br>
		版块是所讨论一系列主题的组。一个版块包括多篇帖子或者多个副版。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D3"></a><b>什么是主题？</b> <br>
		一个主题包括一堆相关的帖子，它们都在讨论一个问题。第一篇帖子成为一个主题，其它回帖都跟在它后面。主题里面还记录一共多少回帖及谁最后回帖等信息。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D4"></a><b>主题图标代表什么意思？</b> 
		<table border=0>
			<tr>
				<td><img src="images/f_New.gif" border="0" alt="有回复的主题"> 
				打开主题</td>
				<td><img src="images/f_top.gif" border="0"> 
		固顶主题</td>
				<td><img src="images/f_poll.gif" border="0"> 
				投票主题</td>
				<td><img src="images/Topicgood.gif"> 精华主题</td>
			</tr>
			<tr>
				<td><img src="images/f_norm.gif" border="0" alt="没有回复的主题"> 
				打开主题</td>
				<td><img src="images/f_locked.gif" border="0"> 
				锁定主题</td>
				<td>
				<img src="images/f_hot.gif" border="0" alt="回复数达到 <%=SiteSettings("PopularPostThresholdPosts")%> 或者点击数达到 <%=SiteSettings("PopularPostThresholdViews")%>"> 
				热门主题 </td>
				<td><img src="images/my.gif"> 自己发表的主题</td>
			</tr>
		</table>
<a href="#top">返回顶部</a></td>
	</tr>

	<tr>
		<td class="a3"><a name="D5"></a><b>我在浏览论坛时，为什么看不见任何主题/帖子？</b> <br>
		如果没人发帖子，当然就看不到。要么就是你设置了帖子过滤器，过滤了你不想看的帖子。比如你可能设置了只显示某个日期以后的帖子，例如只显示两星期以内的帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D6"></a><b>为什么我刚发的帖子看不见？</b> <br>
		论坛可以设置编辑（版主）审查功能。如果你在一个需要审查的论坛发贴，系统会给你一条信息表明你的帖子正在等待审查。一旦审查员通过审查，你的帖子才可见。审查员可以移动、修改或删除你的帖子－－如果他们认为你的帖子不适合在这个论坛发表的话。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D7"></a><b>什么是固顶帖子？</b> <br>
		就是总出现在论坛帖子列表最前面的帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D8"></a><b>什么是锁定贴？</b> <br>
		就是不能回复的帖子。发贴人、管理员、版主都可以设置锁定一个帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D9"></a><b>看贴时，可否排序？</b> <br>
		可以，按作者排序，回复排序，点击量排序，最后回帖排序都行。默认是按照时间排序，最新的帖子在上面。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D10"></a><b>论坛的RSS图标是什么东东？</b> <br>
		这个RSS图标是用来连接论坛的RSS种子。RSS需要使用专门的阅读器来阅读。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D11"></a><b>我怎么连不上论坛？</b> <br>
		如果你登录以前访问过的论坛，却出现一个未知的错误，可能有两个原因。一个原因是这个论坛要需要注册才能登录，另一个原因就是这个论坛已经关闭了。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">发贴 </td>
	</tr>
	<tr>
		<td class="a3"><a name="E1"></a><b>可以用HTML代码吗？</b> <br>
		可以，但是不能直接输入到编辑框内。 如果你用IE浏览，默认编辑框是一个高级HTML编辑器，可以自动加入HTML代码。如果你用其它浏览器浏览，将会使用一个普通的HTML编辑器。 
<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E2"></a><b>什么是YBB代码？</b> <br>
		就是用于把文字搞出很多花样的语法。<a href="?menu=ybb">点击这里可查看YBB代码详细使用方法</a>。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>
	<tr>
		<td class="a3"><a name="E3"></a><b>发贴可以带附件吗？</b> <br>
		可以，但是要管理员开放此功能。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E4"></a><b>什么是表情图标？</b> <br>
		就是在发贴时可以插入的一些小图标，例如：笑脸、哭脸等。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E5"></a><b>怎么发新贴？</b> <br>
		点论坛上的“发表新主题”图片，会打开一个发贴的页面（假设你已经登录，否则会要求你先登录）。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E6"></a><b>怎么回复一个贴子？</b> <br>
		点“回复”按钮或“引用”按钮就可以回复帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E7"></a><b>怎么样修改我的贴子？</b> <br>
		点击“编辑”这个图象按钮你可以编辑你的帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E8"></a><b>怎么删除我的贴子？</b> <br>
		在你发表的帖子旁边看见一个删除图象按钮。如果你发表的帖子已经有了一个或多个回复，你将不能删除这个帖子。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E9"></a><b>为什么我的贴子中有些词被替换成了“***”？</b> <br>
		管理员也许已经为帖子设置了词语过滤器。当词语过滤器启用时，某些被认为无礼或冒犯的词语将被字母***替换。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E10"></a><b>我怎么在我的贴子中加签名？</b> <br>
		请在控制面板－〉资料修改里设置签名。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E11"></a><b>怎么把头像添加到贴子中？</b> <br>
		请在控制面板－〉资料修改里设置头像。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3">
		
		
<a name="E12"></a><b>发帖规则
</b>
<table border="0" cellpadding="3" cellspacing="1">
	<tr>
		<td>1、发表新主题<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralAddThread")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralAddThread")%></font><br>
	</tr>
		<td>2、回复帖子<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralAddPost")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralAddPost")%></font></tr>
		<td>3、加为精华帖<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralAddValuedPost")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralAddValuedPost")%></font></tr>
		<td>4、删除主题帖<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralDeleteThread")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralDeleteThread")%></font></tr>
		<td>5、删除回复帖<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralDeletePost")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralDeletePost")%></font></tr>
		<td>6、取消精华帖<br>
		　　&nbsp; 经验值：<font color="red"><%=SiteSettings("IntegralDeleteValuedPost")%></font><br>
		　　&nbsp; 金币值：<font color="red"><%=SiteSettings("IntegralDeleteValuedPost")%></font></tr>
	</table>
<a href="#top">返回顶部</a></td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">用户权限 </td>
	</tr>
	<tr>
		<td class="a3"><a name="F1"></a><b>什么是管理员？</b> <br>
		管理员拥有论坛的最高级权限。默认情况下，管理员具有在论坛里执行任何操作的所有权限，例如，编辑帖子、批准用户、创建新版块等等。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="F2"></a><b>什么是版主？ </b><br>
		版主拥有论坛的第二级权限。版主能够在所属的版块里执行任何操作。这包括批准帖子、转移帖子、删除帖子、编辑帖子。如果你有关于某版块的问题，最好的解决办法就是同版主联系。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">关于 BBSXP？ </td>
	</tr>
	<tr>
		<td class="a3"><a name="G1"></a><b>什么是 BBSXP？</b> <br>
		BBSXP 是一个使用ASP语言开发的论坛系统。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G2"></a><b>谁在使用 BBSXP？</b> <br>
		很多公开或私人的组织都使用了这种论坛系统。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G3"></a><b>谁创建了 BBSXP？</b> <br>
		BBSXP 是由YUZI开发团队开发。你可以查阅 <a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a> 得到更多的资讯。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G4"></a><b>在哪获取BBSXP的拷贝？ </b><br>
		访问 <a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a> 可以下载 BBSXP 的最新版本。<br><a href="#top">返回顶部</a><br>
		 </td>
	</tr>
</table>
<%end if%>

<%htmlend%>