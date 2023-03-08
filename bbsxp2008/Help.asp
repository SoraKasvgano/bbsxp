<!-- #include file="Setup.asp" -->
<%
HtmlTop
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → <a href="help.asp">使用帮助</a></div>
<%if Request("menu")="Ranks"then%>



<%
	RoleGetRow = FetchEmploymentStatusList("Select RoleID,Name from ["&TablePrefix&"Roles] where RoleID >= 0 order by RoleID")
	If IsArray(RoleGetRow) Then
		For i=0 to ubound(RoleGetRow,2)
			Set Rs=Execute("select * from ["&TablePrefix&"Ranks] where RoleID="&RoleGetRow(0,i)&"")
			If not Rs.eof Then
%>
<table border="0" cellpadding="5" cellspacing="1" class=CommonListArea width=50%>
	<tr class=CommonListTitle>
		<td colspan="4" align="center"><%=RoleGetRow(1,i)%></td>
	</tr>
	<tr class=CommonListHeader>
		<td width="100" align="center">ID</td>
		<td>用户等级</td>
		<td width="150">最少发帖</td>
	</tr>
<%
			Do While Not Rs.eof
%>
	<tr class="CommonListCell">
		<td align="center"><%=Rs("RankID")%></td>
		<td><%=Rs("RankName")%></td>
		<td><%=Rs("PostingCountMin")%></td>
	</tr>
<%
				Rs.movenext
			Loop
%>
</table>
<%
			End If
		Next
	End If
	RoleGetRow=null
%>

<%elseif Request("menu")="BBCode"then%>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>什么是BBCode(Bulletin Board Code)？ </td>
	</tr>
	<tr class="CommonListCell">
		<td>　　 BB 代码是HTML的一个变种。您也许已经对它很熟悉了。一般情况下，如果允许您用HTML，也就可以使用 BB 代码。即使您的讨论区不能让您使用HTML，您也可以使用 BB 代码。 由于要求使用的编码很少，即使可以使用HTML，您可能也想使用 BB 代码，因为代码错误的可能性大大减小了。
		</td>
	</tr>
	<tr class=CommonListTitle>
		<td>ED2K链接</td>
	</tr>
	<tr class="CommonListCell">
		<td>在您的信息里加入ED2K链接，只要按下列方式套入就可以了<br /><br />
		<font color="red">[url]</font><a target="_blank" href="ed2k://|file|eMule.0.49a.exe|3073303|E38526D47AEE487E5FB317F411851631|s=http://www.ed2000.com/download/eMule.exe|/
">ed2k://|file|eMule.exe|3073303|E38526D47AEE487E5FB317F411851631|/</a><font color="red">[/url]</font>
		<br />或者 <br />
		<font color="red">[ed2k]</font><a target="_blank" href="ed2k://|file|eMule.0.49a.exe|3073303|E38526D47AEE487E5FB317F411851631|s=http://www.ed2000.com/download/eMule.exe|/">ed2k://|file|eMule.exe|3073303|E38526D47AEE487E5FB317F411851631|/</a><font color="red">[/ed2k]</font><br /><br />
		注意URL的&quot;ed2k://&quot;这一部分是必须的。</td>
	</tr>
	<tr class=CommonListTitle>
		<td>URL超级链接</td>
	</tr>
	<tr class="CommonListCell">
		<td>在您的信息里加入超级链接，只要按下列方式套入就可以了<br /><br />
		<font color="red">[url]</font><a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a><font color="red">[/url]</font>
		<br />或者 <br />
		<font color="red">[url=http://www.yuzi.net]</font><a target="_blank" href="http://www.yuzi.net">YUZI工作室</a><font color="red">[/url]</font><br /><br />
		按上例套入， BB 代码会自动对URL产生链接，并保证当用户点击新的窗口时这个链接是打开着的。注意URL的&quot;http://&quot;这一部分是随意的。</td>
	</tr>
	<tr class=CommonListTitle>
		<td>电子邮件链接</td>
	</tr>
	<tr class="CommonListCell">
		<td>在您的信息里加入电子邮件的超级链接，只要按照下例套入就可以了<br /><br />
		<font color="red">[email]</font><a href="mailto:yuzi@yuzi.net">yuzi@yuzi.net</a><font color="red">[/email]</font><br /><br />
按上例套入， BB 代码会对电子邮件自动产生链接。 </td>
	</tr>
	<tr class=CommonListTitle>
		<td>粗体与斜体</td>
	</tr>
	<tr class="CommonListCell">
		<td>您可以使用 [b] [i] 等这些标志以达到在帖子中使用相应的效果<br />
		<br />
		您好,<font color="red">[b]</font><strong>管理员</strong><font color="red">[/b]</font><br />
		您好,<font color="red">[i]</font><em>管理员</em><font color="red">[/i]</font><br />
		您好,<font color="red">[u]</font><u>管理员</u><font color="red">[/u]</font><br />
		您好,<font color="red">[strike]</font><strike>管理员</strike><font color="red">[/strike]</font></td>
	</tr>
	<tr class=CommonListTitle>
		<td>移动文字</td>
	</tr>
	<tr class="CommonListCell">
		<td height="42">在您的信息里加入移动文字，只要按下例套入就可以了<br /><br />
		<font color="red">[marquee]</font>移动文字<font color="red">[/marquee]</font><br />
		
		<marquee>移动文字</marquee></td>
	</tr>
	<tr class=CommonListTitle>
		<td>引用其他信息</td>
	</tr>
	<tr class="CommonListCell">
		<td>引用一些人的帖子，只要剪切和拷贝然后按下例套入就可以了<br />
		<br /><font color="red">
		[QUOTE]</font>别问您的国家能为您作什么......<br />
		问您能为您的国家作什么？<font color="red">[/QUOTE]</font>
		<blockquote>别问您的国家能为您作什么......<br />
		问您能为您的国家作什么？</blockquote>
		</td>
	</tr>
	<tr class=CommonListTitle>
		<td>彩色文字</td>
	</tr>
	<tr class="CommonListCell">
		<td><font color="red">[COLOR=green]</font><font color="#008000">彩色文字</font><font color="red">[/COLOR]</font>　　“green”代表文字的颜色<font color="red"><br />
		[FONT=黑体]</font><font face="黑体">彩色文字</font><font color="red">[/FONT]</font>　　“黑体”代表文字的字体<font color="red"><br />
		[SIZE=5]</font><font size="5">彩色文字</font><font color="red">[/SIZE]</font>　　“5”代表文字的大小</td>
	</tr>
	<tr class=CommonListTitle>
		<td>特别注意</td>
	</tr>
	<tr>
		<td height="76" bgcolor="FFFFFF">
		<p>您不可以同时使用<font face="Verdana, Arial">HTML</font>和 BB 代码的同一种功能。并且注意 BB 代码不对大小写敏感。所以您可以用<font color="red">[URL]</font> 
		或 <font color="red">[url]</font> 
		<br />
		<font color="800000">不正确的 BB 代码使用：</font><br />
		<font color="red">[url]</font> www.yuzi.net <font color="red">[/url]</font> 不要在括号和您输入的文字之间有空格<br />
		<font color="red">[email]</font>yuzi@yuzi.net<font color="red">[email]</font>
		在结束时，不要忘了在括号内加入斜杠<font color="red">[/email]</font>
		</p>
		</td>
	</tr>
</table>


<%else%>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>常见问题解答 </td>
	</tr>
	<tr class="CommonListCell">
		<td><b>注册和登录</b> <br />
		<a href="#A1">我为什么要注册？ </a><br />
		<a href="#A2">怎样才能注册？ </a><br />
		<a href="#A3">我已经注册了用户名和密码，怎么登录？ </a><br />
		<a href="#A4">我已经登录，为什么还会自动注销？</a> <br />
		<a href="#A5">我忘记了用户名/密码.</a> <br />
		<a href="#A6">为什么我已注册但仍不能登录？</a> <br />
		<a href="#A7">以前注册，但是现在不能登录？</a>
		<p><b>用户个人资料 &amp; 设置 </b><br />
		<a href="#B1">什么是个人资料？</a> <br />
		<a href="#B2">如何在帖子上添加签名？</a> <br />
		<a href="#B3">什么是头像？</a> <br />
		<a href="#B4">如何设置我的头像？ </a><br />
		<a href="#B5">为什么我需要登录才能发帖、浏览用户资料？</a> <br />
		<a href="#B6">在查看帖子页面的个人资料显示栏中为什么没有声望和活跃天数的图标？ </a> <br />
		<a href="?menu=Ranks">我想知道用户的等级名称？</a></p>
		<p><b>隐私 &amp; 安全</b> <br />
		<a href="#C1">如何更改密码？ </a><br />
		<a href="#C2">如何更改e-mail地址？</a> <br />
		<a href="#C3">要求设置的个人资料？</a> </p>
		<p><b>导航</b> <br />
		<a href="#D1">什么是论坛组？</a> <br />
		<a href="#D2">什么是论坛？</a> <br />
		<a href="#D3">什么是主题？ </a><br />
		<a href="#D4">主题图标是什么意思？</a> <br />
		<a href="#D5">我浏览论坛时看不见任何主题/帖子？</a> <br />
		<a href="#D6">为什么看不见我刚发的帖子？</a> <br />
		<a href="#D7">什么是置顶帖？ </a><br />
		<a href="#D8">什么是锁定帖</a> <br />
		<a href="#D9">当浏览论坛时我能将主题排序吗？ </a><br />
		<a href="#D10">论坛的RSS图标是什么？</a> <br />
		<a href="#D11">我不能登录以前访问过的论坛？</a> </p>
		<p><b>发帖 </b><br />
		<a href="#E1">我能使用HTML吗？</a> <br />
		<a href="#E2">什么是 BB 代码？</a> <br />
		<a href="#E3">我可以在我的帖子里面添加附件吗？</a> <br />
		<a href="#E4">什么是表情图标？</a> <br />
		<a href="#E5">我怎样才能在版块中发表新帖子？</a> <br />
		<a href="#E6">怎么回复一个帖子？</a> <br />
		<a href="#E7">怎么样修改我的帖子？</a> <br />
		<a href="#E8">怎么删除我的帖子？</a> <br />
		<a href="#E9">为什么我的帖子中有些词被替换成了“***”？</a> <br />
		<a href="#E10">我怎么在我的帖子中加签名？</a> <br />
		<a href="#E11">怎么把头像添加到帖子中？</a> <br />
		<a href="#E12">我想知道论坛的发帖规则？</a></p>
		<p><b>用户组和权限</b> <br />
		<a href="#F1">什么是权限？</a> <br />
		<a href="#F2">什么是管理员？</a> <br />
		<a href="#F3">什么是超级版主？</a> <br />
		<a href="#F4">什么是版主？</a><br />
		<a href="#F5">什么是论坛角色？</a> <br />
		<a href="#F6">我如何加入一个论坛角色或组？</a> <br />
		<p><b>关于 BBSXP</b> <br />
		<a href="#G1">什么是 BBSXP？</a> <br />
		<a href="#G2">谁在使用 BBSXP？</a> <br />
		<a href="#G3">谁创建了 BBSXP？ </a><br />
		<a href="#G4">在哪获取 BBSXP 的拷贝？</a> </p>
		</td>
	</tr>
</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>注册 &amp; 登录 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A1"></a><b>我为什么要注册？ </b><br />
		发帖、阅读其他人的帖子前是否需要登录，这取决于论坛管理员是否设置了论坛的游客权限。大多数设置允许不需要注册登录就可以阅读信息。通过注册您可以享受到所有的附加功能，例如：设置您的头像、
		收藏感兴趣的帖子、给用户发送邮件、私人留言、访问一些受保护的论坛等等，注册步骤也非常简单，一分钟就能完成。<br /><a href="#top">返回顶部</a></a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A2"></a><b>怎样才能注册？ </b><br />
		要注册一个新帐号，你需要访问注册并且按照要求填写注册表单。您必须提供一个用户名和一个有效的电子信箱地址。管理员可能要求你指定密码，如果不需要指定密码，那么在注册完成后您将收到确认邮件。 
		<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A3"></a><b>我已经注册了用户名和密码，怎么登录？ </b><br />
		注册成功后，你将拥有一个用户名和密码，你可以访问登录页面并输入您的用户名和密码登录到论坛。 <br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A4"></a><b>我已经登录，为什么还会自动注销？</b> <br />
		当你登录的时候，如果没有选中 &quot;自动登录&quot; 复选框，您会在离开时自动注销。假如您想一直保持登录状态，请在登录的时候选中 &quot;自动登录&quot; 复选框。 
		<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A5"></a><b>我忘记了密码.</b> <br />
		假如您忘记了密码，您可以访问 &quot;找回密码&quot; 
		页输入您注册时的电子信箱地址，会有一份有新密码的邮件发送到您的注册信箱。因为密码是不可逆加密存储，所以我们无法找回您的原始密码。一旦您收到您的新密码，您可以登录并修改您的密码。 
		<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A6"></a><b>已注册但是仍然不能登录？ </b><br />
		如果你已注册但是仍然不能登录,请确认你是否有一个合法的用户名和密码。 
		如果你确认你的用户名和密码是正确的，但仍然无法登入，则可能你的帐号需要激活或你的帐号在停用中。 如果是这个原因，你最好联系版主或是管理员。 <br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="A7"></a><b>已注册，但是现在不能登录？ </b><br />
		请先检查你的用户名和密码是否正确。如果仍然不能登录，你的帐号可能由于长期未登录已被删除。请与论坛管理员或者版主联系。 <br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>用户资料 &amp; 设置 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="B1"></a><b>什么是个人资料？</b> <br />
		个人资料就是你的帐号信息，帐号控制着你如何查看论坛上的帖子。它包含着你所发表过帖子的详细资料，你乐意与他人共享的私人信息如你的网址、博客地址，以及一些控制着你如何与论坛相互作用的设置
		。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="B2"></a><b>如何在帖子里添加签名？ </b><br />
		签名就是附加在你发表于论坛上的每一篇帖子后面的信息. 你可以在个人资料页面编辑你的签名. 此签名将显示在你所发表的任何信息的尾部。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="B3"></a><b>什么是头像？</b> <br />
		头像是一个帖子附带的形象。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="B4"></a><b>我如何设置头像？ </b><br />
		如果管理员已启用头像，当你查看个人资料时可以看见头像区。在这里你可以选择一个你喜欢的头像或是上传或是输入一个图片的网址来作为你的头像。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="B5"></a><b>为什么我需要登录才能发帖、浏览用户资料？ </b>
		<br />
		根据管理员如何配置论坛，你可能需要登录才能浏览/使用论坛的某些区域，这主要是为了保护用户的隐私。<br /><a href="#top">返回顶部</a> 
		<br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="B6"></a><b>在查看帖子页面的个人资料显示栏中为什么没有声望和活跃天数的图标？ </b>
		<br />
		用户声望不少于 4 就会显示声望图标，图标演示：<%=ShowReputation(4)%>&nbsp;<%=ShowReputation(251)%>&nbsp;<%=ShowReputation(10001)%> <br />
		活跃天数不少于 5 就会显示活跃图标，图标演示：<%=ShowUserActivityDay(5)%>&nbsp;<%=ShowUserActivityDay(32)%>&nbsp;<%=ShowUserActivityDay(320)%>
<br /><a href="#top">返回顶部</a> 
		<br />
		 </td>
	</tr>


	</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>私人信息 &amp; 安全 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="C1"></a><b>如何更改密码？ </b><br />
		登录以后，你能在个人资料页面里更改密码。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="C2"></a><b>如何更改email地址？</b> <br />
		你可以在控制面板－〉密码修改里更改Email地址。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="C3"></a><b>要求设置的个人资料？ </b><br />
		只有电子邮箱地址是必须填的。 这个邮箱地址用来向你发送你预定的论坛信息，及向你发送忘记的用户名或密码。其它的资料都是不是必须填的。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>导航 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="D1"></a><b>什么是论坛组？</b> <br />
		论坛组是个包含几个相关版面的大类。一个论坛组包括一个或多个子版面。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="D2"></a><b>什么是版块？ </b><br />
		版块是所讨论一系列主题的组。一个版块包括多篇帖子或者多个副版。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="D3"></a><b>什么是主题？</b> <br />
		一个主题包括一堆相关的帖子，它们都在讨论一个问题。第一篇帖子成为一个主题，其它回帖都跟在它后面。主题里面还记录一共多少回帖及谁最后回帖等信息。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="D4"></a><b>主题图标代表什么意思？</b> 
		<table border=0 width="100%">
			<tr>
				<td><img src="images/topic-announce.gif" border="0" /> 公告主题</td>
				<td><img src="images/topic-pinned.gif" border="0" /> 置顶主题</td>
				<td><img src="images/topic-popular.gif" /> 精华主题</td>
				<td><img src="images/topic-locked.gif" border="0" /> 锁定主题</td>
				<td><img src="images/topic-poll.gif" border="0" /> 投票主题</td>
				<td><img src="images/topic-hot.gif" border="0" title="回复数达到 <%=SiteConfig("PopularPostThresholdPosts")%> 或者点击数达到 <%=SiteConfig("PopularPostThresholdViews")%>" /> 热门主题 </td>
				<td><img src="images/topic.gif" border="0" /> 普通主题</td>

			</tr>
		</table>
<a href="#top">返回顶部</a></td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D5"></a><b>我在浏览论坛时，为什么看不见任何主题/帖子？</b> <br />
		如果没人发帖子，当然就看不到。要么就是你设置了帖子过滤器，过滤了你不想看的帖子。比如你可能设置了只显示某个日期以后的帖子，例如只显示两星期以内的帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D6"></a><b>为什么我刚发的帖子看不见？</b> <br />
		论坛可以设置编辑（版主）审核功能。如果你在一个需要审核的论坛发帖，系统会给你一条信息表明你的帖子正在等待审核。一旦审核员通过审核，你的帖子才可见。审核员可以移动、修改或删除你的帖子－－如果他们认为你的帖子不适合在这个论坛发表的话。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D7"></a><b>什么是置顶帖子？</b> <br />
		就是总出现在论坛帖子列表最前面的帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D8"></a><b>什么是锁定帖？</b> <br />
		就是不能回复的帖子。发帖人、管理员、版主都可以设置锁定一个帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D9"></a><b>看帖时，可否排序？</b> <br />
		可以，按作者排序，回复排序，点击量排序，最后回帖排序都行。默认是按照时间排序，最新的帖子在上面。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D10"></a><b>论坛的RSS图标是什么东东？</b> <br />
		这个RSS图标是用来连接论坛的RSS种子。RSS需要使用专门的阅读器来阅读。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="D11"></a><b>我怎么连不上论坛？</b> <br />
		如果你登录以前访问过的论坛，却出现一个未知的错误，可能有两个原因。一个原因是这个论坛要需要注册才能登录，另一个原因就是这个论坛已经关闭了。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>发帖 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="E1"></a><b>可以用HTML代码吗？</b> <br />
		可以，但是不能直接输入到编辑框内。 如果你用IE浏览，默认编辑框是一个高级HTML编辑器，可以自动加入HTML代码。如果你用其它浏览器浏览，将会使用一个普通的HTML编辑器。 
<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E2"></a><b>什么是 BB 代码？</b> <br />
		就是用于把文字搞出很多花样的语法。<a href="?menu=BBCode">点击这里可查看 BB 代码详细使用方法</a>。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="E3"></a><b>发帖可以带附件吗？</b> <br />
		可以，但是要管理员开放此功能。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E4"></a><b>什么是表情图标？</b> <br />
		就是在发帖时可以插入的一些小图标，例如：笑脸、哭脸等。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E5"></a><b>怎么发新帖？</b> <br />
		点论坛上的“发表新主题”图片，会打开一个发帖的页面（假设你已经登录，否则会要求你先登录）。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E6"></a><b>怎么回复一个帖子？</b> <br />
		点“回复”按钮或“引用”按钮就可以回复帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E7"></a><b>怎么样修改我的帖子？</b> <br />
		点击“编辑”这个图象按钮你可以编辑你的帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E8"></a><b>怎么删除我的帖子？</b> <br />
		点击“删除”这个图象按钮你可以删除你的帖子。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E9"></a><b>为什么我的帖子中有些词被替换成了“***”？</b> <br />
		管理员也许已经为帖子设置了词语过滤器。当词语过滤器启用时，某些被认为无礼或冒犯的词语将被字母***替换。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E10"></a><b>我怎么在我的帖子中加签名？</b> <br />
		请在控制面板－〉资料修改里设置签名。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="E11"></a><b>怎么把头像添加到帖子中？</b> <br />
		请在控制面板－〉资料修改里设置头像。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td>
		
		
<a name="E12"></a><b>发帖规则
</b>
<table border="0" cellpadding="5" cellspacing="1">
	<tr>
		<td>1、发表新主题<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralAddThread")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralAddThread")%></font><br />
	</tr>
		<td>2、回复帖子<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralAddPost")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralAddPost")%></font></tr>
		<td>3、加为精华帖<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralAddValuedPost")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralAddValuedPost")%></font></tr>
		<td>4、删除主题帖<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralDeleteThread")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralDeleteThread")%></font></tr>
		<td>5、删除回复帖<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralDeletePost")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralDeletePost")%></font></tr>
		<td>6、取消精华帖<br />
		　　&nbsp; 经验值：<font color="red"><%=SiteConfig("IntegralDeleteValuedPost")%></font><br />
		　　&nbsp; 金钱值：<font color="red"><%=SiteConfig("IntegralDeleteValuedPost")%></font></tr>
	</table>
<a href="#top">返回顶部</a></td>
	</tr>
</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>用户组和权限</td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="F1"></a><b>什么是权限？</b> <br />
当你浏览论坛时，权限控制着你可以做什么和不可以做什么。当使用论坛的默认观点时，权限被显示在每页的右下角（可应用权限的地方）。你被授予的权限控制着你在论坛浏览的方方面面。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="F2"></a><b>什么是管理员？</b> <br />
		管理员拥有论坛的最高级权限。默认情况下，管理员具有在论坛里执行任何操作的所有权限。<br /><a href="#top">返回顶部</a></td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="F3"></a><b>什么是超级版主？</b> <br />
		超级版主拥有论坛的第二级权限。默认情况下，超级版主具有在论坛前台执行任何操作的所有权限，例如帖子公告、审核主题、查看在线用户IP等等。<br /><a href="#top">返回顶部</a></td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="F4"></a><b>什么是版主？ </b><br />
		版主拥有的版块内的任何权限。版主能够在所属的版块里执行任何操作；论坛组版主能够在所属论坛组下的所有的版块里执行任何的操作。这包括批准帖子、转移帖子、删除帖子、编辑帖子。如果你有关于某版块的问题，最好的解决办法就是同版主联系。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="F5"></a><b>什么是论坛角色？</b> <br />
		论坛角色是指为了指派权限而将普通用户进行分组。由于能给用户指派角色，以及权限被应用到基于这些角色的论坛上，管理员和版主的工作变得更简单起来。 <br /><a href="#top">返回顶部</a></td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="F6"></a><b>我如何加入一个论坛角色或组？</b> <br />
管理员将论坛角色指派给各用户。如果你希望参与一个特定的论坛角色，请与管理员联系以获取更多的信息。<br /><a href="#top">返回顶部</a></td>
	</tr>

	</table>
<p>　</p>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td>关于 BBSXP？ </td>
	</tr>
	<tr class="CommonListCell">
		<td><a name="G1"></a><b>什么是 BBSXP？</b> <br />
		BBSXP 是一个使用ASP语言开发的论坛系统。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="G2"></a><b>谁在使用 BBSXP？</b> <br />
		很多公开或私人的组织都使用了这种论坛系统。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="G3"></a><b>谁创建了 BBSXP？</b> <br />
		BBSXP 是由YUZI开发团队开发。你可以查阅 <a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a> 得到更多的资讯。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>

	<tr class="CommonListCell">
		<td><a name="G4"></a><b>在哪获取BBSXP的拷贝？ </b><br />
		访问 <a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a> 可以下载 BBSXP 的最新版本。<br /><a href="#top">返回顶部</a><br />
		 </td>
	</tr>
</table>
<%end if%>

<%HtmlBottom%>