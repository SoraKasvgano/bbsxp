<!-- #include file="Setup.asp" -->
<%
UserName=HTMLEncode(Request("UserName"))
UserID=RequestInt("UID")


if UserID>0 then
	sql="Select * from ["&TablePrefix&"Users] where UserID="&UserID&""
else
	sql="Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'"
end if
Set Rs=Execute(sql)
	if Rs.eof then Alert("对不起，不存在你要查看的用户资料")


HtmlHeadTitle="查看"&Rs("UserName")&"的资料"
HtmlTop
if SiteConfig("RequireAuthenticationForProfileViewing")=1 and CookieUserName=empty then error("您必须<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>后才能浏览个人资料")

%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> → 查看用户“<%=Rs("UserName")%>”资料</div>

<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td valign=top width="200">
			<table height="100%" width="95%" cellspacing=1 cellpadding=5 border=0 class=CommonListArea>
				<tr class=CommonListTitle>
					<td colspan="2"><%=Rs("UserName")%></td>
				</tr>
				<tr class="CommonListCell">
					<td colspan="2"><%if SiteConfig("EnableAvatars")=1 then%><center><img src="<%=Rs("UserFaceUrl")%>"  style="max-width:<%=SiteConfig("AvatarWidth")%>px;max-height:<%=SiteConfig("AvatarHeight")%>px;" /></center><%end if%>　<table border="0" cellspacing="0">

	<tr><td>用户状态：<%=ShowUserAccountStatus(RS("UserAccountStatus"))%></td></tr>
	<tr><td>角　　色：<%=ShowRole(RS("UserRoleID"))%></td></tr>
	<tr><td>注册日期：<%=FormatDateTime(Rs("UserRegisterTime"),2)%></td></tr>
	<tr><td>最近活动：<%=FormatDateTime(Rs("UserActivityTime"),2)%></td></tr>
	
	<tr><td>活跃天数：<%=Rs("UserActivityDay")%></td></tr>
	
	<tr><td>声　　望：<%=Rs("Reputation")%></td></tr>
	<tr><td>发 帖 数：<%=Rs("TotalPosts")%></td></tr>
	<tr><td>金　　钱：<%=Rs("UserMoney")%></td></tr>
	<tr><td>经 验 值：<%=Rs("experience")%></td></tr>

						</table>
					</td>
				</tr>
				<tr class=CommonListTitle>
					<td colspan="2">选项</td>
				</tr>
				<tr class="CommonListCell">
					<td>
						<table  border="0">




							
							<%if CookieUserName<>"" then
							
							if CookieUserName=Rs("UserMate") or CookieUserName=Rs("UserName") and Rs("UserMate")<>"" then%>
							<tr>
								<td><img src=images/divorce.gif /></td>
								<td><span id="UserMate"><a href="javascript:Ajax_CallBack(false,'UserMate','UserMate.asp?menu=divorce');">与 <%=Rs("UserMate")%> 离婚</a></span></td>
							</tr>
							<%else%>
							<tr>
								<td><img src=images/Marry.gif /></td>
								<td><span id="UserMate"><a href="javascript:Ajax_CallBack(false,'UserMate','UserMate.asp?menu=MarryRequest&TargetUser=<%=escape(Rs("UserName"))%>');">向 <%=Rs("UserName")%> 求婚</a></span></td>
							</tr>
							<%
							end if
							end if
							
							
							if CookieUserName=Rs("UserName") then%>
							<tr>
								<td><img src=images/edit.gif /></td>
								<td><a href="EditProfile.asp">编辑我的资料</a></td>
							</tr>
							<%end if%>
							<tr>
								<td><img src=images/favorite.gif /></td>
								<td><a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("UserName")%>',true);">将 <%=Rs("UserName")%> 加为好友</a></td>
							</tr>
							<%if CookieUserName<>empty then%>
							<tr>
								<td><img src=images/privatemessage.gif /></td>
								<td><a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&RecipientUserName=<%=Rs("UserName")%>',600,350);">给 <%=Rs("UserName")%> 发送讯息</a></td>
							</tr>
							<%	if SiteConfig("EnableReputation")=1 then%>
							<tr>
								<td><img src=images/reputation.gif /></td>
								<td><a href="javascript:BBSXP_Modal.Open('Reputation.asp?CommentFor=<%=Rs("UserName")%>',550,200);">对 <%=Rs("UserName")%> 进行评价</a></td>
							</tr>
							<%
								end if
							end if
							%>
							
							<tr>
								<td><img src=images/email.gif /></td>
								<td><a target="_blank" href="mailto:<%=Rs("UserEmail")%>">给 <%=Rs("UserName")%> 发送邮件</a></td>
							</tr>


							<%if Rs("WebAddress")<>"" then%>
							<tr>
								<td><img src=images/homepage.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebAddress")%>">查看 <%=Rs("UserName")%> 的主页</a></td>
							</tr>
							<%
							end if
							
							if Rs("WebLog")<>"" then%>
							<tr>
								<td><img src=images/weblog.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebLog")%>">查看 <%=Rs("UserName")%> 的博客</a></td>
							</tr>
							<%
							end if
							
							if Rs("WebGallery")<>"" then%>
							<tr>
								<td><img src=images/webgallery.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebGallery")%>">查看 <%=Rs("UserName")%> 的相册</a></td>
							</tr>
							<%end if%>
							<tr>
								<td><img src=images/search.gif /></td>
								<td><a href="ShowBBS.asp?menu=MyTopic&UserName=<%=Rs("UserName")%>">搜索 <%=Rs("UserName")%> 的主题</a></td>
							</tr>
							
							
							<%if Rs("QQ")<>"" then%>
							<tr>
								<td><img src=images/im_qq.gif /></td>
								<td><a target="blank" href="http://wpa.qq.com/msgrd?V=1&Uin=<%=Rs("QQ")%>&menu=yes&Site=<%=Server_Name%>"><%=Rs("QQ")%></a></td>
							</tr>
							<%
							end if
							
							if Rs("MSN")<>"" then%>
							<tr>
								<td><img src=images/im_msn.gif /></td>
								<td><a target="blank" href="msnim:chat?contact=<%=Rs("MSN")%>"><%=Rs("MSN")%></a></td>
							</tr>
							<%
							end if
							
							if Rs("AIM")<>"" then%>
							<tr>
								<td><img src=images/im_aim.gif /></td>
								<td><%=Rs("AIM")%></td>
							</tr>
							<%
							end if
							
							if Rs("Yahoo")<>"" then%>
							<tr>
								<td><img src=images/im_yahoo.gif /></td>
								<td><%=Rs("Yahoo")%></td>
							</tr>
							<%
							end if
							if Rs("ICQ")<>"" then%>
							<tr>
								<td><img src=images/im_icq.gif /></td>
								<td><%=Rs("ICQ")%></td>
							</tr>
							<%
							end if
							if Rs("Skype")<>"" then%>
							<tr>
								<td><img src=images/im_skype.gif /></td>
								<td><%=Rs("Skype")%></td>
							</tr>
							<%end if%>
						</table>
					</td>
				</tr>
			</table>
		</td>
		<td valign=top>

		<table  width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>个人资料</td>
			</tr>
       		<tr class="CommonListCell">
				<td>
				<table border="0" width="100%" cellspacing="0">
	<tr><td>名 字：<%=Rs("RealName")%></td></tr>
	<tr><td>头 衔：<%=Rs("UserTitle")%></td></tr>
	<tr><td>性 别：<%=ShowUserSex(Rs("UserSex"))%></td></tr>
	<tr><td>生 日：<%=Rs("birthday")%></td></tr>
	<tr><td>生 肖：<%=Zodiac(Rs("birthday"))%></td></tr>
	<tr><td>星 座：<%=Horoscope(Rs("birthday"))%></td></tr>
	<tr><td>配 偶：<a href="Profile.asp?UserName=<%=Rs("UserMate")%>"><%=Rs("UserMate")%></a></td></tr>
	<tr><td>职 业：<%=Rs("Occupation")%></td></tr>
	<tr><td>兴 趣：<%=Rs("Interests")%></td></tr>	
	<tr><td>地 址：<%=Rs("Address")%></td></tr>	
				</table>
				</td>
			</tr>
				</td>
			</tr>
		</table>
		



		<%if Rs("UserBio")<>"" then%>
		<table width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>简介</td>
			</tr>
			<tr class="CommonListCell">
				<td><%=Rs("UserBio")%></td>
			</tr>
		</table>
		<br />
		<%
		end if
		
		
		if SiteConfig("EnableSignatures")=1 and Rs("UserSign")<>"" then
		%>
		<table width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>签名</td>
			</tr>
			<tr class="CommonListCell">
				<td><%=BBCode(Rs("UserSign"))%></td>
			</tr>
		</table>
		<%
		end if
		
		if SiteConfig("EnableReputation")=1 then
		%>
			<br />
			<div class="MenuTag" ajaxurl="Loading.asp?menu=Reputation"><li id="CommentArea_btn_0" onclick='AjaxShowPannel(this)' menu="CommentFor=<%=escape(Rs("UserName"))%>" class="NowTag">他人对 <%=Rs("UserName")%> 的评价</li><li id="CommentArea_btn_1" onclick='AjaxShowPannel(this)' menu="CommentBy=<%=escape(Rs("UserName"))%>"><%=Rs("UserName")%> 对他人的评价</li></div>
			<div id=CommentArea>
            <table cellspacing=0 cellpadding=0 width="100%" class="PannelBody"><tr><td><img src="images/loading.gif" border=0 /><script language="javascript" type="text/javascript">Ajax_CallBack(false,'CommentArea','Loading.asp?menu=Reputation&CommentFor=<%=escape(Rs("UserName"))%>');</script></td></tr></table>
			</div>
		<%
		end if
		
		
		if Rs("UserNote")<>"" and BestRole=1 then
		%>
		<table width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>备注（仅超级版主与管理员可见）</td>
			</tr>
			<tr class="CommonListCell">
				<td><%=Rs("UserNote")%></td>
			</tr>
		</table>
		<%end if%>
		</td>
	</tr>
</table>

<center><br /><input onclick="history.back()" type="button" value=" &lt;&lt; 返 回 "> <br /></center>
<%
Set Rs = Nothing
UserID=empty
HtmlBottom
%>