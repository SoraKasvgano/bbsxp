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
	if Rs.eof then Alert("�Բ��𣬲�������Ҫ�鿴���û�����")


HtmlHeadTitle="�鿴"&Rs("UserName")&"������"
HtmlTop
if SiteConfig("RequireAuthenticationForProfileViewing")=1 and CookieUserName=empty then error("������<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">��¼</a>����������������")

%>

<div class="CommonBreadCrumbArea"><%=ClubTree%> �� �鿴�û���<%=Rs("UserName")%>������</div>

<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td valign=top width="200">
			<table height="100%" width="95%" cellspacing=1 cellpadding=5 border=0 class=CommonListArea>
				<tr class=CommonListTitle>
					<td colspan="2"><%=Rs("UserName")%></td>
				</tr>
				<tr class="CommonListCell">
					<td colspan="2"><%if SiteConfig("EnableAvatars")=1 then%><center><img src="<%=Rs("UserFaceUrl")%>"  style="max-width:<%=SiteConfig("AvatarWidth")%>px;max-height:<%=SiteConfig("AvatarHeight")%>px;" /></center><%end if%>��<table border="0" cellspacing="0">

	<tr><td>�û�״̬��<%=ShowUserAccountStatus(RS("UserAccountStatus"))%></td></tr>
	<tr><td>�ǡ���ɫ��<%=ShowRole(RS("UserRoleID"))%></td></tr>
	<tr><td>ע�����ڣ�<%=FormatDateTime(Rs("UserRegisterTime"),2)%></td></tr>
	<tr><td>������<%=FormatDateTime(Rs("UserActivityTime"),2)%></td></tr>
	
	<tr><td>��Ծ������<%=Rs("UserActivityDay")%></td></tr>
	
	<tr><td>����������<%=Rs("Reputation")%></td></tr>
	<tr><td>�� �� ����<%=Rs("TotalPosts")%></td></tr>
	<tr><td>�𡡡�Ǯ��<%=Rs("UserMoney")%></td></tr>
	<tr><td>�� �� ֵ��<%=Rs("experience")%></td></tr>

						</table>
					</td>
				</tr>
				<tr class=CommonListTitle>
					<td colspan="2">ѡ��</td>
				</tr>
				<tr class="CommonListCell">
					<td>
						<table  border="0">




							
							<%if CookieUserName<>"" then
							
							if CookieUserName=Rs("UserMate") or CookieUserName=Rs("UserName") and Rs("UserMate")<>"" then%>
							<tr>
								<td><img src=images/divorce.gif /></td>
								<td><span id="UserMate"><a href="javascript:Ajax_CallBack(false,'UserMate','UserMate.asp?menu=divorce');">�� <%=Rs("UserMate")%> ���</a></span></td>
							</tr>
							<%else%>
							<tr>
								<td><img src=images/Marry.gif /></td>
								<td><span id="UserMate"><a href="javascript:Ajax_CallBack(false,'UserMate','UserMate.asp?menu=MarryRequest&TargetUser=<%=escape(Rs("UserName"))%>');">�� <%=Rs("UserName")%> ���</a></span></td>
							</tr>
							<%
							end if
							end if
							
							
							if CookieUserName=Rs("UserName") then%>
							<tr>
								<td><img src=images/edit.gif /></td>
								<td><a href="EditProfile.asp">�༭�ҵ�����</a></td>
							</tr>
							<%end if%>
							<tr>
								<td><img src=images/favorite.gif /></td>
								<td><a href="javascript:Ajax_CallBack(false,false,'MyFavorites.asp?menu=FavoriteFriend&FriendUserName=<%=Rs("UserName")%>',true);">�� <%=Rs("UserName")%> ��Ϊ����</a></td>
							</tr>
							<%if CookieUserName<>empty then%>
							<tr>
								<td><img src=images/privatemessage.gif /></td>
								<td><a href="javascript:BBSXP_Modal.Open('MyMessage.asp?menu=Post&RecipientUserName=<%=Rs("UserName")%>',600,350);">�� <%=Rs("UserName")%> ����ѶϢ</a></td>
							</tr>
							<%	if SiteConfig("EnableReputation")=1 then%>
							<tr>
								<td><img src=images/reputation.gif /></td>
								<td><a href="javascript:BBSXP_Modal.Open('Reputation.asp?CommentFor=<%=Rs("UserName")%>',550,200);">�� <%=Rs("UserName")%> ��������</a></td>
							</tr>
							<%
								end if
							end if
							%>
							
							<tr>
								<td><img src=images/email.gif /></td>
								<td><a target="_blank" href="mailto:<%=Rs("UserEmail")%>">�� <%=Rs("UserName")%> �����ʼ�</a></td>
							</tr>


							<%if Rs("WebAddress")<>"" then%>
							<tr>
								<td><img src=images/homepage.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebAddress")%>">�鿴 <%=Rs("UserName")%> ����ҳ</a></td>
							</tr>
							<%
							end if
							
							if Rs("WebLog")<>"" then%>
							<tr>
								<td><img src=images/weblog.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebLog")%>">�鿴 <%=Rs("UserName")%> �Ĳ���</a></td>
							</tr>
							<%
							end if
							
							if Rs("WebGallery")<>"" then%>
							<tr>
								<td><img src=images/webgallery.gif /></td>
								<td><a target="_blank" href="<%=Rs("WebGallery")%>">�鿴 <%=Rs("UserName")%> �����</a></td>
							</tr>
							<%end if%>
							<tr>
								<td><img src=images/search.gif /></td>
								<td><a href="ShowBBS.asp?menu=MyTopic&UserName=<%=Rs("UserName")%>">���� <%=Rs("UserName")%> ������</a></td>
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
				<td align=Left>��������</td>
			</tr>
       		<tr class="CommonListCell">
				<td>
				<table border="0" width="100%" cellspacing="0">
	<tr><td>�� �֣�<%=Rs("RealName")%></td></tr>
	<tr><td>ͷ �Σ�<%=Rs("UserTitle")%></td></tr>
	<tr><td>�� ��<%=ShowUserSex(Rs("UserSex"))%></td></tr>
	<tr><td>�� �գ�<%=Rs("birthday")%></td></tr>
	<tr><td>�� Ф��<%=Zodiac(Rs("birthday"))%></td></tr>
	<tr><td>�� ����<%=Horoscope(Rs("birthday"))%></td></tr>
	<tr><td>�� ż��<a href="Profile.asp?UserName=<%=Rs("UserMate")%>"><%=Rs("UserMate")%></a></td></tr>
	<tr><td>ְ ҵ��<%=Rs("Occupation")%></td></tr>
	<tr><td>�� Ȥ��<%=Rs("Interests")%></td></tr>	
	<tr><td>�� ַ��<%=Rs("Address")%></td></tr>	
				</table>
				</td>
			</tr>
				</td>
			</tr>
		</table>
		



		<%if Rs("UserBio")<>"" then%>
		<table width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>���</td>
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
				<td align=Left>ǩ��</td>
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
			<div class="MenuTag" ajaxurl="Loading.asp?menu=Reputation"><li id="CommentArea_btn_0" onclick='AjaxShowPannel(this)' menu="CommentFor=<%=escape(Rs("UserName"))%>" class="NowTag">���˶� <%=Rs("UserName")%> ������</li><li id="CommentArea_btn_1" onclick='AjaxShowPannel(this)' menu="CommentBy=<%=escape(Rs("UserName"))%>"><%=Rs("UserName")%> �����˵�����</li></div>
			<div id=CommentArea>
            <table cellspacing=0 cellpadding=0 width="100%" class="PannelBody"><tr><td><img src="images/loading.gif" border=0 /><script language="javascript" type="text/javascript">Ajax_CallBack(false,'CommentArea','Loading.asp?menu=Reputation&CommentFor=<%=escape(Rs("UserName"))%>');</script></td></tr></table>
			</div>
		<%
		end if
		
		
		if Rs("UserNote")<>"" and BestRole=1 then
		%>
		<table width=100% cellspacing=1 cellpadding=5 align=center border=0 class=CommonListArea>
			<tr class=CommonListTitle>
				<td align=Left>��ע�����������������Ա�ɼ���</td>
			</tr>
			<tr class="CommonListCell">
				<td><%=Rs("UserNote")%></td>
			</tr>
		</table>
		<%end if%>
		</td>
	</tr>
</table>

<center><br /><input onclick="history.back()" type="button" value=" &lt;&lt; �� �� "> <br /></center>
<%
Set Rs = Nothing
UserID=empty
HtmlBottom
%>