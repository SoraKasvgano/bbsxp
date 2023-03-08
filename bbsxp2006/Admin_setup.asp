<!-- #include file="Setup.asp" --><%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))
response.write "<center>"

select case Request("menu")
case "Showbanner"
Showbanner
case "agreement"
agreement
case "variable"
variable
case "SiteSettingsUp"
SiteSettingsUp
case "SiteStat"
SiteStat
case "AutoSiteStat"
TotalUserCount=conn.execute("Select count(id) from [BBSXP_Users]")(0)
TotalThreadCount=conn.execute("Select count(id) from [BBSXP_Threads]")(0)
TotalPostCount=conn.execute("Select sum(replies) from [BBSXP_Threads]")(0)+TotalThreadCount
if TotalThreadCount=0 then TotalPostCount=0
Conn.execute("update [BBSXP_Statistics_Site] set TotalUser="&TotalUserCount&",TotalThread="&TotalThreadCount&",TotalPost="&TotalPostCount&"")
SiteStat
case "SiteStatok"
Rs.Open "[BBSXP_Statistics_Site]",Conn,1,3
for each ho in Request.Form
Rs(ho)=Request(ho)
next
Rs.update
Rs.close
%>更新成功<br><br><a href="javascript:history.back()">返 回</a> <%


end select
sub variable

AutoSiteURL="http://"&Request.ServerVariables("server_name")&Request.ServerVariables("script_name")&""
AutoSiteURL=Left(AutoSiteURL,inStrRev(AutoSiteURL,"/"))

SiteURL=SiteSettings("SiteURL")
if ""&SiteURL&""=empty then SiteURL=AutoSiteURL
%>
<SCRIPT>
function VerifyInput()
{

if (document.form.SiteName.value == "")
{alert("请输入站点名称");
document.form.SiteName.focus();
return false;
}

if (document.form.DefaultSiteStyle.value == "")
{alert("默认风格不能没有设置");
document.form.DefaultSiteStyle.focus();
return false;
}


}
</SCRIPT>
[<b><a href="#普通配置">普通配置</a></b>]
[<b><a href="#首页显示">首页显示</a></b>]
[<b><a href="#投票设置">投票设置</a></b>]
[<b><a href="#上传设置">上传设置</a></b>]
[<b><a href="#水印设置">水印设置</a></b>]
[<b><a href="#SNMP设置">SNMP设置</a></b>]
<br>
[<b><a href="#过滤设置">过滤设置</a></b>]
[<b><a href="#用户选项">用户选项</a></b>]
[<b><a href="#论坛选项">论坛选项</a></b>]
[<b><a href="#帖子设置">帖子设置</a></b>]
[<b><a href="#验证码设置">验证码设置</a></b>]
[<b><a href="#金币和经验设置">金币和经验设置</a></b>]
<form method="POST" action="?menu=SiteSettingsUp" onsubmit="return VerifyInput();" name=form><a name="普通配置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>普通配置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
		<tr>
			<td width="45%" class="a3"><b>站点名称<br>
			</b>论坛的应用程序名称，用在站点显示中</td>
			<td class="a4" width="57%">
			<input size="30" name="SiteName" value="<%=SiteSettings("SiteName")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>站点URL</b>
<script>
if ('<%=AutoSiteURL%>' != '<%=SiteURL%>'){
document.write("（站点URL可能设置错误）<br>系统自动检测到的URL：<font color=FF0000><%=AutoSiteURL%></font>");
}
</script>
			</td>
			<td class="a4" width="57%">
			<input size="30" name="SiteURL" value="<%=SiteURL%>"></td>
		</tr>
		<tr>
			<td class="a3" valign="top" width="45%"><b>META标签-描述信息</b></td>
			<td class="a4" width="57%">
			<input size="40" name="MetaDescription" value="<%=SiteSettings("MetaDescription")%>"></td>
		</tr>
		<tr>
			<td class="a3" valign="top" width="45%"><b>META标签-关键字</b></td>
			<td class="a4" width="57%">
			<input size="40" name="MetaKeywords" value="<%=SiteSettings("MetaKeywords")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>主题/页码</b><br>
			每个网页上显示的主题数</td>
			<td class="a4" width="57%">
			<input value="<%=SiteSettings("ThreadsPerPage")%>" name="ThreadsPerPage" size="20"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>帖子/页码</b><br>
			每个网页上所显示的帖子数</td>
			<td class="a4" width="57%">
			<input value="<%=SiteSettings("PostsPerPage")%>" name="PostsPerPage" size="20"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>统计用户在线时间（分钟）<br>
			</b>默认:20</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("UserOnlineTime")%>" name="UserOnlineTime"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>脚本超时时间（秒）<br>
			</b>默认:60</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("Timeout")%>" name="Timeout"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>默认风格设置</b><br>
			指定<font color="#FF0000">images/skins/</font>目录下的目录名即可 [默认:1]</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("DefaultSiteStyle")%>" name="DefaultSiteStyle"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>论坛使用的缓存名称</b></td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("CacheName")%>" name="CacheName"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>公司/组织名称</b></td>
			<td class="a4" width="57%">
			<input size="30" name="CompanyName" value="<%=SiteSettings("CompanyName")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>公司/组织网址</b></td>
			<td class="a4" width="57%">
			<input size="30" value="<%=SiteSettings("CompanyURL")%>" name="CompanyURL"></td>
		</tr>
</table>
<br>
<a name="首页显示"></a><table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table1">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>首页显示 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>显示在线用户</b><br>
		在首页上显示“在线统计”部分 </td>
		<td class="a4">
		<input type="radio" name="DisplayWhoIsOnline" value="1" <%if sitesettings("displaywhoisonline")=1 then%>checked<%end if%>>是 
		<input type="radio" name="DisplayWhoIsOnline" value="0" <%if sitesettings("displaywhoisonline")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>显示友情链接</b><br>
		在首页上显示“友情链接”部分 </td>
		<td class="a4">
		<input type="radio" name="DisplayLink" value="1" <%if sitesettings("displaylink")=1 then%>checked<%end if%>>是 
		<input type="radio" name="DisplayLink" value="0" <%if sitesettings("displaylink")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>显示论坛样式</b></td>
		<td class="a4">
		<input type="radio" name="DisplayForumFloor" value="1" <%if sitesettings("DisplayForumFloor")=1 then%>checked<%end if%>>简洁 
		<input type="radio" name="DisplayForumFloor" value="0" <%if sitesettings("DisplayForumFloor")=0 then%>checked<%end if%>>详细
	</tr>
</table>
<br>

<a name="投票设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table5">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>投票设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>投票选项最少</b><br>
		投票选项最少不得低于该数值 </td>
		<td class="a4">
		<input size="20" name="MinVoteOptions" value="<%=SiteSettings("MinVoteOptions")%>"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>投票选项最大</b><br>
		投票选项最大不得高于该数值 </td>
		<td class="a4">
		<input size="20" name="MaxVoteOptions" value="<%=SiteSettings("MaxVoteOptions")%>"></td>
	</tr>
</table>


<br>
<a name="上传设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table3">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>上传设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>选择上传组件</b></td>
		<td class="a4"><select name="UpFileOption">
		<option value="">关闭</option>
		<option value="ADODB.Stream" <%if sitesettings("upfileoption")="ADODB.Stream" then%>selected<%end if%>>ADODB.Stream</option>
		<option value="SoftArtisans.FileUp" <%if sitesettings("upfileoption")="SoftArtisans.FileUp" then%>selected<%end if%>>
		SoftArtisans.FileUp</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>附件保存选项</b><br>
				附件是否保存到数据库</td>
		<td class="a4">
		<input type="radio" name="AttachmentsSaveOption" value="1" <%if sitesettings("AttachmentsSaveOption")=1 then%>checked<%end if%>>是 
		<input type="radio" name="AttachmentsSaveOption" value="0" <%if sitesettings("AttachmentsSaveOption")=0 then%>checked<%end if%>>否 </td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>允许上传的附件类型（请用“|”隔开）</b><br>
		例如：gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("UpFileTypes")%>" name="UpFileTypes"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>设置会员附件夹的最大容量 (byte)</b><br>
		默认:10240000</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxPostAttachmentsSize")%>" name="MaxPostAttachmentsSize">（此项只对于保存在数据库中的附件有效）</td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>每次上传附件的大小 (byte)</b><br>
		默认:1024000</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxFileSize")%>" name="MaxFileSize"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>允许照片文件的大小 (byte)</b><br>
		默认:102400</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxPhotoSize")%>" name="MaxPhotoSize"></td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>允许头像文件的大小 (byte)</b><br>
		默认:10240</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxFaceSize")%>" name="MaxFaceSize"></td>
	</tr>
</table>
<br>
<a name="水印设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table6">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>水印设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	
	<tr>
		<td width="45%" class="a3"><b>水印图片组件</b></td>
		<td class="a4"><select name="WatermarkOption" size="1">
		<option value="">关闭</option>
		<option value="Persits.Jpeg" <%if sitesettings("WatermarkOption")="Persits.Jpeg" then%>selected<%end if%>>Persits.Jpeg</option>
		</select>（此项只对于保存在数据库中的附件有效）</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>水印效果</b></td>
		<td class="a4"><select name="WatermarkType" size="1">
		<option value="0" <%if sitesettings("WatermarkType")="0" then%>selected<%end if%>>水印文字</option>
		<option value="1" <%if sitesettings("WatermarkType")="1" then%>selected<%end if%>>水印图片</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>水印文字</b><br>
		该文字将显示于用户上传的图片上(只支持JPEG)</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("WatermarkText")%>" name="WatermarkText"></td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>水印图片的相对路径</b><br>
		该图片将显示于用户上传的图片上(只支持JPEG)</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("WatermarkImage")%>" name="WatermarkImage"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>水印图片显示位置</b></td>
		<td class="a4"><select name="WatermarkPosition" size="1">
		<option value="0" <%if sitesettings("WatermarkPosition")="0" then%>selected<%end if%>>左上</option>
		<option value="1" <%if sitesettings("WatermarkPosition")="1" then%>selected<%end if%>>左下</option>
		<option value="2" <%if sitesettings("WatermarkPosition")="2" then%>selected<%end if%>>居中</option>
		<option value="3" <%if sitesettings("WatermarkPosition")="3" then%>selected<%end if%>>右上</option>
		<option value="4" <%if sitesettings("WatermarkPosition")="4" then%>selected<%end if%>>右下</option>
		
		</select></td>
	</tr></table>
<br>

<a name="SNMP设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table6">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>SNMP设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>发送邮件组件</b></td>
		<td class="a4"><select name="SelectMailMode" size="1">
		<option value="">关闭</option>
		<option value="JMail.Message" <%if sitesettings("selectmailmode")="JMail.Message" then%>selected<%end if%>>
		JMail.Message</option>
		<option value="CDONTS.NewMail" <%if sitesettings("selectmailmode")="CDONTS.NewMail" then%>selected<%end if%>>
		CDONTS.NewMail</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>发件人Email地址</b></td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerMail")%>" name="SmtpServerMail"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP 服务器</b><br>
		SMTP 服务器用来为你的论坛发送邮件</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServer")%>" name="SmtpServer"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP 服务器注册名称</b><br>
		用来登录SMTP服务器的注册名称</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerUserName")%>" name="SmtpServerUserName"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP 服务器密码</b><br>
		用来登录SMTP服务器的密码</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerPassword")%>" name="SmtpServerPassword"></td>
	</tr>
</table>
<br>
<a name="过滤设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table7">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>过滤设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>过滤HTML标签（请用“|”隔开）</b><br>
		例如：iframe|SCRIPT|form|style|div|object|TEXTAREA</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedHtmlLabel")%>" name="BannedHtmlLabel"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>过滤HTML标签内的事件（请用“|”隔开）</b><br>
		例如：javascript:|Document.|onerror|onload|onmouseover</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedHtmlEvent")%>" name="BannedHtmlEvent"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>过滤敏感词（请用“|”隔开）</b><br>
		例如：fuck|shit|你妈</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedText")%>" name="BannedText"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>过滤用户帖子（请用“|”隔开）</b><br>
		例如：yuzi|裕裕</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedUserPost")%>" name="BannedUserPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>禁止注册的用户名（请用“|”隔开）</b><br>
		例如：yuzi|裕裕</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedRegUserName")%>" name="BannedRegUserName"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>禁止IP地址进入论坛（请用“|”隔开）</b><br>
		例如：127.0.0.1|192.168.0.1</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedIP")%>" name="BannedIP"></td>
	</tr>
</table>
<br>
<a name="用户选项"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table8">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>用户选项 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>账号激活方式</b><br>
		控制用户注册后账号激活的方式。 
		&quot;自动&quot;表示允许用户创建自己的账号；&quot;Email发送&quot;表示新注册用户后，会通过邮件方式将密码发送给新用户；&quot;管理员审批&quot;表示账号注册后需要管理员审批。</td>
		<td class="a4">
		<input type="radio" name="EnableUser" value="0" <%if sitesettings("EnableUser")=0 then%>checked<%end if%>>自动 
		<input type="radio" name="EnableUser" value="1" <%if sitesettings("EnableUser")=1 then%>checked<%end if%>>Email发送 
		<input type="radio" name="EnableUser" value="2" <%if sitesettings("EnableUser")=2 then%>checked<%end if%>>管理员审批</td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>关闭用户注册</b></td>
		<td class="a4">
		<input type="radio" name="CloseRegUser" value="1" <%if sitesettings("closereguser")=1 then%>checked<%end if%>>是 
		<input type="radio" name="CloseRegUser" value="0" <%if sitesettings("closereguser")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>一个Email只能注册一个帐号</b></td>
		<td class="a4">
		<input type="radio" name="OnlyMailReg" value="1" <%if sitesettings("onlymailreg")=1 then%>checked<%end if%>>是 
		<input type="radio" name="OnlyMailReg" value="0" <%if sitesettings("onlymailreg")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	</table>
<br>
<a name="论坛选项"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>论坛选项 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>开放申请论坛功能</b></td>
		<td class="a4">
		<input type="radio" name="ForumApply" value="1" <%if sitesettings("forumapply")=1 then%>checked<%end if%>>是 
		<input type="radio" name="ForumApply" value="0" <%if sitesettings("forumapply")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>类别论坛与下属论坛拥有同样功能</b><br>
		如：发帖、浏览</td>
		<td class="a4">
		<input type="radio" name="SortShowForum" value="1" <%if sitesettings("sortshowforum")=1 then%>checked<%end if%>>是 
		<input type="radio" name="SortShowForum" value="0" <%if sitesettings("sortshowforum")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
</table>
<br>
<a name="帖子设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table2">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>帖子设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>发帖间隔（秒）<br>
		</b>同一个用户2次发帖的间隔</td>
		<td class="a4">
		<input size="20" name="DuplicatePostIntervalInMinutes" value="<%=SiteSettings("DuplicatePostIntervalInMinutes")%>"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>注册后多长时间才能发表帖子（秒）</b></td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("RegUserTimePost")%>" name="RegUserTimePost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>热门帖子的回复数</b><br>
		一个主题需要多少篇回帖才能变成热门? </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("PopularPostThresholdPosts")%>" name="PopularPostThresholdPosts"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>热门帖子的浏览数</b><br>
		一个主题需要多少次浏览才能变成热门? </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("PopularPostThresholdViews")%>" name="PopularPostThresholdViews"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>显示编辑记录</b><br>
		启用后，在消息的底部(位于用户签名前)将会显示出帖子的编辑记录 。</td>
		<td class="a4">
		<input type="radio" name="DisplayEditNotes" value="1" <%if sitesettings("DisplayEditNotes")=1 then%>checked<%end if%>>是 
		<input type="radio" name="DisplayEditNotes" value="0" <%if sitesettings("DisplayEditNotes")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>公开显示帖子作者的IP地址<br>
		</b>显示帖子作者的IP地址。</td>
		<td class="a4">
		<input type="radio" name="DisplayPostIP" value="1" <%if sitesettings("DisplayPostIP")=1 then%>checked<%end if%>>是 
		<input type="radio" name="DisplayPostIP" value="0" <%if sitesettings("DisplayPostIP")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
</table>
<br>

<a name="验证码设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>验证码设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>启用注册验证</b><br>
		在用户注册时启用验证码功能。 </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then%>checked<%end if%>>是 
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForRegister")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>启用登录验证</b><br>
		在用户登录时启用验证码功能。 </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForLogin")=1 then%>checked<%end if%>>是 
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForLogin")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>启用发贴验证</b><br>
		在用户发贴时启用验证码功能。 </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>checked<%end if%>>是 
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForPost")=0 then%>checked<%end if%>>否 
		</td>
	</tr>
</table>



<br>
<a name="金币和经验设置"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>金币和经验设置 [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>发主题帖</b><br>
		用户每发一个主题帖所得金币和经验</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddThread")%>" name="IntegralAddThread"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>回帖</b><br>
		用户每回一个帖子所得金币和经验</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddPost")%>" name="IntegralAddPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>加精华</b><br>
		斑竹将某帖子加为精华后，帖子的发帖人所得金币和经验 </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddValuedPost")%>" name="IntegralAddValuedPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>主题帖被删除</b><br>
		斑竹删除某主题帖后，帖子的发帖人所得金币和经验 </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeleteThread")%>" name="IntegralDeleteThread"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>回帖被删除</b><br>
		斑竹删除某回帖后，帖子的发帖人所得金币和经验</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeletePost")%>" name="IntegralDeletePost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>精华被取消</b><br>
		斑竹取消某精华后，帖子的发帖人所得金币和经验</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeleteValuedPost")%>" name="IntegralDeleteValuedPost"></td>
	</tr>
</table>
<table border="0" width="100%" align="center">
	<tr>
		<td align="right"><input type="submit" value=" 保 存 ">
		<input type="reset" value=" 重 置 "></form></td>
	</tr>
</table>
<%
end sub


sub SiteSettingsUp

if Request("EnableUser")="1" and Request("SelectMailMode")="" then error2("您选择了注册用户密码通过Email发送，但是您没有设定发送邮件组件")

filtrate=split(Request("BannedIP"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then error2("请检查您输入的IP地址是否正确")
next

if Request("UpFileOption")<>"" then
if Not IsObjInstalled(Request("UpFileOption")) then error2("本服务器不支持 "&Request("UpFileOption")&" 组件！请关闭此功能！")
end if

if Request("SelectMailMode")<>"" then
if Not IsObjInstalled(Request("SelectMailMode")) then error2("本服务器不支持 "&Request("SelectMailMode")&" 组件！请关闭此功能！")
end if

if Request("WatermarkOption")<>"" then
if Not IsObjInstalled(Request("WatermarkOption")) then error2("本服务器不支持 "&Request("WatermarkOption")&" 组件！请关闭此功能！")
end if

Rs.Open "[BBSXP_SiteSettings]",Conn,1,3
for each ho in Request.Form
Rs(ho)=Request(ho)
next
Rs.update
Rs.close
%>
更新成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub
sub Showbanner
%><form method="POST" action="?menu=SiteSettingsUp">
	<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
		<tr height="25">
			<td class="a1" align="middle" colspan="2"><b>广告代码设置</b></td>
		</tr>
		<tr>
			<td class="a3" align="middle" width="20%">顶部广告代码<br>
			<font color="#FF0000">支持HTML</font> </td>
			<td class="a3">
			<textarea name="TopBanner" rows="6" style="width:100%"><%=SiteSettings("TopBanner")%></textarea></td>
		</tr>
		<tr>
			<td class="a3" align="middle" width="20%">底部版权信息<br>
			<font color="#FF0000">支持HTML</font> </td>
			<td class="a3">
			<textarea name="BottomAD" rows="6" style="width:100%"><%=SiteSettings("BottomAD")%></textarea></td>
		</tr>
	</table>
	<input type="submit" value=" 保 存 ">		<input type="reset" value=" 重 置 ">
</form>
<%
end sub


sub agreement
%><form method="POST" action="?menu=SiteSettingsUp">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle><b>注册用户协议设置</b></td>
  </tr>
    <tr>
    <td class=a3 align=middle>
<textarea name="RegUserAgreement" rows="18" style="width:100%"><%=SiteSettings("RegUserAgreement")%></textarea></td>
  </tr>
        
  </table>
<input type="submit" value=" 保 存 ">		<input type="reset" value=" 重 置 ">
</form>
<%
end sub
sub SiteStat

Set Statistics=Conn.Execute("[BBSXP_Statistics_Site]")

%><form method="POST" action="?menu=SiteStatok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2"><b>社区统计信息</b></td>
  </tr>
	<tr>
		<td width="45%" class="a3"><b>统计注册用户</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalUser")%>" name="TotalUser"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>统计主题数</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalThread")%>" name="TotalThread"></td>
	</tr>
        
	<tr>
		<td width="45%" class="a3"><b>统计帖子数</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalPost")%>" name="TotalPost"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>今日总发帖数</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TodayPost")%>" name="TodayPost"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>最高在线人数</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("BestOnline")%>" name="BestOnline"></td>
	</tr>
        
	<tr>
		<td width="45%" class="a3"><b>最高在线人数发生时间</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("BestOnlineTime")%>" name="BestOnlineTime"></td>
	</tr>
	
	<tr>
		<td width="45%" class="a3"><b>最新注册的用户</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("NewUser")%>" name="NewUser"></td>
	</tr>
  </table>
<input type="submit" value=" 保 存 ">		<input type="reset" value=" 重 置 ">		
<input type="button" value=" 更新统计 " onclick="document.location.href='?menu=AutoSiteStat'">
</form>
<%
Set Statistics=Nothing

end sub

htmlend

%>