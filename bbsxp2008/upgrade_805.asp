<!-- #include file="Conn.asp" -->
<%

On Error Resume Next

Server.ScriptTimeout=999999


if BBSxpDataBaseVersion="" then BBSxpDataBaseVersion=Conn.Execute("Select BBSxpVersion From ["&TablePrefix&"SiteSettings]")(0)
%>
<b><font size="7">BBSXP 8.0.5数据库升级程序</font></b><p>您当前的BBSXP数据库版本是 <font color="#FF0000"><b><%=BBSxpDataBaseVersion%></b></font><br /><br />

<%if BBSxpDataBaseVersion="7.0.0" then %><a href=?Menu=Up700To710>7.0.0 数据库升级到 7.1.0 版本</a><br /><br /></p><%end if%>
<%if BBSxpDataBaseVersion="7.1.0" then %><a href=?Menu=Up710To711>7.1.0 数据库升级到 7.1.1 版本</a><br /><br /></p><%end if%>
<%if BBSxpDataBaseVersion="7.1.1" then %><a href=?Menu=Up711To720>7.1.1 数据库升级到 7.2.0 版本</a><br /><br /></p><%end if%>
<%if BBSxpDataBaseVersion="7.2.0" then %><a href=?Menu=Up720To730>7.2.0 数据库升级到 7.3.0 版本</a><br /><br /></p><%end if%>
<%if BBSxpDataBaseVersion="7.3.0" then %><a href=?Menu=Up730To731>7.3.0 数据库升级到 7.3.1 版本</a><br /><br /></p><%end if%>
<%
if BBSxpDataBaseVersion="7.3.1" then
	if not Conn.Execute("select PostsTableName from ["&TablePrefix&"Threads] where PostsTableName>0").eof then
		Response.Write("您的数据库库含有多个帖子表，请先<a href='?menu=MovePost'>合并帖子表</a>再进行数据库的升级")
	Else
%>
		<a href=?Menu=Up731To800>7.3.1 数据库升级到 8.0.0 版本</a><br /><br /></p>
<%
	end if
end if

if BBSxpDataBaseVersion="8.0.0" then %><a href=?Menu=Up800To801>8.0.0 数据库升级到 8.0.1 版本</a><br /><br /></p><%end if

if BBSxpDataBaseVersion="8.0.1" then %><a href=?Menu=Up801To802>8.0.1 数据库升级到 8.0.2 版本</a><br /><br /></p><%end if
if BBSxpDataBaseVersion="8.0.2" then %><a href=?Menu=Up802To803>8.0.2 数据库升级到 8.0.3 版本</a><br /><br /></p><%end if
if BBSxpDataBaseVersion="8.0.3" then %><a href=?Menu=Up803To804>8.0.3 数据库升级到 8.0.4 版本</a><br /><br /></p><%end if
if BBSxpDataBaseVersion="8.0.4" then %><a href=?Menu=Up804To805>8.0.4 数据库升级到 8.0.5 版本</a><br /><br /></p><%end if

if Request("Menu")="Up700To710" then
	if BBSxpDataBaseVersion<>"7.0.0" then response.write("您当前的数据库版本不是 7.0.0！"):response.end
	AddSiteConfigXMLField "IsShowSonForum","1"
	AddSiteConfigXMLField "ViewMode","1"
	AddSiteConfigXMLField "MaxPrivateMessageSize","100"

	AddSiteConfigXMLField "WatermarkFontFamily","宋体"
	AddSiteConfigXMLField "WatermarkFontSize","25"
	AddSiteConfigXMLField "WatermarkFontColor","#000000"
	AddSiteConfigXMLField "WatermarkFontIsBold","1"

	AddSiteConfigXMLField "WatermarkWidthPosition","left"
	AddSiteConfigXMLField "WatermarkHeightPosition","bottom"

	Sql="CREATE TABLE ["&TablePrefix&"Subscriptions] ("&_
	"SubscriptionID int IDENTITY (1, 1) NOT NULL ,"&_
	"UserName nvarchar(50) NOT NULL ,"&_
	"ThreadID int Default 0  NOT NULL ,"&_
	"Email nvarchar(255) NOT NULL ,"&_
	"SubscriptionDate datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	Conn.Execute(sql)

	'建立关联
	Sql="ALTER TABLE ["&TablePrefix&"Subscriptions] ADD CONSTRAINT [FK_"&TablePrefix&"Subscriptions_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE  ON UPDATE CASCADE"
	Conn.Execute(sql)

	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='7.1.0'")
		response.redirect "?"
	end if
end if


if Request("Menu")="Up710To711" then
	if BBSxpDataBaseVersion<>"7.1.0" then response.write("您当前的数据库版本不是 7.1.0！"):response.end

	Conn.Execute("alter table ["&TablePrefix&"Users] add UserMate nvarchar(50) NULL ")
	Conn.Execute("alter table ["&TablePrefix&"Users] add BankMoney int Default 0  NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add BankDate datetime Default "&SqlNowString&" NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add UserNote ntext NULL ")
	Conn.Execute("alter table ["&TablePrefix&"Threads] add ThreadStyle nvarchar(255) NULL ")
	
	Conn.Execute("Update ["&TablePrefix&"Users] set BankMoney=0,BankDate="&SqlNowString&"")


	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='7.1.1'")
		response.redirect "?"
	end if
end if



if Request("Menu")="Up711To720" then
	if BBSxpDataBaseVersion<>"7.1.1" then response.write("您当前的数据库版本不是 7.1.1！"):response.end
	
	AddSiteConfigXMLField "SiteDisabledReason","论坛维护中，暂时无法访问！"
	AddSiteConfigXMLField "CacheName","BBSXP"
	AddSiteConfigXMLField "CacheUpDateInterval","5"
	AddSiteConfigXMLField "DefaultPasswordFormat","SHA1"

	Sql="CREATE TABLE ["&TablePrefix&"Advertisements] ("&_
	"AdvertisementID int IDENTITY (1, 1) NOT NULL ,"&_
	"Body ntext NULL ,"&_
	"ExpireDate datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	Conn.Execute(sql)

	Conn.Execute("alter table ["&TablePrefix&"Links] add SortOrder int Default 0 NOT NULL")
	Conn.Execute("Update ["&TablePrefix&"Links] set SortOrder=1")


	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='7.2.0'")
		response.redirect "?"
	end if
end if


if Request("Menu")="Up720To730" then
	if BBSxpDataBaseVersion<>"7.2.0" then response.write("您当前的数据库版本不是 7.2.0！"):response.end

	AddSiteConfigXMLField "ReputationDefault","0"
	AddSiteConfigXMLField "ShowUserRates","5"
	AddSiteConfigXMLField "MinReputationPost","50"
	AddSiteConfigXMLField "MinReputationCount","0"
	AddSiteConfigXMLField "MaxReputationPerDay","10"
	AddSiteConfigXMLField "ReputationRepeat","20"
	AddSiteConfigXMLField "AdminReputationPower","10"
	AddSiteConfigXMLField "InPrisonReputation","-10"
	
	AddSiteConfigXMLField "CustomUserTitle","0"
	AddSiteConfigXMLField "UserTitleMaxChars","25"
	AddSiteConfigXMLField "UserTitleCensorWords","admin|forum|moderator"


	Sql="CREATE TABLE ["&TablePrefix&"Reputation] ("&_
	"ReputationID int IDENTITY (1, 1) NOT NULL ,"&_
	"Reputation int Default 0  NOT NULL ,"&_
	"Comment ntext NULL ,"&_
	"CommentFor nvarchar(50) NOT NULL ,"&_
	"CommentBy nvarchar(50) NOT NULL ,"&_
	"IPAddress nvarchar(50) NOT NULL ,"&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	Conn.Execute(sql)

	'建立关联
	Sql="ALTER TABLE ["&TablePrefix&"Reputation] ADD CONSTRAINT [FK_"&TablePrefix&"Reputation_"&TablePrefix&"Users] FOREIGN KEY ([CommentFor]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE  ON UPDATE CASCADE"
	Conn.Execute(sql)

	Conn.Execute("alter table ["&TablePrefix&"Users] add UserTitle nvarchar(255) NULL ")
	Conn.Execute("alter table ["&TablePrefix&"Users] add Reputation int Default 0 NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"SiteSettings] add AdminNotes ntext NULL")

	Conn.Execute("Update ["&TablePrefix&"Users] set Reputation=0")


	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='7.3.0'")
		response.redirect "?"
	end if
end if



if Request("Menu")="Up730To731" then
	if BBSxpDataBaseVersion<>"7.3.0" then response.write("您当前的数据库版本不是 7.3.0！"):response.end

	AddSiteConfigXMLField "CookiePath","/"
	AddSiteConfigXMLField "SiteUrl",""
	AddSiteConfigXMLField "NoCacheHeaders","0"


	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='7.3.1'")
		response.redirect "?"
	end if
end if



if Request("Menu")="Up731To800" then
	if BBSxpDataBaseVersion<>"7.3.1" then response.write("您当前的数据库版本不是 7.3.1！"):response.end

	Conn.Execute("alter table ["&TablePrefix&"Users] add UserRank nvarchar(255) NULL ")
	Conn.Execute("alter table ["&TablePrefix&"Users] add UserActivityDay int Default 0 NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add PasswordQuestion nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add PasswordAnswer nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add UserPassword nvarchar(50) NULL")
	
	Conn.Execute("alter table ["&TablePrefix&"Threads] add Description nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Threads] add StickyDate datetime Default "&SqlNowString&" NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Threads] add HiddenCount int Default 0 NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Threads] add DeletedCount int Default 0 NOT NULL")
	Conn.Execute("alter table ["&TablePrefix&"Threads] add Visible int Default 1 NOT NULL")
	
	Conn.Execute("alter table ["&TablePrefix&"Ranks] add RoleID int Default 0 NOT NULL")
	Conn.Execute("Update ["&TablePrefix&"Users] set UserActivityDay=0")
	Conn.Execute("Update ["&TablePrefix&"Users] set UserPassword=UserPass")
	Conn.Execute("alter table ["&TablePrefix&"Users] drop column UserPass")
	Conn.Execute("alter table ["&TablePrefix&"PostAttachments] add PostID int Default 0 NOT NULL")
	
	Conn.Execute("Update ["&TablePrefix&"Ranks] set RoleID=0")
	Conn.Execute("Update ["&TablePrefix&"PostAttachments] set PostID=0")
	
	Conn.Execute("alter table ["&TablePrefix&"Posts] add Visible int Default 1 NOT NULL")
	Conn.Execute("Update ["&TablePrefix&"Posts] set Visible=1")
	Conn.Execute("Update ["&TablePrefix&"Threads] set Visible=1")
	Conn.Execute("Update ["&TablePrefix&"Threads] set DeletedCount=0")
	Conn.Execute("Update ["&TablePrefix&"Threads] set HiddenCount=0")
	Conn.Execute("Update ["&TablePrefix&"Threads] set StickyDate="&SqlNowString&"")
	Conn.Execute("Update ["&TablePrefix&"Threads] set StickyDate=DateAdd("&SqlChar&"yyyy"&SqlChar&", 1, "&SqlNowString&") where ThreadTop=1")
	Conn.Execute("Update ["&TablePrefix&"Threads] set StickyDate=DateAdd("&SqlChar&"yyyy"&SqlChar&", 2, "&SqlNowString&") where ThreadTop=2")
	
	AddSiteConfigXMLField "EnableReputation","1"
	AddSiteConfigXMLField "EnablePostPreviewPopup","0"
	AddSiteConfigXMLField "RequireEditNotes","0"

	'Conn.Execute("alter table ["&TablePrefix&"Threads] drop column IsDel")	'SQL数据库 定义约束（初始值）的字段没办法删除
	'Conn.Execute("alter table ["&TablePrefix&"Threads] drop column IsApproved")
	

	Sql="CREATE TABLE ["&TablePrefix&"PostTags] ("&_
	"TagID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"TagName nvarchar(255) NOT NULL ,"&_
	"IsEnabled int Default 1 NOT NULL ,"&_
	"TotalPosts int Default 0 NOT NULL ,"&_
	"MostRecentPostDate datetime Default "&SqlNowString&" NOT NULL , "&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	Conn.Execute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"PostInTags] ("&_
	"TagID int Default 0 NOT NULL ,"&_
	"PostID int Default 0 NOT NULL"&_
	")"
	Conn.Execute(sql)

	'建立关联（关系的主键必须得是表中的主键）
	Sql="ALTER TABLE ["&TablePrefix&"PostInTags] ADD CONSTRAINT [FK_"&TablePrefix&"PostInTags_"&TablePrefix&"PostTags] FOREIGN KEY ([TagID]) REFERENCES ["&TablePrefix&"PostTags] ([TagID]) ON DELETE CASCADE ON UPDATE CASCADE"
	Conn.Execute(sql)

	
	'创建帖子编辑记录表
	Sql="CREATE TABLE ["&TablePrefix&"PostEditNotes] ("&_
	"EditNoteID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"PostID int Default 0 NOT NULL ,"&_
	"EditNotes nvarchar(255) NULL"&_
	")"
	Conn.Execute(sql)
	




	'将表"&TablePrefix&"PostRating的名称改为"&TablePrefix&"ThreadRating	Start
	Conn.Execute("select * into ["&TablePrefix&"ThreadRating] from ["&TablePrefix&"PostRating]")	'实现数据的拷贝
	Conn.Execute("ALTER TABLE ["&TablePrefix&"ThreadRating] ADD CONSTRAINT [FK_"&TablePrefix&"ThreadRating_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE ON UPDATE CASCADE")	'建立关联
	If IsSqlDataBase=0 Then		'设置表的默认值
		conn.execute("alter table ["&TablePrefix&"ThreadRating] alter Rating int default 0")
		conn.execute("alter table ["&TablePrefix&"ThreadRating] alter DateCreated datetime default "&SqlNowString&"")
	else
		Conn.Execute("ALTER TABLE ["&TablePrefix&"ThreadRating] ADD CONSTRAINT [DF_"&TablePrefix&"ThreadRating_Rating] DEFAULT 0 FOR [Rating]")
		Conn.Execute("ALTER TABLE ["&TablePrefix&"ThreadRating] ADD CONSTRAINT [DF_"&TablePrefix&"ThreadRating_DateCreated] DEFAULT "&SqlNowString&" FOR [DateCreated]")
	End If
	Conn.Execute("drop table ["&TablePrefix&"PostRating]")	'删除表"&TablePrefix&"PostRating
	'将表"&TablePrefix&"PostRating的名称改为"&TablePrefix&"ThreadRating	End



	
	Sql="CREATE TABLE ["&TablePrefix&"FavoritePosts] ("&_
	"FavoriteID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"PostID int Default 0 NOT NULL ,"&_
	"OwnerUserName nvarchar(50) NOT NULL"&_
	")"
	Conn.Execute(sql)
	
	If Err Then
		Response.Write Err.Description
	else
		If IsSqlDataBase=0 then Conn.Execute("alter table ["&TablePrefix&"Posts] add constraint pk_add primary key (PostID)")	'设置PostID为主键
		'创建关联
		Sql="ALTER TABLE ["&TablePrefix&"PostEditNotes] ADD CONSTRAINT [FK_"&TablePrefix&"Posts_"&TablePrefix&"PostEditNotes] FOREIGN KEY ([PostID]) REFERENCES ["&TablePrefix&"Posts] ([PostID]) ON DELETE CASCADE ON UPDATE CASCADE"
		Conn.Execute(sql)
		
		Sql="ALTER TABLE ["&TablePrefix&"FavoritePosts] ADD CONSTRAINT [FK_"&TablePrefix&"Users_"&TablePrefix&"FavoritePosts] FOREIGN KEY ([OwnerUserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
		Conn.Execute(sql)
		'Conn.Execute("drop table ["&TablePrefix&"FavoriteThreads]")	'删除表"&TablePrefix&"FvaoriteThreads]
		
		If Err Then Err.clear

		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.0'")
		response.redirect "?"
	end if
end if

if Request("Menu")="Up800To801" then
	if BBSxpDataBaseVersion<>"8.0.0" then response.write("您当前的数据库版本不是 8.0.0！"):response.end
	Conn.Execute("alter table ["&TablePrefix&"UserOnline] add ThreadID int Default 0 NULL")
	Conn.Execute("alter table ["&TablePrefix&"UserOnline] add Topic nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"UserOnline] drop column Act")
	Conn.Execute("alter table ["&TablePrefix&"UserOnline] drop column ActUrl")
	
	Conn.Execute("alter table ["&TablePrefix&"PostAttachments] add FileData Image NULL")

	Conn.Execute("alter table ["&TablePrefix&"PostAttachments] add TempField nvarchar(255) NULL")
	Conn.Execute("Update ["&TablePrefix&"PostAttachments] set TempField=FilePath")
	Conn.Execute("alter table ["&TablePrefix&"PostAttachments] drop column FilePath")
	Conn.Execute("alter table ["&TablePrefix&"PostAttachments] add FilePath nvarchar(255) NULL")
	Conn.Execute("Update ["&TablePrefix&"PostAttachments] set FilePath=TempField")
	
	
	AddSiteConfigXMLField "AttachmentsSaveOption","1"
	AddSiteConfigXMLField "DisplayForumUsers","1"
	AddSiteConfigXMLField "DisplayThreadUsers","0"
	
	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.1'")
		response.redirect "?"
	end if
end if

if Request("Menu")="Up801To802" then
	if BBSxpDataBaseVersion<>"8.0.1" then response.write("您当前的数据库版本不是 8.0.1！"):response.end
	AddSiteConfigXMLField "APIEnable","0"
	AddSiteConfigXMLField "APISafeKey",""
	AddSiteConfigXMLField "APIUrlList",""
	AddSiteConfigXMLField "WebMasterEmail",""
	AddSiteConfigXMLField "CookieDomain",""
	
	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.2'")
		response.redirect "?"
	end if
end if

if Request("Menu")="Up802To803" then
	if BBSxpDataBaseVersion<>"8.0.2" then response.write("您当前的数据库版本不是 8.0.2！"):response.end
	
	Conn.Execute("alter table ["&TablePrefix&"Users] add QQ nvarchar(255)")
	Conn.Execute("alter table ["&TablePrefix&"Users] add ICQ nvarchar(255)")
	Conn.Execute("alter table ["&TablePrefix&"Users] add AIM nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add MSN nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add Skype nvarchar(255) NULL")
	Conn.Execute("alter table ["&TablePrefix&"Users] add Yahoo nvarchar(255) NULL")
	
	Set XMLDOM=Server.CreateObject("Microsoft.XMLDOM")
	Rs.open "["&TablePrefix&"Users]",Conn,1,3
	do while not Rs.eof
		XMLDOM.loadxml("<bbsxp>"&Rs("UserInfo")&"</bbsxp>")
		QQ=XMLDOM.documentElement.SelectSingleNode("QQ").text
		ICQ=XMLDOM.documentElement.SelectSingleNode("ICQ").text
		AIM=XMLDOM.documentElement.SelectSingleNode("AIM").text
		MSN=XMLDOM.documentElement.SelectSingleNode("MSN").text
		Yahoo=XMLDOM.documentElement.SelectSingleNode("Yahoo").text
		Skype=XMLDOM.documentElement.SelectSingleNode("Skype").text
		Rs("QQ")=QQ
		Rs("ICQ")=ICQ
		Rs("AIM")=AIM
		Rs("MSN")=MSN
		Rs("Yahoo")=Yahoo
		Rs("Skype")=Skype
		Rs.update
		Rs.movenext
	loop
	Rs.close
	Set XMLDOM=nothing
	
	Conn.Execute("ALTER table ["&TablePrefix&"Groups] add ForumColumns int Default 0")
	Conn.Execute("ALTER table ["&TablePrefix&"Groups] add Moderated nvarchar(255) Null")
	Conn.Execute("Update ["&TablePrefix&"Groups] set ForumColumns=0")
	Conn.Execute("ALTER table ["&TablePrefix&"Forums] add ModerateNewThread int Default 0")
	Conn.Execute("ALTER table ["&TablePrefix&"Forums] add ModerateNewPost int Default 0")
	Conn.Execute("Update ["&TablePrefix&"Forums] set ModerateNewThread=IsModerated")
	Conn.Execute("Update ["&TablePrefix&"Forums] set ModerateNewPost=0")
	
	'If Err Then
	'	Response.Write Err.Description
	'else
		Conn.Execute("alter table ["&TablePrefix&"Users] drop column UserInfo")
		Conn.Execute("alter table ["&TablePrefix&"Forums] drop column IsModerated")
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.3'")
		response.redirect "?"
	'end if
end if


if Request("Menu")="Up803To804" then
	if BBSxpDataBaseVersion<>"8.0.3" then response.write("您当前的数据库版本不是 8.0.3！"):response.end

	AddSiteConfigXMLField "DisplayTags","1"
	
	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.4'")
		response.redirect "?"
	end if
end if

if Request("Menu")="Up804To805" then
	if BBSxpDataBaseVersion<>"8.0.4" then response.write("您当前的数据库版本不是 8.0.4！"):response.end
	Conn.Execute("ALTER table ["&TablePrefix&"Roles] add RoleMaxFileSize int Default 0")
	Conn.Execute("ALTER table ["&TablePrefix&"Roles] add RoleMaxPostAttachmentsSize int Default 0")
	Conn.Execute("Update ["&TablePrefix&"Roles] set RoleMaxFileSize=0,RoleMaxPostAttachmentsSize=0")
	
	If Err Then
		Response.Write Err.Description
	else
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set BBSxpVersion='8.0.5'")
		response.redirect "?"
	end if
end if



if Request("menu")="MovePost" then
	Set Rs=Conn.openSchema(20)
		Do While not Rs.EOF
		if instr(Rs("TABLE_NAME"),""&TablePrefix&"Posts")>0 and Rs("TABLE_TYPE")="TABLE" and Rs("TABLE_NAME")<>""&TablePrefix&"Posts" then
			Conn.Execute("insert into ["&TablePrefix&"Posts] (ThreadID,ParentID,PostAuthor,Subject,Body,IPAddress,PostDate) select ThreadID,ParentID,PostAuthor,Subject,Body,IPAddress,PostDate from ["&Rs("TABLE_NAME")&"] order by PostID")
			PostsTableName=Replace(""&Rs("TABLE_NAME")&"",""&TablePrefix&"Posts","")
			Conn.Execute("Update ["&TablePrefix&"Threads] set PostsTableName=0 where PostsTableName="&PostsTableName&"")
		end if
		Rs.movenext
	Loop
	Rs.close
	response.redirect "?"
end if



sub AddSiteConfigXMLField (NodeName,NodeValue)
	Set XMLRoot=SiteConfigXMLDOM.documentElement

	FoundNodeIndex=-1
	for NodeIndex=0 to XMLRoot.childNodes.length-1
		set childSearch=XMLRoot.childNodes(NodeIndex)
		if childSearch.nodeName=NodeName then
			FoundNodeIndex=NodeIndex
			exit for
		end if
	next

	if FoundNodeIndex=-1 then
		Set NewNode=SiteConfigXMLDOM.CreateElement(NodeName)
		NewNode.Text=NodeValue
		XMLRoot.appendChild NewNode
		XmlStr=replace(replace(SiteConfigXMLDOM.xml,"<bbsxp>",""),"</bbsxp>","")
		Conn.Execute("Update ["&TablePrefix&"SiteSettings] set SiteSettingsXML='"&XmlStr&"'")
	end if
end sub
%>