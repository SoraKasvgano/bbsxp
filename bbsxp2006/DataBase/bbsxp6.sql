if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Threads_BBSXP_Forums]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Threads] DROP CONSTRAINT FK_BBSXP_Threads_BBSXP_Forums
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Posts_BBSXP_Threads]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Posts] DROP CONSTRAINT FK_BBSXP_Posts_BBSXP_Threads
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Vote_BBSXP_Threads]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Vote] DROP CONSTRAINT FK_BBSXP_Vote_BBSXP_Threads
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Calendar_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Calendar] DROP CONSTRAINT FK_BBSXP_Calendar_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Consort_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Consort] DROP CONSTRAINT FK_BBSXP_Consort_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Consortia_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Consortia] DROP CONSTRAINT FK_BBSXP_Consortia_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Favorites_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Favorites] DROP CONSTRAINT FK_BBSXP_Favorites_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Messages_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Messages] DROP CONSTRAINT FK_BBSXP_Messages_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_PostAttachments_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_PostAttachments] DROP CONSTRAINT FK_BBSXP_PostAttachments_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Prison_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Prison] DROP CONSTRAINT FK_BBSXP_Prison_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BBSXP_Threads_BBSXP_Users]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BBSXP_Threads] DROP CONSTRAINT FK_BBSXP_Threads_BBSXP_Users
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Affiche]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Affiche]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Calendar]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Calendar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Consort]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Consort]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Consortia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Consortia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Favorites]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Favorites]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Forums]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Forums]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Link]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Link]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Menu]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Menu]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Messages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Messages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_PostAttachments]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_PostAttachments]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Posts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Posts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Prison]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Prison]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Ranks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Ranks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_SiteSettings]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_SiteSettings]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Statistics_Site]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Statistics_Site]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Threads]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Threads]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Users]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Users]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_UsersOnline]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_UsersOnline]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BBSXP_Vote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BBSXP_Vote]
GO

CREATE TABLE [dbo].[BBSXP_Affiche] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[PostTime] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Calendar] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Title] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[AddDate] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Hide] [int] NOT NULL ,
	[DateCreated] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Consort] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Aim] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Unburden] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[DateCreated] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Consortia] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ConsortiaName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[FullName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Tenet] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[DateCreated] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Favorites] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Name] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[URL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[DateCreated] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Forums] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[FollowID] [int] NULL ,
	[SortNum] [int] NULL ,
	[ForumName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Moderated] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[TolSpecialTopic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumIntro] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumToday] [int] NULL ,
	[ForumThreads] [int] NULL ,
	[ForumPosts] [int] NULL ,
	[ForumPass] [int] NULL ,
	[ForumLogo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumIcon] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumHide] [int] NULL ,
	[LastTopic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastTime] [datetime] NULL ,
	[ForumPassword] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumUserList] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumRules] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Link] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[URL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Logo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Intro] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Log] (
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[IPAddress] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserAgent] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[DateCreated] [datetime] NULL ,
	[HttpVerb] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PathAndQuery] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Menu] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[URL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[FollowID] [int] NOT NULL ,
	[SortNum] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Messages] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Incept] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[DateCreated] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_PostAttachments] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ThreadID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[FileName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ContentType] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ContentSize] [int] NOT NULL ,
	[TotalDownloads] [int] NOT NULL ,
	[Description] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Content] [image] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Posts] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ThreadID] [int] NULL ,
	[IsTopic] [int] NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Subject] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[PostIP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostTime] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Prison] (
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Causation] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Constable] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ComeTime] [datetime] NULL ,
	[PrisonDay] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Ranks] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[RankName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[PostingCountMin] [int] NOT NULL ,
	[RankIconUrl] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_SiteSettings] (
	[AdminPassword] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Nowdate] [datetime] NULL ,
	[DefaultPostsName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[SiteName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SiteURL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[metaDescription] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[metaKeywords] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ThreadsPerPage] [int] NULL ,
	[PostsPerPage] [int] NULL ,
	[Timeout] [int] NULL ,
	[UserOnlineTime] [int] NULL ,
	[DefaultSiteStyle] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[CompanyName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[CompanyURL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[DisplayForumFloor] [int] NULL ,
	[DisplayWhoIsOnline] [int] NULL ,
	[DisplayLink] [int] NULL ,
	[DuplicatePostIntervalInMinutes] [int] NULL ,
	[RegUserTimePost] [int] NULL ,
	[PopularPostThresholdPosts] [int] NULL ,
	[PopularPostThresholdViews] [int] NULL ,
	[UpFileOption] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UpFileTypes] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[MaxFileSize] [int] NULL ,
	[MaxFaceSize] [int] NULL ,
	[MaxPhotoSize] [int] NULL ,
	[MinVoteOptions] [int] NULL ,
	[MaxVoteOptions] [int] NULL ,
	[SelectMailMode] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SmtpServer] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SmtpServerUserName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SmtpServerPassword] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SmtpServerMail] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedHtmlLabel] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedHtmlEvent] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedText] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedUserPost] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedRegUserName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BannedIP] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumApply] [int] NULL ,
	[SortShowForum] [int] NULL ,
	[OnlyMailReg] [int] NULL ,
	[EnableUser] [int] NULL ,
	[CloseRegUser] [int] NULL ,
	[IntegralAddThread] [int] NULL ,
	[IntegralAddPost] [int] NULL ,
	[IntegralAddValuedPost] [int] NULL ,
	[IntegralDeleteThread] [int] NULL ,
	[IntegralDeletePost] [int] NULL ,
	[IntegralDeleteValuedPost] [int] NULL ,
	[TopBanner] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[BottomAD] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[RegUserAgreement] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[CacheName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[DisplayEditNotes] [int] NULL ,
	[EnableAntiSpamTextGenerateForRegister] [int] NULL ,
	[EnableAntiSpamTextGenerateForLogin] [int] NULL ,
	[EnableAntiSpamTextGenerateForPost] [int] NULL ,
	[DisplayPostIP] [int] NULL ,
	[MaxPostAttachmentsSize] [int] NULL ,
	[WatermarkOption] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[WatermarkType] [int] NULL ,
	[WatermarkImage] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[WatermarkText] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[WatermarkPosition] [int] NULL ,
	[AttachmentsSaveOption] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Statistics_Site] (
	[NewUser] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[TotalUser] [int] NOT NULL ,
	[TotalThread] [int] NOT NULL ,
	[TotalPost] [int] NOT NULL ,
	[TodayPost] [int] NOT NULL ,
	[BestOnline] [int] NOT NULL ,
	[BestOnlineTime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Threads] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ForumID] [int] NULL ,
	[Icon] [int] NULL ,
	[Topic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[SpecialTopic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastTime] [datetime] NULL ,
	[IsGood] [int] NULL ,
	[IsTop] [int] NULL ,
	[IsLocked] [int] NULL ,
	[IsDel] [int] NULL ,
	[IsVote] [int] NULL ,
	[Views] [int] NULL ,
	[Replies] [int] NULL ,
	[PostsTableName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Users] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[UserPass] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Membercode] [int] NULL ,
	[UserMail] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserHome] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[PasswordQuestion] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[PasswordAnswer] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserHonor] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Birthday] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserSex] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Consortia] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Consort] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[NewMessage] [int] NULL ,
	[PostTopic] [int] NULL ,
	[PostRevert] [int] NULL ,
	[DelTopic] [int] NULL ,
	[GoodTopic] [int] NULL ,
	[UserMoney] [int] NULL ,
	[SaveMoney] [int] NULL ,
	[Experience] [int] NULL ,
	[UserDegree] [int] NULL ,
	[UserFace] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserPhoto] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserMobile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserLastIP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserRegTime] [datetime] NOT NULL ,
	[UserLandTime] [datetime] NOT NULL ,
	[SaveMoneyTime] [datetime] NULL ,
	[UserSign] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[UserFriend] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[UserInfo] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[UserIM] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_UsersOnline] (
	[ForumID] [int] NULL ,
	[SessionID] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[IP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Eremite] [int] NULL ,
	[UserFace] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ForumName] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Act] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ActURL] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ComeTime] [datetime] NULL ,
	[LastTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BBSXP_Vote] (
	[ThreadID] [int] NULL ,
	[Type] [int] NULL ,
	[Expiry] [datetime] NULL ,
	[Items] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[Result] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[BallotUserList] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[BBSXP_Affiche] ADD 
	CONSTRAINT [DF_BBSXP_Affiche_PostTime] DEFAULT (getdate()) FOR [PostTime]
GO

ALTER TABLE [dbo].[BBSXP_Calendar] ADD 
	CONSTRAINT [DF_BBSXP_Calendar_Hide] DEFAULT (0) FOR [Hide],
	CONSTRAINT [DF_BBSXP_Calendar_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_Consort] ADD 
	CONSTRAINT [DF_BBSXP_Consort_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_Consortia] ADD 
	CONSTRAINT [DF_BBSXP_Consortia_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_Favorites] ADD 
	CONSTRAINT [DF_BBSXP_Favorites_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_Forums] ADD 
	CONSTRAINT [DF_BBSXP_Forums_FollowID] DEFAULT (0) FOR [FollowID],
	CONSTRAINT [DF_BBSXP_Forums_SortNum] DEFAULT (0) FOR [SortNum],
	CONSTRAINT [DF_BBSXP_Forums_ForumToday] DEFAULT (0) FOR [ForumToday],
	CONSTRAINT [DF_BBSXP_Forums_ForumThreads] DEFAULT (0) FOR [ForumThreads],
	CONSTRAINT [DF_BBSXP_Forums_ForumPosts] DEFAULT (0) FOR [ForumPosts],
	CONSTRAINT [DF_BBSXP_Forums_ForumPass] DEFAULT (1) FOR [ForumPass],
	CONSTRAINT [DF_BBSXP_Forums_ForumHide] DEFAULT (0) FOR [ForumHide],
	CONSTRAINT [DF_BBSXP_Forums_LastTime] DEFAULT (getdate()) FOR [LastTime],
	CONSTRAINT [PK_BBSXP_Forums] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BBSXP_Log] ADD 
	CONSTRAINT [DF_BBSXP_Log_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_Menu] ADD 
	CONSTRAINT [DF_BBSXP_Menu_FollowID] DEFAULT (0) FOR [FollowID],
	CONSTRAINT [DF_BBSXP_Menu_SortNum] DEFAULT (0) FOR [SortNum]
GO

ALTER TABLE [dbo].[BBSXP_Messages] ADD 
	CONSTRAINT [DF_BBSXP_Messages_DateCreated] DEFAULT (getdate()) FOR [DateCreated]
GO

ALTER TABLE [dbo].[BBSXP_PostAttachments] ADD 
	CONSTRAINT [DF__BBSXP_Pos__Threa__71D1E811] DEFAULT (0) FOR [ThreadID],
	CONSTRAINT [DF__BBSXP_Pos__Creat__72C60C4A] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__BBSXP_Pos__Conte__73BA3083] DEFAULT (0) FOR [ContentSize],
	CONSTRAINT [DF__BBSXP_Pos__Total__74AE54BC] DEFAULT (0) FOR [TotalDownloads],
	CONSTRAINT [PK_BBSXP_PostAttachments] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BBSXP_Posts] ADD 
	CONSTRAINT [DF_BBSXP_Posts_ThreadID] DEFAULT (0) FOR [ThreadID],
	CONSTRAINT [DF_BBSXP_Posts_IsTopic] DEFAULT (0) FOR [IsTopic],
	CONSTRAINT [DF_BBSXP_Posts_PostTime] DEFAULT (getdate()) FOR [PostTime]
GO

ALTER TABLE [dbo].[BBSXP_Prison] ADD 
	CONSTRAINT [DF_BBSXP_Prison_ComeTime] DEFAULT (getdate()) FOR [ComeTime]
GO

ALTER TABLE [dbo].[BBSXP_Ranks] ADD 
	CONSTRAINT [DF_BBSXP_Ranks_PostingCountMin] DEFAULT (0) FOR [PostingCountMin]
GO

ALTER TABLE [dbo].[BBSXP_SiteSettings] ADD 
	CONSTRAINT [DF_BBSXP_SiteSettings_ThreadsPerPage] DEFAULT (20) FOR [ThreadsPerPage],
	CONSTRAINT [DF_BBSXP_SiteSettings_PostsPerPage] DEFAULT (15) FOR [PostsPerPage],
	CONSTRAINT [DF_BBSXP_SiteSettings_Timeout] DEFAULT (60) FOR [Timeout],
	CONSTRAINT [DF_BBSXP_SiteSettings_UserOnlineTime] DEFAULT (20) FOR [UserOnlineTime],
	CONSTRAINT [DF_BBSXP_SiteSettings_DefaultSiteStyle] DEFAULT (1) FOR [DefaultSiteStyle],
	CONSTRAINT [DF_BBSXP_SiteSettings_DisplayForumFloor] DEFAULT (0) FOR [DisplayForumFloor],
	CONSTRAINT [DF_BBSXP_SiteSettings_DisplayWhoIsOnline] DEFAULT (1) FOR [DisplayWhoIsOnline],
	CONSTRAINT [DF_BBSXP_SiteSettings_DisplayLink] DEFAULT (1) FOR [DisplayLink],
	CONSTRAINT [DF_BBSXP_SiteSettings_DuplicatePostIntervalInMinutes] DEFAULT (0) FOR [DuplicatePostIntervalInMinutes],
	CONSTRAINT [DF_BBSXP_SiteSettings_RegUserTimePost] DEFAULT (0) FOR [RegUserTimePost],
	CONSTRAINT [DF_BBSXP_SiteSettings_PopularPostThresholdPosts] DEFAULT (15) FOR [PopularPostThresholdPosts],
	CONSTRAINT [DF_BBSXP_SiteSettings_PopularPostThresholdViews] DEFAULT (150) FOR [PopularPostThresholdViews],
	CONSTRAINT [DF_BBSXP_SiteSettings_MaxFileSize] DEFAULT (1024000) FOR [MaxFileSize],
	CONSTRAINT [DF_BBSXP_SiteSettings_MaxFaceSize] DEFAULT (10240) FOR [MaxFaceSize],
	CONSTRAINT [DF_BBSXP_SiteSettings_MaxPhotoSize] DEFAULT (102400) FOR [MaxPhotoSize],
	CONSTRAINT [DF_BBSXP_SiteSettings_MinVoteOptions] DEFAULT (2) FOR [MinVoteOptions],
	CONSTRAINT [DF_BBSXP_SiteSettings_MaxVoteOptions] DEFAULT (20) FOR [MaxVoteOptions],
	CONSTRAINT [DF_BBSXP_SiteSettings_ForumApply] DEFAULT (0) FOR [ForumApply],
	CONSTRAINT [DF_BBSXP_SiteSettings_SortShowForum] DEFAULT (0) FOR [SortShowForum],
	CONSTRAINT [DF_BBSXP_SiteSettings_OnlyMailReg] DEFAULT (0) FOR [OnlyMailReg],
	CONSTRAINT [DF_BBSXP_SiteSettings_EnableUser] DEFAULT (0) FOR [EnableUser],
	CONSTRAINT [DF_BBSXP_SiteSettings_CloseRegUser] DEFAULT (0) FOR [CloseRegUser],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralAddThread] DEFAULT (3) FOR [IntegralAddThread],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralAddPost] DEFAULT (1) FOR [IntegralAddPost],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralAddValuedPost] DEFAULT (5) FOR [IntegralAddValuedPost],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralDeleteThread] DEFAULT ((-3)) FOR [IntegralDeleteThread],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralDeletePost] DEFAULT ((-1)) FOR [IntegralDeletePost],
	CONSTRAINT [DF_BBSXP_SiteSettings_IntegralDeleteValuedPost] DEFAULT ((-5)) FOR [IntegralDeleteValuedPost],
	CONSTRAINT [DF_BBSXP_SiteSettings_DisplayEditNotes] DEFAULT (0) FOR [DisplayEditNotes],
	CONSTRAINT [DF_BBSXP_SiteSettings_EnableAntiSpamTextGenerateForRegister] DEFAULT (0) FOR [EnableAntiSpamTextGenerateForRegister],
	CONSTRAINT [DF_BBSXP_SiteSettings_EnableAntiSpamTextGenerateForLogin] DEFAULT (0) FOR [EnableAntiSpamTextGenerateForLogin],
	CONSTRAINT [DF_BBSXP_SiteSettings_EnableAntiSpamTextGenerateForPost] DEFAULT (0) FOR [EnableAntiSpamTextGenerateForPost],
	CONSTRAINT [DF_BBSXP_SiteSettings_DisplayPostIP] DEFAULT (0) FOR [DisplayPostIP],
	CONSTRAINT [DF__BBSXP_Sit__MaxPo__6C190EBB] DEFAULT (10240000) FOR [MaxPostAttachmentsSize],
	CONSTRAINT [DF__BBSXP_Sit__Water__6D0D32F4] DEFAULT (0) FOR [WatermarkType],
	CONSTRAINT [DF__BBSXP_Sit__Water__6E01572D] DEFAULT (0) FOR [WatermarkPosition],
	CONSTRAINT [DF__BBSXP_Sit__Attac__6EF57B66] DEFAULT (1) FOR [AttachmentsSaveOption]
GO

ALTER TABLE [dbo].[BBSXP_Statistics_Site] ADD 
	CONSTRAINT [DF__BBSXP_Sta__Total__76969D2E] DEFAULT (0) FOR [TotalUser],
	CONSTRAINT [DF__BBSXP_Sta__Total__778AC167] DEFAULT (0) FOR [TotalThread],
	CONSTRAINT [DF__BBSXP_Sta__Total__787EE5A0] DEFAULT (0) FOR [TotalPost],
	CONSTRAINT [DF__BBSXP_Sta__Today__797309D9] DEFAULT (0) FOR [TodayPost],
	CONSTRAINT [DF__BBSXP_Sta__BestO__7A672E12] DEFAULT (0) FOR [BestOnline],
	CONSTRAINT [DF__BBSXP_Sta__BestO__7B5B524B] DEFAULT (getdate()) FOR [BestOnlineTime]
GO

ALTER TABLE [dbo].[BBSXP_Threads] ADD 
	CONSTRAINT [DF_BBSXP_Threads_ForumID] DEFAULT (0) FOR [ForumID],
	CONSTRAINT [DF_BBSXP_Threads_Icon] DEFAULT (0) FOR [Icon],
	CONSTRAINT [DF_BBSXP_Threads_IsGood] DEFAULT (0) FOR [IsGood],
	CONSTRAINT [DF_BBSXP_Threads_IsTop] DEFAULT (0) FOR [IsTop],
	CONSTRAINT [DF_BBSXP_Threads_IsLocked] DEFAULT (0) FOR [IsLocked],
	CONSTRAINT [DF_BBSXP_Threads_IsDel] DEFAULT (0) FOR [IsDel],
	CONSTRAINT [DF_BBSXP_Threads_IsVote] DEFAULT (0) FOR [IsVote],
	CONSTRAINT [DF_BBSXP_Threads_Views] DEFAULT (0) FOR [Views],
	CONSTRAINT [DF_BBSXP_Threads_Replies] DEFAULT (0) FOR [Replies],
	CONSTRAINT [DF__BBSXP_Thr__PostT__1332DBDC] DEFAULT (getdate()) FOR [PostTime],
	CONSTRAINT [PK_BBSXP_Threads] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BBSXP_Users] ADD 
	CONSTRAINT [DF_BBSXP_Users_Membercode] DEFAULT (1) FOR [Membercode],
	CONSTRAINT [DF_BBSXP_Users_NewMessage] DEFAULT (0) FOR [NewMessage],
	CONSTRAINT [DF_BBSXP_Users_PostTopic] DEFAULT (0) FOR [PostTopic],
	CONSTRAINT [DF_BBSXP_Users_PostRevert] DEFAULT (0) FOR [PostRevert],
	CONSTRAINT [DF_BBSXP_Users_DelTopic] DEFAULT (0) FOR [DelTopic],
	CONSTRAINT [DF_BBSXP_Users_GoodTopic] DEFAULT (0) FOR [GoodTopic],
	CONSTRAINT [DF_BBSXP_Users_UserMoney] DEFAULT (0) FOR [UserMoney],
	CONSTRAINT [DF_BBSXP_Users_SaveMoney] DEFAULT (0) FOR [SaveMoney],
	CONSTRAINT [DF_BBSXP_Users_Experience] DEFAULT (0) FOR [Experience],
	CONSTRAINT [DF_BBSXP_Users_UserDegree] DEFAULT (0) FOR [UserDegree],
	CONSTRAINT [DF_BBSXP_Users_UserRegTime] DEFAULT (getdate()) FOR [UserRegTime],
	CONSTRAINT [DF_BBSXP_Users_UserLandTime] DEFAULT (getdate()) FOR [UserLandTime],
	CONSTRAINT [DF_BBSXP_Users_SaveMoneyTime] DEFAULT (getdate()) FOR [SaveMoneyTime],
	CONSTRAINT [PK_BBSXP_Users] PRIMARY KEY  CLUSTERED 
	(
		[UserName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BBSXP_UsersOnline] ADD 
	CONSTRAINT [DF_BBSXP_UsersOnline_Eremite] DEFAULT (0) FOR [Eremite],
	CONSTRAINT [DF_BBSXP_UsersOnline_ComeTime] DEFAULT (getdate()) FOR [ComeTime],
	CONSTRAINT [DF_BBSXP_UsersOnline_LastTime] DEFAULT (getdate()) FOR [LastTime]
GO

ALTER TABLE [dbo].[BBSXP_Vote] ADD 
	CONSTRAINT [DF_BBSXP_Vote_ThreadID] DEFAULT (0) FOR [ThreadID],
	CONSTRAINT [DF_BBSXP_Vote_Type] DEFAULT (0) FOR [Type]
GO

ALTER TABLE [dbo].[BBSXP_Calendar] ADD 
	CONSTRAINT [FK_BBSXP_Calendar_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Consort] ADD 
	CONSTRAINT [FK_BBSXP_Consort_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Consortia] ADD 
	CONSTRAINT [FK_BBSXP_Consortia_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Favorites] ADD 
	CONSTRAINT [FK_BBSXP_Favorites_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Messages] ADD 
	CONSTRAINT [FK_BBSXP_Messages_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_PostAttachments] ADD 
	CONSTRAINT [FK_BBSXP_PostAttachments_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Posts] ADD 
	CONSTRAINT [FK_BBSXP_Posts_BBSXP_Threads] FOREIGN KEY 
	(
		[ThreadID]
	) REFERENCES [dbo].[BBSXP_Threads] (
		[ID]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Prison] ADD 
	CONSTRAINT [FK_BBSXP_Prison_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Threads] ADD 
	CONSTRAINT [FK_BBSXP_Threads_BBSXP_Forums] FOREIGN KEY 
	(
		[ForumID]
	) REFERENCES [dbo].[BBSXP_Forums] (
		[ID]
	) ON DELETE CASCADE  ON UPDATE CASCADE ,
	CONSTRAINT [FK_BBSXP_Threads_BBSXP_Users] FOREIGN KEY 
	(
		[UserName]
	) REFERENCES [dbo].[BBSXP_Users] (
		[UserName]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[BBSXP_Vote] ADD 
	CONSTRAINT [FK_BBSXP_Vote_BBSXP_Threads] FOREIGN KEY 
	(
		[ThreadID]
	) REFERENCES [dbo].[BBSXP_Threads] (
		[ID]
	) ON DELETE CASCADE  ON UPDATE CASCADE 
GO

insert into [dbo].[BBSXP_SiteSettings](SiteName) values('BBSxp Board')
insert into [dbo].[BBSXP_Statistics_Site](NewUser) values('')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('风格','#','0','1')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('社区服务','#','0','1')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('默认风格','cookies.asp?menu=skins','1','1')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('蓝色风格','cookies.asp?menu=skins&no=1','1','2')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('橘色风格','cookies.asp?menu=skins&no=2','1','3')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('绿色风格','cookies.asp?menu=skins&no=3','1','4')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('社区银行','bank.asp','2','1')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('社区监狱','prison.asp','2','2')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('社区公会','Consortia.asp','2','3')
insert into [dbo].[BBSxp_menu] (name,url,followid,SortNum) values ('社区配偶','consort.asp','2','4')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('工兵','0','Images/level/1.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('排长','50','Images/level/2.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('连长','150','Images/level/3.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('营长','500','Images/level/4.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('团长','1000','Images/level/5.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('旅长','2000','Images/level/6.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('师长','5000','Images/level/7.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('军长','20000','Images/level/8.gif')
insert into [dbo].[BBSXP_Ranks] (RankName,PostingCountMin,RankIconUrl) values ('司令','50000','Images/level/9.gif')
