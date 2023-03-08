<%@ CodePage="936" Language="VBScript"%>
<!--#include file="Config.asp"-->
<!-- #include file="BBSXP_Class.asp" -->
<%
if Request.ServerVariables("HTTP_USER_AGENT") = "" then Response.End

ForumsDataBaseVersion="8.0.5"
ForumsBuild="8.0.5.081208"
ForumsProgram="BBSXP 2008 SP2 "&IsSqlVer&""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error Resume Next

Set Conn=Server.CreateObject("ADODB.Connection")
Conn.open ConnStr
If Err Then
	if Err.Number=-2147467259 then
		err.Clear
		response.Redirect("Install.asp")
	else
		Response.Write ""&IsSqlVer&"数据库连接出错，请检查连接字串。<br /><br />"&Err.Source&" ("&Err.Number&")"
		Set Conn = Nothing
		err.Clear
		Response.End
	end if
End If

Set Rs=Execute("Select * from ["&TablePrefix&"SiteSettings]")
if Err then
	err.Clear
	response.Redirect("Install.asp")
else
	CreateUserAgreement = Rs("CreateUserAgreement")
	GenericHeader = Rs("GenericHeader")
	GenericTop = Rs("GenericTop")
	GenericBottom = Rs("GenericBottom")
	AdminNotes = Rs("AdminNotes")
	BestOnline = Rs("BestOnline")
	BestOnlineTime = Rs("BestOnlineTime")
	SiteSettingsXML=Rs("SiteSettingsXML")
	BBSxpDataBaseVersion = Rs("BBSxpVersion")
end if
Rs.Close
Set Rs=Nothing

Set SiteConfigXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
SiteConfigXMLDOM.loadxml("<bbsxp>"&SiteSettingsXML&"</bbsxp>")


Server.ScriptTimeout=SiteConfig("Timeout")	'设置脚本超时时间 单位:秒

Request_Method=Request.ServerVariables("Request_method")
REMOTE_ADDR=Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR"))
Script_Name=LCase(Request.ServerVariables("script_name"))
Server_Name=Request.ServerVariables("server_name")
Query_String=Request.ServerVariables("Query_String")
Http_Referer=Request.ServerVariables("http_referer")
ReturnUrl=Request("ReturnUrl")
startime=timer()
CookieUserRoleID=0


if IsNumeric(RequestCookies("UserID")) then
	Set Rs=Execute("Select * from ["&TablePrefix&"Users] where UserID="&RequestCookies("UserID")&"")
	if Rs.eof then
		CleanCookies()
	elseif SiteConfig("AllowLogin")=0 and Rs("UserRoleID")<>1 or Rs("UserAccountStatus")<>1 then
		CleanCookies()
	elseif RequestCookies("UserPassword") <> Rs("UserPassword") then
		CleanCookies()
	else
		CookieUserID=Rs("UserID")
		CookieUserName=Rs("UserName")
		CookieUserPassword=Rs("UserPassword")
		CookieUserEmail=Rs("UserEmail")
		CookieUserMoney=Rs("UserMoney")
		CookieExperience=Rs("Experience")
		CookieNewMessage=Rs("NewMessage")
		CookieUserAccountStatus=Rs("UserAccountStatus")
		CookieUserRoleID=Rs("UserRoleID")
		CookieUserMate=Rs("UserMate")
		CookieUserRegisterTime=Rs("UserRegisterTime")
		CookieUserPostTime=Rs("UserPostTime")
		CookieModerationLevel=Rs("ModerationLevel")
		CookieReputation=Rs("Reputation")
		CookieTotalPosts=Rs("TotalPosts")
		CookieUserActivityTime=Rs("UserActivityTime")		
		
		if CookieUserRoleID=1 or CookieUserRoleID=2 then BestRole=1
	end if
	Rs.Close
	Set Rs=Nothing
	if FormatDateTime(now(),2) <> FormatDateTime(CookieUserActivityTime,2) then Execute("update ["&TablePrefix&"Users] Set UserActivityTime="&SqlNowString&",UserActivityDay=UserActivityDay+1,UserActivityIP='"&REMOTE_ADDR&"' where UserID="&CookieUserID&"")
end if

'第一次加载更新缓存
if ""&RequestApplication("LastUpdate")&""="" then
	UpdateApplication"UserBirthday","Select UserID,UserName From ["&TablePrefix&"Users] where Month(Birthday)="&Month(now())&" and day(Birthday)="&day(now())&""
	UpdateApplication"Links","Select name,Logo,url,Intro From ["&TablePrefix&"Links] where SortOrder>0 order by SortOrder"
	UpdateApplication"Advertisements","select Body from ["&TablePrefix&"Advertisements] where DateDiff("&SqlChar&"d"&SqlChar&","&SqlNowString&",ExpireDate) > 0"
	ResponseApplication"LastUpdate",CStr(Now)
end if

On Error GoTo 0


if ""&RequestCookies("Themes")&""="" then ResponseCookies "Themes",SiteConfig("DefaultSiteStyle"),9999

Set Rs = Server.CreateObject("ADODB.Recordset")
set AutoTerminate=new AutoTerminate_Class

dim SqlQueryNum,ForumsList,ForumTreeList,TotalPage,PageCount,IsResponseTop,MailSubject,MailBody
%>