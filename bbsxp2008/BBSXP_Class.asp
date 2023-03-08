<!-- #include file="Utility/HashPassword_Class.asp" -->
<%
Function Execute(Command)
	SqlQueryNum = SqlQueryNum + 1
	'Response.Write "（"&SqlQueryNum&"）"&Command&"<p>"
	Set Execute = Conn.Execute(Command)
End Function


''''''''''''''''''''''''''''''''''''
Class AutoTerminate_Class
	Private Sub Class_Terminate
		if Err.Number<>0 then
			'955 = 未知的运行时错误
			'-2147217864 = 乐观并发检查失败。该行在此游标之外进行了修改。
			If Err.Number<>995 and Err.Number<>-2147217864 then
				set fso = server.CreateObject("Scripting.FileSystemObject")
				set f = fso.OpenTextFile(Server.MapPath("Log/"&FormatDateTime(now(),2)&"_500.log"),8,1)
				f.WriteLine(now()&"	"&Request.ServerVariables("REMOTE_ADDR")&"	"&Request.ServerVariables("Request_method")&"	"&Request.ServerVariables("server_name")&"	"&Request.ServerVariables("script_name")&"	"&Request.ServerVariables("Query_String")&"	"&Err.Number&"	"&Err.Source&"	"&Err.Description&"	"&Request.ServerVariables("All_Http")) 
				f.Close() 
				set f = nothing 
				set fso = nothing
			end if
		end if
		Conn.Close
		Set Rs = Nothing
		Set Conn = Nothing
		Set SiteConfigXMLDOM = Nothing
	End Sub
End Class




''''''''''''''''''''''''''''''''''''
Function SiteConfig(str)
	TextStr=SiteConfigXMLDOM.documentElement.SelectSingleNode(str).text
	if IsNumeric(TextStr) then
		str=int(TextStr)	'转化为数字类型
		if Len(str)<>Len(TextStr) then	str=TextStr	'防止数字前面的 0 消失掉
	else
		str=TextStr
	End If
	SiteConfig=str
End Function



''''''''''''''''''''''''''''''''''''
Function HTMLEncode(fString)
	fString=Trim(fString)
	fString=Replace(fString,CHR(9),"")
	fString=Replace(fString,CHR(13),"")
	fString=Replace(fString,CHR(22),"")
	fString=Replace(fString,CHR(38),"&#38;")	'“&”
	fString=Replace(fString,CHR(32),"&#32;")	'“ ”
	fString=Replace(fString,CHR(34),"&quot;")	'“"”
	fString=Replace(fString,CHR(37),"&#37;")	'“%”
	fString=Replace(fString,CHR(39),"&#39;")	'“'”
	fString=Replace(fString,CHR(42),"&#42;")	'“*”
	fString=Replace(fString,CHR(43),"&#43;")	'“+”
	fString=Replace(fString,CHR(44),"&#44;")	'“,”
	fString=Replace(fString,CHR(45)&CHR(45),"&#45;&#45;")	'“--”
	fString=Replace(fString,CHR(92),"&#92;")	'“\”
	'fString=Replace(fString,CHR(95),"&#95;")	'“_”
	fString=Replace(fString,CHR(40),"&#40;")	'“(”
	fString=Replace(fString,CHR(41),"&#41;")	'“)”
	fString=Replace(fString,CHR(60),"&#60;")	'“<”
	fString=Replace(fString,CHR(62),"&#62;")	'“>”
	fString=Replace(fString,CHR(123),"&#123;")	'“{”
	fString=Replace(fString,CHR(125),"&#125;")	'“}”
	fString=Replace(fString,CHR(59),"&#59;")	'“;”
	fString=Replace(fString,CHR(10),"<br>")
	fString=ReplaceText(fString,"([&#])([a-z0-9]*)&#59;","$1$2;")

	if SiteConfig("BannedText")<>"" then fString=ReplaceText(fString,"("&SiteConfig("BannedText")&")",string(len("&$1&"),"*"))

	if IsSqlDataBase=0 then '过滤片假名(日文字符)[\u30A0-\u30FF] by yuzi
		fString=escape(fString)
		fString=ReplaceText(fString,"%u30([A-F][0-F])","&#x30$1;")
		fString=unescape(fString)
	end if

	HTMLEncode=fString
End Function
''''''''''''''''''''''''''''''''''''
Function BodyEncode(fString)
	fString=Trim(fString)
	fString=Replace(fString, "<br />",vbCrlf)
	fString=Replace(fString,"\","&#92;")
	fString=Replace(fString,"'","&#39;")
	fString=Replace(fString,CHR(60),"&#60;")	'“<”
	fString=Replace(fString,CHR(62),"&#62;")	'“>”
	fString=Replace(fString,"  ","&nbsp; ")
	fString=Replace(fString,"<a href=","<a target=_blank href=") '点链接打开新窗口
	if SiteConfig("BannedText")<>"" then fString=ReplaceText(fString,"("&SiteConfig("BannedText")&")",string(len("&$1&"),"*"))
	BodyEncode=fString
End Function
''''''''''''''''''''''''''''''''''''
Function BBCode(str)
	str=ReplaceText(str,"\[(\/|)(b|i|u|strike|center|marquee)\]","<$1$2>")
	str=ReplaceText(str,"\[align=(center|left|right)\]((?:(?!\[\/align\])[\s\S])*)\[\/align\]","<p align=""$1"">$2</p>")
	str=ReplaceText(str,"\[float=(center|left|right)\]((?:(?!\[\/float\])[\s\S])*)\[\/float\]","<div style=""float:$1"">$2</style>")
	str=ReplaceText(str,"\[COLOR=([^[]*)\]","<font COLOR=$1>")
	str=ReplaceText(str,"\[FONT=([^[]*)\]","<font face=$1>")
	str=ReplaceText(str,"\[SIZE=([0-9]*[pt|px|%]*)\]","<font size=$1>")
	str=ReplaceText(str,"\[\/(SIZE|FONT|COLOR)\]","</font>")
	str=ReplaceText(str,"\[list\]","<ul>")
	str=ReplaceText(str,"\[list=([\w]+)\]","<ul type=$1>")
	str=ReplaceText(str,"\[\/list\]","</ul>")
	str=ReplaceText(str,"\[\*\]","<li>")
	str=ReplaceText(str,"\[media=([\d]*),([\d]*),(true|false),(true|false)\]((?:(?!\[\/media\])[\s\S])*)\[\/media\]","<embed style=""width: $1px; HEIGHT: $2px"" src=""$5"" type=""audio/x-ms-wma"" autostart=""$3"" ShowStatusBar=""$4"" quality=""high"" />")
	
    '表格
	if RegExpTest("\[table(=.[^\[]*)?\]((?:(?!\[\/table\])[\s\S])*)\[\/table\]",""&str&"") then
		str=ReplaceText(str,"\[td=([0-9]*),([0-9]*),(.*?)\]","<td colspan=""$1"" rowspan=""$2"" width=""$3"">")
		str=ReplaceText(str,"\[td=([0-9]*),([0-9]*)\]","<td colspan=""$1"" rowspan=""$2"">")
		str=ReplaceText(str,"\[tr(=.[^\[]*)?\]","<tr>")
		str=ReplaceText(str,"\[\/tr\]","</tr>")
		str=ReplaceText(str,"\[td\]","<td>")
		str=ReplaceText(str,"\[\/td\]","</td>")
		do while RegExpTest("\[table(=.[^\[]*)?\]((?:(?!\[\/table\])[\s\S])*)\[\/table\]",""&str&"")
			str=ReplaceText(str,"\[table(=.[^\[]*)?\]((?:(?!\[\/table\])[\s\S])*)\[\/table\]","<table width=""$1"" border=""1"" style=""border-collapse:collapse;"">$2</table>")
			str=ReplaceText(str,"\[table\]((?:(?!\[\/table\])[\s\S])*)\[\/table\]","<table border=""1"" style=""border-collapse:collapse;"">$1</table>")
		loop
	end if

	str=ReplaceText(str,"\[(url|ed2k)\]ed2k:\/\/\|file\|([^\\\/:*?<>""|]+[\.]?[^\\\/:*?<>""|]+)\|(\d+)\|([0-9a-zA-Z]{32})((\|[^[]*)?)\|\/\[\/(url|ed2k)\]",""&vbCrlf&_
	"<br /><table align=""center"" cellspacing=""1"" cellpadding=""5"" width=""100%"" class=""CommonListArea"">"&vbCrlf&_
	"<tr align=center class=""CommonListHeader""><td>文件名</td><td width=""100"">大小</td></tr>"&vbCrlf&_
	"<tr class=""CommonListCell""><td><a href=""ed2k://|file|$2|$3|$4$5|/"" target=_blank>$2</a> (<a href=""http://www.ed2000.com/ed2kstats/?hash=$4"" target=""_blank"">查源</a>)</td><td align=center><script language=""javascript"">document.write(gen_size($3, 3, 1));</script></td></tr>"&vbCrlf&_
	"<tr class=""CommonListCell""><td colspan=""2""><input type=""button"" value=""下载该资源"" onClick=""download('ed2k://|file|$2|$3|$4$5|/')"" /> <input type=""button"" value=""复制ED2K链接"" onClick=""copyToClipboard('ed2k://|file|$2|$3|$4$5|/')"" /> <span style=""float:right;margin-top:-17px;""><a href=""http://www.ed2000.com/download/"" target=""_blank"">推荐使用eMule进行下载</a></span></td></tr>"&vbCrlf&_
	"</table><br />")

	str=ReplaceText(str,"\[ed2k\]((?:(?!\[\/(url|ed2k)\])[\s\S])*)\[\/ed2k\]",""&vbCrlf&_
	"<br /><table align=""center"" cellspacing=""1"" cellpadding=""5"" width=""100%"" class=""CommonListArea"">"&vbCrlf&_
	"<tr align=center class=""CommonListHeader""><td width=25><input type=""checkbox"" id=""checkall_ed2klist"" onclick=""checkAll('ed2klist',this.checked)"" /></td><td>文件名</td><td width=""100"">大小</td></tr>$1"&vbCrlf&_
	"<tr class=""CommonListCell""><td colspan=""3""><input type=""button"" value=""下载选中资源"" onClick=""Download('ed2klist',0,1)"" /> <input type=""button"" value=""复制ED2K链接"" onClick=""copy('ed2klist')"" /> <span style=""float:right;margin-top:-17px;""><a href=""http://www.ed2000.com/download/"" target=""_blank"">推荐使用eMule进行下载</a></span></td></tr>"&vbCrlf&_
	"</table><br />")

	str=ReplaceText(str,"\[link\]ed2k:\/\/\|file\|([^\\\/:*?<>""|]+[\.]?[^\\\/:*?<>""|]+)\|(\d+)\|([0-9a-zA-Z]{32})((\|[^[]*)?)\|\/\[\/link\]","<tr class=""CommonListCell""><td align=center><input type=""checkbox"" value=""ed2k://|file|$1|$2|$3$4|/"" onclick=""em_size(''+this.name+'');"" name=""ed2klist"" /></td><td><a href=""ed2k://|file|$1|$2|$3$4|/"" target=_blank>$1</a> (<a href=""http://www.ed2000.com/ed2kstats/?hash=$3"" target=""_blank"">查源</a>)</td><td align=center><script language=""javascript"">document.write(gen_size($2, 3, 1));</script></td></tr>")

	str=ReplaceText(str,"\[URL\]([^[]*)","<a target=_blank href=$1>$1")
	str=ReplaceText(str,"\[URL=([^[]*)\]","<a target=_blank href=$1>")
	str=ReplaceText(str,"\[\/URL\]","</a>")
	str=ReplaceText(str,"\[EMAIL\](\S+\@[^[]*)(\[\/EMAIL\])","<a href=mailto:$1>$1</a>")
	str=ReplaceText(str,"\[EMAIL=(\S+\@[^[]*)\](\S+\@[^[]*)(\[\/EMAIL\])","<a href=mailto:$1>$2</a>")
	str=ReplaceText(str,"\[IMG\]([^("&CHR(34)&"|[|#)]*)(\[\/IMG\])","<img border=0 src=$1 />")
	str=ReplaceText(str,"\[IMG=([0-9]*),([0-9]*)\]([^("&CHR(34)&"|[|#)]*)(\[\/IMG\])","<img width=$1 height=$2 border=0 src=$3 />")
	str=ReplaceText(str,"\[quote\]","<blockquote>")
	str=ReplaceText(str,"\[quote user="&CHR(34)&"([^[]*)"&CHR(34)&"\]","<blockquote><img border=0 src=images/icon-quote.gif> <b>$1:</b><br>")
	str=ReplaceText(str,"\[\/quote\]","</blockquote>")
	str=Replace(str,"<%","&lt;%")
	str=replace(str,vbcrlf,"<br />")

	BBCode=str
End Function

''''''''''''''''''''''''''''''''''''
Function RequestInt(fString)
	RequestInt=Request(fString)
	if IsNumeric(RequestInt) then
		RequestInt=int(RequestInt)
	else
		RequestInt=0
	end if
End Function

''''''''''替换模块START''''''''''''
Function ReplaceText(fString,patrn,replStr)
	Set regEx = New RegExp   	' 建立正则表达式。
		regEx.Pattern = patrn   ' 设置模式。
		regEx.IgnoreCase = True ' 设置是否区分大小写。
		regEx.Global = True     ' 设置全局可用性。 
		ReplaceText = regEx.Replace(""&fString&"",""&replStr&"") ' 作替换。
	Set regEx=nothing
End Function
''''''''''替换模块END''''''''''''
Function RegExpTest(patrn, strng)
	Set regEx = New RegExp   		' 建立正则表达式。
		regEx.Pattern = patrn   	' 设置模式。
		regEx.IgnoreCase = True 	' 设置是否区分大小写。
		regEx.Global = True     	' 设置全局可用性。 
		RegExpTest = regEx.Test(strng) ' 执行搜索。
	Set regEx=nothing
End Function

'''''''''''''''''''Cookies Process Start''''''''''''''''''''
Function ResponseCookies(Key,Value,Expires)
	Response.Cookies(Key) = ""&Value&""
	if ""&SiteConfig("CookieDomain")&""<>"" then Response.Cookies(Key).Domain = SiteConfig("CookieDomain")
	Response.Cookies(Key).Path = SiteConfig("CookiePath")
	if int(Expires) > 365 then Expires=365
	if int(Expires)>0 then Response.Cookies(Key).Expires = date+Expires
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function RequestCookies(CookieName)
	RequestCookies=Request.Cookies(CookieName)
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function CleanCookies()
	For Each objCookie In Request.Cookies
		ResponseCookies objCookie,"",0
	Next
	ResponseCookies "Themes",SiteConfig("DefaultSiteStyle"),365
End Function
'''''''''''''''''''Cookies Process End''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''
'写入Application
Function ResponseApplication(Key,Value)
	Application(SiteConfig("CacheName")&"_"&Key) = Value
End Function

'读取Application
Function RequestApplication(Key)
	RequestApplication=Application(SiteConfig("CacheName")&"_"&Key)
End Function

'删除Application
Function RemoveApplication(Key)
	Application.Contents.Remove(SiteConfig("CacheName")&"_"&Key)
End Function

'追加Application
Function AddApplication(Key,Value)
	Application(SiteConfig("CacheName")&"_"&Key) = Application(SiteConfig("CacheName")&"_"&Key)&Value&"<br>"
End Function

'更新缓存
Function UpdateApplication(Key,SQL)
	Application.Lock
		ResponseApplication Key,FetchEmploymentStatusList(SQL)
	Application.Unlock
End Function

'''''''''''''''''''''''''''''''''''''''''''
Function FetchEmploymentStatusList(SQL)
	Set Rs2=Execute(SQL)
	if Rs2.Eof then
		Rs2.Close
		Set Rs2 = Nothing
		Exit Function
	End if
  	FetchEmploymentStatusList = Rs2.GetRows()
	Rs2.Close
	Set Rs2 = Nothing
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	On Error GoTo 0
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function DelFile(DelFilePath)
	On Error Resume Next
	DelFile = False
	Set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
	MyFileObject.DeleteFile""&Server.MapPath(""&DelFilePath&"")&""
	Set MyFileObject = Nothing
	If 0 = Err or 53 = Err Then
		DelFile = True
	else
		Alert("出错讯息："&Err.Description&"\n"&DelFilePath&" 无法删除！")
	end if
	On Error GoTo 0
End Function

Function DelAttachments(SqlString)
	Set Rs2=Server.CreateObject("Adodb.Recordset")
	Rs2.open SqlString,Conn,1,3
		do while not Rs2.eof
			if ""&Rs2("FilePath")&""<>"" then IsDelFile=DelFile(""&Rs2("FilePath")&"")
			if ""&Rs2("FilePath")&""="" or (""&Rs2("FilePath")&""<>"" and IsDelFile=True) then
				Rs2.Delete()
				Rs2.Update()
			end if
			Rs2.movenext
		loop
	Rs2.Close
	set Rs2=nothing
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function CheckSize(ByteSize)
	if ByteSize=>1073741824 then
		ByteSize=formatnumber(ByteSize/1073741824)&" GB"
	elseif ByteSize=>1048576 then
		ByteSize=formatnumber(ByteSize/1048576)&" MB"
	elseif ByteSize=>1024 then
		ByteSize=formatnumber(ByteSize/1024)&" KB"
	else
		ByteSize=ByteSize&" 字节"
	end if
	CheckSize=ByteSize
End Function
'''''''''''''''''''''''''''''''''''''''''''


Function UpUserRank()
	Set Rs1=Execute("select top 1 RankName from ["&TablePrefix&"Ranks] where (RoleID="&Rs("UserRoleID")&" or RoleID=0) and PostingCountMin<="&Rs("TotalPosts")&" order by RoleID Desc,PostingCountMin Desc")
	if Not Rs1.Eof Then UpUserRank=Rs1("RankName")
	Rs1.close
	Set Rs1=nothing
End Function



Function ShowRole(value)
	select case value
		case "1"
			ShowRole="管理员"
		case "2"
			ShowRole="超级版主"
		case "3"
			ShowRole="注册用户"
		case else
			ShowRole=Execute("Select Name From ["&TablePrefix&"Roles] where RoleID="&value&"")(0)
	end select
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function ShowUserAccountStatus(value)
	select case value
		case "0"
			ShowUserAccountStatus="正等待审核"
		case "1"
			ShowUserAccountStatus="已通过审核"
		case "2"
			ShowUserAccountStatus="已禁用"
		case "3"
			ShowUserAccountStatus="未通过审核"
		case else
			ShowUserAccountStatus="未知状态"
	end select
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function ShowUserSex(value)
if SiteConfig("AllowGender")=1 then
	select case value
		case 0
			ShowUserSex=""
		case 1
			ShowUserSex="<img src=images/Sex_1.gif title='男'>"
		case 2
			ShowUserSex="<img src=images/Sex_2.gif title='女'>"
	end select
end if
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function Zodiac(birthday)
	if IsDate(birthday) then
		birthyear=year(birthday)
		ZodiacList=array("猴(Monkey)","鸡(Rooster)","狗(Dog)","猪(Boar)","鼠(Rat)","牛(Ox)","虎(Tiger)","兔(Rabbit)","龙(Dragon)","蛇(Snake)","马(Horse)","羊(Goat)")
		Zodiac=ZodiacList(birthyear mod 12)
	end if
End Function
'''''''''''''''''''''''''''''''''''''''''''
Function Horoscope(birthday)
	if IsDate(birthday) then
		HoroscopeMon=month(birthday)
		HoroscopeDay=day(birthday)
		if Len(HoroscopeMon)<2 then HoroscopeMon="0"&HoroscopeMon
		if Len(HoroscopeDay)<2 then HoroscopeDay="0"&HoroscopeDay
		MyHoroscope=HoroscopeMon&HoroscopeDay
		if MyHoroscope < 0120 then
			Horoscope="<img src=images/Horoscope/Capricorn.gif title='魔羯座 Capricorn'>"
		elseif MyHoroscope < 0219 then
			Horoscope="<img src=images/Horoscope/Aquarius.gif title='水瓶座 Aquarius'>"
		elseif MyHoroscope < 0321 then
			Horoscope="<img src=images/Horoscope/Pisces.gif title='双鱼座 Pisces'>"
		elseif MyHoroscope < 0420 then
			Horoscope="<img src=images/Horoscope/Aries.gif title='白羊座 Aries'>"
		elseif MyHoroscope < 0521 then
			Horoscope="<img src=images/Horoscope/Taurus.gif title='金牛座 Taurus'>"
		elseif MyHoroscope < 0622 then
			Horoscope="<img src=images/Horoscope/Gemini.gif title='双子座 Gemini'>"
		elseif MyHoroscope < 0723 then
			Horoscope="<img src=images/Horoscope/Cancer.gif title='巨蟹座 Cancer'>"
		elseif MyHoroscope < 0823 then
			Horoscope="<img src=images/Horoscope/Leo.gif title='狮子座 Leo'>"
		elseif MyHoroscope < 0923 then
			Horoscope="<img src=images/Horoscope/Virgo.gif title='处女座 Virgo'>"
		elseif MyHoroscope < 1024 then
			Horoscope="<img src=images/Horoscope/Libra.gif title='天秤座 Libra'>"
		elseif MyHoroscope < 1122 then
			Horoscope="<img src=images/Horoscope/Scorpio.gif title='天蝎座 Scorpio'>"
		elseif MyHoroscope < 1222 then
			Horoscope="<img src=images/Horoscope/Sagittarius.gif title='射手座 Sagittarius'>"
		elseif MyHoroscope > 1221 then
			Horoscope="<img src=images/Horoscope/Capricorn.gif title='魔羯座 Capricorn'>"
		end if
	end if
End Function

Function ShowReputation(value)
		if value < 4 then
			ShowReputation=""
		elseif value < 251 then
			ShowReputation="<img src=images/ReputationA.gif title='声望:"&value&"' >"
		elseif value < 10001 then
			ShowReputation="<img src=images/ReputationB.gif title='声望:"&value&"' >"
		elseif value > 10000 then
			ShowReputation="<img src=images/ReputationC.gif title='声望:"&value&"' >"
		end if
End Function


Function ShowUserActivityDay(value)
		if value < 5 then
			ShowUserActivityDay=""
		elseif value < 32 then
			ShowUserActivityDay="<img src=images/time_star.gif title='活跃天数:"&value&"' >"
		elseif value < 320 then
			ShowUserActivityDay ="<img src=images/time_moon.gif title='活跃天数:"&value&"' >"
		elseif value > 319 then
			ShowUserActivityDay ="<img src=images/time_sun.gif title='活跃天数:"&value&"' >"
		end if
End Function


'针对中文日期格式的服务器
Function FormatTime(value)
	FormatTime=""&FormatDateTime(value, 2)&" "&FormatDateTime(value, 4)&":"&second(value)&""
End Function

Function DefaultPasswordFormat(value)
	if SiteConfig("DefaultPasswordFormat")="MD5" then
		DefaultPasswordFormat=MD5(value)
	Else
		DefaultPasswordFormat=SHA1(value)
	End if
End Function


'''''''''''''''''''''''''''''''''''''''''''
Function ShowTags(PostID)
	sql="select * from ["&TablePrefix&"PostInTags] where PostID="&PostID&""
	Set Rs=Execute(sql)
	do while not Rs.eof
		ShowTags=ShowTags&","&Execute("Select TagName from ["&TablePrefix&"PostTags] where TagID="&Rs("TagID")&"")(0)
		Rs.movenext
	Loop
	Rs.Close
	ShowTags=Mid(ShowTags,2)
End Function

Function AddTags()
	Execute("Delete from ["&TablePrefix&"PostInTags] where PostID="&PostID&"")
	CookiePostTags=RequestCookies("CookiePostTags")
	TempTags=","
	for i=0 to Ubound(TagArray)
		TempTagstr=HTMLEncode(TagArray(i))
		if ""&TagArray(i)&""<>"" and instr(TempTags,","&TempTagstr&",")<=0 then
			Rs.open "select top 1 * from ["&TablePrefix&"PostTags] where TagName='"&TempTagstr&"'",Conn,1,3
				if Rs.eof then Rs.Addnew
				Rs("TagName")=TempTagstr
				Rs("MostRecentPostDate")=""&now()&""
				Rs.update
				TagID=Rs("TagID")
			Rs.Close
			
			Execute("insert into ["&TablePrefix&"PostInTags] (TagID,PostID) values ('"&TagID&"','"&PostID&"')")

			TagsStatic(TagID)
			if instr(CookiePostTags,","&TempTagstr&",")=0 then CookiePostTags=","&TempTagstr&CookiePostTags
			TempTags=TempTags&TempTagstr&","
		end if
	next
	ResponseCookies "CookiePostTags",CookiePostTags,30
End Function

Function TagsStatic(TagID)
	TotalPosts=Execute("Select Count(TagID) from ["&TablePrefix&"PostInTags] where TagID="&TagID&"")(0)
	Execute("Update ["&TablePrefix&"PostTags] set TotalPosts="&TotalPosts&" where TagID="&TagID&"")
	TagsStatic=TotalPosts
End Function

'''''''''''''''''''''''''''''''''''''''''''
Function GroupList(ParentID)
	GroupListGetRow=FetchEmploymentStatusList("Select GroupID,GroupName from ["&TablePrefix&"Groups] where SortOrder>0 order by SortOrder")
	if IsArray(GroupListGetRow) then 
	for i=0 to Ubound(GroupListGetRow,2)
		ForumsList=ForumsList&"<optgroup label='"&GroupListGetRow(1,i)&"'>"
		ii=ii+1
		ForumList GroupListGetRow(0,i),0,ParentID
		ii=ii-1
		ForumsList=ForumsList&"</optgroup>"
	next
	GroupList=ForumsList
	End if
	GroupListGetRow=null
End Function

Function ForumList(GroupID,ParentID,Selected)
	sql="Select ForumID,ForumName From ["&TablePrefix&"Forums] where GroupID="&GroupID&" and ParentID="&ParentID&" and SortOrder>0 and IsActive=1 order by SortOrder"
	Set Rs1=Execute(sql)
		Do While Not Rs1.EOF
			if RS1("ForumID")=Selected then
				ForumsList=ForumsList&"<option value='"&RS1("ForumID")&"' selected>"&string(ii,"　")&"-&raquo; "&RS1("ForumName")&"</option>"
			else
				ForumsList=ForumsList&"<option value='"&RS1("ForumID")&"'>"&string(ii,"　")&"-&raquo; "&RS1("ForumName")&"</option>"
			end if
			ii=ii+1
			ForumList GroupID,RS1("ForumID"),Selected
			ii=ii-1
		Rs1.MoveNext
		loop
	Rs1.Close
	Set Rs1 = Nothing
End Function

Function ForumTree(selec)
	if selec=0 then
		Set Rs1=Execute("Select * from ["&TablePrefix&"Groups] where GroupID="&GroupID&"")
		if not Rs1.eof then
			ForumTreeList="<span id=TempGroup"&GroupID&"><a onmouseover=Ajax_CallBack(false,'TempGroup"&GroupID&"','loading.asp?menu=ForumTree&GroupID="&GroupID&"') href=Default.asp?GroupID="&Rs1("GroupID")&">"&Rs1("GroupName")&"</a></span> → "&ForumTreeList&""
		end if
	else
		Set Rs1=Execute("Select * From ["&TablePrefix&"Forums] where ForumID="&selec&"")
		if not Rs1.eof then
			ForumTreeList="<span id=tempForum"&selec&"><a onmouseover=Ajax_CallBack(false,'tempForum"&selec&"','loading.asp?menu=ForumTree&ParentID="&selec&"') href=ShowForum.asp?ForumID="&Rs1("ForumID")&">"&Rs1("ForumName")&"</a></span> → "&ForumTreeList&""
			ForumTree Rs1("ParentID")
		end if
	end if
	Rs1.Close
	Set Rs1 = Nothing
	ForumTree=ForumTreeList
End Function



Function ClubTree()
	Set Rs1=Execute("Select * From ["&TablePrefix&"Groups] where SortOrder>0 order by SortOrder")
	do while not Rs1.eof
		ClubTreeList=ClubTreeList&"<div class=menuitems><a href=Default.asp?GroupID="&Rs1("GroupID")&">"&Rs1("GroupName")&"</a></div>"
		Rs1.Movenext
	loop
	Rs1.Close
	Set Rs1 = Nothing
	ClubTree="<a onmouseover="&Chr(34)&"showmenu(event,'"&ClubTreeList&"')"&Chr(34)&" href=Default.asp>"&SiteConfig("SiteName")&"</a>" 
End Function


''''''''''''''''''''''''''''''''
Sub ConciseMsg(Message)
	Response.write(Message)	
	Response.End
End Sub
''''''''''''''''''''''''
Sub Log(Message)
	MessageXML=MessageXML&"<Message>"&Message&"</Message>"&vbCrlf
	MessageXML=MessageXML&"<REMOTE_ADDR>"&REMOTE_ADDR&"</REMOTE_ADDR>"&vbCrlf
	MessageXML=MessageXML&"<Request_Method>"&Escape(Request.ServerVariables("Request_method"))&"</Request_Method>"&vbCrlf
	MessageXML=MessageXML&"<Server_Name>"&Escape(Request.ServerVariables("server_name"))&"</Server_Name>"&vbCrlf
	MessageXML=MessageXML&"<Script_Name>"&Escape(Request.ServerVariables("script_name"))&"</Script_Name>"&vbCrlf
	MessageXML=MessageXML&"<Query_String>"&Escape(Request.ServerVariables("Query_String"))&"</Query_String>"&vbCrlf
	MessageXML=MessageXML&"<Request_Form>"&Escape(Request.Form)&"</Request_Form>"&vbCrlf
	MessageXML=MessageXML&"<All_Http>"&Escape(Request.ServerVariables("All_Http"))&"</All_Http>"&vbCrlf

	Execute("insert into ["&TablePrefix&"EventLog] (UserName,ErrNumber,MessageXML) values ('"&CookieUserName&"','"&Err.Number&"','"&MessageXML&"')")
End Sub
''''''''''''''''''''''''''''''''

Function AjaxShowPage(TotalPage,PageIndex,url)
	AjaxShowPage=""
	AjaxShowPage="<span class='PageInation' style='float:right;'><a class=MultiPages>"&PageIndex&"/"&TotalPage&"</a>"
	if PageIndex<6 then
		PageLong=11-PageIndex
	elseif TotalPage-PageIndex<6 then
		PageLong=10-(TotalPage-PageIndex)
	else
		PageLong=5
	end if
	
	for i=1 to TotalPage
		if i < PageIndex+PageLong and i > PageIndex-PageLong or i=1 or i=TotalPage then
			if PageIndex=i then
				AjaxShowPage=AjaxShowPage&"<a class=CurrentPage>"& i &"</a>"
			else
				AjaxShowPage=AjaxShowPage&"<a class=PageNum href=""Javascript:Ajax_CallBack(false,'CommentArea','"&url&"&PageIndex="&i&"')"">"& i &"</a>"
			end if
		end if
	next
	AjaxShowPage=AjaxShowPage&"</span>"
End Function




''''''''''''''''''''''''''''''''

Sub UpdateStatistics(DaysUsers,DaysTopics,DaysPosts)

	sql="Select * from ["&TablePrefix&"Statistics] where DateDiff("&SqlChar&"d"&SqlChar&",DateCreated,"&SqlNowString&")=0"
	Rs.open sql,conn,1,3
	if Rs.eof then
		Rs.Addnew
		
		TotalUsers=Execute("Select count(UserID) from ["&TablePrefix&"Users]")(0)
		TotalTopics=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where Visible=1")(0)
		TotalPosts=Execute("Select sum(TotalReplies) as TotalPosts from ["&TablePrefix&"Threads] where Visible=1")(0)
		
		
		if IsNull(TotalPosts) then
		TotalPosts=0
		else
		NewestUserName=Execute("Select Top 1 UserName from ["&TablePrefix&"Users] order by UserID desc")(0)
		end if

		Rs("TotalUsers")=TotalUsers
		Rs("TotalTopics")=TotalTopics
		Rs("TotalPosts")=TotalPosts
		Rs("NewestUserName")=NewestUserName
		
		Execute("update ["&TablePrefix&"Forums] Set TodayPosts=0")
		Rs("DaysUsers")=Rs("DaysUsers")+int(DaysUsers)
		Rs("DaysTopics")=Rs("DaysTopics")+int(DaysTopics)
		Rs("DaysPosts")=Rs("DaysPosts")+int(DaysPosts)
		Rs("DateCreated")=date()
	else
		Rs("TotalUsers")=Rs("TotalUsers")+DaysUsers
		Rs("TotalTopics")=Rs("TotalTopics")+DaysTopics
		Rs("TotalPosts")=Rs("TotalPosts")+DaysPosts
		Rs("DaysUsers")=Rs("DaysUsers")+DaysUsers
		Rs("DaysTopics")=Rs("DaysTopics")+DaysTopics
		Rs("DaysPosts")=Rs("DaysPosts")+DaysPosts
	end if
	Rs.update
	Rs.close
End Sub



Sub UpForumMostRecent(ForumID)
		sql="Select top 1 * from ["&TablePrefix&"Threads] where ForumID="&ForumID&" and Visible=1 order by LastTime DESC"
		Set Rs2=Execute(sql)
		if Rs2.Eof then Exit sub
		MostRecentThreadID=Rs2("ThreadID")
		MostRecentPostSubject=Rs2("Topic")
		MostRecentPostAuthor=Rs2("LastName")
		MostRecentPostDate=Rs2("LastTime")
		Rs2.close
		Set Rs2 = Nothing
		Execute("update ["&TablePrefix&"Forums] Set MostRecentThreadID="&MostRecentThreadID&",MostRecentPostSubject='"&MostRecentPostSubject&"',MostRecentPostAuthor='"&MostRecentPostAuthor&"',MostRecentPostDate='"&FormatTime(MostRecentPostDate)&"' where ForumID="&ForumID&"")
End Sub



Sub UpdateThreadStatic(ThreadID)
	TotalReplies=Execute("select count(PostID) from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and ParentID>0 and Visible=1")(0)
	DeletedCount=Execute("select count(PostID) from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and Visible=2")(0)
	HiddenCount=Execute("select count(PostID) from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and Visible=0")(0)
	Visible=Execute("select Visible from ["&TablePrefix&"Posts] where ThreadID="&ThreadID&" and ParentID=0")(0)

	Execute("update ["&TablePrefix&"Threads] set TotalReplies="&TotalReplies&",DeletedCount="&DeletedCount&",HiddenCount="&HiddenCount&",Visible="&Visible&" where ThreadID="&ThreadID&"")
End Sub



Sub SendMail(MailAddRecipient,MailSubject,MailBody)
if  MailAddRecipient=""or MailSubject="" or MailBody="" then Exit Sub

on error resume next
MailSubject="("&SiteConfig("SiteName")&")"&MailSubject
MailBody=MailBody&"<br><br><br><a target=_blank href="&SiteConfig("SiteUrl")&"/Default.asp>"&SiteConfig("SiteName")&"</a> 管理团队<br><br><a target=_blank href=http://www.bbsxp.com>BBSXP</a>  &copy; 1998-"&year(now)&" <a target=_blank href=http://www.yuzi.net>YUZI Corporation.</a>"
if SiteConfig("SelectMailMode")="JMail.Message" then
	Set JMail=Server.CreateObject("JMail.Message")
		JMail.Charset=BBSxpCharset
		JMail.ContentType = "text/html"
		'JMail.ContentType = "text/plain"
		JMail.From = SiteConfig("SmtpServerMail")

			AddRecipientArray=split(MailAddRecipient,";")
			For i=0 to Ubound(AddRecipientArray)
				if ""&AddRecipientArray(i)&""<>"" then JMail.AddRecipient AddRecipientArray(i)
			Next

		JMail.Subject = MailSubject
		JMail.Body = MailBody
		JMail.MailServerUserName = SiteConfig("SmtpServerUserName")
		JMail.MailServerPassword = SiteConfig("SmtpServerPassword")
		JMail.Send SiteConfig("SmtpServer")
	Set JMail=nothing
elseif SiteConfig("SelectMailMode")="Persits.MailSender" then
	Set AspEmail = Server.CreateObject("Persits.MailSender")
 		AspEmail.Host = SiteConfig("SmtpServer")
		AspEmail.Username = SiteConfig("SmtpServerUserName")
		AspEmail.Password = SiteConfig("SmtpServerPassword")
		AspEmail.From = SiteConfig("SmtpServerMail")

			AddRecipientArray=split(MailAddRecipient,";")
			For i=0 to Ubound(AddRecipientArray)
				if ""&AddRecipientArray(i)&""<>"" then AspEmail.AddAddress AddRecipientArray(i)
			Next

		AspEmail.Subject = MailSubject
		AspEmail.Body = MailBody
		AspEmail.IsHTML = true
		AspEmail.Charset = BBSxpCharset
		AspEmail.Send
	Set AspEmail=Nothing
elseif SiteConfig("SelectMailMode")="CDO.Message" then
	Set CDO=Server.CreateObject("CDO.Message")
		CDO.From = SiteConfig("SmtpServerMail")
		CDO.To = MailAddRecipient
		CDO.Subject = MailSubject
		CDO.HtmlBody = MailBody
		'CDO.TextBody = MailBody
		CDO.HTMLBodyPart.Charset=BBSxpCharset
		CDO.Send
	Set CDO=Nothing
end if

If Err Then Response.Write ""&MailAddRecipient&"邮件发送失败！错误原因：" & Err.Description & "<br>"
On Error GoTo 0

End Sub

Sub LoadingEmailXml(emailType)
	Set EmailsXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
		EmailsXMLDOM.Load(Server.MapPath("Xml/emails.xml"))
		MailSubject = EmailsXMLDOM.documentElement.selectSingleNode("//emails/"&emailType&"/subject").Text
		Mailbody = EmailsXMLDOM.documentElement.selectSingleNode("//emails/"&emailType&"/body").Text
		Mailbody = Replace(Mailbody,CHR(10),"<br>")
	Set EmailsXMLDOM = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function CheckUser(UserName)
	if Len(UserName) < SiteConfig("UserNameMinLength") then CheckUser=CheckUser&"<li>您的用户名长度不能少于 "&SiteConfig("UserNameMinLength")&" 个字节</li>"
	if Len(UserName) > SiteConfig("UserNameMaxLength") then CheckUser=CheckUser&"<li>您的用户名长度不能超过 "&SiteConfig("UserNameMaxLength")&" 个字节</li>"

	
	
	ErrorChar=array("　","","","","","`","","+","-","#","@",".","%","&","\","/",":","*","?","<",">","|")
	for i=0 to ubound(ErrorChar)
		if instr(username,ErrorChar(i))>0 then CheckUser=CheckUser&"<li>用户名中不能含有“"&ErrorChar(i)&"”特殊符号</li>":Exit For
	next

	if SiteConfig("BannedRegUserName")<>"" then
		filtrate=split(""&SiteConfig("BannedRegUserName")&"","|")
		for i = 0 to ubound(filtrate)
			if filtrate(i)<>"" then
				if instr(UserName,filtrate(i))>0 then CheckUser=CheckUser&"<li>用户名中不能含有系统禁止注册的字符“"&filtrate(i)&"”</li>":Exit For
			End if
		next
	end if
End Function


Function AddUser
	AddUser=False

	if UserEmail="" then Message=Message&"<li>您的Email没有填写</li>"
	if instr(UserEmail,"@")=0 or instr(UserEmail,".")=0 then Message=Message&"<li>您的电子邮箱地址填写错误</li>"

	Message=Message&CheckUser(UserName)

	if Message<>"" then Exit Function
	
	if not Execute("Select UserID From ["&TablePrefix&"Users] where UserName='"&UserName&"'" ).eof Then Message=Message&"<li>"&UserName&" 已经被别人注册了</li>"
	If not Execute("Select UserID From ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"'" ).eof Then Message=Message&"<li>"&UserEmail&" 已经被别人注册了</li>"

	if Message<>"" then Exit Function

	
	Sql="Select * From ["&TablePrefix&"Users] Where UserName='"&UserName&"' or UserEmail='"&UserEmail&"'"
	Rs.Open Sql,Conn,1,3
		CheckStatus = 0
		Rs.addNew
			Rs("UserName")=UserName
			Rs("UserPassword")=DefaultPasswordFormat(""&UserPassword&"")
			Rs("UserEmail")=UserEmail
			if PasswordAnswer<>empty then
				Rs("PasswordQuestion")=PasswordQuestion
				Rs("PasswordAnswer")=md5(""&PasswordAnswer&"")
			end if
			Rs("UserFaceUrl")="images/face/Default.gif"
			Rs("ReferrerName")=ReferrerName

			Rs("ModerationLevel")=SiteConfig("NewUserModerationLevel")
			if SiteConfig("AccountActivation")=3 then Rs("UserAccountStatus")=0

			Rs("UserRegisterTime")=""&now()&""
			Rs("UserRegisterIP")=REMOTE_ADDR
			Rs("UserActivityTime")=""&now()&""
			Rs("UserActivityIP")=REMOTE_ADDR
		Rs.update
		UserID=Rs("UserID")
		UserPassword=Rs("UserPassword")
	Rs.close

	UpdateStatistics 1,0,0
	NowDate=date()
	Execute("update ["&TablePrefix&"Statistics] Set NewestUserName='"&UserName&"' where DateDiff("&SqlChar&"d"&SqlChar&",DateCreated,"&SqlNowString&")=0")
	AddUser=True
End Function

Function ModifyUserPassword(UserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer)
	ModifyUserPassword=False
	
	if ""&UserEmail&""<>"" then
		if instr(UserEmail,"@")=0 then Message=Message&"<li>您的电子邮箱地址填写错误</li>" : Exit Function
		if CookieUserEmail<>empty and UserEmail<>CookieUserEmail then 
			If not Execute("Select UserID From ["&TablePrefix&"Users] where UserEmail='"&UserEmail&"'" ).eof Then Message=Message&"<li>"&UserEmail&" 已经被别人注册了" : Exit Function
		end if
	end if
	
	sql="Select * from ["&TablePrefix&"Users] where UserName='"&UserName&"'"
	Rs.Open sql,Conn,1,3
		if Rs.eof then Message=Message&"<li>该用户不存在</li>" : Exit Function
		if UserPassword<>empty then Rs("UserPassword")=DefaultPasswordFormat(""&UserPassword&"")
		if ""&UserEmail&""<>"" then Rs("UserEmail")=UserEmail
		if PasswordAnswer<>empty then
			Rs("PasswordQuestion")=PasswordQuestion
			Rs("PasswordAnswer")=MD5(""&PasswordAnswer&"")
		end if
	Rs.update
	Rs.close
	ModifyUserPassword=True
End Function

Function UserLogin
	UserLogin=False
	if instr(UserName,"@")>0 and instr(UserName,".")>0 then
		SqlItem="UserEmail='"&UserName&"'"
	else
		SqlItem="UserName='"&UserName&"'"
	end if
	Rs.open "select * from ["&TablePrefix&"Users] where "&SqlItem,Conn,1,3
	if not Rs.eof and not Rs.bof Then
		if Rs("UserAccountStatus")=0 then Message=Message & "您的帐号正在等待审核"
		if Rs("UserAccountStatus")=2 then Message=Message & "您的帐号已被禁用！"
		if Rs("UserAccountStatus")=3 then Message=Message & "您的帐号尚未通过审核"
		
		if SiteConfig("AllowLogin")=0 and Rs("UserRoleID")<>1 then Message=Message & "除管理员外任何人都不允许登录！"
		if IsMD5=False then
			if Len(Rs("UserPassword"))=16 then
				if LCASE(mid(md5(UserPassword),9,16))<>LCASE(Rs("UserPassword")) then Message=Message & "您输入的密码错误"
			elseif Len(Rs("UserPassword"))=32 then
				if MD5(""&UserPassword&"")<>Rs("UserPassword") then Message=Message & "您输入的密码错误"
			elseif Len(Rs("UserPassword"))=40 then
				if SHA1(""&UserPassword&"")<>Rs("UserPassword") then Message=Message & "您输入的密码错误"
			else
				if UserPassword<>Rs("UserPassword") then Message=Message & "您输入的密码错误"
			end if
		ElseIf IsMD5=True then
			if Len(Rs("UserPassword"))=16 then
				if LCASE(UserPassword)<>LCASE(Rs("UserPassword")) then Message=Message & "您输入的密码错误"
			elseif Len(Rs("UserPassword"))=32 then
				if LCase(UserPassword)<>LCase(mid(Rs("UserPassword"),9,16)) then Message=Message & "您输入的密码错误"
			else
				if UserPassword<>Rs("UserPassword") then Message=Message & "您输入的密码错误"
			end if
		End If
		If Message<>"" Then Exit Function
		
		if Request("Invisible")="1" then
			Invisible="1"
		else
			Invisible="0"
		end if
		
		ResponseCookies "UserID",Rs("UserID"),Expired
		ResponseCookies "UserPassword",Rs("UserPassword"),Expired
		ResponseCookies "Invisible",Invisible,Expired
		UserLogin=True
	else
		Message="用户资料不存在！"
	end if
	Rs.close
End Function

Function UserLoginOut
	'Execute("Delete from ["&TablePrefix&"UserOnline] where sessionid='"&session.sessionid&"'")
	CleanCookies()
End Function


'附件水印
Function JpegPersits
	if SiteConfig("WatermarkType")=0 then
		Jpeg.Canvas.Font.Color = Replace(SiteConfig("WatermarkFontColor"),"#","&h")	'颜色
		Jpeg.Canvas.Font.Family = SiteConfig("WatermarkFontFamily")  		'字体
		Jpeg.Canvas.Font.size = SiteConfig("WatermarkFontSize")			'大小
		Jpeg.Canvas.Font.Bold = CBool(SiteConfig("WatermarkFontIsBold"))	'是否加粗
		'Jpeg.Canvas.Font.ShadowXoffset = 10		'水印文字阴影向右偏移的像素值，输入负值则向左偏移
		'Jpeg.Canvas.Font.ShadowYoffset = 10		'水印文字阴影向下偏移的像素值，输入负值则向右偏移
		Title = SiteConfig("WatermarkText") 
		TitleWidth = Jpeg.Canvas.GetTextExtent(Title)
		if Jpeg.Width<TitleWidth then exit function	'图片比水印文字小，则不加水印
		select case SiteConfig("WatermarkWidthPosition")
		case "left"
			PositionWidth=10
		case "center"
			PositionWidth=(Jpeg.Width - TitleWidth) / 2
		case "right"
			PositionWidth= Jpeg.Width - TitleWidth - 10
		end select
		Jpeg.Canvas.Print PositionWidth, 10, Title
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	elseif SiteConfig("WatermarkType")=1 then
		Set Jpeg2 = Server.CreateObject("Persits.Jpeg")
		Jpeg2.Open Server.MapPath(SiteConfig("WatermarkImage"))
		Jpeg2Width=Jpeg2.OriginalWidth
		Jpeg2Height=Jpeg2.OriginalHeight
		if Jpeg.Width<Jpeg2Width or Jpeg.Height<Jpeg2Height*2 then exit function	'图片比水印图片小，则不加水印
		select case SiteConfig("WatermarkWidthPosition")
		case "left"
			PositionWidth=10
		case "center"
			PositionWidth=(Jpeg.Width - Jpeg2Width) / 2
		case "right"
			PositionWidth= Jpeg.Width - Jpeg2Width - 10
		end select
		select case SiteConfig("WatermarkHeightPosition")
		case "top"
			PositionHeight=10
		case "center"
			PositionHeight=(Jpeg.Height - Jpeg2Height) / 2
		case "bottom"
			PositionHeight= Jpeg.Height - Jpeg2Height - 10
		end select
		Jpeg.Canvas.DrawImage PositionWidth, PositionHeight, Jpeg2, 1, &HFFFFFF	'透明度, 透明颜色
	end if
End Function
%>
