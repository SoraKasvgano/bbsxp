<!-- #include file="Setup.asp" -->
<%
ThreadID=RequestInt("ThreadID")
if CookieUserName=empty then error("您必须<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>后才能投票")
if Request("PostVote")="" then error("请选择，您要投票的项目！")
sql="Select * from ["&TablePrefix&"Votes] where ThreadID="&ThreadID&""
Rs.Open sql,Conn,1,3
	if instr("|"&Rs("BallotUserList")&"|","|"&CookieUserName&"|")>0 then error("您已经投过票了，无需重复投票！")
	if instr("|"&Rs("BallotIPList")&"|","|"&REMOTE_ADDR&"|")>0 then error("此IP已经投过票了，无需重复投票！")
	if Rs("Expiry")< now() then error("投票已过期")
	for each ho in Request.Form("PostVote")
		pollresult=split(Rs("Result"),"|")
		for i = 0 to ubound(pollresult)
			if not pollresult(i)="" then
				if int(ho)=i then
					pollresult(i)=pollresult(i)+1
				end if
				allpollresult=""&allpollresult&""&pollresult(i)&"|"
			end if
		next
		Rs("Result")=allpollresult
		Rs.update
		allpollresult=""
	next
	Rs("BallotUserList")=""&Rs("BallotUserList")&""&CookieUserName&"|"
	Rs("BallotIPList")=""&Rs("BallotIPList")&""&REMOTE_ADDR&"|"
Rs.update
Rs.close
Execute("update ["&TablePrefix&"Threads] Set lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where ThreadID="&ThreadID&"")
response.redirect Http_Referer
%>