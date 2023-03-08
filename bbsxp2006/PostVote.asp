<!-- #include file="Setup.asp" -->
<%

id=int(Request("id"))


if CookieUserName=empty then error2("您必须登录后才能投票")
if Request("PostVote")="" then error2("请选择，您要投票的项目！")


sql="select * from [BBSXP_Vote] where ThreadID="&id&""
Rs.Open sql,Conn,1,3
if instr(Rs("BallotUserList"),""&CookieUserName&"|")>0 then error2("您已经投过票了，无需重复投票！")
if Rs("Expiry")< now() then error2("投票已过期")

for each ho in request.form("PostVote")
pollresult=split(Rs("Result"),"|")
for i = 0 to ubound(pollresult)
if not pollresult(i)="" then
if cint(ho)=i then
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
Rs.update
Rs.close
Conn.execute("update [BBSXP_Threads] set lasttime="&SqlNowString&",lastname='"&CookieUserName&"' where id="&id&"")

error2("投票成功!")


%>