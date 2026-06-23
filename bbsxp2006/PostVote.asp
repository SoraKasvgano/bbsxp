<!-- #include file="Setup.asp" -->
<%

id=RequestInt("id")


if CookieUserName=empty then error2("�������¼�����ͶƱ")
if Request("PostVote")="" then error2("��ѡ����ҪͶƱ����Ŀ��")


sql="select * from [BBSXP_Vote] where ThreadID="&id&""
Rs.Open sql,Conn,1,3
if instr(Rs("BallotUserList"),""&CookieUserName&"|")>0 then error2("���Ѿ�Ͷ��Ʊ�ˣ������ظ�ͶƱ��")
if Rs("Expiry")< now() then error2("ͶƱ�ѹ���")

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
Conn.execute("update [BBSXP_Threads] set lasttime="&SqlNowString&",lastname='"&SqlString(CookieUserName)&"' where id="&id&"")

error2("ͶƱ�ɹ�!")


%>