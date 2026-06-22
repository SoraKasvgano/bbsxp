<!-- #include file="Conn.asp" -->
<%
Const JET_3X = 4
Dim SiteSettings
Set SiteSettings=Conn.Execute("[BBSXP_SiteSettings]")
if Request.Cookies("UserName")="" or SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
if ""&Request.Form("sessionid")&""<>""&session.sessionid&"" then error2("效验码错误")

Dim dbpath,boolIs97,RawDbPath
RawDbPath = Trim(""&Request.Form("dbpath")&"")
If RawDbPath="" Or RawDbPath="__current__" Then
dbpath = SafeCompactDbPath(DB)
Else
dbpath = SafeCompactDbPath(RawDbPath)
End If
boolIs97 = Request.Form("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
Else
error2("数据库路径非法")
End If

Function SafeCompactDbPath(Value)
Dim TempPath, LowerPath
TempPath=Trim(""&Value&"")
TempPath=Replace(TempPath,"\","/")
TempPath=Replace(TempPath,Chr(0),"")
TempPath=Replace(TempPath,Chr(9),"")
TempPath=Replace(TempPath,Chr(10),"")
TempPath=Replace(TempPath,Chr(13),"")
LowerPath=LCase(TempPath)
If TempPath="" Or InStr(LowerPath,"..")>0 Or InStr(LowerPath,":")>0 Or Left(LowerPath,1)="/" Or Left(LowerPath,2)="//" Or Right(LowerPath,4)<>".mdb" Then
SafeCompactDbPath=""
Else
SafeCompactDbPath=TempPath
End If
End Function

Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = Left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

On Error Resume Next
If boolIs97 = "True" Then
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
& "Jet OLEDB:Engine Type=" & JET_3X
Else
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
End If
If Err.Number<>0 Then error2("不能识别指定的数据库格式")

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing
CompactDB = "<script>alert('压缩成功！');history.back();</script>"
Else
CompactDB = "<script>alert('找不到数据库！\n请检查数据库路径是否设置正确');history.back();</script>"
End If
End Function

sub error2(Message)
%><script>alert('<%=Message%>');history.back();</script><script>window.close();</script><%
response.end
end sub
%>