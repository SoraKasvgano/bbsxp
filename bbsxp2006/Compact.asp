<%
option explicit
Const JET_3X = 4
if ""&Request.Form("sessionid")&""<>""&session.sessionid&"" then error2("Ч�������")

Dim dbpath,boolIs97
dbpath = Request("dbpath")
boolIs97 = Request("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
End If

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
If Err Then error2("����ʶ������ݿ��ʽ")

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing
CompactDB = "<script>alert('ѹ���ɹ���');history.back();</script>"
Else
CompactDB = "<script>alert('�Ҳ������ݿ⣡\n�������ݿ�·���Ƿ��������');history.back();</script>"
End If
End Function

sub error2(Message)
%><script>alert('<%=Message%>');history.back();</script><script>window.close();</script><%
response.end
end sub
%>