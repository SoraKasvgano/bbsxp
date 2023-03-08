<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer=True
IsSqlDataBase=0     '定义数据库类别，0为Access数据库，1为SQL数据库
If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access数据库 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Db = "database/bbsxp6.mdb"       '数据库路径
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(db)
SqlNowString="Now()"
SqlChar="'"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL数据库 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SqlLocalName   ="(local)"     '连接IP  [ 本地用 (local) 外地用IP ]
SqlUserName    ="sa"          '用户名
SqlPassword    ="1"           '用户密码
SqlDatabaseName="bbsxp"       '数据库名
ConnStr = "Provider=Sqloledb;User ID=" & SqlUserName & "; Password=" & SqlPassword & "; Initial CataLog = " & SqlDatabaseName & "; Data Source=" & SqlLocalName & ";"
SqlNowString="GetDate()"
IsSqlVer="SQL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
END IF
ForumsVersion="6.00 "&IsSqlVer&""

On Error Resume Next
Set Conn=Server.CreateObject("ADODB.Connection")
Conn.open ConnStr
If Err Then
err.Clear
Set Conn = Nothing
Response.Write "数据库连接出错，请检查连接字串。"
Response.End
End If
On Error GoTo 0

''''''''''替换模块START''''''''''''
Function ReplaceText(fString,patrn,replStr)
Set regEx = New RegExp  ' 建立正则表达式。
regEx.Pattern = patrn   ' 设置模式。
regEx.IgnoreCase = True ' 设置是否区分大小写。
regEx.Global = True     ' 设置全局可用性。 
ReplaceText = regEx.Replace(""&fString&"",""&replStr&"") ' 作替换。
Set reg=nothing
End Function
''''''''''替换模块END''''''''''''
sub CloseDatabase
Conn.Close
set Rs=nothing
set Rs1=nothing
set Conn=nothing
set SiteSettings=nothing
Response.End
end sub
%>