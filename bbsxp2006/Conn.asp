<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer=True
IsSqlDataBase=0     '�������ݿ����0ΪAccess���ݿ⣬1ΪSQL���ݿ�
If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access���ݿ� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Db = "database/bbsxp6.mdb"       '���ݿ�·��
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(db)
SqlNowString="Now()"
SqlChar="'"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL���ݿ� ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SqlLocalName   ="(local)"     '����IP  [ ������ (local) �����IP ]
SqlUserName    ="sa"          '�û���
SqlPassword    ="1"           '�û�����
SqlDatabaseName="bbsxp"       '���ݿ���
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
Response.Write "���ݿ����ӳ������������ִ���"
Response.End
End If
On Error GoTo 0

''''''''''�滻ģ��START''''''''''''
Function ReplaceText(fString,patrn,replStr)
Set regEx = New RegExp  ' ����������ʽ��
regEx.Pattern = patrn   ' ����ģʽ��
regEx.IgnoreCase = True ' �����Ƿ����ִ�Сд��
regEx.Global = True     ' ����ȫ�ֿ����ԡ� 
ReplaceText = regEx.Replace(""&fString&"",""&replStr&"") ' ���滻��
Set reg=nothing
End Function
''''''''''�滻ģ��END''''''''''''
sub CloseDatabase
Conn.Close
set Rs=nothing
set Rs1=nothing
set Conn=nothing
set SiteSettings=nothing
Response.End
end sub
%>