<%
InstallIPAddress="127.0.0.1"		'��װBBSXP��IP��ַ�����install.asp�ķ���Ȩ��

TablePrefix="BBSXP_"			'���ݿ���ǰ�����һ�㲻�ø��ģ�

IsSqlDataBase=0				'�������ݿ����0ΪAccess���ݿ⣬1ΪSQL���ݿ�

If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access���ݿ����� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	SqlDataBase	= "database/bbsxp2008.mdb"		'���ݿ�·��
	SqlProvider	= "Microsoft.Jet.OLEDB.4.0"		'��������[ Microsoft.Jet.OLEDB.4.0  Microsoft.ACE.OLEDB.12.0 ]
	SqlPassword	= ""							'ACCESS���ݿ�����
	Connstr="Provider="&SqlProvider&";Jet Oledb:Database Password="&SqlPassword&"; Data Source="&Server.MapPath(SqlDataBase)
	SqlNowString="Now()"
	SqlChar="'"
	IsSqlVer="ACCESS"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL���ݿ����� ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	SqlLocalName	= "(local)"		'����IP  [ ������ (local) �����IP ]
	SqlUserName	= "sa"			'SQL�û���
	SqlPassword	= "1234"			'SQL�û�����
	SqlDataBase	= "bbsxp"		'���ݿ���
	SqlProvider	= "SQLOLEDB"		'�������� [ SQLOLEDB  SQLNCLI ]
	ConnStr="Provider="&SqlProvider&"; User ID="&SqlUserName&"; Password="&SqlPassword&"; Initial CataLog="&SqlDataBase&"; Data Source="&SqlLocalName&";"
	SqlNowString="GetDate()"
	IsSqlVer="MSSQL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

'''''''''''''''''''''''''' ����Ϊרҵ��Ա����ѡ���ͨ�û������޸� ''''''''''''''''''''''''''
Session.CodePage="936"			'936(��������)		950(��������)	65001(Unicode)
BBSxpCharset="GB2312"			'GB2312(��������)	Big5(��������)	UTF-8(Unicode)
Response.Charset=BBSxpCharset
Response.Buffer=True
%>