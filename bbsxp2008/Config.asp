<%
InstallIPAddress="127.0.0.1"		'安装BBSXP的IP地址，针对install.asp的访问权限

TablePrefix="BBSXP_"			'数据库表的前辍名（一般不用更改）

IsSqlDataBase=0				'定义数据库类别，0为Access数据库，1为SQL数据库

If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access数据库设置 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	SqlDataBase	= "database/bbsxp2008.mdb"		'数据库路径
	SqlProvider	= "Microsoft.Jet.OLEDB.4.0"		'驱动程序[ Microsoft.Jet.OLEDB.4.0  Microsoft.ACE.OLEDB.12.0 ]
	SqlPassword	= ""							'ACCESS数据库密码
	Connstr="Provider="&SqlProvider&";Jet Oledb:Database Password="&SqlPassword&"; Data Source="&Server.MapPath(SqlDataBase)
	SqlNowString="Now()"
	SqlChar="'"
	IsSqlVer="ACCESS"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL数据库设置 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	SqlLocalName	= "(local)"		'连接IP  [ 本地用 (local) 外地用IP ]
	SqlUserName	= "sa"			'SQL用户名
	SqlPassword	= "1234"			'SQL用户密码
	SqlDataBase	= "bbsxp"		'数据库名
	SqlProvider	= "SQLOLEDB"		'驱动程序 [ SQLOLEDB  SQLNCLI ]
	ConnStr="Provider="&SqlProvider&"; User ID="&SqlUserName&"; Password="&SqlPassword&"; Initial CataLog="&SqlDataBase&"; Data Source="&SqlLocalName&";"
	SqlNowString="GetDate()"
	IsSqlVer="MSSQL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

'''''''''''''''''''''''''' 以下为专业人员设置选项，普通用户请勿修改 ''''''''''''''''''''''''''
Session.CodePage="936"			'936(简体中文)		950(繁体中文)	65001(Unicode)
BBSxpCharset="GB2312"			'GB2312(简体中文)	Big5(繁体中文)	UTF-8(Unicode)
Response.Charset=BBSxpCharset
Response.Buffer=True
%>