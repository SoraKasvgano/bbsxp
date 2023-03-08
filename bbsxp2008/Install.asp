<!-- #include file="BBSXP_Class.asp" -->
<!--#include file="Config.asp"-->
<%
DataBaseVer="8.0.5"

if Request.ServerVariables("REMOTE_ADDR")<>InstallIPAddress then
	response.Write("为了安全起见，请编辑 <font color=red>Config.asp</font> 内的安装IP地址<li>请把它设置成 <font color=red>"&Request.ServerVariables("REMOTE_ADDR")&"</font>")
	Response.end
end if


Set Rs = Server.CreateObject("ADODB.Recordset")
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=<%=BBSxpCharset%>" />
	<meta name="GENERATOR" content="BBSXP 2008 (<%=DataBaseVer%>)" />
	<link rel="stylesheet" type="text/css" href="Themes/Default/Common.css" />
	<title>BBSXP2008 安装程序</title>
</head>

<body>





<table cellspacing="0" cellpadding="5" width="100%">
<tr class="CommonListTitle">
	<td width="160" background="Themes/Default/header.gif"><img src="Themes/Default/Logo.gif"></td>
    <td background="Themes/Default/header.gif">BBSXP2008 (<%=DataBaseVer%>) 安装程序</br><%=IsSqlVer%>数据库</td>
  </tr>
</table>




<%

if Request("menu")="ServerConfig" then
	ServerConfig
elseif Request("menu")="CheckDatabase" then
	CheckDatabase
elseif Request("menu")="" then
	Default
else
	Set Conn=Server.CreateObject("ADODB.Connection")
	
	On Error Resume Next
	Conn.open ConnStr
	If Err Then ErrMsg
	On Error GoTo 0


		select case Request("menu")
			case "AutoCreateTable"
				FoundError=False
				AutoCreateTable
			case "SiteSetting"
				Set SiteConfigXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
				SiteConfigXMLDOM.loadxml("<bbsxp>"&Execute("Select SiteSettingsXML from ["&TablePrefix&"SiteSettings]")(0)&"</bbsxp>")
				SiteSetting
			case "SiteSettingUp"
				Set XMLDOM=Server.CreateObject("Microsoft.XMLDOM")
				Rs.Open "["&TablePrefix&"SiteSettings]",Conn,1,3
	
				XMLDOM.loadxml("<bbsxp>"&Rs("SiteSettingsXML")&"</bbsxp>")
					SiteSettingsXMLStrings=""
					Set objNodes = XMLDOM.documentElement.ChildNodes
					Set objRoot = XMLDOM.documentElement
					for each ho in Request.Form
						objRoot.SelectSingleNode(ho).text = ""&server.HTMLEncode(Request(""&ho&""))&""
					next
					for each element in objNodes	
						SiteSettingsXMLStrings=SiteSettingsXMLStrings&"<"&element.nodename&">"&element.text&"</"&element.nodename&">"&vbCrlf
					next
					Set XMLDOM=nothing
					Rs("SiteSettingsXML")=SiteSettingsXMLStrings
				Rs.update
				Rs.close
				
				Response.redirect("?menu=AddAdmin")
			case "AddAdmin"
				AddAdmin
			case "AddAdminUp"
				SiteSettingsXML=Execute("Select SiteSettingsXML from ["&TablePrefix&"SiteSettings]")(0)
				Set SiteConfigXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
				SiteConfigXMLDOM.loadxml("<bbsxp>"&SiteSettingsXML&"</bbsxp>")
				dim UserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer,Message
				
				AddAdminUp
			case "DeleteAdmin"
				SiteSettingsXML=Execute("Select SiteSettingsXML from ["&TablePrefix&"SiteSettings]")(0)
				Set SiteConfigXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
				SiteConfigXMLDOM.loadxml("<bbsxp>"&SiteSettingsXML&"</bbsxp>")
				DeleteAdmin
		end select
end if
InstallEndHtml

Sub InstallEndHtml
%>



<div class="CommonFooter">
	Powered by <a target="_blank" href="http://www.bbsxp.com">BBSXP 2008 <%=IsSqlVer%></a> &copy; 1998-<%=year(now())%><br />
	Server Time <%=now()%><br />
</div>




</body>
</html>
<%
End Sub




Sub ErrMsg
%>

<div style='text-align:left;padding:40px;width:auto; background:#FFFFFF'>


<div style="padding-bottom:30px; font-size:12pt; color:#FF0000;"><b>安装程序遇到错误</b></div>


<%=Err.Description%><br /></div>

<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="button" value="下一步"  disabled="disabled" /></div>
</div>


<%
InstallEndHtml
response.end
End Sub





Sub Default
%>
<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
<b>欢迎使用 BBSXP2008</b><br /><br />
您即将开始安装本程序。<br /><br />
<font color=red>如果您是首次使用，请点击页面右下角 <strong>[下一步]</strong> 按钮开始全新安装进程。</font><br /><br />
为了避免安装时浏览器崩溃，我们强烈建议您关闭浏览器上的第三方工具栏，例如 Google、MSN、Alexa 工具栏等。<br /><br /><br /><br />
如果您需要重新设置论坛配置，请点击 <input type="button" value="论坛配置" onClick="window.location.href = '?menu=SiteSetting';" /> ，直接设置论坛配置。<br /><br />
如果您忘记管理员密码，请点击 <input type="button" value="设置管理员" onClick="window.location.href = '?menu=AddAdmin';" /> ，直接设置管理员。

</div>

<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:right"><input type="button" value="下一步" onClick="window.location.href = '?menu=ServerConfig';" /></div>
</div>



<%
End Sub


Sub ServerConfig
%>

<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
<p style="font-size:10pt;"><b><u>第 1 步）检查组件支持情况</u></b></p>
	服务器组件支持情况(以下组件ADOX为必需，其余为非必需的)：<br /><br />

	　　ADOX.Catalog (ADOX 用于创建数据库) 组件支持：<%If Not IsObjInstalled("ADOX.Catalog") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />　　-----------------------------------------------<br />

	　　Scripting.FileSystemObject (FSO 文本文件读写) 组件支持：<%If Not IsObjInstalled("Scripting.FileSystemObject") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />

	　　CDO.Message (WINDOWS 邮件服务器) 组件支持：<%If Not IsObjInstalled("CDO.Message") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />　　-----------------------------------------------<br />

	　　Persits.Jpeg (AspJpeg 图像处理) 组件支持：<%If Not IsObjInstalled("Persits.Jpeg") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />

	　　Persits.Upload (AspUpload 文件上传) 组件支持：<%If Not IsObjInstalled("Persits.Upload") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />

	　　Persits.MailSender (AspEmail SMTP邮件发送) 组件支持：<%If Not IsObjInstalled("Persits.MailSender") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />　　-----------------------------------------------<br />

	　　SoftArtisans.FileUp (SA-FileUp 上传组件) 组件支持：<%If Not IsObjInstalled("SoftArtisans.FileUp") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />

	　　JMail.Message (JMail SMTP邮件发送) 组件支持：<%If Not IsObjInstalled("JMail.Message") Then%><font color="red"><b>×</b></font><%else%><b>√</b><%end if%><br />
</div>
    
    
<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="button" value="下一步" onClick="window.location.href = '?menu=CheckDatabase';" /></div>
</div>
    
<%
End Sub

Sub CheckDatabase
%>

<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
<p style="font-size:10pt;"><b><u>第 2 步）新建数据库</u></b></p>
<%
	On Error Resume Next
	
	If IsSqlDataBase=0 Then
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		SqlDataBase=split(SqlDataBase,"/")
		CurrentPath=""
		for i=0 to Ubound(SqlDataBase)-1
			if not fso.folderexists(Server.MapPath(CurrentPath&SqlDataBase(i))) then fso.CreateFolder(Server.MapPath(CurrentPath&SqlDataBase(i)))
			if Err then
				Response.Write(""&SqlDataBase(i)&" 文件夹没有被创建！<br /><br />出错原因："&Err.Description&"("&Err.Number&")<br /><br />")
				if Err.Number = 70 then Response.Write("建议解决方法：请给 "&Server.MapPath(".")&" 目录添加一个有写入权限的 Authenticated Users 用户。<br /><br /><br /><br />")
				err.Clear
				exit for
			end if
			CurrentPath=CurrentPath&SqlDataBase(i)&"/"
		next
		Set fso = nothing
	End If

	Set Ca = Server.CreateObject("ADOX.Catalog")
	call Ca.Create(""&Connstr&"")
	Set Ca = Nothing	
	If Err Then
		Response.Write(""&SqlDataBase&" 数据库创建失败 "&Err.Description&"("&Err.Number&")<br /><br />如果数据库已经存在，请直接点击下一步。")
		if Err.Number = -2147467259 then Response.Write("<br /><br />建议解决方法：请给 "&Request.ServerVariables("TEMP")&" 临时目录添加一个有写入权限的 Authenticated Users 用户。")
		if Err.Number = -2147217897 then Response.Write("<br /><br />请单击 下一步 继续")
	else
		Response.Write(""&SqlDataBase&" 数据库创建成功，请单击 下一步 创建数据表<br /><br />")
	end if
	err.Clear
%></div>


<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="button" value="下一步" onClick="window.location.href = '?menu=AutoCreateTable';" /></div>
</div>

<%
End Sub

Sub AddAdminUp
	UserName=HTMLEncode(Request("UserName"))
	UserPassword=HTMLEncode(Request("UserPassword"))
	UserEmail=HTMLEncode(Request("UserEmail"))
	PasswordQuestion=HTMLEncode(Request.Form("PasswordQuestion"))
	PasswordQuestion2=HTMLEncode(Request.Form("PasswordQuestion2"))
	PasswordAnswer=Request("PasswordAnswer")
	if ""&PasswordQuestion2&""<>"" then PasswordQuestion=PasswordQuestion2
	

	Message=""
	if Execute("select UserID from ["&TablePrefix&"Users] where UserName='"&UserName&"'").eof then
		AddUser
	else
		ModifyUserPassword UserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer
	end if
	if ""&Message&""<>"" then
		HeaderText="注册/更新管理员遇到如下错误"
		Messages=Message
		Messages=Messages&"<li>请返回重新设置管理员帐号</li>"
	else
		Execute("Update ["&TablePrefix&"Users] Set UserRoleID=1 where UserName='"&UserName&"'")
		HeaderText="第 6 步）安装完成"
		Messages="您已成功安装了 BBSXP2008（"&IsSqlVer&"）。<br /><br />"
		Messages=Messages&"<font size=+1><b>您运行此论坛前建议删除如下文件:</b><br>install.asp</font><br /><br />"
		Messages=Messages&"现在，您可以进入您的控制面板。<br>进入控制面板可以<a href=Admin_Default.asp target=_top>点击这里</a>。"
	end if
%>

<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
<p style="font-size:10pt;"><b><u><%=HeaderText%></u></b></p>
    <%=Messages%>
</div>
<%
End Sub


Sub DeleteAdmin
	AdminUserName=HTMLEncode(Request("UserName"))
	Execute("Update ["&TablePrefix&"Users] Set UserRoleID=3 where UserName='"&AdminUserName&"'")
%>

<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
	<li>删除成功</li>
    <li>请<a href=javascript:history.go(-1)>返回</a></li>
</div>
<%
end Sub

Sub AddAdmin
%>
<form method="POST" action="?menu=AddAdminUp" style="margin:0px" onSubmit="return chkInput(this);">
<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">

<p style="font-size:10pt;"><b><u>第 5 步）设置管理员</u></b></p>

<table cellspacing=1 cellpadding=5 width=90% class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan="2">添加(更新)管理员</td>
	</tr>
	<tr class="CommonListCell"><td align=right width="40%"><b>用户名：</b></td>
		<td><input name=UserName size="40" /></td>
	</tr>
    
    
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>密码：</b><br />密码必须至少包含 6 个字符</td>
	  <td align="Left" valign="middle" width="77%"><input type="password" name="UserPassword" size="40" maxlength="16" onkeyup=EvalPwd(this.value); onchange=EvalPwd(this.value); /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>密码强度：</b></td>
		<td align="Left" valign="middle" width="77%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">弱</td>
					<td id=iMedium align="center">中</td>
					<td id=iStrong align="center">强</td>
				</tr>
			</table>		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" valign="middle" width="23%"><b>重新键入密码：</b><br />请与您的密码保持一致</td>
		<td align="Left" valign="middle" width="77%"><input type="password" name="RetypePassword" size="40" /></td>
	</tr> 

	<tr class="CommonListCell"><td align=right width="40%"><b>Email：</b></td>
		<td><input name=UserEmail size="40" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" valign="middle"><b>密码提示问题：</b><br>如果您忘记了密码，系统会向您询问机密答案</td>
		<td>
		<select name="PasswordQuestion" onChange="document.getElementById('PasswordQuestion2').style.display = this.options[this.selectedIndex].value=='Define' ? '':'none';">
			<option value="" selected="selected">选择一个</option>
			<option value="母亲的出生地点">母亲的出生地点</option>
			<option value="儿童时期最好的朋友">儿童时期最好的朋友</option>
			<option value="第一个宠物的名字">第一个宠物的名字</option>
			<option value="最喜欢的老师">最喜欢的老师</option>
			<option value="最喜欢的历史人物">最喜欢的历史人物</option>
			<option value="祖父的职业">祖父的职业</option>
			<option value="Define">自定义</option>
		</select>　<span id=PasswordQuestion2 style="display:none"><input type=text name=PasswordQuestion2 size=20 /></span>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right"><b>机密答案：</b></td>
		<td><input type="text" name="PasswordAnswer" size="40" /></td>
	</tr>
	<tr class=CommonListTitle>
		<td align=center colspan="2">目前已存在的管理员</td>
	</tr>

  <tr class="CommonListCell">
    <td colspan="2">
	<%
    sql="Select * from ["&TablePrefix&"Users] where UserRoleID=1"
Rs.open sql,conn,1
do while not rs.eof
%>
    
    <a target="_blank" href="Profile.asp?UID=<%=Rs("UserID")%>"><%=Rs("UserName")%></a>(<a href="?menu=DeleteAdmin&UserName=<%=Rs("UserName")%>" onClick="return window.confirm('您确定要去掉该管理员的权限?');">删除</a>)
<%
	Rs.movenext
loop
Rs.close
%>    
    
    </td>
  </tr>

</table>
</div>


<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="submit" value=" 完成 " /></div>
</div>
</form>
<br />

<script type="text/javascript" src="Utility/global.js"></script>
<script type="text/javascript" src="Utility/pswdplc.js"></script>
<script language="javascript" type="text/javascript">
function chkInput(form){
	if (form.UserName.value==''){
		alert('请输入管理员名称');
		return false;
	}
	if (form.UserPassword.value=='' || form.UserPassword.value.length<6){
		alert('请输入管理员密码(至少6个字符)！');
		return false;
	}
	if (form.UserPassword.value != form.RetypePassword.value){
		alert('您两次输入的密码不一样！');
		return false;
	}
	if (form.UserEmail.value=='' || form.UserEmail.value.indexOf('@')<1){
		alert('请正确输入你的邮箱');
		return false;
	}
	if (form.PasswordQuestion.value=='' && form.PasswordQuestion2.value==''){
		alert('请选择密码提示问题');
		return false;
	}
	if (form.PasswordAnswer.value=='' && form.PasswordAnswer.value==''){
		alert('请输入密码提示问题的答案');
		return false;
	}
	return true;
}
</script>
<%


End Sub





Sub SiteSetting
Script_Name=LCase(Request.ServerVariables("script_name"))
Server_Name=Request.ServerVariables("server_name")
DomainPath=Left(Script_Name,inStrRev(Script_Name,"/"))
if SiteConfig("SiteUrl")="" then
	SiteUrl="http://"&Server_Name&DomainPath
	SiteUrl=left(SiteUrl,len(SiteUrl)-1)
else
	SiteUrl=SiteConfig("SiteUrl")
end if
%>
<form method="POST" action="?menu=SiteSettingUp" style="margin:0px" onSubmit="return chkInput(this);">
<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">

<p style="font-size:10pt;"><b><u>第 4 步）获取默认设置</u></b></p>

<table cellspacing=1 cellpadding=5 width=90% class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align=center colspan="2">常规设置</td>
	</tr>
	<tr class=CommonListCell>
		<td width="50%"><b>论坛名称</b><br>论坛名称，它将在论坛所有页面的浏览器窗口标题中显示。</td>
		<td><input size="30" name="SiteName" value="<%=SiteConfig("SiteName")%>" /></td>
	</tr>


	<tr class=CommonListCell>
		<td><b>论坛网址</b><br>您访问这个论坛的网址。注意: 不要以斜杠 (“/”) 结尾。</td>
		<td><input size="30" name="SiteUrl" value="<%=SiteUrl%>" /></td>
	</tr>
	<tr class=CommonListCell>
		<td><b>网站管理员电子邮件地址</b><br>网站管理员的电子邮件地址。</td>
		<td><input size="30" value="<%=SiteConfig("WebMasterEmail")%>" name="WebMasterEmail" /></td>
	</tr>

	<tr class=CommonListCell>
		<td><b>公司/组织名称</b><br>您的主页名称，显示在论坛每一页底部。</td>
		<td><input size="30" name="CompanyName" value="<%=SiteConfig("CompanyName")%>" /></td>
	</tr>
	<tr class=CommonListCell>
		<td><b>公司/组织网址</b><br>您主页的网址，显示在论坛每一页底部。</td>
		<td><input size="30" value="<%=SiteConfig("CompanyURL")%>" name="CompanyURL" /></td>
	</tr>
	

	<tr class=CommonListCell>
		<td><b>Cookies 保存路径</b><br>Cookie 保存的路径。如果您在同一个域名下运行了多个论坛，便需要将它设置为每个论坛所在的目录名。否则，填写 / 便可以了。<br><br><font color=red>输入错误的路径，会导致您无法登录论坛。</font></td>
		<td>
		<select name="CookiePath" size="1">
		<option value="/">/</option>
		<option value="<%=DomainPath%>" <%if SiteConfig("CookiePath")=DomainPath then%>selected<%end if%>><%=DomainPath%></option>
		</select></td>
	</tr>
	<tr class=CommonListCell>
		<td><b>Cookies 域名</b><br>Cookie 所影响的域名。修改此选项默认值最常见的原因是，您的论坛有两个不同的网址，例如 bbsxp.com 和 forums.bbsxp.com。要使用户在以两个不同的域名访问论坛时，都能保持登录状态，您需要将此选项设置为 .bbsxp.com (注意域名需要以点开头)。<br><br><font color=red>若您没有把握，最好在这里什么也不填写，因为错误的设置会导致您无法登录论坛。</font></td>
		<td><input size="50" value="<%=SiteConfig("CookieDomain")%>" name="CookieDomain" /></td>
	</tr>
</table>
</div>


<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="submit" value=" 下一步 " /></div>
</div>
</form>
<br />

<script type="text/javascript" src="Utility/global.js"></script>
<script language="javascript" type="text/javascript">
function chkInput(form){
	if (form.SiteName.value==''){
		alert('请输入站点名称');
		return false;
	}
	return true;
}
</script>
<%
	Set SiteConfigXMLDOM=nothing
End Sub




Function InstallExecute(SqlCommand)
	On Error Resume Next
	Execute(SqlCommand)
	If Err Then
		Response.write ""&Err.Description&""&CHR(10)&""
		FoundError=True
	else
		if instr(SqlCommand,"[FK_"&TablePrefix&"")>0 then
			Response.write "关系创建成功：[FK_"&ReplaceText(SqlCommand,"(.)*\[FK_([^[]*)\](.)*","$2")&"]"&CHR(10)&""
		elseif instr(SqlCommand,"insert into")>0 then
			Response.write "数据添加成功：["&ReplaceText(SqlCommand,"(.)*\[([^[]*)\](.|\n)*","$2")&"]"&CHR(10)&""
		else
			Response.write "表创建成功：["&ReplaceText(SqlCommand,"(.)*\[([^[]*)\](.|\n)*","$2")&"]"&CHR(10)&""
		end if
	end if
	On Error goto 0
End Function

Sub AutoCreateTable
%>
<div style="text-align:left;padding:40px;width:auto; background:#FFFFFF">
<p style="font-size:10pt;"><b><u>第 3 步）创建数据表、建立关系及添加数据</u></b></p>
    	<font color=red>正在创建表...</font><br /><br />
    	<textarea cols="100" rows="10">
<%

	Sql="CREATE TABLE ["&TablePrefix&"Advertisements] ("&_
	"AdvertisementID int IDENTITY (1, 1) NOT NULL ,"&_
	"Body ntext NULL ,"&_
	"ExpireDate datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"EventLog] ("&_
	"UserName nvarchar(50) NOT NULL ,"&_
	"ErrNumber int NULL ,"&_
	"EventDate datetime Default "&SqlNowString&" NOT NULL ,"&_
	"MessageXML ntext NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"FavoriteForums] ("&_
	"FavoriteID int IDENTITY (1, 1) NOT NULL ,"&_
	"OwnerUserName nvarchar(50) NOT NULL ,"&_
	"ForumID int Default 0 NOT NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"FavoritePosts] ("&_
	"FavoriteID int IDENTITY (1, 1) NOT NULL ,"&_
	"OwnerUserName nvarchar(50) NOT NULL ,"&_
	"PostID int Default 0 NOT NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"FavoriteUsers] ("&_
	"FavoriteID int IDENTITY (1, 1) NOT NULL ,"&_
	"OwnerUserName nvarchar(50) NOT NULL ,"&_
	"FriendUserName nvarchar(50) NOT NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"ForumPermissions] ("&_
	"ForumID int Default 0 NOT NULL ,"&_
	"RoleID int Default 0 NOT NULL ,"&_
	"PermissionView int Default 1 NOT NULL ,"&_
	"PermissionRead int Default 1 NOT NULL ,"&_
	"PermissionPost int Default 1 NOT NULL ,"&_
	"PermissionReply int Default 1 NOT NULL ,"&_
	"PermissionEdit int Default 1 NOT NULL ,"&_
	"PermissionDelete int Default 0 NOT NULL ,"&_
	"PermissionCreatePoll int Default 1 NOT NULL ,"&_
	"PermissionVote int Default 1 NOT NULL ,"&_
	"PermissionAttachment int Default 1 NOT NULL ,"&_
	"PermissionManage int Default 0 NOT NULL "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"Forums] ("&_
	"ForumID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"GroupID int Default 0 NOT NULL ,"&_
	"ParentID int Default 0 NOT NULL ,"&_
	"SortOrder int Default 1 NOT NULL ,"&_
	"ForumName nvarchar(255) ,"&_
	"Moderated nvarchar(255) ,"&_
	"TotalCategorys nvarchar(255) ,"&_
	"ForumDescription nvarchar(255) ,"&_
	"TodayPosts int Default 0 ,"&_
	"TotalThreads int Default 0 ,"&_
	"TotalPosts int Default 0 ,"&_
	"ForumUrl nvarchar(255) ,"&_
	"IsActive int Default 1 ,"&_
	"MostRecentThreadID int Default 0 ,"&_
	"MostRecentPostSubject nvarchar(255) ,"&_
	"MostRecentPostAuthor nvarchar(50) ,"&_
	"MostRecentPostDate datetime Default "&SqlNowString&" ,"&_
	"ForumRules ntext ,"&_
	"ModerateNewThread int Default 0 ,"&_
	"ModerateNewPost int Default 0 ,"&_
	"DateCreated datetime Default "&SqlNowString&""&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"Groups] ("&_
	"GroupID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"GroupName nvarchar(255) ,"&_
	"SortOrder int Default 1 NOT NULL ,"&_
	"GroupDescription ntext ,"&_
	"ForumColumns int Default 0 NOT NULL ,"&_
	"Moderated nvarchar(255) "&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"Links] ("&_
	"LinkID int IDENTITY (1, 1) NOT NULL ,"&_
	"Name nvarchar(255) ,"&_
	"URL nvarchar(255) ,"&_
	"Logo nvarchar(255) ,"&_
	"Intro nvarchar(255) ,"&_
	"SortOrder int Default 1 "&_
	")"
	InstallExecute(sql)
	

	
	
	
	Sql="CREATE TABLE ["&TablePrefix&"PostAttachments] ("&_
	"UpFileID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"UserName nvarchar(50) NOT NULL ,"&_
	"FileName nvarchar(255) NOT NULL ,"&_
	"FileData image NULL ,"&_
	"ContentType nvarchar(255) NOT NULL ,"&_
	"FilePath nvarchar(255) NULL ,"&_
	"ContentSize int Default 0 NOT NULL ,"&_
	"PostID int Default 0 ,"&_
	"Description nvarchar(255) NULL ,"&_
	"Created datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"PostEditNotes] ("&_
	"EditNoteID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"PostID int Default 0 NOT NULL ,"&_
	"EditNotes nvarchar(255) NULL"&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"PostTags] ("&_
	"TagID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"TagName nvarchar(255) NOT NULL ,"&_
	"IsEnabled int Default 1 NOT NULL ,"&_
	"TotalPosts int Default 0 NOT NULL ,"&_
	"MostRecentPostDate datetime Default "&SqlNowString&" NOT NULL , "&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL"&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"PostInTags] ("&_
	"TagID int Default 0 NOT NULL ,"&_
	"PostID int Default 0 NOT NULL"&_
	")"
	InstallExecute(sql)
	
	Sql="CREATE TABLE ["&TablePrefix&"Posts] ("&_
	"PostID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"ThreadID int Default 0 ,"&_
	"ParentID int Default 0 ,"&_
	"PostAuthor nvarchar(50) ,"&_
	"Subject nvarchar(255) ,"&_
	"Body ntext ,"&_
	"IPAddress nvarchar(50) ,"&_
	"Visible int Default 1 ,"&_
	"PostDate datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"PrivateMessages] ("&_
	"MessageID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"SenderUserName nvarchar(50) ,"&_
	"RecipientUserName nvarchar(50) ,"&_
	"Subject nvarchar(255) ,"&_
	"Body ntext ,"&_
	"IsRead int Default 0 ,"&_
	"IsRecipientDelete int Default 0 ,"&_
	"IsSenderDelete int Default 0 ,"&_
	"CreateTime datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Ranks] ("&_
	"RankID int IDENTITY (1, 1) NOT NULL ,"&_
	"RoleID int Default 0 ,"&_
	"RankName nvarchar(255) ,"&_
	"PostingCountMin int Default 0 "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Reputation] ("&_
	"ReputationID int IDENTITY (1, 1) NOT NULL ,"&_
	"Reputation int Default 0 NOT NULL ,"&_
	"CommentFor nvarchar(50) NOT NULL ,"&_
	"CommentBy nvarchar(50) NOT NULL ,"&_
	"Comment ntext ,"&_
	"IPAddress nvarchar(50) ,"&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Roles] ("&_
	"RoleID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"Name nvarchar(255) NOT NULL ,"&_
	"Description nvarchar(255) ,"&_
	"RoleMaxFileSize int Default 0 ,"&_
	"RoleMaxPostAttachmentsSize int Default 0 "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"SiteSettings] ("&_
	"SiteSettingsXML ntext ,"&_
	"CreateUserAgreement ntext ,"&_
	"GenericHeader ntext ,"&_
	"GenericTop ntext ,"&_
	"GenericBottom ntext ,"&_
	"BestOnline int Default 0 ,"&_
	"BestOnlineTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"AdminNotes ntext ,"&_
	"BBSxpVersion nvarchar(50) "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Statistics] ("&_
	"DaysUsers int Default 0 ,"&_
	"DaysPosts int Default 0 ,"&_
	"DaysTopics int Default 0 ,"&_
	"TotalUsers int Default 0 ,"&_
	"TotalPosts int Default 0 ,"&_
	"TotalTopics int Default 0 ,"&_
	"NewestUserName nvarchar(50) ,"&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL "&_	
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Subscriptions] ("&_
	"SubscriptionID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"ThreadID int Default 0 ,"&_
	"UserName nvarchar(50) NOT NULL ,"&_
	"Email nvarchar(255) NOT NULL ,"&_
	"SubscriptionDate datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"ThreadRating] ("&_
	"UserName nvarchar(50) NOT NULL ,"&_
	"ThreadID int Default 0 ,"&_
	"Rating int Default 0 ,"&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Threads] ("&_
	"ThreadID int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"&_
	"ForumID int Default 0 ,"&_
	"ThreadEmoticonID int Default 0 ,"&_
	"ThreadStyle nvarchar(255) ,"&_
	"Topic nvarchar(255) ,"&_
	"Category nvarchar(255) ,"&_
	"Description nvarchar(255) ,"&_
	"PostAuthor nvarchar(50) ,"&_
	"LastName nvarchar(50) ,"&_
	"PostTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"LastTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"LastViewedDate datetime Default "&SqlNowString&" NOT NULL ,"&_
	"IsGood int Default 0 ,"&_
	"ThreadTop int Default 0 ,"&_
	"StickyDate datetime Default "&SqlNowString&" NOT NULL ,"&_
	"IsLocked int Default 0 ,"&_
	"Visible int Default 1 ,"&_
	"HiddenCount int Default 0 ,"&_
	"DeletedCount int Default 0 ,"&_
	"IsVote int Default 0 ,"&_
	"TotalViews int Default 0 ,"&_
	"TotalReplies int Default 0 ,"&_
	"RatingSum int Default 0 ,"&_
	"TotalRatings int Default 0 ,"&_
	"ThreadStatus int Default 0 "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"UserActivation] ("&_
	"ActivationKey nvarchar(50) NOT NULL ,"&_
	"UserName nvarchar(50) ,"&_
	"Email nvarchar(255) ,"&_
	"DateCreated datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"UserOnline] ("&_
	"ForumID int Default 0 ,"&_
	"SessionID nvarchar(255) ,"&_
	"UserName nvarchar(50) ,"&_
	"IPAddress nvarchar(50) ,"&_
	"IsInvisible int Default 0 ,"&_
	"ForumName nvarchar(255) ,"&_
	"ThreadID int Default 0 ,"&_
	"Topic nvarchar(255) ,"&_
	"ComeTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"LastTime datetime Default "&SqlNowString&" NOT NULL "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Users] ("&_
	"UserID int IDENTITY (1, 1) NOT NULL ,"&_
	"UserName nvarchar(50) PRIMARY KEY NOT NULL ,"&_
	"UserPassword nvarchar(50) NOT NULL ,"&_
	"UserSex int Default 0 ,"&_
	"UserRoleID int Default 3 NOT NULL ,"&_
	"UserAccountStatus int Default 1 ,"&_
	"ModerationLevel int Default 0 ,"&_
	"UserEmail nvarchar(255) NOT NULL ,"&_
	"PasswordQuestion nvarchar(255) ,"&_
	"PasswordAnswer nvarchar(255) ,"&_
	"WebAddress nvarchar(255) ,"&_
	"WebLog nvarchar(255) ,"&_
	"WebGallery nvarchar(255) ,"&_
	"UserRank nvarchar(255) ,"&_
	"Birthday datetime ,"&_
	"NewMessage int Default 0 ,"&_
	"TotalPosts int Default 0 ,"&_
	"UserMoney int Default 0 ,"&_
	"Experience int Default 0 ,"&_
	"UserFaceUrl nvarchar(255) ,"&_
	"UserRegisterIP nvarchar(50) ,"&_
	"UserActivityIP nvarchar(50) ,"&_
	"UserRegisterTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"UserPostTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"UserActivityTime datetime Default "&SqlNowString&" NOT NULL ,"&_
	"ReferrerName nvarchar(50) ,"&_
	"RealName nvarchar(50) ,"&_
	"Occupation nvarchar(255) ,"&_
	"Address nvarchar(255) ,"&_
	"QQ nvarchar(255) ,"&_
	"ICQ nvarchar(255) ,"&_
	"AIM nvarchar(255) ,"&_
	"MSN nvarchar(255) ,"&_
	"Yahoo nvarchar(255) ,"&_
	"Skype nvarchar(255) ,"&_
	"Interests nvarchar(255) ,"&_
	"UserMate nvarchar(50) ,"&_
	"UserTitle nvarchar(255) ,"&_
	"Reputation int Default 0 ,"&_
	"UserActivityDay int Default 0 ,"&_
	"BankMoney int Default 0 ,"&_
	"BankDate datetime Default "&SqlNowString&" NOT NULL ,"&_
	"UserSign ntext ,"&_
	"UserBio ntext ,"&_
	"UserNote ntext "&_
	")"
	InstallExecute(sql)

	Sql="CREATE TABLE ["&TablePrefix&"Votes] ("&_
	"ThreadID int Default 0 NOT NULL ,"&_
	"IsMultiplePoll int Default 0 ,"&_
	"Expiry datetime Default "&SqlNowString&" NOT NULL ,"&_
	"Items ntext ,"&_
	"Result ntext ,"&_
	"BallotUserList ntext ,"&_
	"BallotIPList ntext "&_
	")"
	InstallExecute(sql)

	Response.write("</textarea><br /><br /><font color=red>正在添加建立关系...</font><br /><br />")
	Response.write("<textarea cols=100 rows=10>")
	 
	'创建关联
	Sql="ALTER TABLE ["&TablePrefix&"Forums] ADD CONSTRAINT [FK_"&TablePrefix&"Forums_"&TablePrefix&"Groups] FOREIGN KEY ([GroupID]) REFERENCES ["&TablePrefix&"Groups] ([GroupID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	Sql="ALTER TABLE ["&TablePrefix&"ForumPermissions] ADD CONSTRAINT [FK_"&TablePrefix&"ForumPermissions_"&TablePrefix&"Forums] FOREIGN KEY ([ForumID]) REFERENCES ["&TablePrefix&"Forums] ([ForumID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	Sql="ALTER TABLE ["&TablePrefix&"Threads] ADD CONSTRAINT [FK_"&TablePrefix&"Threads_"&TablePrefix&"Forums] FOREIGN KEY ([ForumID]) REFERENCES ["&TablePrefix&"Forums] ([ForumID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	

	

	
	Sql="ALTER TABLE ["&TablePrefix&"Posts] ADD CONSTRAINT [FK_"&TablePrefix&"Posts_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"Votes] ADD CONSTRAINT [FK_"&TablePrefix&"Votes_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"Subscriptions] ADD CONSTRAINT [FK_"&TablePrefix&"Subscriptions_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)



	Sql="ALTER TABLE ["&TablePrefix&"ThreadRating] ADD CONSTRAINT [FK_"&TablePrefix&"ThreadRating_"&TablePrefix&"Threads] FOREIGN KEY ([ThreadID]) REFERENCES ["&TablePrefix&"Threads] ([ThreadID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	

	Sql="ALTER TABLE ["&TablePrefix&"PostEditNotes] ADD CONSTRAINT [FK_"&TablePrefix&"PostEditNotes_"&TablePrefix&"Posts] FOREIGN KEY ([PostID]) REFERENCES ["&TablePrefix&"Posts] ([PostID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	Sql="ALTER TABLE ["&TablePrefix&"PostInTags] ADD CONSTRAINT [FK_"&TablePrefix&"PostInTags_"&TablePrefix&"Posts] FOREIGN KEY ([PostID]) REFERENCES ["&TablePrefix&"Posts] ([PostID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	Sql="ALTER TABLE ["&TablePrefix&"PostInTags] ADD CONSTRAINT [FK_"&TablePrefix&"PostInTags_"&TablePrefix&"PostTags] FOREIGN KEY ([TagID]) REFERENCES ["&TablePrefix&"PostTags] ([TagID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)


	
	Sql="ALTER TABLE ["&TablePrefix&"ForumPermissions] ADD CONSTRAINT [FK_"&TablePrefix&"ForumPermissions_"&TablePrefix&"Roles] FOREIGN KEY ([RoleID]) REFERENCES ["&TablePrefix&"Roles] ([RoleID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	
	Sql="ALTER TABLE ["&TablePrefix&"Ranks] ADD CONSTRAINT [FK_"&TablePrefix&"Ranks_"&TablePrefix&"Roles] FOREIGN KEY ([RoleID]) REFERENCES ["&TablePrefix&"Roles] ([RoleID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"Users] ADD CONSTRAINT [FK_"&TablePrefix&"Users_"&TablePrefix&"Roles] FOREIGN KEY ([UserRoleID]) REFERENCES ["&TablePrefix&"Roles] ([RoleID]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	


	
	Sql="ALTER TABLE ["&TablePrefix&"Threads] ADD CONSTRAINT [FK_"&TablePrefix&"Threads_"&TablePrefix&"Users] FOREIGN KEY ([PostAuthor]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
	
	
		
	Sql="ALTER TABLE ["&TablePrefix&"PostAttachments] ADD CONSTRAINT [FK_"&TablePrefix&"PostAttachments_"&TablePrefix&"Users] FOREIGN KEY ([UserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"PrivateMessages] ADD CONSTRAINT [FK_"&TablePrefix&"PrivateMessages_"&TablePrefix&"Users] FOREIGN KEY ([SenderUserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)	
		
	Sql="ALTER TABLE ["&TablePrefix&"FavoriteForums] ADD CONSTRAINT [FK_"&TablePrefix&"FavoriteForums_"&TablePrefix&"Users] FOREIGN KEY ([OwnerUserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"FavoritePosts] ADD CONSTRAINT [FK_"&TablePrefix&"FavoritePosts_"&TablePrefix&"Users] FOREIGN KEY ([OwnerUserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
			
	Sql="ALTER TABLE ["&TablePrefix&"FavoriteUsers] ADD CONSTRAINT [FK_"&TablePrefix&"FavoriteUsers_"&TablePrefix&"Users] FOREIGN KEY ([OwnerUserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)

	Sql="ALTER TABLE ["&TablePrefix&"Reputation] ADD CONSTRAINT [FK_"&TablePrefix&"Reputation_"&TablePrefix&"Users] FOREIGN KEY ([CommentFor]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	InstallExecute(sql)
			
	'SQL中以下2句不能创建，冲突
	'Sql="ALTER TABLE ["&TablePrefix&"Subscriptions] ADD CONSTRAINT [FK_"&TablePrefix&"Subscriptions_"&TablePrefix&"Users] FOREIGN KEY ([UserName]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	'InstallExecute(sql)	
	'Sql="ALTER TABLE ["&TablePrefix&"Posts] ADD CONSTRAINT [FK_"&TablePrefix&"Posts_"&TablePrefix&"Users] FOREIGN KEY ([PostAuthor]) REFERENCES ["&TablePrefix&"Users] ([UserName]) ON DELETE CASCADE ON UPDATE CASCADE"
	'InstallExecute(sql)	
			
	Response.write("</textarea><br /><br /><font color=red>正在添加数据...</font><br /><br />")
	Response.write("<textarea cols=100 rows=10>")
	if FoundError=True then
		Response.write("由于有些 表/关系 创建失败，所以不能添加数据。")
	else
	InstallExecute("insert into ["&TablePrefix&"Groups] (GroupName) values ('主分类')")

	InstallExecute("insert into ["&TablePrefix&"Forums] (GroupID,ForumName) values ('1','主版块')")
	

	DomainPath=Left(Request.ServerVariables("script_name"),inStrRev(Request.ServerVariables("script_name"),"/")-1)
	SiteURL="http://"&Request.ServerVariables("server_name")&DomainPath
	If IsObjInstalled("SoftArtisans.FileUp") Then
		UpFileOption="SoftArtisans.FileUp"
	ElseIf IsObjInstalled("Scripting.FileSystemObject") Then
		UpFileOption="ADODB.Stream"
	Else
		UpFileOption=""
	End If
	
	SiteSettingsXML="<SiteName>BBSXP Forum Server</SiteName>"&vbCrlf&_
"<MetaDescription></MetaDescription>"&vbCrlf&_
"<MetaKeywords></MetaKeywords>"&vbCrlf&_
"<UserOnlineTime>20</UserOnlineTime>"&vbCrlf&_
"<Timeout>60</Timeout>"&vbCrlf&_
"<DefaultSiteStyle>default</DefaultSiteStyle>"&vbCrlf&_
"<CompanyName>Yuzi Corporation</CompanyName>"&vbCrlf&_
"<CompanyURL>http://www.yuzi.net</CompanyURL>"&vbCrlf&_
"<SiteDisabled>0</SiteDisabled>"&vbCrlf&_
"<SiteDisabledReason>论坛维护中，暂时无法访问！</SiteDisabledReason>"&vbCrlf&_
"<CacheName>BBSXP</CacheName>"&vbCrlf&_
"<CacheUpDateInterval>5</CacheUpDateInterval>"&vbCrlf&_
"<DefaultPasswordFormat>MD5</DefaultPasswordFormat>"&vbCrlf&_
"<DisplayWhoIsOnline>1</DisplayWhoIsOnline>"&vbCrlf&_
"<DisplayStatistics>1</DisplayStatistics>"&vbCrlf&_
"<DisplayLink>1</DisplayLink>"&vbCrlf&_
"<DisplayTags>1</DisplayTags>"&vbCrlf&_
"<DisplayBirthdays>1</DisplayBirthdays>"&vbCrlf&_
"<MinVoteOptions>2</MinVoteOptions>"&vbCrlf&_
"<MaxVoteOptions>20</MaxVoteOptions>"&vbCrlf&_
"<DisplayForumUsers>1</DisplayForumUsers>"&vbCrlf&_
"<DisplayThreadUsers>0</DisplayThreadUsers>"&vbCrlf&_
"<UpFileOption>"&UpFileOption&"</UpFileOption>"&vbCrlf&_
"<AttachmentsSaveOption>1</AttachmentsSaveOption>"&vbCrlf&_
"<UpFileTypes>gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar</UpFileTypes>"&vbCrlf&_
"<MaxPostAttachmentsSize>10240</MaxPostAttachmentsSize>"&vbCrlf&_
"<MaxFileSize>200</MaxFileSize>"&vbCrlf&_
"<MaxFaceSize>100</MaxFaceSize>"&vbCrlf&_
"<WatermarkOption></WatermarkOption>"&vbCrlf&_
"<WatermarkType>0</WatermarkType>"&vbCrlf&_
"<WatermarkText>BBSXP</WatermarkText>"&vbCrlf&_
"<WatermarkImage>Images/Watermark.gif</WatermarkImage>"&vbCrlf&_
"<WatermarkPosition>4</WatermarkPosition>"&vbCrlf&_
"<SelectMailMode></SelectMailMode>"&vbCrlf&_
"<SmtpServerMail></SmtpServerMail>"&vbCrlf&_
"<SmtpServer></SmtpServer>"&vbCrlf&_
"<SmtpServerUserName></SmtpServerUserName>"&vbCrlf&_
"<SmtpServerPassword></SmtpServerPassword>"&vbCrlf&_
"<BannedRegUserName>fuck|shit</BannedRegUserName>"&vbCrlf&_
"<BannedText>fuck|shit</BannedText>"&vbCrlf&_
"<BannedIP></BannedIP>"&vbCrlf&_
"<AccountActivation>0</AccountActivation>"&vbCrlf&_
"<AllowNewUserRegistration>1</AllowNewUserRegistration>"&vbCrlf&_
"<RequireAuthenticationForProfileViewing>0</RequireAuthenticationForProfileViewing>"&vbCrlf&_
"<NewUserModerationLevel>0</NewUserModerationLevel>"&vbCrlf&_
"<UserNameMinLength>3</UserNameMinLength>"&vbCrlf&_
"<UserNameMaxLength>15</UserNameMaxLength>"&vbCrlf&_
"<EnableReputation>1</EnableReputation>"&vbCrlf&_
"<ReputationDefault>0</ReputationDefault>"&vbCrlf&_
"<ShowUserRates>5</ShowUserRates>"&vbCrlf&_
"<MinReputationPost>50</MinReputationPost>"&vbCrlf&_
"<MinReputationCount>0</MinReputationCount>"&vbCrlf&_
"<MaxReputationPerDay>10</MaxReputationPerDay>"&vbCrlf&_
"<ReputationRepeat>20</ReputationRepeat>"&vbCrlf&_
"<AdminReputationPower>10</AdminReputationPower>"&vbCrlf&_
"<InPrisonReputation>-9</InPrisonReputation>"&vbCrlf&_
"<CustomUserTitle>0</CustomUserTitle>"&vbCrlf&_
"<UserTitleMaxChars>25</UserTitleMaxChars>"&vbCrlf&_
"<UserTitleCensorWords>admin|forum|moderator</UserTitleCensorWords>"&vbCrlf&_
"<EnableBannedUsersToLogin>0</EnableBannedUsersToLogin>"&vbCrlf&_
"<EnableAntiSpamTextGenerateForRegister>0</EnableAntiSpamTextGenerateForRegister>"&vbCrlf&_
"<EnableAntiSpamTextGenerateForLogin>0</EnableAntiSpamTextGenerateForLogin>"&vbCrlf&_
"<EnableAntiSpamTextGenerateForPost>0</EnableAntiSpamTextGenerateForPost>"&vbCrlf&_
"<PostInterval>10</PostInterval>"&vbCrlf&_
"<RegUserTimePost>0</RegUserTimePost>"&vbCrlf&_
"<PostEditBodyAgeInMinutes>0</PostEditBodyAgeInMinutes>"&vbCrlf&_
"<PopularPostThresholdPosts>15</PopularPostThresholdPosts>"&vbCrlf&_
"<PopularPostThresholdViews>200</PopularPostThresholdViews>"&vbCrlf&_
"<DisplayEditNotes>0</DisplayEditNotes>"&vbCrlf&_
"<DisplayThreadStatus>1</DisplayThreadStatus>"&vbCrlf&_
"<DisplayPostTags>1</DisplayPostTags>"&vbCrlf&_
"<ThreadsPerPage>20</ThreadsPerPage>"&vbCrlf&_
"<PostsPerPage>15</PostsPerPage>"&vbCrlf&_
"<RSSDefaultThreadsPerFeed>20</RSSDefaultThreadsPerFeed>"&vbCrlf&_
"<EnableForumsRSS>1</EnableForumsRSS>"&vbCrlf&_
"<AllowDuplicatePosts>0</AllowDuplicatePosts>"&vbCrlf&_
"<PopularPostThresholdDays>3</PopularPostThresholdDays>"&vbCrlf&_
"<AllowLogin>1</AllowLogin>"&vbCrlf&_
"<EnableSignatures>1</EnableSignatures>"&vbCrlf&_
"<SignatureMaxLength>256</SignatureMaxLength>"&vbCrlf&_
"<AllowSignatures>1</AllowSignatures>"&vbCrlf&_
"<AllowGender>1</AllowGender>"&vbCrlf&_
"<AllowAvatars>1</AllowAvatars>"&vbCrlf&_
"<PublicMemberList>1</PublicMemberList>"&vbCrlf&_
"<MemberListAdvancedSearch>1</MemberListAdvancedSearch>"&vbCrlf&_
"<MemberListPageSize>20</MemberListPageSize>"&vbCrlf&_
"<EnableAvatars>1</EnableAvatars>"&vbCrlf&_
"<EnableRemoteAvatars>1</EnableRemoteAvatars>"&vbCrlf&_
"<AvatarHeight>150</AvatarHeight>"&vbCrlf&_
"<AvatarWidth>150</AvatarWidth>"&vbCrlf&_
"<RequireEditNotes>0</RequireEditNotes>"&vbCrlf&_
"<EnablePostPreviewPopup>1</EnablePostPreviewPopup>"&vbCrlf&_
"<IntegralAddThread>+2</IntegralAddThread>"&vbCrlf&_
"<IntegralAddPost>+1</IntegralAddPost>"&vbCrlf&_
"<IntegralAddValuedPost>+3</IntegralAddValuedPost>"&vbCrlf&_
"<IntegralDeleteThread>-2</IntegralDeleteThread>"&vbCrlf&_
"<IntegralDeletePost>-1</IntegralDeletePost>"&vbCrlf&_
"<IntegralDeleteValuedPost>-3</IntegralDeleteValuedPost>"&vbCrlf&_
"<IsShowSonForum>0</IsShowSonForum>"&vbCrlf&_
"<ViewMode>1</ViewMode>"&vbCrlf&_
"<AllowUserToSelectTheme>1</AllowUserToSelectTheme>"&vbCrlf&_
"<MaxPrivateMessageSize>100</MaxPrivateMessageSize>"&vbCrlf&_
"<WatermarkFontFamily>宋体</WatermarkFontFamily>"&vbCrlf&_
"<WatermarkFontSize>25</WatermarkFontSize>"&vbCrlf&_
"<WatermarkFontColor>#000000</WatermarkFontColor>"&vbCrlf&_
"<WatermarkFontIsBold>1</WatermarkFontIsBold>"&vbCrlf&_
"<WatermarkWidthPosition>right</WatermarkWidthPosition>"&vbCrlf&_
"<WatermarkHeightPosition>bottom</WatermarkHeightPosition>"&vbCrlf&_
"<CookiePath>/</CookiePath>"&vbCrlf&_
"<SiteUrl>"&SiteURL&"</SiteUrl>"&vbCrlf&_
"<APIEnable>0</APIEnable>"&vbCrlf&_
"<APISafeKey></APISafeKey>"&vbCrlf&_
"<APIUrlList></APIUrlList>"&vbCrlf&_
"<WebMasterEmail>webmaster@localhost</WebMasterEmail>"&vbCrlf&_
"<CookieDomain></CookieDomain>"&vbCrlf&_
"<NoCacheHeaders>0</NoCacheHeaders>"

	GenericTop="<a target=""_blank"" href=""http://www.ed2000.com"" title=""ED2000资源共享""><img border=""0"" src=""images/ED2000_468x60.gif""/></a>"
	GenericBottom="<a target=""_blank"" href=""http://www.ed2000.com"" title=""ED2000资源共享""><img border=""0"" src=""images/ED2000_468x60.gif""/></a>"
	InstallExecute("insert into ["&TablePrefix&"SiteSettings] (SiteSettingsXML,GenericTop,GenericBottom,BBSxpVersion) values ('"&SiteSettingsXML&"','"&GenericTop&"','"&GenericBottom&"','"&DataBaseVer&"')")
	
	
	InstallExecute("insert into ["&TablePrefix&"Links] (Name,URL,Logo,Intro) values ('BBSXP官方网','http://www.bbsxp.com','Images/BBSXP_88x31.gif','BBSXP官方技术支持')")
	InstallExecute("insert into ["&TablePrefix&"Links] (Name,URL,Intro) values ('多次搜索','http://www.duoci.com','多线程搜索引擎集合')")
	InstallExecute("insert into ["&TablePrefix&"Links] (Name,URL,Intro) values ('资源共享','http://www.ed2000.com','ED2000为您提供最新电影、电视剧、综艺、动漫、音乐、游戏、图书、软件、资料、教育等各类资源。')")
	InstallExecute("insert into ["&TablePrefix&"Links] (Name,URL,Intro) values ('QQ表情网','http://www.biaoqing.com','中文QQ表情网站')")


	if IsSqlDataBase=1 then InstallExecute("SET IDENTITY_INSERT "&TablePrefix&"Roles ON")
	InstallExecute("insert into ["&TablePrefix&"Roles] (RoleID,Name,Description) values ('0','游客','不需要添加用户到该角色，该角色仅仅是作为权限隐射，所有的匿名用户和注册用户都属于该角色。')")
	InstallExecute("insert into ["&TablePrefix&"Roles] (RoleID,Name,Description) values ('1','管理员','享有论坛的最高权限，可以管理整个论坛。')")
	InstallExecute("insert into ["&TablePrefix&"Roles] (RoleID,Name,Description) values ('2','超级版主','可以管理论坛上的所有版块。')")
	InstallExecute("insert into ["&TablePrefix&"Roles] (RoleID,Name,Description) values ('3','注册会员','所有注册用户自动属于该角色。')")
	if IsSqlDataBase=1 then InstallExecute("SET IDENTITY_INSERT "&TablePrefix&"Roles OFF")
	
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','工兵','0')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','排长','50')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','连长','150')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','营长','500')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','团长','1000')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','旅长','2000')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','师长','5000')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','军长','20000')")
	InstallExecute("insert into ["&TablePrefix&"Ranks] (RoleID,RankName,PostingCountMin) values ('0','司令','50000')")
	
    end if
%></textarea>
</div>
<div style="width:auto; padding:10px;background:#698CC3; color:#FFFFFF; font-weight:bold; height:20px">
<div style="float:left"><input type="button" value="返回安装向导首页" onClick="window.location.href = '?';" /></div>
<div style="float:right"><input type="button" value="下一步" onClick="window.location.href = '?menu=SiteSetting';" /></div>
</div>


<%
End Sub

%>