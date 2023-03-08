<!-- #include file="Setup.asp" -->
<!-- #include file="API_Request.asp" -->
<!--#include FILE="Utility/UpFile_Class.asp"-->

<%
if Request("menu")="upface" then
	if CookieUserName=empty then AlertForModal("您还未登录论坛")
	UpFace
elseif Request("menu")="PostUpFace" then
	if CookieUserName=empty then AlertForModal("您还未登录论坛")
	UserID=Execute("Select UserID From ["&TablePrefix&"Users] where UserID="&CookieUserID&"")(0)
	UpFolder="UpFile/UpFace/"&CookieUserID&""	'上传路径及文件名
	UpClass="Face"
	UpMaxFileSize=SiteConfig("MaxFaceSize")*1024		'最大上传文件大小
%>
<!--#include FILE="Utility/UpFile.asp"-->
	<script language="JavaScript" type="text/javascript">
		parent.document.form.UserFaceUrl.value='<%=SaveFile%>';
		parent.document.getElementById("tus").src='<%=SaveFile%>';
		parent.BBSXP_Modal.Close();
	</script>
<%
else
if CookieUserName=empty then error("您还未<a href=""javascript:BBSXP_Modal.Open('Login.asp',380,170);"">登录</a>论坛")

HtmlTop
%>

<%

select case Request("menu")
	case "editProfileok"
		editProfileok
	case "passok"
		passok
end select
%>
<script type="text/javascript" src="Utility/pswdplc.js"></script>
<script type="text/javascript" src="Utility/calendar.js"></script>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 编辑资料</div>

<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<tr class=CommonListTitle>
		<td align="center"><a href="EditProfile.asp">资料修改</a></td>
		<td align="center"><a href="EditProfile.asp?menu=pass">密码修改</a></td>
		<td align="center"><a href="MyUpFiles.asp">上传管理</a></td>
		<td align="center"><a href="MyFavorites.asp">收 藏 夹</a></td>
		<td align="center"><a href="MyMessage.asp">短信服务</a></td>
	</tr>
</table>
<br />
<%
select case Request("menu")
	case "pass"
		pass
	case else
		default
end select

Sub default
	sql="Select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&""
	Set Rs=Execute(sql)
		if Rs.eof then error("该用户资料不存在")

		UserSign=replace(""&Rs("UserSign")&"","<br>",vbCrlf)
		UserBio=replace(""&Rs("UserBio")&"","<br>",vbCrlf)

		birthday=Rs("birthday")

%>
<form method="POST" name="form" action="EditProfile.asp">
<input type=hidden name=menu value="editProfileok" />


<div class="MenuTag"><li id="Tag_btn_0" onclick="ShowPannel(this)" class="NowTag">基本资料</li><li id="Tag_btn_1" onclick="ShowPannel(this)">签名＆简介</li><li id="Tag_btn_2" onclick="ShowPannel(this)">即时通信</li><li id="Tag_btn_3" onclick="ShowPannel(this)">头像设置</li></div>
<table id="Tag_tab_0" style="display:block" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
      <td width="100"><b>名字</b></td>
      <td><input type="text" name="RealName" size="20" value="<%=Rs("RealName")%>" /></td>
    </tr>
<%if SiteConfig("CustomUserTitle")=1 then%>
	<tr>
      <td width="100"><b>头衔</b></td>
      <td><input type="text" name="UserTitle" size="20" value="<%=Rs("UserTitle")%>" /></td>
    </tr>
<%end if%>
	<tr>
    	<td width="100"><b>性别</b>		</td>
    	<td><select size=1 name=UserSex>
				<option value=0 selected="selected">[请选择]</option>
				<option value=1 <%if Rs("UserSex")=1 then%>selected<%end if%>>男</option>
				<option value=2 <%if Rs("UserSex")=2 then%>selected<%end if%>>女</option>
			</select></td>
	</tr>
	<tr>
		<td width="100"><b>生日</b></td>
		<td><input onclick="showcalendar(event, this)" onfocus="showcalendar(event, this)" size=24 value="<%=birthday%>" name=birthday /></td>
	</tr>
	<tr>
	    <td width="100"><b>职业</b></td>
	    <td><input type="text" name="Occupation" size="30" value="<%=Rs("Occupation")%>" /></td>
	</tr>
	<tr>
		<td width="100"><b>兴趣</b></td>
		<td><input type="text" name="Interests" size="30" value="<%=Rs("Interests")%>" /></td>
	</tr>
	<tr>
		<td width="100"><b>地址</b></td>	
		<td><input type="text" name="Address" size="30" value="<%=Rs("Address")%>" /></td>	
	</tr>
	<tr>
		<td width="100"><b>主页</b></td>
		<td><input type="text" name="WebAddress" size="30" value="<%=Rs("WebAddress")%>" /></td>
	</tr>
	<tr>
		<td width="100"><b>博客</b></td>
		<td><input type="text" name="WebLog" size="30" value="<%=Rs("WebLog")%>" /></td>
	</tr>
	<tr>
		<td width="100"><b>相册</b></td>
		<td><input type="text" name="WebGallery" size="30" value="<%=Rs("WebGallery")%>" /></td>
	</tr>
</table>
<table id="Tag_tab_1" style="display:none;" cellspacing=0 cellpadding=0 width=100% class=PannelBody>
	<tr>
		<td>
			<p><b>签名</b><br /><textarea name=UserSign rows=7 cols=80><%=UserSign%></textarea></p>
			<p><b>简介</b><br /><textarea name=UserBio rows=8 cols=80><%=UserBio%></textarea></p>
		</td>
	</tr>
</table>
	<table id="Tag_tab_2" style="display:none" cellspacing=0 cellpadding=0 width=100% class=PannelBody>
		<tr>
			<td width="100"><img src=images/im_qq.gif align=absmiddle /> <b>QQ</b></td>
			<td><b><input type="text" name="QQ" size="20" value="<%=Rs("QQ")%>" /></b></td>
		</tr>
		<tr>
			<td><img src=images/im_icq.gif align=absmiddle /> <b>ICQ</b></td>
			<td><b><input type="text" name="ICQ" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=Rs("ICQ")%>" /></b></td>
		</tr>
		<tr>
			<td><img src=images/im_aim.gif align=absmiddle /> <b>AIM</b></td>
			<td><b><input type="text" name="AIM" size="20" value="<%=Rs("AIM")%>" /></b></td>
		</tr>
		<tr>
			<td><img src=images/im_msn.gif align=absmiddle /> <b>MSN</b></td>
			<td><b><input type="text" name="MSN" size="20" value="<%=Rs("MSN")%>" /></b></td>
		</tr>
		<tr>
			<td><img src=images/im_yahoo.gif align=absmiddle /> <b>Yahoo</b></td>
			<td><b><input type="text" name="Yahoo" size="20" value="<%=Rs("Yahoo")%>" /></b></td>
		</tr>
		<tr>
			<td><img src=images/im_skype.gif align=absmiddle /> <b>Skype</b></td>
			<td><b><input type="text" name="Skype" size="20" value="<%=Rs("Skype")%>" /></b></td>
		</tr>
	</table>
<table id="Tag_tab_3" style="display:none" cellspacing=0 cellpadding=0 width=100% class=PannelBody>
	<tr>
		<td width="50">
			<table cellspacing=0 cellpadding=0 border=0 width=62 style='cursor:pointer' onclick="javascript:open('Utility/Face.htm','','width=500,height=500,resizable,scrollbars')"><tr><td style="padding-left:1px; padding-top:1px; background-repeat:no-repeat" background=Images/FaceBackground.gif><img id="tus" src="<%=Rs("UserFaceUrl")%>" width=48 height=48 name=tus border="0" title="选择系统给定的头像" /></td></tr></table>
		</td>
		<td>
			<b>头像地址</b> <input type="text" value="<%=Rs("UserFaceUrl")%>" name="UserFaceUrl" size="40" <%if SiteConfig("EnableRemoteAvatars")=0 then%> readonly="readonly"<%end if%> /> <a href="javascript:BBSXP_Modal.Open('EditProfile.asp?menu=upface',500,150);" class="CommonTextButton" title="上传本地自定义头像">上传头像</a>
	 	</td>
	</tr>
</table>


<div style="padding-top:5px"><input type="submit" value=" 确 定 " /></div>
</form>

<%
End Sub

Sub editProfileok
	UserSign=HTMLEncode(Request.Form("UserSign"))
	UserFaceUrl=HTMLEncode(Request("UserFaceUrl"))
	birthday=Request("birthday")
	UserTitle=HTMLEncode(Request.Form("UserTitle"))
	
	if birthday<>"" then
		if Not IsDate(birthday) then Message=Message&"<li>出生日期输入格式错误"
	end if
	
	if instr(UserFaceUrl,";")>0 or instr(UserFaceUrl,"%")>0 or instr(UserFaceUrl,"javascript:")>0 then Message=Message&"<li>头像URL中不能含有特殊符号"

	if Len(UserSign)>SiteConfig("SignatureMaxLength") then Message=Message&"<li>签名档不能大于 "&SiteConfig("SignatureMaxLength")&" 个字节"
	if Len(UserTitle)>SiteConfig("UserTitleMaxChars") then Message=Message&"<li>头衔名称不能大于 "&SiteConfig("UserTitleMaxChars")&" 个字节"
	
	
	if Message<>"" then error(""&Message&"")

	sql="Select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&""
	Rs.Open sql,Conn,1,3

		if IsDate(birthday) then
			Rs("birthday")=birthday
		else	
			Rs("birthday")=null
		end if

		if SiteConfig("CustomUserTitle")=1 then
			if SiteConfig("UserTitleCensorWords")<>"" then UserTitle=ReplaceText(UserTitle,"("&SiteConfig("UserTitleCensorWords")&")",string(len("&$1&"),"*"))
			Rs("UserTitle")=UserTitle
		end if
		
		Rs("UserFaceUrl")=UserFaceUrl
		Rs("UserSex")=RequestInt("UserSex")
		Rs("UserBio")=HTMLEncode(Request("UserBio"))
		Rs("RealName")=HTMLEncode(Request("RealName"))
		Rs("Occupation")=HTMLEncode(Request("Occupation"))
		Rs("Address")=HTMLEncode(Request("Address"))
		Rs("Interests")=HTMLEncode(Request("Interests"))
		Rs("WebAddress")=HTMLEncode(Request("WebAddress"))
		Rs("WebLog")=HTMLEncode(Request("WebLog"))
		Rs("WebGallery")=HTMLEncode(Request("WebGallery"))
		Rs("UserSign")=UserSign

		Rs("QQ")=HTMLEncode(Request("QQ"))
		Rs("ICQ")=HTMLEncode(Request("ICQ"))
		Rs("AIM")=HTMLEncode(Request("AIM"))
		Rs("MSN")=HTMLEncode(Request("MSN"))
		Rs("Yahoo")=HTMLEncode(Request("Yahoo"))
		Rs("Skype")=HTMLEncode(Request("Skype"))

	Rs.update
	Rs.close

	UpdateApplication"UserBirthday","Select UserID,UserName From ["&TablePrefix&"Users] where Month(Birthday)="&Month(now())&" and day(Birthday)="&day(now())&""

	Message=Message&"<li>修改资料成功</li><li><a href=Profile.asp?UID="&CookieUserID&">查看我的资料</a></li>"
	succeed Message,"Profile.asp?UID="&CookieUserID&""
	
	
End Sub

Sub pass
	sql="Select * from ["&TablePrefix&"Users] where UserID="&CookieUserID&""
	Set Rs=Execute(sql)
%>
<form method="POST" name="form"action="EditProfile.asp" onsubmit="return VerifyInput();">
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<input type=hidden name="menu" value="passok" />
	<tr class=CommonListTitle>
		<td align="center" colspan="2" valign="bottom">用户密码修改</td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b> 原密码：</b></td>
	    <td width="55%"> <input type="password" name="OldPassword" size="40" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b> 新密码：</b><br />如不更改密码此处请留空</td>
		<td width="55%"> <input type="password" name="UserPassword" size="40" maxlength="16" onkeyup="EvalPwd(this.value);" onchange="EvalPwd(this.value);" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b>密码强度：</b></td>
	    <td width="55%">
			<table border="0" width="250" cellspacing="1" cellpadding="2">
				<tr bgcolor="#f1f1f1">
					<td id=iWeak align="center">弱</td>
					<td id=iMedium align="center">中</td>
					<td id=iStrong align="center">强</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b>重新键入新密码：</b><br />请与您的新密码保持一致</td>
	    <td valign="middle" width="55%"> <input type="password" name="RetypePassword" size="40" maxlength="16" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="right" width="45%"><b>邮箱地址：</b></td>
	    <td width="55%"> <input type="text" name="UserEmail" size="40" value="<%=Rs("UserEmail")%>" /></td>
	</tr>
	
	<tr class="CommonListCell">
    	<td align="right" width="45%"><b>密码提示问题：</b></td>
   		<td width="55%"><input type="text" name="PasswordQuestion" size="40" value="<%=Rs("PasswordQuestion")%>" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="45%"><b>机密答案：</b><br>如不更改答案此处请留空</td>
    	<td width="55%" class=a4><input type="text" name="PasswordAnswer" size="40" value="" /></td>
	</tr>
	<tr class="CommonListCell">
	    <td align="center" width="100%" colspan="2"> <input type="submit" value=" 确 定 " /></td>
	</tr>
</table>
</form>
<script language="JavaScript" type="text/javascript">
function VerifyInput()
{
	var Mail = document.form.UserEmail.value;
	if(Mail.indexOf('@',0) == -1 || Mail.indexOf('.',0) == -1){
		alert("您输入的Email有错误\n请重新检查您的Email");
		document.form.UserEmail.focus();
		return false;
	}
	if (document.form.UserEmail.value == "")
	{
		alert("请输入您的Email地址");
		document.form.UserEmail.focus();
		return false;
	}
	if (document.form.UserPassword.value != document.form.RetypePassword.value)
	{
		alert("您二次键入的新密码不同！");
		document.form.UserPassword.focus();
		return false;
	}
	return true;
}
</script>
<%
End Sub

Sub passok
	OldPassword=Trim(Request("OldPassword"))
	if Len(CookieUserPassword)=32 then
		if MD5(OldPassword)<>CookieUserPassword then Message=Message&"<li>您的原密码错误"
	else
		if SHA1(OldPassword)<>CookieUserPassword then Message=Message&"<li>您的原密码错误"
	end if

	UserPassword=Trim(Request("UserPassword"))
	RetypePassword=Trim(Request("RetypePassword"))
	UserEmail=HTMLEncode(Request("UserEmail"))
	PasswordQuestion=HTMLEncode(Request("PasswordQuestion"))
	PasswordAnswer=HTMLEncode(Request("PasswordAnswer"))
	if UserPassword<>RetypePassword then Message=Message&"<li>您的新密码和确认新密码不同"
	if Len(UserPassword)<6 then Message=Message&"<li>新密码必须至少包含 6 个字符"
	
	if Message<>"" then error(""&Message&"")
	
	'--------------------  API Start  --------------------
	If SiteConfig("APIEnable")=1 Then APIUpdateUser CookieUserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer
	if Message<>"" then error(""&Message&"")
	'--------------------  API End  ----------------------


	if ModifyUserPassword(CookieUserName,UserPassword,UserEmail,PasswordQuestion,PasswordAnswer)=False then
		error(""&Message&"")
	else
		Message=Message&"<li>修改资料成功</li><li><a href=Profile.asp?UID="&CookieUserID&">查看我的资料</a></li>"
		succeed Message,"Profile.asp?UID="&CookieUserID&""
	end if

End Sub

HtmlBottom

end if

Sub UpFace
Response.clear
%>
<title>上传头像</title>
<style type="text/css">body,table{FONT-SIZE:9pt;}</style>
<form enctype=multipart/form-data method=Post action="?menu=PostUpFace">
<input type=file name="file" size=50 /><br /><br />
支持文件类型：GIF、JPG、JPEG、PNG<br />
文件大小限制：<%=CheckSize(SiteConfig("MaxFaceSize")*1024)%><br /><br />
<center><input type=submit value=" 上 传 " /> <input type="button" onclick="javascript:parent.BBSXP_Modal.Close()" value=" 取 消 " /></center>
</form>
<%
End Sub
%>