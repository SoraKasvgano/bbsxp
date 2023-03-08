<!-- #include file="Setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if Request("menu")="face" then
%>
<script>
function insertface(userface){
self.opener.form.Userface.value=userface;
self.opener.document.images.tus.src=userface;
window.close();
}
</script>
<body topmargin=0>
<title>BBSXP--头像列表 - Powered By BBSXP</title>
<table border=0 width=100% cellspacing=1 cellpadding=1>
<tr class=a1><td colspan="10" height="25" align=center>BBSXP头像</td></tr>
<tr align=center>
<script>
var ii=1
for (var i=1; i <= 95; i++) {
ii++
document.write("<td class=a3><img src=images/face/"+i+".gif onclick=insertface('images/face/"+i+".gif') style=CURSOR:hand><br>"+i+"</td>");
if(ii >5){document.write("</tr><tr align=center>");ii=1}
}
</script>
</tr>
</table>
<%
CloseDatabase
end if


UserName=HTMLEncode(Request.Form("UserName"))

errorchar=array(" ","　","","","	","","","#","`","|","%","&","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error2("用户名中不能含有特殊符号")
next

if SiteSettings("BannedRegUserName")<>"" then
filtrate=split(""&SiteSettings("BannedRegUserName")&"","|")
for i = 0 to ubound(filtrate)
if instr(UserName,filtrate(i))>0 then error2("你的用户名中包含有系统禁止注册的字符")
next
end if






if SiteSettings("CloseRegUser")= 1 then error("<li>本论坛暂时不开放新用户注册！")

top
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request("menu")="AddUserName" then

if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then
if Request.Form("VerifyCode")<>Session("VerifyCode") then Message=Message&"<li>验证码错误"
end if

UserSign=HTMLEncode(Request.Form("UserSign"))

password=Trim(Request.Form("password"))
Userpass2=Trim(Request.Form("Userpass2"))
UserMail=HTMLEncode(Request.Form("UserMail"))
Userhome=HTMLEncode(Request.Form("Userhome"))
PasswordQuestion=HTMLEncode(Request.Form("PasswordQuestion"))
PasswordAnswer=HTMLEncode(Request.Form("PasswordAnswer"))
birthday=HTMLEncode(Request.Form("birthday"))
Userface=HTMLEncode(Request.Form("Userface"))
UserSex=HTMLEncode(Request.Form("UserSex"))


if SiteSettings("BannedText")<>empty then
filtrate=split(SiteSettings("BannedText"),"|")
for i = 0 to ubound(filtrate)
UserSign=ReplaceText(UserSign,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
next
end if


if UserName="" then Message=Message&"<li>您的用户名没有填写"

if Len(UserName)>16 then Message=Message&"<li>您的用户名中不能超过16个字节"

if SiteSettings("EnableUser")=1 then 
Randomize
password=int(rnd*999999)+1
else
if password="" then Message=Message&"<li>您没有输入密码"
if password<>Userpass2 then Message=Message&"<li>您2次输入的密码不相同"
end if

if instr(UserMail,"@")=0 then Message=Message&"<li>您的电子邮件地址填写错误"

if Len(UserSign)>255 then Message=Message&"<li>签名档不能大于 255 个字节"

if instr(Userface,";")>0 then Message=Message&"<li>头像URL中不能含有特殊符号"

If not Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof Then Message=Message&"<li>此用户名已经被别人注册了"

if SiteSettings("OnlyMailReg") = 1 then
If not Conn.Execute("Select id From [BBSXP_Users] where UserMail='"&UserMail&"'" ).eof Then Message=Message&"<li>此Email已经被别人注册了"
end if

if Message<>"" then error(""&Message&"")

for each ho in request.form("character")
allcharacter=""&allcharacter&""&ho&""
next

if SiteSettings("EnableUser")=2 then
membercode=0
else
membercode=1
end if

Rs.Open "select * from [BBSXP_Users]",Conn,1,3
Rs.addNew
Rs("UserName")=UserName
Rs("Userpass")=md5(password)
Rs("UserMail")=UserMail
Rs("Userhome")=Userhome
Rs("PasswordQuestion")=PasswordQuestion
Rs("membercode")=membercode
if Request("PasswordAnswer")<>empty then Rs("PasswordAnswer")=md5(Request("PasswordAnswer"))
Rs("birthday")=birthday
Rs("Userface")=Userface
Rs("UserSex")=UserSex
Rs("UserSign")=UserSign
Rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""
Rs("UserFriend")="|"
Rs("UserRegTime")=""&now()&""
Rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
Rs("UserLandTime")=""&now()&""
Rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("Postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
Rs("UserIM")=""&HTMLEncode(Request("qq"))&"\"&HTMLEncode(Request("icq"))&"\"&HTMLEncode(Request("uc"))&"\"&HTMLEncode(Request("aim"))&"\"&HTMLEncode(Request("msn"))&"\"&HTMLEncode(Request("Yahoo"))&""
Rs.update
Rs.close

Conn.execute("update [BBSXP_Statistics_Site] set TotalUser=TotalUser+1,NewUser='"&UserName&"'")


Mailaddress=UserMail
MailTopic="用户名注册成功"
body=""&vbCrlf&"亲爱的"&UserName&", 您好!"&vbCrlf&""&vbCrlf&"　　恭喜! 您已经成功地注册了您的资料, 非常感谢您使用"&SiteSettings("CompanyName")&"的服务!"&vbCrlf&""&vbCrlf&"　* 您的帐号是:"&UserName&"　密码是:"&password&""&vbCrlf&""&vbCrlf&"　* "&SiteSettings("SiteName")&"("&SiteSettings("SiteURL")&"Default.asp)"&vbCrlf&""&vbCrlf&"　* 最后, 有几点注意事项请您牢记"&vbCrlf&"1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定。"&vbCrlf&"2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"&vbCrlf&"3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") 提供　程序制作：YUZI工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%

if SiteSettings("EnableUser")=0 then
Response.Cookies("UserName")=escape(UserName)
Response.Cookies("Userpass")=md5(password)
end if


Message=Message&"<li>注册新用户资料成功<li><a href=Default.asp>返回论坛首页</a>"
succeed(""&Message&"<meta http-equiv=refresh content=3;url=Default.asp>")



elseif Request("menu")="write" then


If not Conn.Execute("Select id From [BBSXP_Users] where UserName='"&UserName&"'" ).eof or UserName="" Then error2("您所选的用户名 "&UserName&" 已经有人使用，请另外选择一个用户名。")


%>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="CreateUser.asp">注册协议</a> → 
<a href="CreateUser.asp?menu=Check">检查用户名</a> → 填写用户资料</td>
</tr>
</table><br>

<table width=100% align=center cellspacing=0 cellpadding=0 border=0 class=a1>
<form method="POST" name="form" action=?menu=AddUserName onsubmit="return VerifyInput();">
<input type=hidden name=UserName value="<%=UserName%>">
<input type=hidden name=Userface value="images/face/1.gif">
<tr>
<td>
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr class=a1>
<td colspan="2" height="25" valign="middle" align="Left"><b>&nbsp;个人社区资料</b>（以下内容必填）</td>
</tr>
<tr>
<td class=a3 height="25" align="right" width="46%"><b>用户名：</b></td>
<td class=a3 height="25" align="Left" width="54%">
&nbsp;
<%=UserName%></td>
</tr>

<%if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then%>
<tr>
<td class=a3 height="12" align="right" width="46%"><b>验证码：</b></td>
<td class=a3 align="Left">
&nbsp;
<input name="VerifyCode" size="10"> <img src="VerifyCode.asp" alt="验证码,看不清楚?请点击刷新验证码" style=cursor:pointer onclick="this.src='VerifyCode.asp'"></td>
</tr>
<%end if%>


<%if SiteSettings("EnableUser")<>1 then%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>密码：</b></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;
<input type="password" name="password" size="40" maxlength="16">
</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>确认密码：</b><br>请与您的密码保持一致</td>
<td height="2" align="Left" valign="middle" width="54%" class=a3>
&nbsp;
<input type="password" name="Userpass2" size="40">
</td>
</tr>
<%end if%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>您的Email地址：</b><br>
<%if SiteSettings("EnableUser")=1 then%><font color="FF0000">密码将通过Email发送</font><%end if%></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;
<input type="text" name="UserMail" size="40">
</td>
</tr>
<tr class=a1>
<td height="25" align="Left" valign="middle" colspan="2"><b>&nbsp;个人生活资料</b>（以下内容建议填写）</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a3 width="46%"> 　<b>真实姓名：</b>
<input type="text" name="realname" size="18">
</td>
<td height="71" align="Left" valign="top" class=a4 rowspan="16" width="54%">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr>
<td>

<table border=0>
	<tr>
		<td><b>头　像： &nbsp;</b></td>
		<td>
		<table title="选择头像" cellspacing=1 cellpadding=0 border=0 background=Images/FaceBackground.gif width=54 style=cursor:hand onclick="javascript:open('CreateUser.asp?menu=face','','width=500,height=500,resizable,scrollbars')"><tr><td><img src="images/face/1.gif" width=40 height=40 name=tus border="0" alt="选择头像"></td></tr></table>
		</td>
	</tr>
</table>


</td>



</tr>
<tr>
<td><b>生　日：<font color="000000">　</font></b> <input onfocus="calendar()" value="" name="birthday"></td>
</tr>
<tr>
<td><b>个人主页： </b> 
<input type="text" name="Userhome" size="40" value="http://"></td>
</tr>
<tr>
<td><b><font color="000000">性　格</font>：</b>
<br>
<SCRIPT>
var moderated="成熟稳重 幼稚调皮 温柔体贴 活泼可爱 脾气暴躁 内向害羞 外向开朗 心地善良 风趣幽默 慢条斯理 积极进取 郁郁寡欢 处事洒脱 圆滑老练 豪放不羁 异想天开 多愁善感 淡泊名利 情绪多变 胆小怕事 循规蹈矩 热心助人 快言快语 爱管闲事 追求刺激"
var list= moderated.split (' '); 
for(i=0;i<list.length;i++) {
if (i==5 || i==10 || i==15 || i==20) document.write("<br>")
if (list[i] !=""){document.write("　 <input type=checkbox value=' "+list[i]+"' name=character id=character"+i+"> <label for=character"+i+">"+list[i]+"</label>")}
}
</SCRIPT>

</td>
</tr>
<tr>
<td><b>个人简介：</b><br>
<textarea name=personal rows=6 cols=65></textarea>
</td>
</tr>
<tr>
<td height="5"><b>签名档：</b><br>
<textarea name=UserSign rows=6 cols=65></textarea>  </td>
</tr>
</table>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%"><b> 　性　　别：</b>
<select size=1 name=UserSex>
<option value="" selected>[请选择]</option>
<option value=male>男</option>
<option value=female>女</option>
</select>
</td>
</tr>

<tr class=a3>
<td height="2" valign="top" width="46%">　<b>国　　家：</b>
<b>
<input type="text" name="country" size="18">
</b> </td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>省　　份：</b>
<input type="text" name="province" size="18">
</td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">　<b>城　　市：
</b>
<input type="text" name="city" size="18">
</td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">　<b>邮政编号： 
<input type="text" name="Postcode" size="18"></b></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">
　<b>血　　型：</b>
<select size=1 name=blood>
<option value="" selected>[请选择]</option>
<option
value=A>A</option>
<option value=B>B</option>
<option
value=AB>AB</option>
<option value=O>O</option>
<option
value=其他>其他</option>
</select>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>信　　仰：</b>
<select size=1 name=belief>
<option value="" selected>[请选择]</option>
<option value=佛教>佛教</option>
<option
value=道教>道教</option>
<option value=基督教>基督教</option>
<option
value=天主教>天主教</option>
<option value=回教>回教</option>
<option
value=无神论者>无神论者</option>
<option value=共产主义者>共产主义者</option>
<option
value=其他>其他</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>职　　业： </b>
<select name="occupation">
<option value="" selected>[请选择]</option>
<option value="财会/金融">财会/金融</option>
<option value="工程师">工程师</option>
<option value="顾问">顾问</option>
<option value="计算机相关行业">计算机相关行业</option>
<option value="家庭主妇">家庭主妇</option>
<option value="教育/培训">教育/培训</option>
<option value="客户服务/支持">客户服务/支持</option>
<option value="零售商/手工工人">零售商/手工工人</option>
<option value="退休">退休</option>
<option value="无职业">无职业</option>
<option value="销售/市场/广告">销售/市场/广告</option>
<option value="学生">学生</option>
<option value="研究和开发">研究和开发</option>
<option value="一般管理/监督">一般管理/监督</option>
<option value="政府/军队">政府/军队</option>
<option value="执行官/高级管理">执行官/高级管理</option>
<option value="制造/生产/操作">制造/生产/操作</option>
<option value="专业人员">专业人员</option>
<option value="自雇/业主">自雇/业主</option>
<option value="其他">其他</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>婚姻状况：</b>
<select size=1 name=marital>
<option value="" selected>[请选择]</option>
<option value=未婚>未婚</option>
<option
value=已婚>已婚</option>
<option value=离异>离异</option>
<option
value=丧偶>丧偶</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>最高学历：</b>
<select size=1 name=education>
<option value="" selected>[请选择]</option>
<option value=小学>小学</option>
<option
value=初中>初中</option>
<option value=高中>高中</option>
<option
value=大学>大学</option>
<option value=硕士>硕士</option>
<option
value=博士>博士</option>
</select></td>
</tr>
<tr class=a3>
<td height="1" valign="top" width="46%">　<b>毕业院校：</b>
<input type="text" name="college" size="18"></td>
</tr>

<tr class=a1>
<td valign="top" height="25"><b>&nbsp;私人联系资料</b>（以下资料不公开）</td>
</tr>

<tr class=a3>
<td valign="top" width="46%">　<b>联系地址： 
<input type="text" name="address" size="18"></b></td>
</tr>

<tr class=a3>
<td valign="top" width="46%">　<b>电话号码： 
<input type="text" name="phone" size="18"></b></td>
</tr>

<tr class=a3>
<td valign="top" width="46%">　<b>手机号码： 
<input type="text" name="UserMobile" size="18"></b></td>
</tr>

<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;即时通讯资料</b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">　<b>QQ号码： 
<input type="text" name="qq" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
<td height="4" valign="middle" align="Left"><b>ICQ：　　　<input type="text" name="icq" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
</tr>
<tr class=a4>
<td height="4" valign="middle" align="Left">　<b>UC号码： <input type="text" name="uc" size="20" onkeyup=if(isNaN(this.value))this.value=''></b></td>
<td height="4" valign="middle" align="Left"><b>AIM：　　　<input type="text" name="aim" size="20"></b></td>
</tr>
<tr class=a3>
<td height="4" valign="middle" align="Left">　<b>MSN IM： 
<input type="text" name="msn" size="20"></b>　</td>
<td height="4" valign="middle" align="Left"><b>Yahoo IM： <input type="text" name="Yahoo" size="20"></b></td>
</tr>
<tr class=a1>
<td colspan="2" height="25" align="Left"><b>&nbsp;密码保护资料</b></td>
</tr>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>密码提示问题：</b></td>
<td class=a4 height="2" align="Left" valign="middle" width="54%">
&nbsp;


<select name="PasswordQuestion">
<option value="" selected>[请选择]</option>
<option value="最喜欢的宠物？">最喜欢的宠物？</option>
<option value="最喜爱的电影？">最喜爱的电影？</option>
<option value="周年纪念日 [年/月/日]？">周年纪念日 [年/月/日]？</option>
<option value="父亲的名字？">父亲的名字？</option>
<option value="配偶的名字？">配偶的名字？</option>
<option value="第一个孩子的爱称？">第一个孩子的爱称？</option>
<option value="中学的校名？">中学的校名？</option>
<option value="最尊敬的老师？">最尊敬的老师？</option>
<option value="最喜欢的运行队？">最喜欢的运行队？</option>
</select>

</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>密码提示答案：</b></td>
<td height="2" align="Left" valign="middle" width="54%"> &nbsp;
<input type="text" name="PasswordAnswer" size="40">
</td>
</tr>

<tr align="center" class=a4>
<td height="2" valign="middle" colspan="2">
<input type="submit" value=" 注 册 >>下 一 步 ">
</td>
</tr>
</table>
</td>
</tr></table>



<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="中国"}
else
if (language=="zh-hk"){country="中国香港"}
else
if (language=="zh-tw"){country="中国台湾"}
else
if (language=="zh-sg"){country="新加坡"}
else
country=""
document.form.country.value=country


function VerifyInput()
{
var Mail = document.form.UserMail.value;
if(Mail.indexOf('@',0) == -1 || Mail.indexOf('.',0) == -1){
alert("您输入的Email有错误\n请重新检查您的Email");
document.form.UserMail.focus();
return false;
}

if (document.form.UserMail.value == "")
{
alert("请输入您的Email地址");
document.form.UserMail.focus();
return false;
}

return true;
}
</SCRIPT>
<%
elseif Request("menu")="" then
%>
<center>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 注册协议</td></tr></tbody></table>
<br>
<textarea name="textarea" rows="18" readOnly cols="100"><%=SiteSettings("RegUserAgreement")%></textarea><br><br>
<input type="submit" value=" 同 意 " onclick="window.location.href='CreateUser.asp?menu=Check';">
<input type="submit" value=" 不同意 " onclick="history.back()">
<%
elseif Request("menu")="Check" then
%>
<center><form method="POST" name="form" action=?menu=write>
<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="CreateUser.asp">注册协议</a> → 检查用户名</td></tr></tbody></table>
<br>
<br>请输入您要注册的用户名<BR>
<input type="text" name="UserName" size="28" maxlength="12" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')">
<br><br>
<input type="button" value=" 上一步 " onclick="history.back()">
<input type="submit" value=" 下一步 " id='btnadd' disabled></form>
<%
end if

htmlend
%>