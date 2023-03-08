<!-- #include file="Setup.asp" -->
<%
ForumName=HTMLEncode(Request.Form("ForumName"))
ForumIcon=HTMLEncode(Request.Form("ForumIcon"))
ForumLogo=HTMLEncode(Request.Form("ForumLogo"))
ForumIntro=HTMLEncode(Request.Form("ForumIntro"))
ForumRules=HTMLEncode(Request.Form("ForumRules"))


if SiteSettings("ForumApply")=0 then error("<li>对不起，本系统关闭论坛申请功能")

top

if Request("menu")="Apply" then

if CookieUserName=empty then error("<li>您还未<a href=Login.asp>登录</a>论坛")
if ForumName="" then error("<li>请输入论坛名称")
if Len(ForumName)>30 then  error("<li>论坛名称不能大于 30 个字符")
if Len(ForumIntro)>255 then  error("<li>论坛简介不能大于 255 个字节")
if Request.Cookies("ApplyForum")="1" then error("<li>您已经申请过论坛了，请不要再次申请！")


Rs.Open "[BBSXP_Forums]",Conn,1,3
Rs.addNew
Rs("ForumName")=ForumName
Rs("moderated")=CookieUserName
Rs("ForumIntro")=ForumIntro
Rs("ForumRules")=ForumRules
Rs("ForumIcon")=ForumIcon
Rs("ForumLogo")=ForumLogo
Rs("ForumHide")=1
Rs.update
id=Rs("id")
Rs.close

Conn.execute("insert into [BBSXP_Favorites](UserName,name,url)values('"&CookieUserName&"','"&id&"','forum')")


Mailaddress=Conn.Execute("Select UserMail From [BBSXP_Users] where UserName='"&CookieUserName&"'")(0)
MailTopic="论坛系统开通通知"
body=""&vbCrlf&"亲爱的"&CookieUserName&", 您好!"&vbCrlf&""&vbCrlf&"　　恭喜! 您已经成功地申请了"&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&")的论坛系统, 非常感谢您使用"&SiteSettings("CompanyName")&"的服务!"&vbCrlf&""&vbCrlf&"　* 我们免费为您的论坛提供了一个比较好记的地址,请您试试"&vbCrlf&""&vbCrlf&"　* URL: "&Request("ForumName")&"("&SiteSettings("SiteURL")&"Default.asp?id="&id&")"&vbCrlf&""&vbCrlf&"　* 最后, 有几点注意事项请您牢记"&vbCrlf&"1、不得使用本论坛系统建立任何包含色情、非法、以及危害国家安全的内容的论坛。"&vbCrlf&"2、不得在本系统用户所拥有的论坛内发布任何色情、非法、或者危害国家安全的言论。"&vbCrlf&"3、以上规则违者责任自负，本站有权删除该类用户或者内容，并追究其法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&SiteSettings("CompanyName")&"("&SiteSettings("CompanyURL")&") 提供　程序制作：YUZI工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/Mail.asp" -->
<%

Response.Cookies("ApplyForum")="1"


Message="<li>恭喜您，申请成功！<li>我们为您的新论坛免费提供了一个属于您自己的域名<li><a target=_blank href=Default.asp?id="&id&">"&SiteSettings("SiteURL")&"Default.asp?id="&id&"</a><li> 请在 <a target=_blank href=Default.asp?id="&id&">"&Request("ForumName")&"</a> 上点击您鼠标的右键，把这个域名加入您的书签或者收藏夹中"
succeed(""&Message&"")


end if


%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 申请论坛</td>
</tr>
</table><br>

<center>
<form name="form" method="POST" action="ApplyForum.asp">
<input type=hidden name=menu value="Apply">


<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>申请论坛</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">论坛名称：</td>
<td class=a3 height="5" align="Left" valign="middle" width="78%">
&nbsp;
<input type="text" name="ForumName" size="30" maxlength="12">
</td>
</tr>

<tr class=a4>
<td height="2" align="right" width="20%">论坛介绍：</td>
<td height="2" align="Left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="ForumIntro" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td height="2" align="right" width="20%">论坛规则：</td>
<td height="2" align="Left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="ForumRules" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a4>
<td height="1" align="right" valign="middle" width="20%">小图标URL：</td>
<td height="1" align="Left" valign="middle" width="78%">
&nbsp;
<input size="30" name="icon">　显示在论坛介绍右边</td>
</tr>
<tr class=a3>
<td align="right" valign="bottom" width="20%">大图标URL：</td>
<td align="Left" valign="bottom" width="78%">
&nbsp;
<input size="30" name="ForumLogo">　显示在论坛左上角</td>
</tr>
<tr class=a3>
<td align="right" width="100%" colspan="2">
<input type="submit" value=" 确 定 &gt;&gt;下 一 步 "> </td>
</tr>
</table>

</form>
<%
htmlend
%>