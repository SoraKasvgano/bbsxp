<!-- #include file="Setup.asp" --><%
if SiteSettings("AdminPassword")<>session("pass") then response.redirect "Admin.asp?menu=Login"
Log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))
response.write "<center>"

select case Request("menu")
case "Showbanner"
Showbanner
case "agreement"
agreement
case "variable"
variable
case "SiteSettingsUp"
SiteSettingsUp
case "SiteStat"
SiteStat
case "AutoSiteStat"
TotalUserCount=conn.execute("Select count(id) from [BBSXP_Users]")(0)
TotalThreadCount=conn.execute("Select count(id) from [BBSXP_Threads]")(0)
TotalPostCount=conn.execute("Select sum(replies) from [BBSXP_Threads]")(0)+TotalThreadCount
if TotalThreadCount=0 then TotalPostCount=0
Conn.execute("update [BBSXP_Statistics_Site] set TotalUser="&TotalUserCount&",TotalThread="&TotalThreadCount&",TotalPost="&TotalPostCount&"")
SiteStat
case "SiteStatok"
Rs.Open "[BBSXP_Statistics_Site]",Conn,1,3
for each ho in Request.Form
Rs(ho)=Request(ho)
next
Rs.update
Rs.close
%>���³ɹ�<br><br><a href="javascript:history.back()">�� ��</a> <%


end select
sub variable

AutoSiteURL="http://"&Request.ServerVariables("server_name")&Request.ServerVariables("script_name")&""
AutoSiteURL=Left(AutoSiteURL,inStrRev(AutoSiteURL,"/"))

SiteURL=SiteSettings("SiteURL")
if ""&SiteURL&""=empty then SiteURL=AutoSiteURL
%>
<SCRIPT>
function VerifyInput()
{

if (document.form.SiteName.value == "")
{alert("������վ������");
document.form.SiteName.focus();
return false;
}

if (document.form.DefaultSiteStyle.value == "")
{alert("Ĭ�Ϸ����û������");
document.form.DefaultSiteStyle.focus();
return false;
}


}
</SCRIPT>
[<b><a href="#��ͨ����">��ͨ����</a></b>]
[<b><a href="#��ҳ��ʾ">��ҳ��ʾ</a></b>]
[<b><a href="#ͶƱ����">ͶƱ����</a></b>]
[<b><a href="#�ϴ�����">�ϴ�����</a></b>]
[<b><a href="#ˮӡ����">ˮӡ����</a></b>]
[<b><a href="#SNMP����">SNMP����</a></b>]
<br>
[<b><a href="#��������">��������</a></b>]
[<b><a href="#�û�ѡ��">�û�ѡ��</a></b>]
[<b><a href="#��̳ѡ��">��̳ѡ��</a></b>]
[<b><a href="#��������">��������</a></b>]
[<b><a href="#��֤������">��֤������</a></b>]
[<b><a href="#��Һ;�������">��Һ;�������</a></b>]
<form method="POST" action="?menu=SiteSettingsUp" onsubmit="return VerifyInput();" name=form><a name="��ͨ����"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>��ͨ���� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
		<tr>
			<td width="45%" class="a3"><b>վ������<br>
			</b>��̳��Ӧ�ó������ƣ�����վ����ʾ��</td>
			<td class="a4" width="57%">
			<input size="30" name="SiteName" value="<%=SiteSettings("SiteName")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>վ��URL</b>
<script>
if ('<%=AutoSiteURL%>' != '<%=SiteURL%>'){
document.write("��վ��URL�������ô���<br>ϵͳ�Զ���⵽��URL��<font color=FF0000><%=AutoSiteURL%></font>");
}
</script>
			</td>
			<td class="a4" width="57%">
			<input size="30" name="SiteURL" value="<%=SiteURL%>"></td>
		</tr>
		<tr>
			<td class="a3" valign="top" width="45%"><b>META��ǩ-������Ϣ</b></td>
			<td class="a4" width="57%">
			<input size="40" name="MetaDescription" value="<%=SiteSettings("MetaDescription")%>"></td>
		</tr>
		<tr>
			<td class="a3" valign="top" width="45%"><b>META��ǩ-�ؼ���</b></td>
			<td class="a4" width="57%">
			<input size="40" name="MetaKeywords" value="<%=SiteSettings("MetaKeywords")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>����/ҳ��</b><br>
			ÿ����ҳ����ʾ��������</td>
			<td class="a4" width="57%">
			<input value="<%=SiteSettings("ThreadsPerPage")%>" name="ThreadsPerPage" size="20"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>����/ҳ��</b><br>
			ÿ����ҳ������ʾ��������</td>
			<td class="a4" width="57%">
			<input value="<%=SiteSettings("PostsPerPage")%>" name="PostsPerPage" size="20"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>ͳ���û�����ʱ�䣨���ӣ�<br>
			</b>Ĭ��:20</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("UserOnlineTime")%>" name="UserOnlineTime"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>�ű���ʱʱ�䣨�룩<br>
			</b>Ĭ��:60</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("Timeout")%>" name="Timeout"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>Ĭ�Ϸ������</b><br>
			ָ��<font color="#FF0000">images/skins/</font>Ŀ¼�µ�Ŀ¼������ [Ĭ��:1]</td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("DefaultSiteStyle")%>" name="DefaultSiteStyle"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>��̳ʹ�õĻ�������</b></td>
			<td class="a4" width="57%">
			<input size="20" value="<%=SiteSettings("CacheName")%>" name="CacheName"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>��˾/��֯����</b></td>
			<td class="a4" width="57%">
			<input size="30" name="CompanyName" value="<%=SiteSettings("CompanyName")%>"></td>
		</tr>
		<tr>
			<td class="a3" width="45%"><b>��˾/��֯��ַ</b></td>
			<td class="a4" width="57%">
			<input size="30" value="<%=SiteSettings("CompanyURL")%>" name="CompanyURL"></td>
		</tr>
</table>
<br>
<a name="��ҳ��ʾ"></a><table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table1">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>��ҳ��ʾ [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��ʾ�����û�</b><br>
		����ҳ����ʾ������ͳ�ơ����� </td>
		<td class="a4">
		<input type="radio" name="DisplayWhoIsOnline" value="1" <%if sitesettings("displaywhoisonline")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="DisplayWhoIsOnline" value="0" <%if sitesettings("displaywhoisonline")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��ʾ��������</b><br>
		����ҳ����ʾ���������ӡ����� </td>
		<td class="a4">
		<input type="radio" name="DisplayLink" value="1" <%if sitesettings("displaylink")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="DisplayLink" value="0" <%if sitesettings("displaylink")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��ʾ��̳��ʽ</b></td>
		<td class="a4">
		<input type="radio" name="DisplayForumFloor" value="1" <%if sitesettings("DisplayForumFloor")=1 then%>checked<%end if%>>��� 
		<input type="radio" name="DisplayForumFloor" value="0" <%if sitesettings("DisplayForumFloor")=0 then%>checked<%end if%>>��ϸ
	</tr>
</table>
<br>

<a name="ͶƱ����"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table5">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>ͶƱ���� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ͶƱѡ������</b><br>
		ͶƱѡ�����ٲ��õ��ڸ���ֵ </td>
		<td class="a4">
		<input size="20" name="MinVoteOptions" value="<%=SiteSettings("MinVoteOptions")%>"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ͶƱѡ�����</b><br>
		ͶƱѡ����󲻵ø��ڸ���ֵ </td>
		<td class="a4">
		<input size="20" name="MaxVoteOptions" value="<%=SiteSettings("MaxVoteOptions")%>"></td>
	</tr>
</table>


<br>
<a name="�ϴ�����"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table3">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>�ϴ����� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ѡ���ϴ����</b></td>
		<td class="a4"><select name="UpFileOption">
		<option value="">�ر�</option>
		<option value="ADODB.Stream" <%if sitesettings("upfileoption")="ADODB.Stream" then%>selected<%end if%>>ADODB.Stream</option>
		<option value="SoftArtisans.FileUp" <%if sitesettings("upfileoption")="SoftArtisans.FileUp" then%>selected<%end if%>>
		SoftArtisans.FileUp</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��������ѡ��</b><br>
				�����Ƿ񱣴浽���ݿ�</td>
		<td class="a4">
		<input type="radio" name="AttachmentsSaveOption" value="1" <%if sitesettings("AttachmentsSaveOption")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="AttachmentsSaveOption" value="0" <%if sitesettings("AttachmentsSaveOption")=0 then%>checked<%end if%>>�� </td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�����ϴ��ĸ������ͣ����á�|��������</b><br>
		���磺gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("UpFileTypes")%>" name="UpFileTypes"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>���û�Ա�����е�������� (byte)</b><br>
		Ĭ��:10240000</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxPostAttachmentsSize")%>" name="MaxPostAttachmentsSize">������ֻ���ڱ��������ݿ��еĸ�����Ч��</td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>ÿ���ϴ������Ĵ�С (byte)</b><br>
		Ĭ��:1024000</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxFileSize")%>" name="MaxFileSize"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>������Ƭ�ļ��Ĵ�С (byte)</b><br>
		Ĭ��:102400</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxPhotoSize")%>" name="MaxPhotoSize"></td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>����ͷ���ļ��Ĵ�С (byte)</b><br>
		Ĭ��:10240</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("MaxFaceSize")%>" name="MaxFaceSize"></td>
	</tr>
</table>
<br>
<a name="ˮӡ����"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table6">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>ˮӡ���� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	
	<tr>
		<td width="45%" class="a3"><b>ˮӡͼƬ���</b></td>
		<td class="a4"><select name="WatermarkOption" size="1">
		<option value="">�ر�</option>
		<option value="Persits.Jpeg" <%if sitesettings("WatermarkOption")="Persits.Jpeg" then%>selected<%end if%>>Persits.Jpeg</option>
		</select>������ֻ���ڱ��������ݿ��еĸ�����Ч��</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ˮӡЧ��</b></td>
		<td class="a4"><select name="WatermarkType" size="1">
		<option value="0" <%if sitesettings("WatermarkType")="0" then%>selected<%end if%>>ˮӡ����</option>
		<option value="1" <%if sitesettings("WatermarkType")="1" then%>selected<%end if%>>ˮӡͼƬ</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ˮӡ����</b><br>
		�����ֽ���ʾ���û��ϴ���ͼƬ��(ֻ֧��JPEG)</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("WatermarkText")%>" name="WatermarkText"></td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>ˮӡͼƬ�����·��</b><br>
		��ͼƬ����ʾ���û��ϴ���ͼƬ��(ֻ֧��JPEG)</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("WatermarkImage")%>" name="WatermarkImage"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ˮӡͼƬ��ʾλ��</b></td>
		<td class="a4"><select name="WatermarkPosition" size="1">
		<option value="0" <%if sitesettings("WatermarkPosition")="0" then%>selected<%end if%>>����</option>
		<option value="1" <%if sitesettings("WatermarkPosition")="1" then%>selected<%end if%>>����</option>
		<option value="2" <%if sitesettings("WatermarkPosition")="2" then%>selected<%end if%>>����</option>
		<option value="3" <%if sitesettings("WatermarkPosition")="3" then%>selected<%end if%>>����</option>
		<option value="4" <%if sitesettings("WatermarkPosition")="4" then%>selected<%end if%>>����</option>
		
		</select></td>
	</tr></table>
<br>

<a name="SNMP����"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table6">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>SNMP���� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�����ʼ����</b></td>
		<td class="a4"><select name="SelectMailMode" size="1">
		<option value="">�ر�</option>
		<option value="JMail.Message" <%if sitesettings("selectmailmode")="JMail.Message" then%>selected<%end if%>>
		JMail.Message</option>
		<option value="CDONTS.NewMail" <%if sitesettings("selectmailmode")="CDONTS.NewMail" then%>selected<%end if%>>
		CDONTS.NewMail</option>
		</select></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>������Email��ַ</b></td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerMail")%>" name="SmtpServerMail"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP ������</b><br>
		SMTP ����������Ϊ�����̳�����ʼ�</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServer")%>" name="SmtpServer"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP ������ע������</b><br>
		������¼SMTP��������ע������</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerUserName")%>" name="SmtpServerUserName"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>SMTP ����������</b><br>
		������¼SMTP������������</td>
		<td class="a4">
		<input size="30" value="<%=SiteSettings("SmtpServerPassword")%>" name="SmtpServerPassword"></td>
	</tr>
</table>
<br>
<a name="��������"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table7">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>�������� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����HTML��ǩ�����á�|��������</b><br>
		���磺iframe|SCRIPT|form|style|div|object|TEXTAREA</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedHtmlLabel")%>" name="BannedHtmlLabel"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����HTML��ǩ�ڵ��¼������á�|��������</b><br>
		���磺javascript:|Document.|onerror|onload|onmouseover</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedHtmlEvent")%>" name="BannedHtmlEvent"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�������дʣ����á�|��������</b><br>
		���磺fuck|shit|����</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedText")%>" name="BannedText"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�����û����ӣ����á�|��������</b><br>
		���磺yuzi|ԣԣ</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedUserPost")%>" name="BannedUserPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��ֹע����û��������á�|��������</b><br>
		���磺yuzi|ԣԣ</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedRegUserName")%>" name="BannedRegUserName"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��ֹIP��ַ������̳�����á�|��������</b><br>
		���磺127.0.0.1|192.168.0.1</td>
		<td class="a4">
		<input size="40" value="<%=SiteSettings("BannedIP")%>" name="BannedIP"></td>
	</tr>
</table>
<br>
<a name="�û�ѡ��"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table8">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>�û�ѡ�� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�˺ż��ʽ</b><br>
		�����û�ע����˺ż���ķ�ʽ�� 
		&quot;�Զ�&quot;��ʾ�����û������Լ����˺ţ�&quot;Email����&quot;��ʾ��ע���û��󣬻�ͨ���ʼ���ʽ�����뷢�͸����û���&quot;����Ա����&quot;��ʾ�˺�ע�����Ҫ����Ա������</td>
		<td class="a4">
		<input type="radio" name="EnableUser" value="0" <%if sitesettings("EnableUser")=0 then%>checked<%end if%>>�Զ� 
		<input type="radio" name="EnableUser" value="1" <%if sitesettings("EnableUser")=1 then%>checked<%end if%>>Email���� 
		<input type="radio" name="EnableUser" value="2" <%if sitesettings("EnableUser")=2 then%>checked<%end if%>>����Ա����</td>
	</tr>	<tr>
		<td width="45%" class="a3"><b>�ر��û�ע��</b></td>
		<td class="a4">
		<input type="radio" name="CloseRegUser" value="1" <%if sitesettings("closereguser")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="CloseRegUser" value="0" <%if sitesettings("closereguser")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>һ��Emailֻ��ע��һ���ʺ�</b></td>
		<td class="a4">
		<input type="radio" name="OnlyMailReg" value="1" <%if sitesettings("onlymailreg")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="OnlyMailReg" value="0" <%if sitesettings("onlymailreg")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	</table>
<br>
<a name="��̳ѡ��"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>��̳ѡ�� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����������̳����</b></td>
		<td class="a4">
		<input type="radio" name="ForumApply" value="1" <%if sitesettings("forumapply")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="ForumApply" value="0" <%if sitesettings("forumapply")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�����̳��������̳ӵ��ͬ������</b><br>
		�磺���������</td>
		<td class="a4">
		<input type="radio" name="SortShowForum" value="1" <%if sitesettings("sortshowforum")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="SortShowForum" value="0" <%if sitesettings("sortshowforum")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
</table>
<br>
<a name="��������"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table2">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>�������� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����������룩<br>
		</b>ͬһ���û�2�η����ļ��</td>
		<td class="a4">
		<input size="20" name="DuplicatePostIntervalInMinutes" value="<%=SiteSettings("DuplicatePostIntervalInMinutes")%>"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ע���೤ʱ����ܷ������ӣ��룩</b></td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("RegUserTimePost")%>" name="RegUserTimePost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�������ӵĻظ���</b><br>
		һ��������Ҫ����ƪ�������ܱ������? </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("PopularPostThresholdPosts")%>" name="PopularPostThresholdPosts"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�������ӵ������</b><br>
		һ��������Ҫ���ٴ�������ܱ������? </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("PopularPostThresholdViews")%>" name="PopularPostThresholdViews"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>��ʾ�༭��¼</b><br>
		���ú�����Ϣ�ĵײ�(λ���û�ǩ��ǰ)������ʾ�����ӵı༭��¼ ��</td>
		<td class="a4">
		<input type="radio" name="DisplayEditNotes" value="1" <%if sitesettings("DisplayEditNotes")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="DisplayEditNotes" value="0" <%if sitesettings("DisplayEditNotes")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>������ʾ�������ߵ�IP��ַ<br>
		</b>��ʾ�������ߵ�IP��ַ��</td>
		<td class="a4">
		<input type="radio" name="DisplayPostIP" value="1" <%if sitesettings("DisplayPostIP")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="DisplayPostIP" value="0" <%if sitesettings("DisplayPostIP")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
</table>
<br>

<a name="��֤������"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>��֤������ [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����ע����֤</b><br>
		���û�ע��ʱ������֤�빦�ܡ� </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForRegister")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForRegister")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>���õ�¼��֤</b><br>
		���û���¼ʱ������֤�빦�ܡ� </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForLogin")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForLogin")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>���÷�����֤</b><br>
		���û�����ʱ������֤�빦�ܡ� </td>
		<td class="a4">
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value="1" <%if sitesettings("EnableAntiSpamTextGenerateForPost")=1 then%>checked<%end if%>>�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value="0" <%if sitesettings("EnableAntiSpamTextGenerateForPost")=0 then%>checked<%end if%>>�� 
		</td>
	</tr>
</table>



<br>
<a name="��Һ;�������"></a>
<table cellspacing="1" cellpadding="5" width="100%" border="0" class="a2" align="center" id="table9">
	<tr height="25" id="TableTitleLink">
		<td class="a1" colspan="2"><b>��Һ;������� [<a href="#TOP">TOP</a>]</b></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��������</b><br>
		�û�ÿ��һ�����������ý�Һ;���</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddThread")%>" name="IntegralAddThread"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>����</b><br>
		�û�ÿ��һ���������ý�Һ;���</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddPost")%>" name="IntegralAddPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>�Ӿ���</b><br>
		����ĳ���Ӽ�Ϊ���������ӵķ��������ý�Һ;��� </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralAddValuedPost")%>" name="IntegralAddValuedPost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>��������ɾ��</b><br>
		����ɾ��ĳ�����������ӵķ��������ý�Һ;��� </td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeleteThread")%>" name="IntegralDeleteThread"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>������ɾ��</b><br>
		����ɾ��ĳ���������ӵķ��������ý�Һ;���</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeletePost")%>" name="IntegralDeletePost"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>������ȡ��</b><br>
		����ȡ��ĳ���������ӵķ��������ý�Һ;���</td>
		<td class="a4">
		<input size="20" value="<%=SiteSettings("IntegralDeleteValuedPost")%>" name="IntegralDeleteValuedPost"></td>
	</tr>
</table>
<table border="0" width="100%" align="center">
	<tr>
		<td align="right"><input type="submit" value=" �� �� ">
		<input type="reset" value=" �� �� "></form></td>
	</tr>
</table>
<%
end sub


sub SiteSettingsUp

if Request("EnableUser")="1" and Request("SelectMailMode")="" then error2("��ѡ����ע���û�����ͨ��Email���ͣ�������û���趨�����ʼ����")

filtrate=split(Request("BannedIP"),"|")
for i = 0 to ubound(filtrate)
if instr("|"&Request.ServerVariables("REMOTE_ADDR")&"","|"&filtrate(i)&"") > 0 then error2("�����������IP��ַ�Ƿ���ȷ")
next

if Request("UpFileOption")<>"" then
if Not IsObjInstalled(Request("UpFileOption")) then error2("����������֧�� "&Request("UpFileOption")&" �������رմ˹��ܣ�")
end if

if Request("SelectMailMode")<>"" then
if Not IsObjInstalled(Request("SelectMailMode")) then error2("����������֧�� "&Request("SelectMailMode")&" �������رմ˹��ܣ�")
end if

if Request("WatermarkOption")<>"" then
if Not IsObjInstalled(Request("WatermarkOption")) then error2("����������֧�� "&Request("WatermarkOption")&" �������رմ˹��ܣ�")
end if

Rs.Open "[BBSXP_SiteSettings]",Conn,1,3
for each ho in Request.Form
Rs(ho)=Request(ho)
next
Rs.update
Rs.close
%>
���³ɹ�<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub
sub Showbanner
%><form method="POST" action="?menu=SiteSettingsUp">
	<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
		<tr height="25">
			<td class="a1" align="middle" colspan="2"><b>����������</b></td>
		</tr>
		<tr>
			<td class="a3" align="middle" width="20%">����������<br>
			<font color="#FF0000">֧��HTML</font> </td>
			<td class="a3">
			<textarea name="TopBanner" rows="6" style="width:100%"><%=SiteSettings("TopBanner")%></textarea></td>
		</tr>
		<tr>
			<td class="a3" align="middle" width="20%">�ײ���Ȩ��Ϣ<br>
			<font color="#FF0000">֧��HTML</font> </td>
			<td class="a3">
			<textarea name="BottomAD" rows="6" style="width:100%"><%=SiteSettings("BottomAD")%></textarea></td>
		</tr>
	</table>
	<input type="submit" value=" �� �� ">		<input type="reset" value=" �� �� ">
</form>
<%
end sub


sub agreement
%><form method="POST" action="?menu=SiteSettingsUp">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle><b>ע���û�Э������</b></td>
  </tr>
    <tr>
    <td class=a3 align=middle>
<textarea name="RegUserAgreement" rows="18" style="width:100%"><%=SiteSettings("RegUserAgreement")%></textarea></td>
  </tr>
        
  </table>
<input type="submit" value=" �� �� ">		<input type="reset" value=" �� �� ">
</form>
<%
end sub
sub SiteStat

Set Statistics=Conn.Execute("[BBSXP_Statistics_Site]")

%><form method="POST" action="?menu=SiteStatok">
<table cellspacing="1" cellpadding="2" width="100%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2"><b>����ͳ����Ϣ</b></td>
  </tr>
	<tr>
		<td width="45%" class="a3"><b>ͳ��ע���û�</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalUser")%>" name="TotalUser"></td>
	</tr>
	<tr>
		<td width="45%" class="a3"><b>ͳ��������</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalThread")%>" name="TotalThread"></td>
	</tr>
        
	<tr>
		<td width="45%" class="a3"><b>ͳ��������</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TotalPost")%>" name="TotalPost"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>�����ܷ�����</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("TodayPost")%>" name="TodayPost"></td>
	</tr>

	<tr>
		<td width="45%" class="a3"><b>�����������</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("BestOnline")%>" name="BestOnline"></td>
	</tr>
        
	<tr>
		<td width="45%" class="a3"><b>���������������ʱ��</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("BestOnlineTime")%>" name="BestOnlineTime"></td>
	</tr>
	
	<tr>
		<td width="45%" class="a3"><b>����ע����û�</b></td>
		<td class="a4">
		<input size="20" value="<%=Statistics("NewUser")%>" name="NewUser"></td>
	</tr>
  </table>
<input type="submit" value=" �� �� ">		<input type="reset" value=" �� �� ">		
<input type="button" value=" ����ͳ�� " onclick="document.location.href='?menu=AutoSiteStat'">
</form>
<%
Set Statistics=Nothing

end sub

htmlend

%>