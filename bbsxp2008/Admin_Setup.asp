<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")

select case Request("menu")
	case "AdvancedSettings"
		AdvancedSettings
	case "variable"
		variable
	case "AdvancedSettingsUp"
		AdvancedSettingsUp
	case "SiteInfo"
		SiteInfo
end select

Sub variable
DomainPath=Left(Script_Name,inStrRev(Script_Name,"/"))
if SiteConfig("SiteUrl")="" then
	SiteUrl="http://"&Server_Name&DomainPath
	SiteUrl=left(SiteUrl,len(SiteUrl)-1)
else
	SiteUrl=SiteConfig("SiteUrl")
end if

%>
<script type="text/javascript" language="JavaScript">
function VerifyInput() {
	if (document.form.SiteName.value == ""){
		alert("������վ������");
		return false;
	}
	if (document.form.DefaultSiteStyle.value == ""){
		alert("Ĭ�Ϸ����û������");
		return false;
	}
	if (document.form.elements["APIEnable"][0].checked == true){
		if (document.form.elements["DefaultPasswordFormat"][0].checked == false){
			alert("������ϵͳ���ϣ��û�ע���µġ�Ĭ����������㷨��Ӧ����ΪMD5");
			return false;
		}
		if (document.form.APISafeKey.value == ""){
			alert("���������ϰ�ȫ��");
			return false;
		}
		if (document.form.APIUrlList.value == ""){
			alert("���������ϵĽӿ��ļ���ַ");
			return false;
		}
	}
}
</script>
<script type="text/javascript" src="Utility/selectColor.js"></script>

<form method="POST" action="?menu=SiteInfo" onsubmit="return VerifyInput();" name=form>

<div class="MenuTag"><li id="Tag_btn_0" onclick="ShowPannel(this)" class="NowTag">��ͨ����</li><li id="Tag_btn_1" onclick="ShowPannel(this)">��ҳ��ʾ</li><li id="Tag_btn_2" onclick="ShowPannel(this)">�û�ע��</li><li id="Tag_btn_3" onclick="ShowPannel(this)">��̳����</li><li id="Tag_btn_4" onclick="ShowPannel(this)">�ϴ�����</li><li id="Tag_btn_5" onclick="ShowPannel(this)">�ʼ�����</li><li id="Tag_btn_6" onclick="ShowPannel(this)">��������</li><li id="Tag_btn_7" onclick="ShowPannel(this)">��֤��</li><li id="Tag_btn_8" onclick="ShowPannel(this)">��Ǯ�뾭��</li><li id="Tag_btn_9" onclick="ShowPannel(this)">��������</li></div>


<table id="Tag_tab_0" style="display: block" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='��̳���ƣ���������̳����ҳ�����������ڱ�������ʾ��'><img src="images/spacer.gif" border="0" /></a><b>��̳����</b></td>
		<td><input size="30" name="SiteName" value="<%=SiteConfig("SiteName")%>" /></td>
	</tr>


	<tr>
		<td width="30%"><a class="helpicon" title='�����������̳����ַ��ע��: ��Ҫ��б�� (��/��) ��β��'><img src="images/spacer.gif" border="0" /></a><b>��̳��ַ</b></td>
		<td><input size="30" name="SiteUrl" value="<%=SiteUrl%>" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���������Աemail��ַ'><img src="images/spacer.gif" border="0" /></a><b>����Ա�����ʼ���ַ</b></td>
		<td><input size="30" value="<%=SiteConfig("WebMasterEmail")%>" name="WebMasterEmail" /></td>
	</tr>

	<tr>
		<td><a class="helpicon" title='�����빫˾/��֯����'><img src="images/spacer.gif" border="0" /></a><b>��˾/��֯����</b></td>
		<td><input size="30" name="CompanyName" value="<%=SiteConfig("CompanyName")%>" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����빫˾/��֯��ַ'><img src="images/spacer.gif" border="0" /></a><b>��˾/��֯��ַ</b></td>
		<td><input size="30" value="<%=SiteConfig("CompanyURL")%>" name="CompanyURL" /></td>
	</tr>
	

	<tr>
		<td><a class="helpicon" title='�����̵���ҳ˵���������վ��������������վ����ά�����������ݣ��������������п�����˵����'><img src="images/spacer.gif" border="0" /></a><b>META��ǩ-������Ϣ</b></td>
		<td><input size="60" name="MetaDescription" value="<%=SiteConfig("MetaDescription")%>" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����������ҳ���ݵ�һ�������ؼ��֣���һ�����ź�һ���ո�ָ�����һЩ Internet ��������ʹ����Щ�ؼ�����ȷ����ҳ�Ƿ�����������ѯ������'><img src="images/spacer.gif" border="0" /></a><b>META��ǩ-�ؼ���</b></td>
		<td><input size="60" name="MetaKeywords" value="<%=SiteConfig("MetaKeywords")%>" /></td>
	</tr>
	
	
	<tr>
		<td><a class="helpicon" title='Ĭ��:20'><img src="images/spacer.gif" border="0" /></a><b>ͳ���û�����ʱ�䣨���ӣ�</b></td>
		<td><input size="20" value="<%=SiteConfig("UserOnlineTime")%>" name="UserOnlineTime" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ��:60'><img src="images/spacer.gif" border="0" /></a><b>�ű���ʱʱ�䣨�룩</b></td>
		<td><input size="20" value="<%=SiteConfig("Timeout")%>" name="Timeout" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ѡ����̳��Ĭ�Ϸ��'><img src="images/spacer.gif" border="0" /></a><b>�������</b></td>
		<td>
        <select name="DefaultSiteStyle">
        <%
		XmlFilePath=Server.MapPath("Xml/Themes.xml")
		Set XMLDOM=Server.CreateObject("Microsoft.XMLDOM")
		XMLDOM.load(XmlFilePath)
		Set XMLRoot = XMLDOM.documentElement
		For NodeIndex=0 To XMLRoot.childNodes.length-1
			Set ParentNode=XMLRoot.childNodes(NodeIndex)
		%>
        	<option value="<%=ParentNode.getAttribute("Name")%>"<%if SiteConfig("DefaultSiteStyle")=ParentNode.getAttribute("Name") then Response.Write(" selected")%>><%=ParentNode.text%></option>
        <%
		Next
		Set ParentNode=nothing
		Set XMLRoot=nothing
		Set XMLDOM=nothing
		%>
        </select>
        </td>
	</tr>

	<tr>
		<td><a class="helpicon" title='Ĭ��:BBSXP'><img src="images/spacer.gif" border="0" /></a><b>վ�㻺������</b></td>
		<td><input size="20" value="<%=SiteConfig("CacheName")%>" name="CacheName" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ��:5����'><img src="images/spacer.gif" border="0" /></a><b>������¼�������ӣ�</b></td>
		<td><input size="20" value="<%=SiteConfig("CacheUpDateInterval")%>" name="CacheUpDateInterval" /></td>
	</tr>
	

	<tr>
		<td width="40%"><a class="helpicon" title='Cookie �����·�����������ͬһ�������������˶����̳������Ҫ��������Ϊÿ����̳���ڵ�Ŀ¼����������д / ������ˡ�'><img src="images/spacer.gif" border="0" /></a><b>Cookies ����·��</b></td>
		<td>
		<select name="CookiePath" size="1">
		<option value="/">/</option>
		<option value="<%=DomainPath%>" <%if SiteConfig("CookiePath")=DomainPath then%>selected<%end if%>><%=DomainPath%></option>
		</select></td>
	</tr>
	<tr>
		<td width="40%"><a class="helpicon" title='Cookie ��Ӱ����������޸Ĵ�ѡ��Ĭ��ֵ�����ԭ���ǣ�������̳��������ͬ����ַ������ bbsxp.com �� forums.bbsxp.com��Ҫʹ�û�����������ͬ������������̳ʱ�����ܱ��ֵ�¼״̬������Ҫ����ѡ������Ϊ .bbsxp.com (ע��������Ҫ�Ե㿪ͷ)��

����û�а��գ����������ʲôҲ����д����Ϊ��������ûᵼ�����޷���¼��̳'><img src="images/spacer.gif" border="0" /></a><b>Cookies ����</b></td>
		<td><input size="50" value="<%=SiteConfig("CookieDomain")%>" name="CookieDomain" /></td>
	</tr>


	<tr>
		<td><a class="helpicon" title='ά���ڼ�����ùر���̳'><img src="images/spacer.gif" border="0" /></a><b>��̳״̬</b></td>
		<td><input type="radio" name="SiteDisabled" value=1 <%if SiteConfig("SiteDisabled")=1 then%>checked<%end if%> />�ر� <input type="radio" name="SiteDisabled" value=0 <%if SiteConfig("SiteDisabled")=0 then%>checked<%end if%> />��</td>
	</tr>
	
	
	<tr>
		<td><a class="helpicon" title='��̳�ر�ʱ���ֵ���ʾ��Ϣ'><img src="images/spacer.gif" border="0" /></a><b>��̳�رյ�ԭ��</b></td>
		<td><input size="50" value="<%=SiteConfig("SiteDisabledReason")%>" name="SiteDisabledReason" /></td>
	</tr>



	<tr>
		<td width="30%"><a class="helpicon" title='ѡ���ǡ���ʹ��̳���롰no-cache���� HTTP ͷ��Ϣ���ͻ��˲��ٱ���ҳ�滺�档�������ҳ���������Ӷ����·������������ӡ�'><img src="images/spacer.gif" border="0" /></a><b>��ӡ�No-cache��HTTP ͷ��Ϣ</b></td>
		<td>
		<input type="radio" name="NoCacheHeaders" value=1 <%if SiteConfig("NoCacheHeaders")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="NoCacheHeaders" value=0 <%if SiteConfig("NoCacheHeaders")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>	
</table>

<!--��ҳ��ʾ start-->
<table id="Tag_tab_1" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='����ҳ����ʾ�����ű�ǩ������'><img src="images/spacer.gif" border="0" /></a><b>��ʾ��ǩ</b></td>
		<td>
		<input type="radio" name="DisplayTags" value=1 <%if SiteConfig("DisplayTags")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayTags" value=0 <%if SiteConfig("DisplayTags")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td width="30%"><a class="helpicon" title='����ҳ����ʾ���û�������Ϣ������'><img src="images/spacer.gif" border="0" /></a><b>��ʾ�����û�</b></td>
		<td>
		<input type="radio" name="DisplayWhoIsOnline" value=1 <%if SiteConfig("DisplayWhoIsOnline")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayWhoIsOnline" value=0 <%if SiteConfig("DisplayWhoIsOnline")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>	
	<tr>
		<td><a class="helpicon" title='����ҳ����ʾ��ͳ����Ϣ������'><img src="images/spacer.gif" border="0" /></a><b>��ʾͳ����Ϣ</b></td>
		<td>
		<input type="radio" name="DisplayStatistics" value=1 <%if SiteConfig("DisplayStatistics")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayStatistics" value=0 <%if SiteConfig("DisplayStatistics")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ҳ����ʾ����������յ��û�������'><img src="images/spacer.gif" border="0" /></a><b>��ʾ��������հ�</b></td>
		<td>
		<input type="radio" name="DisplayBirthdays" value=1 <%if SiteConfig("DisplayBirthdays")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayBirthdays" value=0 <%if SiteConfig("DisplayBirthdays")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ҳ����ʾ���������ӡ�����'><img src="images/spacer.gif" border="0" /></a><b>��ʾ��������</b></td>
		<td>
		<input type="radio" name="DisplayLink" value=1 <%if SiteConfig("DisplayLink")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayLink" value=0 <%if SiteConfig("DisplayLink")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�ڰ���б�����ʾ�Ӱ��'><img src="images/spacer.gif" border="0" /></a><b>����б�����ʾ�Ӱ��</b></td>
		<td>
		<input type="radio" name="IsShowSonForum" value=1 <%if SiteConfig("IsShowSonForum")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="IsShowSonForum" value=0 <%if SiteConfig("IsShowSonForum")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
</table>
<!--��ҳ��ʾ end-->


<!--�ϴ����� start-->
<table id="Tag_tab_4" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>�ϴ����ѡ��</b></font></td></tr>
	<tr>
		<td width="40%"><a class="helpicon" title='���� ����֧�ֵ���� ѡ���ϴ����'><img src="images/spacer.gif" border="0" /></a><b>ѡ���ϴ����</b></td>
		<td>
		<select name="UpFileOption" size="1" onchange="IsObjInstalled('IsUpFileInstalled',this.options[this.selectedIndex].value)">
		<option value="">�ر�</option>
		<option value="ADODB.Stream" <%if SiteConfig("UpFileOption")="ADODB.Stream" then%>selected<%end if%>>ADODB.Stream</option>
		<option value="Persits.Upload" <%if SiteConfig("UpFileOption")="Persits.Upload" then%>selected<%end if%>>AspUpload</option>
		<option value="SoftArtisans.FileUp" <%if SiteConfig("UpFileOption")="SoftArtisans.FileUp" then%>selected<%end if%>>SoftArtisans.FileUp</option>
		</select> <span id="IsUpFileInstalled"></span></td>
	</tr>
	<tr>
		<td width="40%"><a class="helpicon" title='�����Ǳ��������ݿ��л���ֱ�ӱ�����Ӳ����'><img src="images/spacer.gif" border="0" /></a><b>��������ģʽ</b></td>
		<td>
        <select name="AttachmentsSaveOption">
		<option value="0" <%if SiteConfig("AttachmentsSaveOption")=0 then%>selected<%end if%>>���ݿ�</option>
		<option value="1" <%if SiteConfig("AttachmentsSaveOption")=1 then%>selected<%end if%>>Ӳ��</option>
		</select>
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ��:100 KB'><img src="images/spacer.gif" border="0" /></a><b>������ͷ���ļ��Ĵ�С (KB)</b></td>
		<td><input size="20" value="<%=SiteConfig("MaxFaceSize")%>" name="MaxFaceSize" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ��:200 KB'><img src="images/spacer.gif" border="0" /></a><b>���������Ӹ����Ĵ�С (KB)</b></td>
		<td><input size="20" value="<%=SiteConfig("MaxFileSize")%>" name="MaxFileSize" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ��:10240 KB��10M��'><img src="images/spacer.gif" border="0" /></a><b>�����û��ϴ��ļ��е�������� (KB)</b></td>
		<td><input size="20" value="<%=SiteConfig("MaxPostAttachmentsSize")%>" name="MaxPostAttachmentsSize" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���磺gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar'><img src="images/spacer.gif" border="0" /></a><b>�����ϴ��ĸ������ͣ����á�|���ָ���</b></td>
		<td><input size="60" value="<%=SiteConfig("UpFileTypes")%>" name="UpFileTypes" /></td>
	</tr>	
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ˮӡ���ѡ��</b></font></td></tr>
	<tr>
		<td width="30%"><a class="helpicon" title='ˮӡ������ã��ر�/����'><img src="images/spacer.gif" border="0" /></a><b>ˮӡͼƬ���</b></td>
		<td>
		<select name="WatermarkOption" size="1" onchange="IsObjInstalled('IsWatermarkInstalled',this.options[this.selectedIndex].value)">
		<option value="">�ر�</option>
		<option value="Persits.Jpeg" <%if SiteConfig("WatermarkOption")="Persits.Jpeg" then%>selected<%end if%>>AspJpeg</option>
		</select> <span id="IsWatermarkInstalled"></span>
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ѡ��ˮӡЧ����ֻ֧��JPEG��'><img src="images/spacer.gif" border="0" /></a><b>ˮӡЧ��</b></td>
		<td><select name="WatermarkType" size="1">
		<option value="0" <%if SiteConfig("WatermarkType")="0" then%>selected<%end if%>>ˮӡ����</option>
		<option value="1" <%if SiteConfig("WatermarkType")="1" then%>selected<%end if%>>ˮӡͼƬ</option>
		</select></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='ˮӡˮƽλ��'><img src="images/spacer.gif" border="0" /></a><b>ˮӡˮƽλ��</b></td>
		<td>
		<input type="radio" name="WatermarkWidthPosition" value="left" <%if SiteConfig("WatermarkWidthPosition")="left" then%>checked<%end if%> />��
		<input type="radio" name="WatermarkWidthPosition" value="center" <%if SiteConfig("WatermarkWidthPosition")="center" then%>checked<%end if%> />��
		<input type="radio" name="WatermarkWidthPosition" value="right" <%if SiteConfig("WatermarkWidthPosition")="right" then%>checked<%end if%> />��
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='ˮӡ��ֱλ�ã�����ֻ���ͼƬ��'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ��ֱλ��</b></td>
		<td>
		<input type="radio" name="WatermarkHeightPosition" value="top" <%if SiteConfig("WatermarkHeightPosition")="top" then%>checked<%end if%> />��
		<input type="radio" name="WatermarkHeightPosition" value="center" <%if SiteConfig("WatermarkHeightPosition")="center" then%>checked<%end if%> />��
		<input type="radio" name="WatermarkHeightPosition" value="bottom" <%if SiteConfig("WatermarkHeightPosition")="bottom" then%>checked<%end if%> />��
		</td>
	</tr>
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ˮӡ����</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='�����ֽ���ʾ���û��ϴ���ͼƬ��'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ����</b></td>
		<td><input size="30" value="<%=SiteConfig("WatermarkText")%>" name="WatermarkText" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��������ˮӡ������������'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ��������</b></td>
		<td>
		<select name="WatermarkFontFamily">
		<option value="<%=SiteConfig("WatermarkFontFamily")%>" selected="selected"><%=SiteConfig("WatermarkFontFamily")%></option>
		<option value="����">����</option>
		<option value="����">����</option>
		<option value="Arial">Arial</option> 
		<option value="Arial Black">Arial Black</option> 
		<option value="Times New Roman" >Times New Roman</option>
		<option value="Garamond">Garamond</option>
		<option value="Lucida Handwriting">Lucida Handwriting</option>
		<option value="Verdana">Verdana</option>
		</select>
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ˮӡ���������С'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ���ִ�С</b></td>
		<td><input size="2" value="<%=SiteConfig("WatermarkFontSize")%>" name="WatermarkFontSize" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ˮӡ����������ɫ'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ������ɫ</b></td>
		<td><img title="���ѡȡ��ɫ!" id="WatermarkFontColor" onclick="OpenColorPicker('WatermarkFontColor', event)" style="CURSOR:pointer;BACKGROUND-COLOR:<%=SiteConfig("WatermarkFontColor")%>" src="images/rect.gif" /><input type="hidden" value="<%=SiteConfig("WatermarkFontColor")%>" name="WatermarkFontColor" />
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ˮӡ�����Ƿ���ʾ����'><img src="images/spacer.gif" border="0" /></a><b>ˮӡ�����Ƿ����</b></td>
		<td>
		<input type="radio" name="WatermarkFontIsBold" value=1 <%if SiteConfig("WatermarkFontIsBold")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="WatermarkFontIsBold" value=0 <%if SiteConfig("WatermarkFontIsBold")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>

	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ˮӡͼƬ</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='��ͼƬ����ʾ���û��ϴ���ͼƬ��'><img src="images/spacer.gif" border="0" /></a><b>ˮӡͼƬ�����·��</b></td>
		<td><input size="30" value="<%=SiteConfig("WatermarkImage")%>" name="WatermarkImage" /></td>
	</tr>	
</table>
<!--�ϴ����� End-->


<!--�ʼ����� start-->
<table id="Tag_tab_5" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='�����ʼ�������ã�����/�ر�'><img src="images/spacer.gif" border="0" /></a><b>�����ʼ����</b></td>
		<td>
		<select name="SelectMailMode" size="1" onchange="IsObjInstalled('IsMailObjInstalled',this.options[this.selectedIndex].value)">
		<option value="">�ر�</option>
		<option value="CDO.Message" <%if SiteConfig("SelectMailMode")="CDO.Message" then%>selected<%end if%>>CDO.Message</option>
		<option value="Persits.MailSender" <%if SiteConfig("SelectMailMode")="Persits.MailSender" then%>selected<%end if%>>AspEmail</option>
		<option value="JMail.Message" <%if SiteConfig("SelectMailMode")="JMail.Message" then%>selected<%end if%>>JMail</option>

		</select> <span id=IsMailObjInstalled></span>
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�������͵����ʼ��ĵ�ַ'><img src="images/spacer.gif" border="0" /></a><b>������Email��ַ</b></td>
		<td><input size="30" value="<%=SiteConfig("SmtpServerMail")%>" name="SmtpServerMail" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����������� SMTP �����ʼ�����������ָ�� SMTP ��������һ����������������� IP ��ַ������������������졣'><img src="images/spacer.gif" border="0" /></a><b>SMTP ������</b></td>
		<td><input size="30" value="<%=SiteConfig("SmtpServer")%>" name="SmtpServer" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����������� SMTP �����ʼ������ҷ�������Ҫ��֤���������������û�����'><img src="images/spacer.gif" border="0" /></a><b>SMTP �û���</b></td>
		<td><input size="30" value="<%=SiteConfig("SmtpServerUserName")%>" name="SmtpServerUserName" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����������� SMTP �����ʼ������ҷ�������Ҫ��֤�����������������롣'><img src="images/spacer.gif" border="0" /></a><b>SMTP ����</b></td>
		<td><input type=password size="30" value="<%=SiteConfig("SmtpServerPassword")%>" name="SmtpServerPassword" /></td>
	</tr>
</table>
<!--�ʼ����� End-->


<!--�������� Start-->
<table id="Tag_tab_6" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td><a class="helpicon" title='���磺fuck|shit|����'><img src="images/spacer.gif" border="0" /></a><b>���˷���ʱ���е����дʣ����á�|���ָ���</b></td>
		<td><input size="60" value="<%=SiteConfig("BannedText")%>" name="BannedText" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���磺fuck|shit'><img src="images/spacer.gif" border="0" /></a><b>��ֹע����ַ������á�|���ָ���</b></td>
		<td><input size="60" value="<%=SiteConfig("BannedRegUserName")%>" name="BannedRegUserName" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���磺127.0.0.1|192.168.0.1'><img src="images/spacer.gif" border="0" /></a><b>��ֹIP��ַ������̳�����á�|���ָ���</b></td>
		<td><input size="60" value="<%=SiteConfig("BannedIP")%>" name="BannedIP" /></td>
	</tr>
</table>
<!--�������� End-->

<!--�û�ע�� Start-->
<table id="Tag_tab_2" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>��ע������</b></font></td></tr>
	<tr>
		<td width="30%"><a class="helpicon" title='��ֹ��¼ʱ��������Ա���Ե�¼վ�㡣'><img src="images/spacer.gif" border="0" /></a><b>�����¼</b></td>
		<td>
		<input type="radio" name="AllowLogin" value=1 <%if SiteConfig("AllowLogin")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowLogin" value=0 <%if SiteConfig("AllowLogin")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����Ƿ����������ʺŵ��û���¼����������ͼ����֮ǰ����ע�⵽�Լ��ʺ��ѱ����á�'><img src="images/spacer.gif" border="0" /></a><b>���������ʺŵ��û���¼</b></td>
		<td>
		<input type="radio" name="EnableBannedUsersToLogin" value=1 <%if SiteConfig("EnableBannedUsersToLogin")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableBannedUsersToLogin" value=0 <%if SiteConfig("EnableBannedUsersToLogin")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ֹ�����û�������ע�ᡣ'><img src="images/spacer.gif" border="0" /></a><b>�������û�ע��</b></td>
		<td>
		<input type="radio" name="AllowNewUserRegistration" value=1 <%if SiteConfig("AllowNewUserRegistration")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowNewUserRegistration" value=0 <%if SiteConfig("AllowNewUserRegistration")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���ڿ�����ע����û��� ������Ҫ��˵İ�� �������������Ƿ���Ҫ���'><img src="images/spacer.gif" border="0" /></a><b>���û����εȼ�</b></td>
		<td>
		<select name="NewUserModerationLevel">
			<option value="1" <%if SiteConfig("NewUserModerationLevel")=1 then%>selected<%end if%>>�����û�</option>
			<option value="0" <%if SiteConfig("NewUserModerationLevel")=0 then%>selected<%end if%>>�������û�</option>
		</select> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�û������ַ���������ڻ��ߵ��ڸ���ֵ'><img src="images/spacer.gif" border="0" /></a><b>�û�����С����</b></td>
		<td><input type="text" name="UserNameMinLength" value="<%=SiteConfig("UserNameMinLength")%>" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�û������ַ�������С�ڻ��ߵ��ڸ���ֵ'><img src="images/spacer.gif" border="0" /></a><b>�û�����󳤶�</b></td>
		<td><input type="text" name="UserNameMaxLength" value="<%=SiteConfig("UserNameMaxLength")%>" /></td>
	</tr>
	
	<tr>
		<td><a class="helpicon" title='�����û�ע����ʺż���ķ�ʽ�� &quot;�Զ�&quot;��ʾ�����û������Լ����ʺţ�&quot;Email&quot;��ʾ��ע���û��󣬻�ͨ���ʼ���ʽ�����뷢�͸����û���&quot;ֻ��������;��ʾ��ע���û�����ӵ����ע���û����͵����������ע�᣻&quot;����Ա����&quot;��ʾ�ʺ�ע�����Ҫ����Ա������'><img src="images/spacer.gif" border="0" /></a><b>�ʺż��ʽ</b></td>
		<td>
		<input type="radio" name="AccountActivation" value="0" <%if SiteConfig("AccountActivation")=0 then%>checked<%end if%> />�Զ� 
		<input type="radio" name="AccountActivation" value="1" <%if SiteConfig("AccountActivation")=1 then%>checked<%end if%> />Email  
		<input type="radio" name="AccountActivation" value="2" <%if SiteConfig("AccountActivation")=2 then%>checked<%end if%> />�������� 
		<input type="radio" name="AccountActivation" value="3" <%if SiteConfig("AccountActivation")=3 then%>checked<%end if%> />����Ա����</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�û�����ļ����㷨������ѡ��MD5��SHA1'><img src="images/spacer.gif" border="0" /></a><b>Ĭ����������㷨</b></td>
		<td>
		<input type="radio" name="DefaultPasswordFormat" value="MD5" <%if SiteConfig("DefaultPasswordFormat")="MD5" then%>checked<%end if%> />MD5 
		<input type="radio" name="DefaultPasswordFormat" value="SHA1" <%if SiteConfig("DefaultPasswordFormat")="SHA1" then%>checked<%end if%> />SHA1 
</td>
	</tr>
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ȫ���û���������</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='�����û�������ǩ����ʾ���û�����ÿ������ĩ�ˡ�'><img src="images/spacer.gif" border="0" /></a><b>���ø���ǩ��</b></td>
		<td>
		<input type="radio" name="EnableSignatures" value=1 <%if SiteConfig("EnableSignatures")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableSignatures" value=0 <%if SiteConfig("EnableSignatures")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�������ǩ������ַ���'><img src="images/spacer.gif" border="0" /></a><b>����ǩ����󳤶�
		���ֽڣ�</b></td>
		<td>
		<input type="text" name="SignatureMaxLength" value="<%=SiteConfig("SignatureMaxLength")%>" />
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�û����ö���Ϣ���������'><img src="images/spacer.gif" border="0" /></a><b>�û�����Ϣ���������������</b></td>
		<td>
		<input type="text" name="MaxPrivateMessageSize" value="<%=SiteConfig("MaxPrivateMessageSize")%>" />
		</td>
	</tr>

	<tr>
		<td><a class="helpicon" title='���øù��ܺ�����û������Լ��Ա𣬽���ʾ����Ϣ��'><img src="images/spacer.gif" border="0" /></a><b>������ʾ�û��Ա�</b></td>
		<td>
		<input type="radio" name="AllowGender" value=1 <%if SiteConfig("AllowGender")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowGender" value=0 <%if SiteConfig("AllowGender")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����û�ͷ����ʾ���û�����ÿ�������С�'><img src="images/spacer.gif" border="0" /></a><b>������������ʾͷ��</b></td>
		<td>
		<input type="radio" name="AllowAvatars" value=1 <%if SiteConfig("AllowAvatars")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowAvatars" value=0 <%if SiteConfig("AllowAvatars")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����û�������ǩ����ʾ���û�����ÿ������ĩ�ˡ�'><img src="images/spacer.gif" border="0" /></a><b>������������ʾ����ǩ��</b></td>
		<td>
		<input type="radio" name="AllowSignatures" value=1 <%if SiteConfig("AllowSignatures")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowSignatures" value=0 <%if SiteConfig("AllowSignatures")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ж���ҳ�����⣬�Ƿ������û�����ѡ���Լ�ϲ����ҳ����'><img src="images/spacer.gif" border="0" /></a><b>�����û�ѡ��ҳ������</b></td>
		<td>
		<input type="radio" name="AllowUserToSelectTheme" value=1 <%if SiteConfig("AllowUserToSelectTheme")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowUserToSelectTheme" value=0 <%if SiteConfig("AllowUserToSelectTheme")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�ÿ��ڵ�¼ǰ�޷�����û���������'><img src="images/spacer.gif" border="0" /></a><b>����û�������Ҫ��¼</b></td>
		<td>
		<input type="radio" name="RequireAuthenticationForProfileViewing" value=1 <%if SiteConfig("RequireAuthenticationForProfileViewing")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="RequireAuthenticationForProfileViewing" value=0 <%if SiteConfig("RequireAuthenticationForProfileViewing")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>�û��б�����</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='�رպ��û��б����ӽ��ӵ����˵����Ƴ���'><img src="images/spacer.gif" border="0" /></a><b>��ʾ�û��б�</b></td>
		<td>
		<input type="radio" name="PublicMemberList" value=1 <%if SiteConfig("PublicMemberList")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="PublicMemberList" value=0 <%if SiteConfig("PublicMemberList")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����û�ʹ�ø��߼��ĵ��û�������'><img src="images/spacer.gif" border="0" /></a><b>�����û�ʹ�ø߼�����ѡ��</b></td>
		<td>
		<input type="radio" name="MemberListAdvancedSearch" value=1 <%if SiteConfig("MemberListAdvancedSearch")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="MemberListAdvancedSearch" value=0 <%if SiteConfig("MemberListAdvancedSearch")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����û��б�ʱ��ÿҳ��ʾ���û�����'><img src="images/spacer.gif" border="0" /></a><b>ÿҳ��ʾ���û�</b></td>
		<td>
		<input type="text" name="MemberListPageSize" value="<%=SiteConfig("MemberListPageSize")%>" /> 
		</td>
	</tr>
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ȫ���û�ͷ������</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='���ú��û�����ѡ����Ի�ͷ��'><img src="images/spacer.gif" border="0" /></a><b>����ͷ��</b></td>
		<td>
		<input type="radio" name="EnableAvatars" value=1 <%if SiteConfig("EnableAvatars")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableAvatars" value=0 <%if SiteConfig("EnableAvatars")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����û�ʹ���ⲿ���ӵ�ͼƬ��Ϊͷ��'><img src="images/spacer.gif" border="0" /></a><b>����Զ��ͷ��</b></td>
		<td>
		<input type="radio" name="EnableRemoteAvatars" value=1 <%if SiteConfig("EnableRemoteAvatars")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableRemoteAvatars" value=0 <%if SiteConfig("EnableRemoteAvatars")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ʾ�û�ͷ������ߴ磬�������Ƶ�ͼƬ����������С��'><img src="images/spacer.gif" border="0" /></a><b>ͷ��ߴ�</b></td>
		<td>
		<input size=3 name="AvatarHeight" value="<%=SiteConfig("AvatarHeight")%>"> X <input size=3 name="AvatarWidth" value="<%=SiteConfig("AvatarWidth")%>" />
		</td>
	</tr>
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>��������</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='�Ƿ�������������'><img src="images/spacer.gif" border="0" /></a><b>��������</b></td>
		<td>
		<input type="radio" name="EnableReputation" value=1 <%if SiteConfig("EnableReputation")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableReputation" value=0 <%if SiteConfig("EnableReputation")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ע��Ļ�Ա����������Ĭ�ϣ�0��'><img src="images/spacer.gif" border="0" /></a><b>Ĭ������</b></td>
		<td>
		<input type="text" name="ReputationDefault" value="<%=SiteConfig("ReputationDefault")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���û����ϲ鿴ҳ������ʾ�����������û����������ۡ���Ĭ�ϣ�5��'><img src="images/spacer.gif" border="0" /></a><b>��ʾ������������</b></td>
		<td>
		<input type="text" name="ShowUserRates" value="<%=SiteConfig("ShowUserRates")%>" /> 
		</td>
	</tr>

	<tr>
		<td><a class="helpicon" title='�û��ķ���������ﵽ��һ��Ŀ�����ܶ������û������������ۡ���Ĭ�ϣ�50��'><img src="images/spacer.gif" border="0" /></a><b>���ٷ�����</b></td>
		<td>
		<input type="text" name="MinReputationPost" value="<%=SiteConfig("MinReputationPost")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�û�������ֵ����ﵽ��һ��Ŀ�����ܶ������û�������������ʱ����Ĭ�ϣ�0��'><img src="images/spacer.gif" border="0" /></a><b>��С����ֵ</b></td>
		<td>
		<input type="text" name="MinReputationCount" value="<%=SiteConfig("MinReputationCount")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='һ���û�һ������ܶ����������û������������ۣ�����Ա�������������������ơ���Ĭ�ϣ�10��'><img src="images/spacer.gif" border="0" /></a><b>ÿ���������۴�������</b></td>
		<td>
		<input type="text" name="MaxReputationPerDay" value="<%=SiteConfig("MaxReputationPerDay")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='��ĳ�û��ٴθ����������۹����û�������������֮ǰ��������������ٸ��û������������ۣ�����Ա�������������������ơ���Ĭ�ϣ�20��'><img src="images/spacer.gif" border="0" /></a><b>�û�������Χ</b></td>
		<td>
		<input type="text" name="ReputationRepeat" value="<%=SiteConfig("ReputationRepeat")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����Ա������������ÿ�ζ������û�������������ʱ�������û��õ�����ʧȥ������ֵ����Ĭ�ϣ�10��'><img src="images/spacer.gif" border="0" /></a><b>����Ա������������������</b></td>
		<td>
		<input type="text" name="AdminReputationPower" value="<%=SiteConfig("AdminReputationPower")%>" /> 
		</td>
	</tr>
	
	<tr>
		<td><a class="helpicon" title='����û�����ֵ������һ��Ŀ�����Զ���������������ٷ������ӣ�֮ǰ���������ӽ������Ρ���Ĭ�ϣ�-9��'><img src="images/spacer.gif" border="0" /></a><b>������������ֵ</b></td>
		<td>
		<input type="text" name="InPrisonReputation" value="<%=SiteConfig("InPrisonReputation")%>" /> 
		</td>
	</tr>
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ͷ������</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='���ú��û������Զ���ͷ��'><img src="images/spacer.gif" border="0" /></a><b>�����û��Զ���ͷ��</b></td>
		<td>
		<input type="radio" name="CustomUserTitle" value=1 <%if SiteConfig("CustomUserTitle")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="CustomUserTitle" value=0 <%if SiteConfig("CustomUserTitle")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����û��Զ���ͷ�ε���󳤶ȡ�'><img src="images/spacer.gif" border="0" /></a><b>�û��Զ���ͷ�ε���󳤶�</b></td>
		<td>
		<input type="text" name="UserTitleMaxChars" value="<%=SiteConfig("UserTitleMaxChars")%>" /> 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�Զ���ͷ����Ҫ���εĴ���'><img src="images/spacer.gif" border="0" /></a><b>�Զ���ͷ����Ҫ���εĴ���</b></td>
		<td>
		<input type="text" name="UserTitleCensorWords" value="<%=SiteConfig("UserTitleCensorWords")%>" size="60" /> 
		</td>
	</tr>
</table>
<!--�û�ע�� End-->


<!--��̳���� Start-->
<table id="Tag_tab_3" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">

	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>������������</b></font></td></tr>


	<tr>
	  <td width=30%><a class="helpicon" title='��Ϊ���й�����̳���� RSS ����Դ��'><img src="images/spacer.gif" border="0" /></a><b>���ù�����̳�� RSS ����Դ</b></td>
	  <td><input type="radio" name="EnableForumsRSS" value=1 <%if SiteConfig("EnableForumsRSS")=1 then%>checked<%end if%> />�� <input type="radio" name="EnableForumsRSS" value=0 <%if SiteConfig("EnableForumsRSS")=0 then%>checked<%end if%> />�� </td>
	</tr>

	<tr>
	  <td><a class="helpicon" title='�������ⷢ�����Ƿ���Խ������������״̬����Ϊ�ѽ��/δ���'><img src="images/spacer.gif" border="0" /></a><b>��������״̬����</b></td>
	  <td><input type="radio" name="DisplayThreadStatus" value=1 <%if SiteConfig("DisplayThreadStatus")=1 then%>checked<%end if%> />�� <input type="radio" name="DisplayThreadStatus" value=0 <%if SiteConfig("DisplayThreadStatus")=0 then%>checked<%end if%> />�� </td>
	</tr>
	<tr>
	  <td><a class="helpicon" title='�����û�ʹ�ñ�ǩ�ӵ������к�ʹ�ñ�ǩ�������������'><img src="images/spacer.gif" border="0" /></a><b>�������Ӻ��б�ǩ</b></td>
	  <td><input type="radio" name="DisplayPostTags" value=1 <%if SiteConfig("DisplayPostTags")=1 then%>checked<%end if%> />�� <input type="radio" name="DisplayPostTags" value=0 <%if SiteConfig("DisplayPostTags")=0 then%>checked<%end if%> />�� </td>
	</tr>
	
	<tr>
	  <td><a class="helpicon" title='��������ƹ�����ʱ������������'><img src="images/spacer.gif" border="0" /></a><b>����ƹ�������������</b></td>
	  <td><input type="radio" name="EnablePostPreviewPopup" value=1 <%if SiteConfig("EnablePostPreviewPopup")=1 then%>checked<%end if%> />�� <input type="radio" name="EnablePostPreviewPopup" value=0 <%if SiteConfig("EnablePostPreviewPopup")=0 then%>checked<%end if%> />�� </td>
	</tr>
	<tr>
		<td><a class="helpicon" title='Ĭ�ϲ鿴���ӵ�ģʽ'><img src="images/spacer.gif" border="0" /></a><b>Ĭ�ϲ鿴���ӵ�ģʽ</b></td>
		<td>
		<input type="radio" name="ViewMode" value=0 <%if SiteConfig("ViewMode")=0 then%>checked<%end if%> />��� 
		<input type="radio" name="ViewMode" value=1 <%if SiteConfig("ViewMode")=1 then%>checked<%end if%> />���� 
		</td>
	</tr>
	
	
	<tr>
		<td><a class="helpicon" title='ÿ��RSS����Դ������'><img src="images/spacer.gif" border="0" /></a><b>Ĭ��RSSԴ������</b></td>
		<td><input value="<%=SiteConfig("RSSDefaultThreadsPerFeed")%>" name="RSSDefaultThreadsPerFeed" size="20" /></td>
	</tr>
	
	</tr>
	<tr>
		<td><a class="helpicon" title='ÿ����ҳ����ʾ��������'><img src="images/spacer.gif" border="0" /></a><b>����/ҳ��</b></td>
		<td><input value="<%=SiteConfig("ThreadsPerPage")%>" id="ThreadsPerPage" name="ThreadsPerPage" size="20" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='ÿ����ҳ������ʾ��������'><img src="images/spacer.gif" border="0" /></a><b>����/ҳ��</b></td>
		<td><input value="<%=SiteConfig("PostsPerPage")%>" name="PostsPerPage" size="20" /></td>
	</tr>
	
	<tr>
		<td><a class="helpicon" title='ͶƱѡ�����ٲ��õ��ڸ���ֵ'><img src="images/spacer.gif" border="0" /></a><b>ͶƱѡ������</b></td>
		<td>
		<input size="20" name="MinVoteOptions" value="<%=SiteConfig("MinVoteOptions")%>" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='ͶƱѡ����󲻵ø��ڸ���ֵ'><img src="images/spacer.gif" border="0" /></a><b>ͶƱѡ�����</b></td>
		<td><input size="20" name="MaxVoteOptions" value="<%=SiteConfig("MaxVoteOptions")%>" /></td>
	</tr>
	
	<tr>
		<td><a class="helpicon" title='�򿪴�ѡ���ʾ�������ĳ����̳�����û�'><img src="images/spacer.gif" border="0" /></a><b>��ʾ������������û�</b></td>
		<td><input type="radio" name="DisplayForumUsers" value=1 <%if SiteConfig("DisplayForumUsers")=1 then%>checked<%end if%> />�� <input type="radio" name="DisplayForumUsers" value=0 <%if SiteConfig("DisplayForumUsers")=0 then%>checked<%end if%> />��</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�򿪴�ѡ���ʾ��ǰ�������ĳ��������û���'><img src="images/spacer.gif" border="0" /></a><b>��ʾ�������������û�</b></td>
		<td><input type="radio" name="DisplayThreadUsers" value=1 <%if SiteConfig("DisplayThreadUsers")=1 then%>checked<%end if%> />�� <input type="radio" name="DisplayThreadUsers" value=0 <%if SiteConfig("DisplayThreadUsers")=0 then%>checked<%end if%> />��</td>
	</tr>

		
	
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>�༭����</b></font></td></tr>
	
	<tr>
		<td><a class="helpicon" title='���ú�����Ϣ�ĵײ�(λ���û�ǩ��ǰ)������ʾ�����ӵı༭��¼ ��'><img src="images/spacer.gif" border="0" /></a><b>��ʾ�޸�ԭ��</b></td>
		<td>
		<input type="radio" name="DisplayEditNotes" value=1 <%if SiteConfig("DisplayEditNotes")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="DisplayEditNotes" value=0 <%if SiteConfig("DisplayEditNotes")=0 then%>checked<%end if%> />��</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���ú��û��༭����ʱҪ�������޸�ԭ��'><img src="images/spacer.gif" border="0" /></a><b>Ҫ���޸�ԭ��</b></td>
		<td>
		<input type="radio" name="RequireEditNotes" value=1 <%if SiteConfig("RequireEditNotes")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="RequireEditNotes" value=0 <%if SiteConfig("RequireEditNotes")=0 then%>checked<%end if%> />��</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='�����øù��ܣ��ھ���ָ��ʱ��������༭����(����Ϊ 0 ʱ������)'><img src="images/spacer.gif" border="0" /></a><b>������೤ʱ�����ܱ༭���ӣ����ӣ�</b></td>
		<td><input size="20" value="<%=SiteConfig("PostEditBodyAgeInMinutes")%>" name="PostEditBodyAgeInMinutes" /></td>
	</tr>	
	
	
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>ˮ������</b></font></td></tr>
	<tr>
		<td><a class="helpicon" title='����ͬһ���û��ظ���������ͬ�����ӡ�'><img src="images/spacer.gif" border="0" /></a><b>�����ظ�����</b></td>
		<td>
		<input type="radio" name="AllowDuplicatePosts" value=1 <%if SiteConfig("AllowDuplicatePosts")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="AllowDuplicatePosts" value=0 <%if SiteConfig("AllowDuplicatePosts")=0 then%>checked<%end if%> />��</td>
	</tr>	
	<tr>
		<td><a class="helpicon" title='���η���ʱ�����ټ��'><img src="images/spacer.gif" border="0" /></a><b>ÿ�η���������룩</b></td>
		<td><input size="20" name="PostInterval" value="<%=SiteConfig("PostInterval")%>" /></td>
	</tr>	
	<tr>
		<td><a class="helpicon" title='�����ע���û��ڴ�ʱ����ܷ�������'><img src="images/spacer.gif" border="0" /></a><b>ע���೤ʱ����ܷ������ӣ����ӣ�</b></td>
		<td><input size="20" value="<%=SiteConfig("RegUserTimePost")%>" name="RegUserTimePost" /></td>
	</tr>
	
	<tr><td colspan="2"><font size=4 color="#C0C0C0"><b>��������׼</b></font></td></tr>
	
	<tr>
		<td><a class="helpicon" title='һ��������Ҫ����ƪ�������ܱ������?'><img src="images/spacer.gif" border="0" /></a><b>�������ӵĻظ���</b></td>
		<td><input size="20" value="<%=SiteConfig("PopularPostThresholdPosts")%>" name="PopularPostThresholdPosts" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='һ��������Ҫ���ٴ�������ܱ������?'><img src="images/spacer.gif" border="0" /></a><b>�������ӵ������</b></td>
		<td><input size="20" value="<%=SiteConfig("PopularPostThresholdViews")%>" name="PopularPostThresholdViews" /></td>
	</tr>	
	<tr>
		<td><a class="helpicon" title='��ͨ�����Ϊ����������뱣�ֻ�Ծ״̬������'><img src="images/spacer.gif" border="0" /></a><b>����ʱ�� (��)</b></td>
		<td><input size="20" value="<%=SiteConfig("PopularPostThresholdDays")%>" name="PopularPostThresholdDays" /></td>
	</tr>	
</table>
<!--��̳���� End-->

<!--��֤������ Start-->
<table id="Tag_tab_7" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='���û�ע��ʱ������֤�빦�ܡ�'><img src="images/spacer.gif" border="0" /></a><b>����ע����֤</b></td>
		<td>
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value=1 <%if SiteConfig("EnableAntiSpamTextGenerateForRegister")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForRegister" value=0 <%if SiteConfig("EnableAntiSpamTextGenerateForRegister")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���û���¼ʱ������֤�빦�ܡ�'><img src="images/spacer.gif" border="0" /></a><b>���õ�¼��֤</b></td>
		<td>
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value=1 <%if SiteConfig("EnableAntiSpamTextGenerateForLogin")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForLogin" value=0 <%if SiteConfig("EnableAntiSpamTextGenerateForLogin")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���û�����ʱ������֤�빦�ܡ�'><img src="images/spacer.gif" border="0" /></a><b>���÷�����֤</b></td>
		<td>
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value=1 <%if SiteConfig("EnableAntiSpamTextGenerateForPost")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="EnableAntiSpamTextGenerateForPost" value=0 <%if SiteConfig("EnableAntiSpamTextGenerateForPost")=0 then%>checked<%end if%> />�� 
		</td>
	</tr>
</table>
<!--��֤������ End-->

<!--��Ǯ�;������� Strat-->
<table id="Tag_tab_8" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='�û�ÿ����һ�����������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>��������</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralAddThread")%>" name="IntegralAddThread" /></td>
	</tr>
	<tr>
	  <td><a class="helpicon" title='�û�ÿ�ظ�һ���������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>����</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralAddPost")%>" name="IntegralAddPost" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='������ĳ���Ӽ�Ϊ���������ӵ��������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>�Ӿ���</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralAddValuedPost")%>" name="IntegralAddValuedPost" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ɾ��ĳ�����������ӵ��������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>��������ɾ��</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralDeleteThread")%>" name="IntegralDeleteThread" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ɾ��ĳ���������ӵ��������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>������ɾ��</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralDeletePost")%>" name="IntegralDeletePost" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='����ȡ��ĳ���������ӵ��������ý�Ǯ�;���'><img src="images/spacer.gif" border="0" /></a><b>������ȡ��</b></td>
		<td><input size="20" value="<%=SiteConfig("IntegralDeleteValuedPost")%>" name="IntegralDeleteValuedPost" /></td>
	</tr>
</table>
<!--��Ǯ�;������� End-->


<!--�������� Strat-->
<table id="Tag_tab_9" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td width="30%"><a class="helpicon" title='�Ƿ�����ϵͳ����'><img src="images/spacer.gif" border="0" /></a><b>��������</b></td>
		<td>
		<input type="radio" name="APIEnable" value=1 <%if SiteConfig("APIEnable")=1 then%>checked<%end if%> />�� 
		<input type="radio" name="APIEnable" value=0 <%if SiteConfig("APIEnable")=0 then%>checked<%end if%> />�� </td>
	</tr>
	<tr>
	  <td><a class="helpicon" title='����ϵͳ��ȫ��Կ,��BBSXP��ע�⣺ϵͳ���ϣ����뱣֤������ϵͳ���õ���Կһ�¡���'><img src="images/spacer.gif" border="0" /></a><b>��ȫ��Կ</b></td>
		<td><input size="40" value="<%=SiteConfig("APISafeKey")%>" name="APISafeKey" /></td>
	</tr>
	<tr>
		<td><a class="helpicon" title='���ϵ���������Ľӿ��ļ�·�����磺http://bbs.yuzi.net/API_Response.asp��ע���������ӿ�֮����Ӣ�İ��"|"�ָ���'><img src="images/spacer.gif" border="0" /></a><b>�ӿ��ļ���ַ�����á�|���ָ���</b></td>
		<td><input value="<%=SiteConfig("APIUrlList")%>" name="APIUrlList" style="width:90%" /></td>
	</tr>
</table>
<!--�������� End-->



<br />
<table border="0"  width="100%" align="center">
	<tr>
		<td align="right" height="20"><input type="button" onclick="SiteSettingsInit()" value=" ��ԭĬ��ֵ " />��<input type="submit" value=" �� �� " />��<input type="reset" value=" �� �� " /></td>
	</tr>
</table>
</form>
<script language="JavaScript" type="text/javascript">
var FormObject = document.form;

function SiteSettingsInit() {

	FormObject.elements["UserOnlineTime"].value = 20;
	FormObject.elements["Timeout"].value = 60;
	FormObject.elements["DefaultSiteStyle"].value = "default";
	FormObject.elements["CacheName"].value = "BBSXP";
	FormObject.elements["CacheUpDateInterval"].value = 5;
	FormObject.elements["SiteDisabled"][1].checked = true;
	FormObject.elements["SiteDisabledReason"].value = "��̳ά���У���ʱ�޷����ʣ�";
	FormObject.elements["CacheName"].value = "BBSXP";
	FormObject.elements["CacheUpDateInterval"].value = 5;
	FormObject.elements["NoCacheHeaders"][1].checked = true;

	FormObject.elements["DisplayWhoIsOnline"][0].checked = true;
	FormObject.elements["DisplayStatistics"][0].checked = true;
	FormObject.elements["DisplayBirthdays"][0].checked = true;
	FormObject.elements["DisplayLink"][0].checked = true;
	FormObject.elements["IsShowSonForum"][0].checked = true;

	FormObject.elements["AllowLogin"][0].checked = true;
	FormObject.elements["EnableBannedUsersToLogin"][1].checked = true;
	FormObject.elements["AllowNewUserRegistration"][0].checked = true;
	FormObject.elements["NewUserModerationLevel"].options[1].selected = true;
	FormObject.elements["UserNameMinLength"].value = 3;
	FormObject.elements["UserNameMaxLength"].value = 15;
	FormObject.elements["DefaultPasswordFormat"][0].checked = true;
	FormObject.elements["AccountActivation"][0].checked = true;
	FormObject.elements["EnableSignatures"][0].checked = true;
	FormObject.elements["SignatureMaxLength"].value = 256;
	FormObject.elements["MaxPrivateMessageSize"].value = 100;
	FormObject.elements["AllowSignatures"][0].checked = true;
	FormObject.elements["AllowGender"][0].checked = true;
	FormObject.elements["AllowAvatars"][0].checked = true;
	FormObject.elements["AllowUserToSelectTheme"][0].checked = true;
	FormObject.elements["RequireAuthenticationForProfileViewing"][1].checked = true;
	FormObject.elements["PublicMemberList"][0].checked = true;
	FormObject.elements["MemberListAdvancedSearch"][0].checked = true;
	FormObject.elements["MemberListPageSize"].value = 20;
	FormObject.elements["EnableAvatars"][0].checked = true;
	FormObject.elements["EnableRemoteAvatars"][0].checked = true;
	FormObject.elements["AvatarHeight"].value = 150;
	FormObject.elements["AvatarWidth"].value = 150;
	FormObject.elements["EnableReputation"][0].checked = true;
	FormObject.elements["ReputationDefault"].value = 0;
	FormObject.elements["ShowUserRates"].value = 5;
	FormObject.elements["MinReputationPost"].value = 50;
	FormObject.elements["MinReputationCount"].value = 0;
	FormObject.elements["MaxReputationPerDay"].value = 10;
	FormObject.elements["ReputationRepeat"].value = 20;
	FormObject.elements["AdminReputationPower"].value = 10;
	FormObject.elements["InPrisonReputation"].value = "-9";
	FormObject.elements["CustomUserTitle"][0].checked = true;
	FormObject.elements["UserTitleMaxChars"].value = 25;
	FormObject.elements["UserTitleCensorWords"].value = "admin|forum|moderator";
	

	FormObject.elements["EnableForumsRSS"][0].checked = true;
	FormObject.elements["DisplayThreadStatus"][0].checked = true;
	FormObject.elements["DisplayPostTags"][0].checked = true;
	FormObject.elements["EnablePostPreviewPopup"][1].checked = true;
	FormObject.elements["ViewMode"][1].checked = true;
	FormObject.elements["RSSDefaultThreadsPerFeed"].value = 20;
	FormObject.elements["ThreadsPerPage"].value = 20;
	FormObject.elements["PostsPerPage"].value = 15;
	FormObject.elements["MinVoteOptions"].value = 2;
	FormObject.elements["MaxVoteOptions"].value = 20;
	FormObject.elements["DisplayEditNotes"][1].checked = true;
	FormObject.elements["RequireEditNotes"][1].checked = true;
	FormObject.elements["PostEditBodyAgeInMinutes"].value = 0;
	FormObject.elements["AllowDuplicatePosts"][1].checked = true;
	FormObject.elements["PostInterval"].value = 10;
	FormObject.elements["RegUserTimePost"].value = 0;
	FormObject.elements["PopularPostThresholdPosts"].value = 15;
	FormObject.elements["PopularPostThresholdViews"].value = 200;
	FormObject.elements["PopularPostThresholdDays"].value = 3;
	FormObject.elements["DisplayForumUsers"][0].checked = true;
	FormObject.elements["DisplayThreadUsers"][1].checked = true;

	FormObject.elements["UpFileTypes"].value = "gif|jpg|png|bmp|swf|txt|mid|doc|xls|zip|rar";
	FormObject.elements["AttachmentsSaveOption"].options[1].selected = true;
	FormObject.elements["MaxPostAttachmentsSize"].value = 10240;
	FormObject.elements["MaxFileSize"].value = 200;
	FormObject.elements["MaxFaceSize"].value = 100;
	
	
	FormObject.elements["WatermarkFontFamily"].value = "����";
	FormObject.elements["WatermarkFontSize"].value = 25;
	FormObject.elements["WatermarkFontColor"].value = "#000000";
	$("WatermarkFontColor").style.backgroundColor = "#000000";
	FormObject.elements["WatermarkFontIsBold"][0].checked = true;
	FormObject.elements["WatermarkWidthPosition"][0].checked = true;
	FormObject.elements["WatermarkHeightPosition"][2].checked = true;

	FormObject.elements["BannedText"].value = "fuck|shit";
	FormObject.elements["BannedRegUserName"].value = "fuck|shit";

	FormObject.elements["EnableAntiSpamTextGenerateForRegister"][0].checked = true;
	FormObject.elements["EnableAntiSpamTextGenerateForLogin"][1].checked = true;
	FormObject.elements["EnableAntiSpamTextGenerateForPost"][1].checked = true;

	FormObject.elements["IntegralAddThread"].value = "+2";
	FormObject.elements["IntegralAddPost"].value = "+1";
	FormObject.elements["IntegralAddValuedPost"].value = "+3";
	FormObject.elements["IntegralDeleteThread"].value = "-2";
	FormObject.elements["IntegralDeletePost"].value = "-1";
	FormObject.elements["IntegralDeleteValuedPost"].value = "-3";
	
	FormObject.elements["APIEnable"][1].checked = true;
}

function IsObjInstalled (OjbSpanID,OjbStr) {
	OjbStr == "" ? $(OjbSpanID).innerHTML = "" : Ajax_CallBack(false,OjbSpanID,'Loading.asp?menu=IsObjInstalled&Object='+OjbStr);
}
</script>
<%
End Sub

Sub SiteInfo
	If IsObjInstalled("Scripting.FileSystemObject") Then
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not fso.folderexists(Server.MapPath("Themes/"&Request("DefaultSiteStyle")&"")) then Alert("Ĭ�Ϸ�����ô���Themes/"&Request("DefaultSiteStyle")&" Ŀ¼�����ڣ�")
		Set fso = nothing
	end if

	if Request("UpFileOption")<>"" then
		if Not IsObjInstalled(Request("UpFileOption")) then Alert("����������֧�� "&Request("UpFileOption")&" �������رմ˹��ܣ�")
	end if
	
	if Request("SelectMailMode")<>"" then
		if Not IsObjInstalled(Request("SelectMailMode")) then Alert("����������֧�� "&Request("SelectMailMode")&" �������رմ˹��ܣ�")
		if Request("SmtpServerMail")="" then Alert("����д�����˵�Email��ַ��")
	end if
	
	if Request("WatermarkOption")<>"" then
		if Not IsObjInstalled(Request("WatermarkOption")) then Alert("����������֧�� "&Request("WatermarkOption")&" �������رմ˹��ܣ�")
	end if

	if Request("AccountActivation")="1" and Request("SelectMailMode")="" then Alert("��ѡ����ע���û�����ͨ��Email���ͣ�������û���趨�����ʼ������")
	if left(Request("BannedText"),1)="|" or right(Request("BannedText"),1)="|" then Alert("���˸�ʽ���ô���ǰ�����С�|��")
	if left(Request("BannedRegUserName"),1)="|" or right(Request("BannedRegUserName"),1)="|" then Alert("���˸�ʽ���ô���ǰ�����С�|��")
	if left(Request("BannedIP"),1)="|" or right(Request("BannedIP"),1)="|" then Alert("���˸�ʽ���ô���ǰ�����С�|��")
	
	filtrate=split(Request("BannedIP"),"|")
	for i = 0 to ubound(filtrate)
		if instr("|"&REMOTE_ADDR&"","|"&filtrate(i)&"") > 0 then Alert("�����������IP��ַ�Ƿ���ȷ")
	next

	
	FielsStr="UserOnlineTime|Timeout|UserNameMinLength|UserNameMaxLength|ThreadsPerPage|PostsPerPage|RSSDefaultThreadsPerFeed|PostInterval|PostEditBodyAgeInMinutes|RegUserTimePost|PopularPostThresholdPosts|PopularPostThresholdViews|MinVoteOptions|MaxVoteOptions|MaxPostAttachmentsSize|MaxFileSize|MaxFaceSize|AvatarHeight|AvatarWidth|SignatureMaxLength|MemberListPageSize|PopularPostThresholdDays"
	FielsStrArray=split(FielsStr,"|")
	for i=0 to ubound(FielsStrArray)
		if IsNumeric(Request.Form(""&FielsStrArray(i)&""))=false then Alert(FielsStrArray(i)&"������ڵ�ֵ���������������ͣ�")
	next
	
	Rs.Open "["&TablePrefix&"SiteSettings]",Conn,1,3
	
	SiteConfigXMLDOM.loadxml("<bbsxp>"&Rs("SiteSettingsXML")&"</bbsxp>")
		SiteSettingsXMLStrings=""
		Set objNodes = SiteConfigXMLDOM.documentElement.ChildNodes
		Set objRoot = SiteConfigXMLDOM.documentElement
		for each ho in Request.Form
			objRoot.SelectSingleNode(ho).text = ""&server.HTMLEncode(Request(""&ho&""))&""
		next
		for each element in objNodes	
			SiteSettingsXMLStrings=SiteSettingsXMLStrings&"<"&element.nodename&">"&element.text&"</"&element.nodename&">"&vbCrlf
		next
	Rs("SiteSettingsXML")=SiteSettingsXMLStrings
	Rs.update
	Rs.close
	Response.Write("���³ɹ�")
End Sub

Sub AdvancedSettingsUp
	Rs.Open "["&TablePrefix&"SiteSettings]",Conn,1,3
		for each ho in Request.Form
			Rs(ho)=Request(ho)
		next
	Rs.update
	Rs.close
	Response.Write("���³ɹ�")
End Sub

Sub AdvancedSettings
%>
<form method="POST" action="?menu=AdvancedSettingsUp">
<div class="MenuTag">
	<li id="Tag_btn_0" onclick="ShowPannel(this)" class="NowTag">ע���û�Э��</li>
	<li id="Tag_btn_1" onclick="ShowPannel(this)">HTML��������</li>
</div>
<table id="Tag_tab_0" style="display:block" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
	<tr>
		<td align=center><textarea name="CreateUserAgreement" rows="18" style="width:90%"><%=CreateUserAgreement%></textarea></td>
	</tr>
</table>
	<table id="Tag_tab_1" style="display:none" cellspacing=0 cellpadding=0 width=100% class="PannelBody">
		<tr>
			<td align="center" width="110">ҳ�涥����Ϣ<br /><font color="#FF0000">����&lt;Body&gt;��</font></td>
			<td><textarea name="GenericHeader" rows="6" style="width:95%"><%=GenericHeader%></textarea></td>
		</tr>
		<tr>
			<td align="center" width="110">Banner (468 x 60)<br /><font color="#FF0000">�����Ҳ�λ��</font> </td>
			<td><textarea name="GenericTop" rows="6" style="width:95%"><%=GenericTop%></textarea></td>
		</tr>

		<tr>
			<td align="center" width="110">ҳ��ײ���Ϣ <br /><font color="#FF0000">����&lt;Body&gt;��</font> </td>
			<td><textarea name="GenericBottom" rows="6" style="width:95%"><%=GenericBottom%></textarea></td>
		</tr>
	</table>
    <br />
    <table border="0"  width="100%" align="center">
	<tr>
		<td align="right" height="20"><input type="submit" value=" �� �� " />		<input type="reset" value=" �� �� " /></td>
	</tr>
	</table>
</form>
<%
End Sub

AdminFooter
%>