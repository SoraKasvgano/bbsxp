<!-- #include file="Setup.asp" -->
<%top%>

<table border=0 width=100% align=center cellspacing=1 cellpadding=4 class=a2><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
	<a href="help.asp">ʹ�ð���</a></td></tr></table>
<br>

<%if Request("menu")="Ranks"then%>
<table border="0" cellpadding="3" cellspacing="1" class=a2 width=100%>
	<tr class=a1>
		<td width="30%" align="center">�ȼ�����</td>
		<td width="30%">
		<p align="center">����ֵ</p>
		</td>
		<td width="30%">����ͼ��</td>
	</tr>
<%
sql="select * from [BBSXP_Ranks] order by PostingCountMin"
Set Rs=Conn.Execute(sql)
do while not Rs.eof
%>
	<tr class=a3>
		<td width="30%" align="center"><%=Rs("RankName")%></td>
		<td width="30%" align="center">��<%=Rs("PostingCountMin")%></td>
		<td width="30%">
		<img src="<%=Rs("RankIconUrl")%>">
		</td>
	</tr>
<%
Rs.Movenext
loop
Set Rs = Nothing
%>
	<tr class=a3>
		<td width="30%" align="center">��δ����</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode0.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">�����Ա</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode2.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">����</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode3.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">��������</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode4.gif" /></td>
	</tr>
	<tr class=a3>
		<td width="30%" align="center">����Ա</td>
		<td width="30%" align="center">--</td>
		<td width="30%">
		<img border="0" src="images/level/MemberCode5.gif" /></td>
	</tr>
</table>

<%elseif Request("menu")="popedom"then%>

<table cellspacing="1" cellpadding="0" border="0" class=a2 width=100%>
	<tr height="25" class=a1>
		<td align="middle" width="20%">����Ȩ��</td>
		<td align="middle" width="10%">�ÿ�</td>
		<td align="middle" width="10%">ע���Ա</td>
		<td align="middle" width="10%">�����Ա</td>
		<td align="middle" width="10%">��̳����</td>
		<td align="middle" width="10%">��������</td>
		<td align="center" width="10%">����Ա</td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">���������̳</td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�����Ա��̳</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">���Ͷ�Ѷ</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">����ͶƱ</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�༭����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">��������</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">��������</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�����ɫ����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">��ǰ����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�ر�����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�ƶ�����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">ɾ������</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�̶�����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">��Ӿ�����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�����̶ܹ�</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">������������</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">ɾ��������־</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">�鿴���߻�ԱIP</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b>��</b></td>
		<td align="center"><b>��</b></td>
	</tr>
	<tr height="25" class=a3>
		<td align="middle">��̨����</td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="middle"><b><font color="red">��</font></b></td>
		<td align="center"><b>��</b></td>
	</tr>
	</table>
<%elseif Request("menu")="ybb"then%>

<table cellspacing="1" cellpadding="5" width="100%" align="center" border="0" class="a2">
	<tr>
		<td class="a1" height="25"><b>ʲô��YBB(Yuzi Bulletin Board)���룿 </b></td>
	</tr>
	<tr class=a3>
		<td>YBB������HTML��һ�����֡���Ҳ���Ѿ���������Ϥ�ˡ�һ������£������������HTM��Ҳ�Ϳ���ʹ��YBB���롣��ʹ������������������ʹ��HTML����Ҳ����ʹ��YBB���롣 
		����Ҫ��ʹ�õı�����٣���ʹ����ʹ��HTML��������Ҳ��ʹ��YBB���룬��Ϊ�������Ŀ����Դ���С�ˡ�
		</td>
	</tr>
	<tr>
		<td class="a1" height="25"><b>URL��������</b></td>
	</tr>
	<tr class=a3>
		<td>��������Ϣ����볬�����ӣ�ֻҪ�����з�ʽ����Ϳ�����<br><br>
		<font color="red">[url]</font><a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a><font color="red">[/url]</font>
		<br>���� <br>
		<font color="red">[url=http://www.yuzi.net]</font><a target="_blank" href="http://www.yuzi.net">YUZI������</a><font color="red">[/url]</font><br><br>
		���������룬YBB������Զ���URL�������ӣ�����֤���û�����µĴ���ʱ��������Ǵ��ŵġ�ע��URL��&quot;http://&quot;��һ����������ġ�</td>
	</tr>
	<tr class="a1">
		<td height="25"><b>�����ʼ�����</b></td>
	</tr>
	<tr class=a3>
		<td>��������Ϣ���������ʼ��ĳ������ӣ�ֻҪ������������Ϳ�����<br>
<br>
		<font color="red">[email]</font><a href="mailto:yuzi@yuzi.net">yuzi@yuzi.net</a><font color="red">[/email]</font><br>
<br>
���������룬YBB�����Ե����ʼ��Զ��������ӡ� </td>
	</tr>
	<tr class="a1">
		<td height="25"><b>������б��</b></td>
	</tr>
	<tr class=a3>
		<td>������ʹ�� [b] [i] ����Щ��־�Դﵽ��������ʹ����Ӧ��Ч��<br>
		<br>
		����,<font color="red">[b]</font><strong>����Ա</strong><font color="red">[/b]</font><br>
		����,<font color="red">[i]</font><em>����Ա</em><font color="red">[/i]</font><br>
		����,<font color="red">[u]</font><u>����Ա</u><font color="red">[/u]</font><br>
		����,<font color="red">[strike]</font><strike>����Ա</strike><font color="red">[/strike]</font></td>
	</tr>
	<tr class="a1">
		<td><b>�ƶ����� </b></td>
	</tr>
	<tr class=a3>
		<td height="42">��������Ϣ������ƶ����֣�ֻҪ����������Ϳ�����<br><br>
		<font color="red">[marquee]</font>�ƶ�����<font color="red">[/marquee]</font><br>
		
		<marquee>�ƶ�����</marquee></td>
	</tr>
	<tr class="a1">
		<td><b>����������Ϣ </b></td>
	</tr>
	<tr class=a3>
		<td>����һЩ�˵����ӣ�ֻҪ���кͿ���Ȼ����������Ϳ�����<br>
		<br><font color="red">
		[QUOTE]</font>�������Ĺ�����Ϊ����ʲô......<br>
		������Ϊ���Ĺ�����ʲô��<font color="red">[/QUOTE]</font>
		<BLOCKQUOTE><strong>����</strong>��<HR Size=1>�������Ĺ�����Ϊ����ʲô......<br>
		������Ϊ���Ĺ�����ʲô�� <HR SIZE=1></BLOCKQUOTE>
		</td>
	</tr>
	<tr class="a1">
		<td><b>��ɫ����</b></td>
	</tr>
	<tr class=a3>
		<td><font color="red">[COLOR=green]</font><font COLOR="#008000">��ɫ����</font><font COLOR="red">[/COLOR]</font>������green���������ֵ���ɫ<font color="red"><br>
		[FONT=����]</font><font face="����">��ɫ����</font><font color="red">[/FONT]</font>���������塱�������ֵ�����<font color="red"><br>
		[SIZE=5]</font><font size="5">��ɫ����</font><font color="red">[/SIZE]</font>������5���������ֵĴ�С</td>
	</tr>
	<tr>
		<td height="14" class="a1"><b>�ر�ע��</b></td>
	</tr>
	<tr>
		<td height="76" bgcolor="FFFFFF">
		<p>��������ͬʱʹ��<font face="Verdana, Arial">HTML</font>��YBB�����ͬһ�ֹ��ܡ�����ע��YBB���벻�Դ�Сд���С�������������<font color="red">[URL]</font> 
		�� <font color="red">[url]</font> <font color="800000">
		<br>
		����ȷ��</font><font color="800000" face="Verdana, Arial">YBB</font><font color="800000">����ʹ�ã�</font><font face="Verdana, Arial"><font color="red"><br>
		[url]</font> www.yuzi.net <font color="red">[/url]</font> </font>��Ҫ�����ź������������֮���пո�<font face="Verdana, Arial"><font color="red"><br>
		[email]</font>yuzi@yuzi.net<font color="red">[email]</font>
		</font>�ڽ���ʱ����Ҫ�����������ڼ���б��<font color="red">[/email]</font>
		</p>
		</td>
	</tr>
</tbody>
</table>


<%else%>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">���������� </td>
	</tr>
	<tr>
		<td class="a3"><b>ע��͵�¼</b> <br>
		<a href="#A1">��ΪʲôҪע�᣿ </a><br>
		<a href="#A2">��������ע�᣿ </a><br>
		<a href="#A3">���Ѿ�ע�����û��������룬��ô��¼�� </a><br>
		<a href="#A4">���Ѿ���¼��Ϊʲô�����Զ�ע����</a> <br>
		<a href="#A5">���������û���/����.</a> <br>
		<a href="#A6">Ϊʲô����ע�ᵫ�Բ��ܵ�¼��</a> <br>
		<a href="#A7">��ǰע�ᣬ�������ڲ��ܵ�¼��</a>
		<p><b>�û��������� &amp; ���� </b><br>
		<a href="#B1">ʲô�Ǹ������ϣ�</a> <br>
		<a href="#B2">��������������ǩ����</a> <br>
		<a href="#B3">ʲô��ͷ��</a> <br>
		<a href="#B4">��������ҵ�ͷ�� </a><br>
		<a href="#B5">Ϊʲô����Ҫ��¼���ܷ����������Ա���ϣ�</a> <br>
		<a href="?menu=Ranks">����֪����Ա�ĵȼ����ƣ�</a></p>
		<p><b>��˽ &amp; ��ȫ</b> <br>
		<a href="#C1">��θ������룿 </a><br>
		<a href="#C2">��θ���e-mail��ַ��</a> <br>
		<a href="#C3">Ҫ�����õĸ������ϣ�</a> </p>
		<p><b>����</b> <br>
		<a href="#D1">ʲô����̳�飿</a> <br>
		<a href="#D2">ʲô����̳��</a> <br>
		<a href="#D3">ʲô�����⣿ </a><br>
		<a href="#D4">����ͼ����ʲô��˼��</a> <br>
		<a href="#D5">�������̳ʱ�������κ�����/���ӣ�</a> <br>
		<a href="#D6">Ϊʲô�������Ҹշ������ӣ�</a> <br>
		<a href="#D7">ʲô�ǹ̶����� </a><br>
		<a href="#D8">ʲô��������</a> <br>
		<a href="#D9">�������̳ʱ���ܽ����������� </a><br>
		<a href="#D10">��̳��RSSͼ����ʲô��</a> <br>
		<a href="#D11">�Ҳ��ܵ�¼��ǰ���ʹ�����̳��</a> </p>
		<p><b>���� </b><br>
		<a href="#E1">����ʹ��HTML��</a> <br>
		<a href="#E2">ʲô��YBB���룿</a> <br>
		<a href="#E3">�ҿ������ҵ�����������Ӹ�����</a> <br>
		<a href="#E4">ʲô�Ǳ���ͼ�ꣿ</a> <br>
		<a href="#E5">�����������ڰ���з��������ӣ�</a> <br>
		<a href="#E6">��ô�ظ�һ�����ӣ�</a> <br>
		<a href="#E7">��ô���޸��ҵ����ӣ�</a> <br>
		<a href="#E8">��ôɾ���ҵ����ӣ�</a> <br>
		<a href="#E9">Ϊʲô�ҵ���������Щ�ʱ��滻���ˡ�***����</a> <br>
		<a href="#E10">����ô���ҵ������м�ǩ����</a> <br>
		<a href="#E11">��ô��ͷ����ӵ������У�</a> <br>
		<a href="#E12">����֪����̳�ķ�������</a></p>
		<p><b>�û�Ȩ��</b> <br>
		<a href="#F1">ʲô�ǹ���Ա��</a> <br>
		<a href="#F2">ʲô�ǰ�����</a> <br>
		<a href="?menu=popedom">����֪����̳�����˵�Ȩ�ޣ�</a> </p>
		<p><b>���� BBSXP</b> <br>
		<a href="#G1">ʲô�� BBSXP��</a> <br>
		<a href="#G2">˭��ʹ�� BBSXP��</a> <br>
		<a href="#G3">˭������ BBSXP�� </a><br>
		<a href="#G4">���Ļ�ȡ BBSXP �Ŀ�����</a> </p>
		</td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">ע�� &amp; ��¼ </td>
	</tr>
	<tr>
		<td class="a3"><a name="A1"></a><b>��ΪʲôҪע�᣿ </b><br>
		�������Ķ������˵�����ǰ�Ƿ���Ҫ��¼����ȡ������̳����Ա�Ƿ���������̳���ο�Ȩ�ޡ����������������Ҫע���¼�Ϳ����Ķ���Ϣ��ͨ��ע�����������ܵ����еĸ��ӹ��ܣ����磺��������ͷ��
		�ղظ���Ȥ�����ӡ����û������ʼ���˽�����ԡ�����һЩ�ܱ�������̳�ȵȣ�ע�Ჽ��Ҳ�ǳ��򵥣�һ���Ӿ�����ɡ�<br><a href="#top">���ض���</a></a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A2"></a><b>��������ע�᣿ </b><br>
		Ҫע��һ�����ʺţ�����Ҫ����ע�Ტ�Ұ���Ҫ����дע������������ṩһ���û�����һ����Ч�ĵ��������ַ������Ա����Ҫ����ָ�����룬�������Ҫָ�����룬��ô��ע����ɺ������յ�ȷ���ʼ��� 
		<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A3"></a><b>���Ѿ�ע�����û��������룬��ô��¼�� </b><br>
		ע��ɹ����㽫ӵ��һ���û��������룬����Է��ʵ�¼ҳ�沢���������û����������¼����̳�� <br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A4"></a><b>���Ѿ���¼��Ϊʲô�����Զ�ע����</b> <br>
		�����¼��ʱ�����û��ѡ�� &quot;�Զ���¼&quot; ��ѡ���������뿪ʱ�Զ�ע������������һֱ���ֵ�¼״̬�����ڵ�¼��ʱ��ѡ�� &quot;�Զ���¼&quot; ��ѡ�� 
		<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A5"></a><b>������������.</b> <br>
		���������������룬�����Է��� &quot;�һ�����&quot; 
		ҳ������ע��ʱ�ĵ��������ַ������һ������������ʼ����͵�����ע�����䡣��Ϊ�����ǲ�������ܴ洢�����������޷��һ�����ԭʼ���롣һ�����յ����������룬�����Ե�¼���޸��������롣 
		<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A6"></a><b>��ע�ᵫ����Ȼ���ܵ�¼�� </b><br>
		�������ע�ᵫ����Ȼ���ܵ�¼,��ȷ�����Ƿ���һ���Ϸ����û��������롣 
		�����ȷ������û�������������ȷ�ģ�����Ȼ�޷����룬���������ʺ���Ҫ���������ʺ���ͣ���С� ��������ԭ���������ϵ�������ǹ���Ա�� <br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="A7"></a><b>��ע�ᣬ�������ڲ��ܵ�¼�� </b><br>
		���ȼ������û����������Ƿ���ȷ�������Ȼ���ܵ�¼������ʺſ������ڳ���δ��¼�ѱ�ɾ����������̳����Ա���߰�����ϵ�� <br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">�û����� &amp; ���� </td>
	</tr>
	<tr>
		<td class="a3"><a name="B1"></a><b>ʲô�Ǹ������ϣ�</b> <br>
		�������Ͼ�������ʺ���Ϣ���ʺſ���������β鿴��̳�ϵ����ӡ���������������������ӵ���ϸ���ϣ������������˹����˽����Ϣ�������ַ�����͵�ַ���Լ�һЩ���������������̳�໥���õ�����
		��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B2"></a><b>��������������ǩ���� </b><br>
		ǩ�����Ǹ������㷢������̳�ϵ�ÿһƪ���Ӻ������Ϣ. ������ڸ�������ҳ��༭���ǩ��. ��ǩ������ʾ������������κ���Ϣ��β����<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B3"></a><b>ʲô��ͷ��</b> <br>
		ͷ����һ�����Ӹ���������<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B4"></a><b>���������ͷ�� </b><br>
		�������Ա������ͷ�񣬵���鿴��������ʱ���Կ���ͷ�����������������ѡ��һ����ϲ����ͷ������ϴ���������һ��ͼƬ����ַ����Ϊ���ͷ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="B5"></a><b>Ϊʲô����Ҫ��¼���ܷ����������Ա���ϣ� </b>
		<br>
		���ݹ���Ա���������̳���������Ҫ��¼�������/ʹ����̳��ĳЩ��������Ҫ��Ϊ�˱����û�����˽��<br><a href="#top">���ض���</a> 
		<br>
		 </td>
	</tr>

	</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">˽����Ϣ &amp; ��ȫ </td>
	</tr>
	<tr>
		<td class="a3"><a name="C1"></a><b>��θ������룿 </b><br>
		��¼�Ժ������ڸ�������ҳ����������롣<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="C2"></a><b>��θ���email��ַ��</b> <br>
		������ڿ�����壭�������޸������Email��ַ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="C3"></a><b>Ҫ�����õĸ������ϣ� </b><br>
		ֻ�е����ʼ���ַ�Ǳ�����ġ� ����ʼ���ַ�������㷢����Ԥ������̳��Ϣ�������㷢�����ǵ��û��������롣���������϶��ǲ��Ǳ�����ġ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">���� </td>
	</tr>
	<tr>
		<td class="a3"><a name="D1"></a><b>ʲô����̳�飿</b> <br>
		��̳���Ǹ�����������ذ���Ĵ��ࡣһ����̳�����һ�������Ӱ��档<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D2"></a><b>ʲô�ǰ�飿 </b><br>
		�����������һϵ��������顣һ����������ƪ���ӻ��߶�����档<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D3"></a><b>ʲô�����⣿</b> <br>
		һ���������һ����ص����ӣ����Ƕ�������һ�����⡣��һƪ���ӳ�Ϊһ�����⣬�������������������档�������滹��¼һ�����ٻ�����˭����������Ϣ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D4"></a><b>����ͼ�����ʲô��˼��</b> 
		<table border=0>
			<tr>
				<td><img src="images/f_New.gif" border="0" alt="�лظ�������"> 
				������</td>
				<td><img src="images/f_top.gif" border="0"> 
		�̶�����</td>
				<td><img src="images/f_poll.gif" border="0"> 
				ͶƱ����</td>
				<td><img src="images/Topicgood.gif"> ��������</td>
			</tr>
			<tr>
				<td><img src="images/f_norm.gif" border="0" alt="û�лظ�������"> 
				������</td>
				<td><img src="images/f_locked.gif" border="0"> 
				��������</td>
				<td>
				<img src="images/f_hot.gif" border="0" alt="�ظ����ﵽ <%=SiteSettings("PopularPostThresholdPosts")%> ���ߵ�����ﵽ <%=SiteSettings("PopularPostThresholdViews")%>"> 
				�������� </td>
				<td><img src="images/my.gif"> �Լ����������</td>
			</tr>
		</table>
<a href="#top">���ض���</a></td>
	</tr>

	<tr>
		<td class="a3"><a name="D5"></a><b>���������̳ʱ��Ϊʲô�������κ�����/���ӣ�</b> <br>
		���û�˷����ӣ���Ȼ�Ϳ�������Ҫô���������������ӹ��������������㲻�뿴�����ӡ����������������ֻ��ʾĳ�������Ժ�����ӣ�����ֻ��ʾ���������ڵ����ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D6"></a><b>Ϊʲô�Ҹշ������ӿ�������</b> <br>
		��̳�������ñ༭����������鹦�ܡ��������һ����Ҫ������̳������ϵͳ�����һ����Ϣ��������������ڵȴ���顣һ�����Աͨ����飬������Ӳſɼ������Ա�����ƶ����޸Ļ�ɾ��������ӣ������������Ϊ������Ӳ��ʺ��������̳����Ļ���<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D7"></a><b>ʲô�ǹ̶����ӣ�</b> <br>
		�����ܳ�������̳�����б���ǰ������ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D8"></a><b>ʲô����������</b> <br>
		���ǲ��ܻظ������ӡ������ˡ�����Ա��������������������һ�����ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D9"></a><b>����ʱ���ɷ�����</b> <br>
		���ԣ����������򣬻ظ����򣬵�������������������С�Ĭ���ǰ���ʱ���������µ����������档<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D10"></a><b>��̳��RSSͼ����ʲô������</b> <br>
		���RSSͼ��������������̳��RSS���ӡ�RSS��Ҫʹ��ר�ŵ��Ķ������Ķ���<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="D11"></a><b>����ô��������̳��</b> <br>
		������¼��ǰ���ʹ�����̳��ȴ����һ��δ֪�Ĵ��󣬿���������ԭ��һ��ԭ���������̳Ҫ��Ҫע����ܵ�¼����һ��ԭ����������̳�Ѿ��ر��ˡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">���� </td>
	</tr>
	<tr>
		<td class="a3"><a name="E1"></a><b>������HTML������</b> <br>
		���ԣ����ǲ���ֱ�����뵽�༭���ڡ� �������IE�����Ĭ�ϱ༭����һ���߼�HTML�༭���������Զ�����HTML���롣�������������������������ʹ��һ����ͨ��HTML�༭���� 
<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E2"></a><b>ʲô��YBB���룿</b> <br>
		�������ڰ����ָ���ܶ໨�����﷨��<a href="?menu=ybb">�������ɲ鿴YBB������ϸʹ�÷���</a>��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>
	<tr>
		<td class="a3"><a name="E3"></a><b>�������Դ�������</b> <br>
		���ԣ�����Ҫ����Ա���Ŵ˹��ܡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E4"></a><b>ʲô�Ǳ���ͼ�ꣿ</b> <br>
		�����ڷ���ʱ���Բ����һЩСͼ�꣬���磺Ц���������ȡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E5"></a><b>��ô��������</b> <br>
		����̳�ϵġ����������⡱ͼƬ�����һ��������ҳ�棨�������Ѿ���¼�������Ҫ�����ȵ�¼����<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E6"></a><b>��ô�ظ�һ�����ӣ�</b> <br>
		�㡰�ظ�����ť�����á���ť�Ϳ��Իظ����ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E7"></a><b>��ô���޸��ҵ����ӣ�</b> <br>
		������༭�����ͼ��ť����Ա༭������ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E8"></a><b>��ôɾ���ҵ����ӣ�</b> <br>
		���㷢��������Ա߿���һ��ɾ��ͼ��ť������㷢��������Ѿ�����һ�������ظ����㽫����ɾ��������ӡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E9"></a><b>Ϊʲô�ҵ���������Щ�ʱ��滻���ˡ�***����</b> <br>
		����ԱҲ���Ѿ�Ϊ���������˴�������������������������ʱ��ĳЩ����Ϊ�����ð���Ĵ��ｫ����ĸ***�滻��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E10"></a><b>����ô���ҵ������м�ǩ����</b> <br>
		���ڿ�����壭�������޸�������ǩ����<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="E11"></a><b>��ô��ͷ����ӵ������У�</b> <br>
		���ڿ�����壭�������޸�������ͷ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3">
		
		
<a name="E12"></a><b>��������
</b>
<table border="0" cellpadding="3" cellspacing="1">
	<tr>
		<td>1������������<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralAddThread")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralAddThread")%></font><br>
	</tr>
		<td>2���ظ�����<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralAddPost")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralAddPost")%></font></tr>
		<td>3����Ϊ������<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralAddValuedPost")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralAddValuedPost")%></font></tr>
		<td>4��ɾ��������<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralDeleteThread")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralDeleteThread")%></font></tr>
		<td>5��ɾ���ظ���<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralDeletePost")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralDeletePost")%></font></tr>
		<td>6��ȡ��������<br>
		����&nbsp; ����ֵ��<font color="red"><%=SiteSettings("IntegralDeleteValuedPost")%></font><br>
		����&nbsp; ���ֵ��<font color="red"><%=SiteSettings("IntegralDeleteValuedPost")%></font></tr>
	</table>
<a href="#top">���ض���</a></td>
	</tr>
</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">�û�Ȩ�� </td>
	</tr>
	<tr>
		<td class="a3"><a name="F1"></a><b>ʲô�ǹ���Ա��</b> <br>
		����Աӵ����̳����߼�Ȩ�ޡ�Ĭ������£�����Ա��������̳��ִ���κβ���������Ȩ�ޣ����磬�༭���ӡ���׼�û��������°��ȵȡ�<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="F2"></a><b>ʲô�ǰ����� </b><br>
		����ӵ����̳�ĵڶ���Ȩ�ޡ������ܹ��������İ����ִ���κβ������������׼���ӡ�ת�����ӡ�ɾ�����ӡ��༭���ӡ�������й���ĳ�������⣬��õĽ���취����ͬ������ϵ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	</table>
<p></p>
<table cellspacing="1" cellpadding="3" width="100%" class=a2>
	<tr class=a1>
		<td height="25">���� BBSXP�� </td>
	</tr>
	<tr>
		<td class="a3"><a name="G1"></a><b>ʲô�� BBSXP��</b> <br>
		BBSXP ��һ��ʹ��ASP���Կ�������̳ϵͳ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G2"></a><b>˭��ʹ�� BBSXP��</b> <br>
		�ܶ๫����˽�˵���֯��ʹ����������̳ϵͳ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G3"></a><b>˭������ BBSXP��</b> <br>
		BBSXP ����YUZI�����Ŷӿ���������Բ��� <a target="_blank" href="http://www.yuzi.net">http://www.yuzi.net</a> �õ��������Ѷ��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>

	<tr>
		<td class="a3"><a name="G4"></a><b>���Ļ�ȡBBSXP�Ŀ����� </b><br>
		���� <a target="_blank" href="http://www.bbsxp.com">http://www.bbsxp.com</a> �������� BBSXP �����°汾��<br><a href="#top">���ض���</a><br>
		 </td>
	</tr>
</table>
<%end if%>

<%htmlend%>