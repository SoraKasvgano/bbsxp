<!-- #include file="setup.asp" -->
<%
if Http_Referer="" then Http_Referer="Default.asp"
if Request_Method <> "POST" then error("<li>�ύ��ʽ����</li><li>������ʹ�õ���"&Request_Method&"�ύ��ʽ��</li>")

select case Request("menu")
	case "Invisible"
		ResponseCookies "Invisible","1","999"
	case "Online"
		ResponseCookies "Invisible","0","999"
end select
response.redirect Http_Referer
%>