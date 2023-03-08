<!--#include File="Conn.asp"-->
<%
AttachmentID=RequestInt("AttachmentID")
Sql="select top 1 * from ["&TablePrefix&"PostAttachments] where UpFileID="&AttachmentID&""
Rs.Open Sql,conn,1,1
if not Rs.Eof then
	response.contenttype=Rs("ContentType")
	response.addheader "content-disposition","attachment;filename="&Rs("FileName")&""
	
	if IsObjInstalled("Persits.Jpeg") and SiteConfig("WatermarkOption")="Persits.Jpeg" and Rs("ContentType")="image/pjpeg" then
		Set Jpeg = Server.CreateObject("Persits.Jpeg")
		Jpeg.OpenBinary Rs("FileData").Value
		JpegPersits
		Jpeg.SendBinary
		Set Jpeg = Nothing
	else
		Response.BinaryWrite Rs("FileData")
	end if
end if
Rs.Close
%>