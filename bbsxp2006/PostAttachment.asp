<!-- #include file="conn.asp" -->
<%
SiteSettings=Conn.Execute("[BBSXP_SiteSettings]")
set Rs=server.CreateObject("adodb.recordset")
sql="SELECT top 1 * FROM [BBSXP_PostAttachments] where id="&int(Request("AttachmentID"))
Rs.Open sql,conn,1,3
if Rs.Eof then CloseDatabase
response.contenttype=Rs("ContentType")
response.addheader "content-disposition","attachment;filename="&Rs("FileName")&""

if sitesettings("WatermarkOption")="Persits.Jpeg" and Rs("ContentType")="image/pjpeg" then
Set Jpeg = Server.CreateObject("Persits.Jpeg")
Jpeg.OpenBinary rs("Content").Value

if sitesettings("WatermarkType")=0 then
'Jpeg.Canvas.Font.Color = &H000000	'颜色
'Jpeg.Canvas.Font.Family = "黑体"  	'字体
'Jpeg.Canvas.Font.size = "15"		'大小
Jpeg.Canvas.Print 10, 10, ""&SiteSettings("WatermarkText")&""
elseif sitesettings("WatermarkType")=1 then
Set Jpeg2 = Server.CreateObject("Persits.Jpeg")
Jpeg2.Open Server.MapPath(sitesettings("WatermarkImage"))
select case sitesettings("WatermarkPosition")
case "0"
ImageWidth=10
ImageHeight=10
case "1"
ImageWidth=10
ImageHeight=Jpeg.OriginalHeight-Jpeg2.OriginalHeight-10
case "2"
ImageWidth=(Jpeg.OriginalWidth-Jpeg2.OriginalWidth)/2
ImageHeight=(Jpeg.OriginalHeight-Jpeg2.OriginalHeight)/2
case "3"
ImageWidth=Jpeg.OriginalWidth-Jpeg2.OriginalWidth-10
ImageHeight=10
case "4"
ImageWidth=Jpeg.OriginalWidth-Jpeg2.OriginalWidth-10
ImageHeight=Jpeg.OriginalHeight-Jpeg2.OriginalHeight-10
end select
Jpeg.Canvas.DrawImage ImageWidth, ImageHeight, Jpeg2, 0.5, &HFFFFFF	'0.5透明度, 透明颜色FFFFFF
Set Jpeg2 = Nothing
end if
Jpeg.SendBinary
Set Jpeg = Nothing

else
Response.BinaryWrite Rs("Content")
end if

Rs("TotalDownloads")=Rs("TotalDownloads")+1
Rs.update
Rs.Close
CloseDatabase
%>