<%
on error resume next
if SiteSettings("SelectMailMode")="JMail.Message" then
Set JMail=Server.CreateObject("JMail.Message")
JMail.Charset="gb2312"
JMail.AddRecipient Mailaddress
JMail.Subject = MailTopic
JMail.Body = body
JMail.From = SiteSettings("SmtpServerMail")
JMail.MailServerUserName = SiteSettings("SmtpServerUserName")
JMail.MailServerPassword = SiteSettings("SmtpServerPassword")
JMail.Send SiteSettings("SmtpServer")
Set JMail=nothing

elseif SiteSettings("SelectMailMode")="CDONTS.NewMail" then
Set MailObject = Server.CreateObject("CDONTS.NewMail")
MailObject.Send SiteSettings("SmtpServerMail"),Mailaddress,MailTopic,body
Set MailObject=nothing
end if

If Err Then Response.Write ""&Mailaddress&"邮件发送失败！错误原因：" & Err.Description & "<br>"
On Error GoTo 0
%>